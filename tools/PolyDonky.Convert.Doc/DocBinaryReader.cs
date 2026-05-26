using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using OpenMcdf;
using PolyDonky.Core;

namespace PolyDonky.Convert.Doc;

/// <summary>
/// Word 97-2003 binary (.doc, OLE2 Compound File) → IWPF 변환기 (Phase 1a/1b — 텍스트·단락·기본 서식).
///
/// 참고 공식 공개 명세:
///   - [MS-DOC]      Word (.doc) Binary File Format — FIB·CLX·piece table·PAPX/CHPX·SEPX 등
///   - [MS-CFB]      Compound File Binary File Format — OLE2 컨테이너 (OpenMcdf 가 처리)
///   - [MS-OLEPS]    Object Linking and Embedding (OLE) Property Set Format — SummaryInformation
///                   / DocumentSummaryInformation 의 PID·Variant 인코딩
///   - [MS-OFFCRYPTO] Office Document Cryptography Structure — fEncrypted 비트 / EncryptionInfo
///                   stream / RC4·AES 키 유도 (Phase 1a 는 인식 후 거부만)
///
/// 본 단계 범위:
///   - FIB(File Information Block) 파싱 + 암호화/obfuscation 사전 거부 (MS-OFFCRYPTO)
///   - CLX / Piece Table 따라 WordDocument stream 에서 텍스트 추출 (MS-DOC §2.3, §2.8)
///   - 단락 마커 0x0D 기준 단락 분리; 표 셀 마커 0x07 / 줄바꿈 0x0B / 페이지 break 0x0C /
///     필드 제어 0x13..0x15 등은 공백·줄바꿈으로 정규화하거나 폐기
///   - PAPX 운영 → 단락 정렬 (sprmPJc80 0x2461)                                    (Phase 1b)
///   - CHPX 운영 → 굵게(0x0835)·이탤릭(0x0836)·취소선(0x0837)·밑줄(0x2A3E)·
///     글자 크기(sprmCHps 0x4A43) 로 Run 분할                                       (Phase 1b)
///   - SummaryInformation PropertySet (MS-OLEPS §2.18) 에서 Title/Subject/Author/Keywords/
///     LastSavedBy/Created/Modified/AppName 추출 (VT_LPSTR/VT_LPWSTR/VT_FILETIME)
///   - 비-OLE2 입력·손상 CFB 컨테이너는 한국어 진단으로 명확히 거부
///
/// 의도적으로 다음 작업의 범위 밖 (다음 PR 들에서 다룸):
///   - 폰트 패밀리·색상·하이라이트 (CHPX 추가 sprm)                                  (Phase 1c)
///   - 표 / 이미지 / 필드 / 헤더·푸터 / 섹션                                          (Phase 2)
///   - MS-OFFCRYPTO 의 실제 해독 (RC4/AES 키 유도·복호화)
///
/// 자체 구현인 이유는 CLAUDE.md: HWP 코덱과 마찬가지로 비-OSS 상용 라이브러리(Aspose 등) 의존을 피한다.
/// </summary>
public class DocBinaryReader
{
    public PolyDonkyument Read(Stream input)
    {
        // OpenMcdf 의 RootStorage 는 파일 경로 또는 Seekable Stream 을 받는데, 안전성을 위해
        // 임시 파일에 복사 후 OpenRead 로 다룬다 (HwpReader 와 같은 패턴).
        var tmpPath = Path.GetTempFileName();
        try
        {
            using (var fs = File.Create(tmpPath))
                input.CopyTo(fs);

            RootStorage root;
            try
            {
                // [MS-CFB] 컨테이너 열기. OpenMcdf 가 sector/FAT 무결성을 검사하므로 비-OLE2 / 손상 파일은 여기서 거부됨.
                root = RootStorage.OpenRead(tmpPath);
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException(
                    "이 파일은 Word 97-2003 (.doc) 형식이 아니거나 OLE2 컨테이너가 손상되었습니다. " +
                    "RTF 는 .rtf 로, DOCX 는 .docx 확장자로 시도해 보세요.",
                    ex);
            }
            using var _root = root;  // dispose 보장

            // WordDocument stream 전체를 메모리에 적재 — Word97-2003 문서는 보통 수 MB 이하.
            byte[] wd = ReadAll(root, "WordDocument")
                ?? throw new InvalidOperationException("이 파일은 Word 97-2003 (.doc) 형식이 아닙니다 — WordDocument stream 이 없습니다.");

            var fib = ParseFib(wd);

            // [MS-OFFCRYPTO] 암호화·obfuscation 사전 거부.
            // fEncrypted 비트가 켜져 있으면 본문 byte 가 RC4/AES 로 변조되어 piece table 도 의미 없음.
            // 일부 변종은 EncryptionInfo / EncryptedSummary stream 으로도 신호.
            if (fib.Encrypted || HasEncryptionStream(root))
            {
                throw new InvalidOperationException(
                    "암호화된 Word 문서는 현재 지원하지 않습니다 (MS-OFFCRYPTO). " +
                    "Word 에서 암호를 풀어 다시 저장한 뒤 시도해 주세요.");
            }

            byte[] table = ReadAll(root, fib.TableStreamName)
                ?? throw new InvalidOperationException($"{fib.TableStreamName} stream 이 없습니다.");

            var (text, fcs) = ExtractTextWithFcs(wd, table, fib);
            // Phase 1b — PAPX/CHPX 바인 테이블을 한 번 로드해서 단락·문자 서식 조회에 재사용.
            var fmt = FormatStyles.Build(wd, table, fib);
            var doc = BuildDocument(text, fcs, fmt);

            // 메타데이터 (best-effort)
            var summary = ReadAll(root, "SummaryInformation");
            if (summary is { Length: > 0 })
                ApplySummaryInformation(summary, doc);

            return doc;
        }
        finally
        {
            try { File.Delete(tmpPath); } catch { }
        }
    }

    // ─────────────────────────────── FIB ────────────────────────────────────────

    private sealed record Fib(
        string TableStreamName,
        uint   FcMin,
        uint   CcpText,
        uint   FcClx,
        uint   LcbClx,
        ushort NFib,
        bool   Encrypted,
        bool   Obfuscated,
        // Phase 1b — PAPX / CHPX bin tables (FIB §2.5.5 FibRgFcLcb97)
        uint   FcPlcfBteChpx,
        uint   LcbPlcfBteChpx,
        uint   FcPlcfBtePapx,
        uint   LcbPlcfBtePapx);

    // FIB (File Information Block) — WordDocument stream 의 첫 부분. 크기는 nFib 에 따라 다르지만
    // 우리가 필요한 모든 필드는 첫 0x200 byte 안에 있다.
    private static Fib ParseFib(byte[] wd)
    {
        if (wd.Length < 0x200)
            throw new InvalidOperationException("WordDocument stream 이 너무 짧음 — FIB 헤더 부족.");

        ushort magic = BitConverter.ToUInt16(wd, 0x0000);
        if (magic != 0xA5EC)
            throw new InvalidOperationException($"Word 97-2003 시그니처 불일치 (0x{magic:X4}). 다른 형식(RTF/DOCX/HWP)이 .doc 로 잘못 명명되었을 수 있습니다.");

        ushort nFib  = BitConverter.ToUInt16(wd, 0x0002);
        ushort flags = BitConverter.ToUInt16(wd, 0x000A);
        // [MS-DOC] §2.5.1 FibBase flag bits @ 0x000A:
        //   bit 8  (0x0100) = fEncrypted        — RC4/AES 암호화
        //   bit 9  (0x0200) = fWhichTblStm      — 1 → "1Table", 0 → "0Table"
        //   bit 15 (0x8000) = fObfuscated       — XOR obfuscation (legacy)
        bool   encrypted   = (flags & 0x0100) != 0;
        bool   obfuscated  = (flags & 0x8000) != 0;
        string tableName   = (flags & 0x0200) != 0 ? "1Table" : "0Table";

        // fcMin (보통 0) — 본문 텍스트 시작 file character position. Complex 문서에선 piece table 가 우선.
        uint fcMin   = BitConverter.ToUInt32(wd, 0x0018);
        // ccpText — main text 길이 (CP 단위)
        uint ccpText = BitConverter.ToUInt32(wd, 0x004C);

        // fcClx / lcbClx — Complex File Information (piece table 포함)
        uint fcClx   = BitConverter.ToUInt32(wd, 0x01A2);
        uint lcbClx  = BitConverter.ToUInt32(wd, 0x01A6);

        // PAPX / CHPX bin tables — Phase 1b
        // [MS-DOC] §2.5.5 FibRgFcLcb97:
        //   fcPlcfBteChpx @ 0x00FA, lcbPlcfBteChpx @ 0x00FE
        //   fcPlcfBtePapx @ 0x0102, lcbPlcfBtePapx @ 0x0106
        uint fcPlcfBteChpx  = BitConverter.ToUInt32(wd, 0x00FA);
        uint lcbPlcfBteChpx = BitConverter.ToUInt32(wd, 0x00FE);
        uint fcPlcfBtePapx  = BitConverter.ToUInt32(wd, 0x0102);
        uint lcbPlcfBtePapx = BitConverter.ToUInt32(wd, 0x0106);

        return new Fib(tableName, fcMin, ccpText, fcClx, lcbClx, nFib, encrypted, obfuscated,
                       fcPlcfBteChpx, lcbPlcfBteChpx, fcPlcfBtePapx, lcbPlcfBtePapx);
    }

    // [MS-OFFCRYPTO] EncryptionInfo / EncryptedSummary stream 존재 검사 — fEncrypted 비트가
    // 없더라도 일부 변종이 stream 만으로 암호화를 표시할 수 있어 보조 신호로 본다.
    private static bool HasEncryptionStream(RootStorage root)
    {
        foreach (var entry in root.EnumerateEntries())
        {
            if (entry.Name.Equals("EncryptionInfo",   StringComparison.OrdinalIgnoreCase) ||
                entry.Name.Equals("EncryptedSummary", StringComparison.OrdinalIgnoreCase))
                return true;
        }
        return false;
    }

    // ─────────────────────────────── 텍스트 추출 ────────────────────────────────

    // 본문 main text 영역(CP 0..ccpText)에 대해 piece table 따라가며 raw 텍스트를 구성.
    // Phase 1b 부터 각 char 에 대응하는 file character position(FC byte offset) 도 함께 반환 —
    // PAPX/CHPX 바인 테이블이 FC 기준이므로 단락/문자 서식 조회에 필수.
    private static (string Text, int[] Fcs) ExtractTextWithFcs(byte[] wd, byte[] table, Fib fib)
    {
        if (fib.LcbClx == 0)
        {
            // 비-복합 문서 — 매우 드물지만 fall-back. fcMin..fcMin+ccpText 를 1252 로 본다.
            var s = DecodeAnsi(wd, (int)fib.FcMin, (int)fib.CcpText);
            var fc = new int[s.Length];
            for (int i = 0; i < s.Length; i++) fc[i] = (int)fib.FcMin + i;
            return (s, fc);
        }

        var pcds = ParsePieceTable(table, (int)fib.FcClx, (int)fib.LcbClx, out int[] cps);

        var sb  = new StringBuilder((int)fib.CcpText);
        var fcs = new List<int>((int)fib.CcpText);
        for (int i = 0; i < pcds.Count; i++)
        {
            int cpStart = cps[i];
            int cpEnd   = cps[i + 1];
            // 본문 main text 범위(0..ccpText) 와 교집합만 채택 — 그 너머는 헤더/푸터/각주/주석 등.
            if (cpStart >= fib.CcpText) break;
            int effEnd = Math.Min(cpEnd, (int)fib.CcpText);
            int len    = effEnd - cpStart;
            if (len <= 0) continue;

            uint fcRaw = pcds[i].Fc;
            bool compressed = (fcRaw & 0x40000000u) != 0;
            int  fc   = (int)(fcRaw & 0x3FFFFFFFu);
            if (compressed)
            {
                // ANSI (CP1252) 1-byte chars, fc /= 2.
                fc /= 2;
                if (fc < 0 || fc + len > wd.Length) continue;
                var piece = DecodeAnsi(wd, fc, len);
                for (int j = 0; j < piece.Length; j++)
                {
                    sb.Append(piece[j]);
                    fcs.Add(fc + j);
                }
            }
            else
            {
                // UTF-16LE 2-byte chars.
                int byteLen = len * 2;
                if (fc < 0 || fc + byteLen > wd.Length) continue;
                var piece = Encoding.Unicode.GetString(wd, fc, byteLen);
                for (int j = 0; j < piece.Length; j++)
                {
                    sb.Append(piece[j]);
                    fcs.Add(fc + j * 2);
                }
            }
        }

        return (sb.ToString(), fcs.ToArray());
    }

    private sealed record Pcd(uint Fc);

    // CLX 영역에서 PCDT(0x02) 를 찾아 piece table 을 파싱.
    private static List<Pcd> ParsePieceTable(byte[] table, int fcClx, int lcbClx, out int[] cps)
    {
        if (fcClx < 0 || lcbClx < 5 || fcClx + lcbClx > table.Length)
            throw new InvalidOperationException("CLX 영역이 Table stream 범위를 벗어남.");

        int pos = fcClx;
        int end = fcClx + lcbClx;
        while (pos < end)
        {
            byte clxt = table[pos];
            if (clxt == 0x01)
            {
                // PRC — skip (2-byte length + content)
                if (pos + 3 > end) break;
                int cbGrpprl = BitConverter.ToUInt16(table, pos + 1);
                pos += 3 + cbGrpprl;
                continue;
            }
            if (clxt == 0x02)
            {
                if (pos + 5 > end) break;
                uint lcb = BitConverter.ToUInt32(table, pos + 1);
                int plcStart = pos + 5;
                if (plcStart + lcb > end || lcb < 4) break;

                // PlcPcd: aCP[n+1] + aPcd[n] 으로 구성, lcb = 4(n+1) + 8n = 12n + 4
                int n = (int)((lcb - 4) / 12);
                if (n < 0) break;

                cps = new int[n + 1];
                for (int i = 0; i <= n; i++)
                    cps[i] = BitConverter.ToInt32(table, plcStart + i * 4);

                int pcdBase = plcStart + (n + 1) * 4;
                var list = new List<Pcd>(n);
                for (int i = 0; i < n; i++)
                {
                    // PCD: 2 byte flags, 4 byte fc, 2 byte prm. fc 는 LE u32.
                    uint fc = BitConverter.ToUInt32(table, pcdBase + i * 8 + 2);
                    list.Add(new Pcd(fc));
                }
                return list;
            }
            // 알 수 없는 clxt — 안전하게 중단
            break;
        }
        throw new InvalidOperationException("Piece Table (PCDT 0x02) 를 CLX 영역에서 찾을 수 없음.");
    }

    private static string DecodeAnsi(byte[] data, int offset, int length)
    {
        if (offset < 0 || length <= 0 || offset + length > data.Length) return string.Empty;
        // CP1252 는 .NET 10 CodePagesEncodingProvider 가 BCL 에 포함되어 별도 등록 불필요.
        try
        {
            return Encoding.GetEncoding(1252).GetString(data, offset, length);
        }
        catch
        {
            return Encoding.Latin1.GetString(data, offset, length);
        }
    }

    // ─────────────────────────────── 단락 분리 ──────────────────────────────────

    // Word 의 내부 텍스트 제어 문자 처리:
    //   0x0D '\r'  → 단락 끝
    //   0x0B       → soft line break (단락 내 줄바꿈) → '\n' 으로 보존
    //   0x07       → 표 셀 끝 / 표 행 끝 → 일단 탭으로 평탄화 (표 처리는 Phase 2)
    //   0x0C       → 페이지 break → 단락 분리로 처리
    //   0x13/0x14/0x15 → 필드 시작/구분/끝 → Phase 2 까지는 제거
    //   0x02       → 각주/주석 참조 → 제거
    //   0x05       → 주석 참조 → 제거
    //   0x08       → drawing anchor → 제거
    //   0x01       → picture anchor → 제거 (Phase 2 에서 이미지 처리)
    //   기타 0x00..0x1F 중 \t(0x09) 외 → 제거
    // Phase 1b — 단락 경계(\r 또는 \f) 도달 시 그 FC 로 PAPX 를 조회해 단락 정렬 적용.
    //          단락 내 각 char 의 FC 로 CHPX 를 조회해 같은 RunStyle 끼리 묶어 Run 분할.
    private static PolyDonkyument BuildDocument(string raw, int[] fcs, FormatStyles fmt)
    {
        var doc     = new PolyDonkyument();
        var section = new Section();
        doc.Sections.Add(section);

        var paraChars = new List<char>();
        var paraFcs   = new List<int>();
        int lastFc    = 0;

        for (int i = 0; i < raw.Length; i++)
        {
            char c  = raw[i];
            int  fc = i < fcs.Length ? fcs[i] : lastFc;
            lastFc = fc;

            switch (c)
            {
                case '\r':
                case '\f':  // page break — 단락 분리로 처리
                    FlushParagraph(section, paraChars, paraFcs, fc, fmt);
                    break;
                case '\v':  // soft line break
                    paraChars.Add('\n'); paraFcs.Add(fc);
                    break;
                case '':  // cell mark
                    paraChars.Add('\t'); paraFcs.Add(fc);
                    break;
                case '': // field begin
                case '': // field separator
                case '': // field end
                case '': // footnote ref
                case '': // comment ref
                case '': // drawing
                case '': // picture
                    // 무시 (Phase 2 까지)
                    break;
                case '\t':
                case '\n':
                    paraChars.Add(c); paraFcs.Add(fc);
                    break;
                default:
                    if (c < 0x20) break;  // 그 외 제어 문자 폐기
                    paraChars.Add(c); paraFcs.Add(fc);
                    break;
            }
        }
        FlushParagraph(section, paraChars, paraFcs, lastFc, fmt);

        if (section.Blocks.Count == 0) section.Blocks.Add(new Paragraph());
        return doc;
    }

    // Phase 1b — 누적된 (char, fc) 쌍을 한 단락으로 묶어 Section.Blocks 에 추가.
    // 단락 정렬은 PAPX (paragraph end FC) 에서, run 분할은 CHPX (char FC) 에서 best-effort 로 적용.
    private static void FlushParagraph(
        Section section, List<char> paraChars, List<int> paraFcs, int paraEndFc, FormatStyles fmt)
    {
        var para = new Paragraph();
        var ps = fmt.GetParagraphStyle(paraEndFc);
        if (ps is not null) para.Style = ps;

        if (paraChars.Count > 0)
        {
            RunStyle? curStyle = null;
            var curText = new StringBuilder();
            for (int i = 0; i < paraChars.Count; i++)
            {
                var rs = fmt.GetRunStyle(paraFcs[i]) ?? new RunStyle();
                if (curStyle is null || !RunStyleEquals(curStyle, rs))
                {
                    if (curText.Length > 0 && curStyle is not null)
                        para.AddText(curText.ToString(), curStyle);
                    curStyle = rs;
                    curText.Clear();
                }
                curText.Append(paraChars[i]);
            }
            if (curText.Length > 0 && curStyle is not null)
                para.AddText(curText.ToString(), curStyle);
        }

        section.Blocks.Add(para);
        paraChars.Clear();
        paraFcs.Clear();
    }

    private static bool RunStyleEquals(RunStyle a, RunStyle b)
        => a.Bold == b.Bold
        && a.Italic == b.Italic
        && a.Underline == b.Underline
        && a.Strikethrough == b.Strikethrough
        && a.FontSizePt == b.FontSizePt;

    // ─────────────────────────────── 메타데이터 ─────────────────────────────────

    // [MS-OLEPS] §2.18 SummaryInformation PropertySet → DocumentMetadata 매핑.
    // PIDSI 코드: 0x02 Title · 0x03 Subject · 0x04 Author · 0x05 Keywords · 0x06 Comments ·
    //             0x08 LastSavedBy · 0x09 RevisionNumber · 0x0C CreateTime · 0x0D LastSavedTime ·
    //             0x12 AppName.
    // Variant 타입: VT_LPSTR (0x001E, CP1252) · VT_LPWSTR (0x001F, UTF-16LE) · VT_FILETIME (0x0040)
    // 만 지원 — 그 외 PID/Variant 는 무시 (메타데이터는 본문 무효화 사유가 되지 않음).
    private static void ApplySummaryInformation(byte[] data, PolyDonkyument doc)
    {
        try
        {
            // PropertySetStream header: byteOrder(2)·version(2)·sysID(4)·CLSID(16)·numPropertySets(4)
            //                          + FormatID(16) + sectionOffset(4) = 48 byte until first section.
            if (data.Length < 48) return;
            int sectionOffset = BitConverter.ToInt32(data, 44);
            if (sectionOffset < 0 || sectionOffset + 8 > data.Length) return;

            // PropertySection header: cb(4), cProperties(4)
            int cProps = BitConverter.ToInt32(data, sectionOffset + 4);
            if (cProps <= 0 || cProps > 1024) return;

            int idxBase = sectionOffset + 8;
            for (int i = 0; i < cProps; i++)
            {
                int entryOff = idxBase + i * 8;
                if (entryOff + 8 > data.Length) break;
                int pid    = BitConverter.ToInt32(data, entryOff);
                int offset = BitConverter.ToInt32(data, entryOff + 4);
                int valOff = sectionOffset + offset;
                if (valOff < 0 || valOff + 4 > data.Length) continue;

                int type = BitConverter.ToInt32(data, valOff);
                switch (type)
                {
                    case 0x001E:  // VT_LPSTR (CP1252)
                    case 0x001F:  // VT_LPWSTR (UTF-16LE)
                    {
                        var str = ReadLpStr(data, valOff + 4, ansi: type == 0x001E);
                        if (string.IsNullOrEmpty(str)) break;
                        ApplyStringPid(pid, str, doc);
                        break;
                    }
                    case 0x0040:  // VT_FILETIME (Win32 FILETIME, 8 bytes LE)
                    {
                        if (valOff + 12 > data.Length) break;
                        long ft = BitConverter.ToInt64(data, valOff + 4);
                        if (ft <= 0) break;
                        DateTimeOffset when;
                        try { when = new DateTimeOffset(DateTime.FromFileTimeUtc(ft), TimeSpan.Zero); }
                        catch { break; }
                        switch (pid)
                        {
                            case 0x0C: doc.Metadata.Created  = when; break;
                            case 0x0D: doc.Metadata.Modified = when; break;
                        }
                        break;
                    }
                }
            }
        }
        catch
        {
            // 메타데이터는 best-effort. 어떤 예외도 본문 텍스트 추출을 무효화하지 않는다.
        }
    }

    private static void ApplyStringPid(int pid, string value, PolyDonkyument doc)
    {
        switch (pid)
        {
            case 0x02: doc.Metadata.Title       = value; break;
            case 0x04: doc.Metadata.Author      = value; break;
            case 0x08: doc.Metadata.Editor      = value; break;
            case 0x12: doc.Metadata.Application = value; break;
            case 0x03: doc.Metadata.Custom["Subject"]        = value; break;
            case 0x05: doc.Metadata.Custom["Keywords"]       = value; break;
            case 0x06: doc.Metadata.Custom["Comments"]       = value; break;
            case 0x09: doc.Metadata.Custom["RevisionNumber"] = value; break;
        }
    }

    private static string? ReadLpStr(byte[] data, int offset, bool ansi)
    {
        if (offset + 4 > data.Length) return null;
        int len = BitConverter.ToInt32(data, offset);
        if (len <= 0) return null;
        int byteLen = ansi ? len : len * 2;
        if (offset + 4 + byteLen > data.Length) return null;
        string s;
        if (ansi)
        {
            try { s = Encoding.GetEncoding(1252).GetString(data, offset + 4, byteLen); }
            catch { s = Encoding.Latin1.GetString(data, offset + 4, byteLen); }
        }
        else
        {
            s = Encoding.Unicode.GetString(data, offset + 4, byteLen);
        }
        return s.TrimEnd('\0');
    }

    // ─────────────────────────────── PAPX / CHPX 운영 (Phase 1b) ────────────────
    //
    // [MS-DOC] §2.4.6 BTE (Bin Table Entry) plex: aFC[n+1] + aPnFkp[n], 각 PnFkp 는 4-byte 페이지 번호.
    // [MS-DOC] §2.8.1 PapxFkp (512 byte): rgfc[cpara+1] + rgbx[cpara] (각 13 byte BXPap) + cpara(1 byte at off 511).
    //                                     BXPap = bOffset(1) + PHE(12); bOffset 단위는 2-byte.
    //                                     bOffset 위치의 PapxInFkp = cb(1) + grpprlInPapx(...).
    //                                       cb!=0 → 크기 = cb*2 - 1, grpprl 은 +1 에서 시작
    //                                       cb==0 → 다음 1 byte cb'; 크기 = cb'*2, grpprl 은 +2 에서 시작
    //                                     grpprlInPapx 의 첫 2 byte 는 istd; 그 뒤가 sprm 배열.
    // [MS-DOC] §2.8.2 ChpxFkp (512 byte): rgfc[crun+1] + rgb[crun] (각 1 byte bOffset) + crun(1 byte at off 511).
    //                                     ChpxInFkp = cb(1) + grpprlInChpx(cb byte). istd 없음.
    //
    // 본 단계에서 인식하는 sprm (Word 97+):
    //   PAPX:  sprmPJc80 (0x2461)  — 단락 정렬 1 byte (0=L, 1=C, 2=R, 3=J)
    //   CHPX:  sprmCFBold (0x0835), sprmCFItalic (0x0836)        — 1 byte toggle/on/off
    //          sprmCFUnderline (0x2A3E)                          — 1 byte (0=none, 그 외=ON)
    //          sprmCHps (0x4A43)                                 — 2 byte 글자크기 (half-points)
    // 그 외 sprm 은 본 단계에서 무시 — Phase 2 에서 점진적으로 추가한다.

    private sealed class FormatStyles
    {
        private readonly byte[] _wd;
        private readonly List<BteEntry> _papxBte;
        private readonly List<BteEntry> _chpxBte;

        // FKP 페이지(512 byte) 파싱은 매번 동일 데이터를 다시 만지지 않도록 page → grpprl 캐시.
        private readonly Dictionary<(int Pn, int RgIdx), byte[]?> _papxGrpprlCache = new();
        private readonly Dictionary<(int Pn, int RgIdx), byte[]?> _chpxGrpprlCache = new();

        private FormatStyles(byte[] wd, List<BteEntry> papx, List<BteEntry> chpx)
        {
            _wd = wd; _papxBte = papx; _chpxBte = chpx;
        }

        public static FormatStyles Build(byte[] wd, byte[] table, Fib fib)
        {
            var papx = ReadBte(table, (int)fib.FcPlcfBtePapx, (int)fib.LcbPlcfBtePapx);
            var chpx = ReadBte(table, (int)fib.FcPlcfBteChpx, (int)fib.LcbPlcfBteChpx);
            return new FormatStyles(wd, papx, chpx);
        }

        public ParagraphStyle? GetParagraphStyle(int paraEndFc)
        {
            var grpprl = LoadGrpprl(_papxBte, _papxGrpprlCache, paraEndFc, isPapx: true);
            if (grpprl is null || grpprl.Length == 0) return null;

            var style = new ParagraphStyle();
            bool touched = false;
            WalkSprms(grpprl, (sprm, operand) =>
            {
                if (sprm == 0x2461 && operand.Length >= 1)
                {
                    style.Alignment = operand[0] switch
                    {
                        1 => Alignment.Center,
                        2 => Alignment.Right,
                        3 => Alignment.Justify,
                        _ => Alignment.Left,
                    };
                    touched = true;
                }
            });
            return touched ? style : null;
        }

        public RunStyle? GetRunStyle(int charFc)
        {
            var grpprl = LoadGrpprl(_chpxBte, _chpxGrpprlCache, charFc, isPapx: false);
            if (grpprl is null || grpprl.Length == 0) return null;

            var rs = new RunStyle();
            bool touched = false;
            WalkSprms(grpprl, (sprm, operand) =>
            {
                switch (sprm)
                {
                    case 0x0835:  // sprmCFBold
                        if (operand.Length >= 1) { rs.Bold = operand[0] != 0; touched = true; }
                        break;
                    case 0x0836:  // sprmCFItalic
                        if (operand.Length >= 1) { rs.Italic = operand[0] != 0; touched = true; }
                        break;
                    case 0x0837:  // sprmCFStrike
                        if (operand.Length >= 1) { rs.Strikethrough = operand[0] != 0; touched = true; }
                        break;
                    case 0x2A3E:  // sprmCKul (underline kind)
                        if (operand.Length >= 1) { rs.Underline = operand[0] != 0; touched = true; }
                        break;
                    case 0x4A43:  // sprmCHps (font size in half-points)
                        if (operand.Length >= 2)
                        {
                            ushort halfPt = BitConverter.ToUInt16(operand, 0);
                            if (halfPt > 0 && halfPt < 1000) { rs.FontSizePt = halfPt / 2.0; touched = true; }
                        }
                        break;
                }
            });
            return touched ? rs : null;
        }

        private byte[]? LoadGrpprl(
            List<BteEntry> bte,
            Dictionary<(int Pn, int RgIdx), byte[]?> cache,
            int fc, bool isPapx)
        {
            if (bte.Count == 0) return null;
            // BTE 내에서 fc 가 [FcStart, FcEnd) 에 속하는 entry 찾기 — 보통 페이지 수가 적어 선형 충분.
            int pn = -1;
            foreach (var e in bte)
            {
                if (fc >= e.FcStart && fc < e.FcEnd) { pn = e.PnFkp; break; }
            }
            if (pn < 0) return null;

            int fkpOff = pn * 512;
            if (fkpOff < 0 || fkpOff + 512 > _wd.Length) return null;

            int cRun = _wd[fkpOff + 511];
            if (cRun == 0) return null;

            // rgfc[cRun+1] 에서 i 를 찾아 fc ∈ [rgfc[i], rgfc[i+1])
            int rgIdx = -1;
            for (int i = 0; i < cRun; i++)
            {
                int fc0 = BitConverter.ToInt32(_wd, fkpOff + i * 4);
                int fc1 = BitConverter.ToInt32(_wd, fkpOff + (i + 1) * 4);
                if (fc >= fc0 && fc < fc1) { rgIdx = i; break; }
            }
            if (rgIdx < 0) return null;

            if (cache.TryGetValue((pn, rgIdx), out var cached)) return cached;

            int rgbxBase = fkpOff + 4 * (cRun + 1);
            int bOffset;
            if (isPapx)
            {
                // BXPap = 13 byte; first byte = bOffset
                bOffset = _wd[rgbxBase + rgIdx * 13];
            }
            else
            {
                // ChpxFkp rgb = 1 byte each
                bOffset = _wd[rgbxBase + rgIdx];
            }
            byte[]? result;
            if (bOffset == 0)
            {
                result = null;
            }
            else
            {
                int xinFkpOff = fkpOff + bOffset * 2;
                result = isPapx ? ReadPapxInFkp(_wd, xinFkpOff) : ReadChpxInFkp(_wd, xinFkpOff);
            }
            cache[(pn, rgIdx)] = result;
            return result;
        }

        private static byte[]? ReadPapxInFkp(byte[] data, int off)
        {
            if (off < 0 || off >= data.Length) return null;
            byte cb = data[off];
            int grpprlStart, grpprlSize;
            if (cb == 0)
            {
                if (off + 1 >= data.Length) return null;
                byte cb2 = data[off + 1];
                grpprlSize = cb2 * 2;
                grpprlStart = off + 2;
            }
            else
            {
                grpprlSize = cb * 2 - 1;
                grpprlStart = off + 1;
            }
            if (grpprlSize < 2 || grpprlStart + grpprlSize > data.Length) return null;
            // 처음 2 byte 는 istd — sprm 만 잘라낸다.
            int sprmsLen = grpprlSize - 2;
            var sprms = new byte[sprmsLen];
            Buffer.BlockCopy(data, grpprlStart + 2, sprms, 0, sprmsLen);
            return sprms;
        }

        private static byte[]? ReadChpxInFkp(byte[] data, int off)
        {
            if (off < 0 || off >= data.Length) return null;
            byte cb = data[off];
            if (cb == 0) return Array.Empty<byte>();
            if (off + 1 + cb > data.Length) return null;
            var sprms = new byte[cb];
            Buffer.BlockCopy(data, off + 1, sprms, 0, cb);
            return sprms;
        }

        // [MS-DOC] §2.6.2 Sprm: 2-byte sprm 헤더 → ispmd(9) + fSpec(1) + sgc(3) + spra(3).
        // spra 가 operand 크기를 결정: 0/1 → 1 byte, 2/4/5 → 2 byte, 3 → 4 byte, 7 → 3 byte,
        // 6 → variable (다음 1 byte 가 길이).
        private static void WalkSprms(byte[] sprms, Action<ushort, byte[]> onSprm)
        {
            int i = 0;
            while (i + 2 <= sprms.Length)
            {
                ushort sprm = BitConverter.ToUInt16(sprms, i);
                i += 2;
                int spra = (sprm >> 13) & 0x07;
                int operandSize;
                switch (spra)
                {
                    case 0: case 1: operandSize = 1; break;
                    case 2: case 4: case 5: operandSize = 2; break;
                    case 3: operandSize = 4; break;
                    case 7: operandSize = 3; break;
                    case 6:
                        if (i >= sprms.Length) return;
                        operandSize = sprms[i];
                        i += 1;
                        break;
                    default: return;
                }
                if (i + operandSize > sprms.Length) return;
                var operand = new byte[operandSize];
                Buffer.BlockCopy(sprms, i, operand, 0, operandSize);
                i += operandSize;
                onSprm(sprm, operand);
            }
        }

        private readonly record struct BteEntry(int FcStart, int FcEnd, int PnFkp);

        private static List<BteEntry> ReadBte(byte[] table, int fc, int lcb)
        {
            var list = new List<BteEntry>();
            if (lcb <= 4 || fc < 0 || fc + lcb > table.Length) return list;
            // PlcBteFcN: aFC[n+1] (4 byte each) + aPnFkp[n] (4 byte each) → 8n + 4 = lcb
            int n = (lcb - 4) / 8;
            if (n <= 0) return list;
            for (int i = 0; i < n; i++)
            {
                int fcStart = BitConverter.ToInt32(table, fc + i * 4);
                int fcEnd   = BitConverter.ToInt32(table, fc + (i + 1) * 4);
                // PnFkp 의 하위 22 bit 가 페이지 번호 (전체 4 byte 중 bit 0~21).
                int pn      = BitConverter.ToInt32(table, fc + (n + 1) * 4 + i * 4) & 0x003FFFFF;
                list.Add(new BteEntry(fcStart, fcEnd, pn));
            }
            return list;
        }
    }

    // ─────────────────────────────── 유틸 ───────────────────────────────────────

    private static byte[]? ReadAll(RootStorage root, string streamName)
    {
        try
        {
            using var s = root.OpenStream(streamName);
            var ms = new MemoryStream();
            // OpenMcdf v3 stream 은 Seekable. 한 번에 읽는다.
            s.CopyTo(ms);
            return ms.ToArray();
        }
        catch
        {
            return null;
        }
    }
}
