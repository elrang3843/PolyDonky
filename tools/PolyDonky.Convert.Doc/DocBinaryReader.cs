using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using OpenMcdf;
using PolyDonky.Core;

namespace PolyDonky.Convert.Doc;

/// <summary>
/// Word 97-2003 binary (.doc, OLE2 Compound File) → IWPF 변환기 (Phase 1a — 텍스트·단락 추출).
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
///   - SummaryInformation PropertySet (MS-OLEPS §2.18) 에서 Title/Subject/Author/Keywords/
///     LastSavedBy/Created/Modified/AppName 추출 (VT_LPSTR/VT_LPWSTR/VT_FILETIME)
///   - 비-OLE2 입력·손상 CFB 컨테이너는 한국어 진단으로 명확히 거부
///
/// 의도적으로 다음 작업의 범위 밖 (다음 PR 들에서 다룸):
///   - PAPX / CHPX FKP 운영 → 단락 정렬·굵게/이탤릭/밑줄/폰트 크기 (Phase 1b)
///   - 표 / 이미지 / 필드 / 헤더·푸터 / 섹션 (Phase 2)
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

            string text = ExtractText(wd, table, fib);
            var doc = BuildDocument(text);

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
        bool   Obfuscated);

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

        return new Fib(tableName, fcMin, ccpText, fcClx, lcbClx, nFib, encrypted, obfuscated);
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
    private static string ExtractText(byte[] wd, byte[] table, Fib fib)
    {
        if (fib.LcbClx == 0)
            // 비-복합 문서 — 매우 드물지만 fall-back. fcMin..fcMin+ccpText 를 1252 로 본다.
            return DecodeAnsi(wd, (int)fib.FcMin, (int)fib.CcpText);

        var pcds = ParsePieceTable(table, (int)fib.FcClx, (int)fib.LcbClx, out int[] cps);

        var sb = new StringBuilder((int)fib.CcpText);
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
                sb.Append(DecodeAnsi(wd, fc, len));
            }
            else
            {
                // UTF-16LE 2-byte chars.
                int byteLen = len * 2;
                if (fc < 0 || fc + byteLen > wd.Length) continue;
                sb.Append(Encoding.Unicode.GetString(wd, fc, byteLen));
            }
        }

        return sb.ToString();
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
    private static PolyDonkyument BuildDocument(string raw)
    {
        var doc     = new PolyDonkyument();
        var section = new Section();
        doc.Sections.Add(section);

        var cur = new StringBuilder();
        foreach (char c in raw)
        {
            switch (c)
            {
                case '\r':
                case '\f':  // page break — 단락 분리로 처리
                    AppendParagraph(section, cur);
                    break;
                case '\v':  // soft line break
                    cur.Append('\n');
                    break;
                case '':  // cell mark
                    cur.Append('\t');
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
                    cur.Append(c);
                    break;
                default:
                    if (c < 0x20) break;  // 그 외 제어 문자 폐기
                    cur.Append(c);
                    break;
            }
        }
        AppendParagraph(section, cur);

        if (section.Blocks.Count == 0) section.Blocks.Add(new Paragraph());
        return doc;
    }

    private static void AppendParagraph(Section section, StringBuilder cur)
    {
        var text = cur.ToString();
        cur.Clear();
        var para = new Paragraph();
        if (text.Length > 0) para.AddText(text);
        section.Blocks.Add(para);
    }

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
