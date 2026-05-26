using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using OpenMcdf;
using PolyDonky.Core;

namespace PolyDonky.Convert.Doc;

/// <summary>
/// Word 97-2003 binary (.doc, OLE2 Compound File) → IWPF 변환기 (Phase 1a~1g — 텍스트·단락·서식·폰트·스타일·체인 상속).
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
///   - PAPX 운영 → 정렬(sprmPJc80 0x2461) / 들여쓰기(sprmPDxaLeft 0x845D, sprmPDxaRight 0x845E,
///     sprmPDxaLeft1 0x8460) / 단락 간격(sprmPDyaBefore 0xA413, sprmPDyaAfter 0xA415) /
///     줄 간격(sprmPDyaLine 0x6412 LSPD) / istd → STSH lookup → OutlineLevel(H1..H6)  (Phase 1b/1c/1e)
///   - CHPX 운영 → 굵게(0x0835)·이탤릭(0x0836)·취소선(0x0837)·밑줄(0x2A3E)·
///     글자 크기(sprmCHps 0x4A43) / 전경색(sprmCIco 0x2A42 팔레트, sprmCCv 0x6870 RGB) /
///     하이라이트(sprmCHighlight 0x2A0C) / 폰트 패밀리(sprmCRgFtc0/1/2 0x4A4F/0x4A50/0x4A51
///     + STTB FFN 운영) 로 Run 분할                                                  (Phase 1b/1c/1d)
///   - SummaryInformation PropertySet (MS-OLEPS §2.18) 에서 Title/Subject/Author/Keywords/
///     LastSavedBy/Created/Modified/AppName 추출 (VT_LPSTR/VT_LPWSTR/VT_FILETIME)
///   - 비-OLE2 입력·손상 CFB 컨테이너는 한국어 진단으로 명확히 거부
///
/// 의도적으로 다음 작업의 범위 밖 (다음 PR 들에서 다룸):
///   - 표 / 이미지 / 필드 / 헤더·푸터 / 섹션                                          (Phase 2)
///   - MS-OFFCRYPTO 의 실제 해독 (RC4/AES 키 유도·복호화)
///   - 리스트·번호 (LST/LFO) 와 STSH 의 stk=4 numbering 스타일                       (Phase 3)
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
        uint   LcbPlcfBtePapx,
        // Phase 1d — STTB FFN
        uint   FcSttbfFfn,
        uint   LcbSttbfFfn,
        // Phase 1e — Style Sheet (STSH)
        uint   FcStshf,
        uint   LcbStshf);

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

        // Phase 1d — STTB FFN: [MS-DOC] FibRgFcLcb97 fcSttbfFfn @ 0x0112, lcbSttbfFfn @ 0x0116
        uint fcSttbfFfn  = BitConverter.ToUInt32(wd, 0x0112);
        uint lcbSttbfFfn = BitConverter.ToUInt32(wd, 0x0116);

        // Phase 1e — STSH (Style Sheet): [MS-DOC] FibRgFcLcb97 fcStshf @ 0x00A2, lcbStshf @ 0x00A6
        uint fcStshf  = BitConverter.ToUInt32(wd, 0x00A2);
        uint lcbStshf = BitConverter.ToUInt32(wd, 0x00A6);

        return new Fib(tableName, fcMin, ccpText, fcClx, lcbClx, nFib, encrypted, obfuscated,
                       fcPlcfBteChpx, lcbPlcfBteChpx, fcPlcfBtePapx, lcbPlcfBtePapx,
                       fcSttbfFfn, lcbSttbfFfn, fcStshf, lcbStshf);
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
        // Phase 2a — 표 누적 상태 (pending: 진행 중인 표/행).
        Table? pendingTable = null;
        TableRow? pendingRow = null;

        for (int i = 0; i < raw.Length; i++)
        {
            char c  = raw[i];
            int  fc = i < fcs.Length ? fcs[i] : lastFc;
            lastFc = fc;

            switch (c)
            {
                case '\r':
                case '\f':  // page break — 단락 분리로 처리
                    FlushParagraph(section, paraChars, paraFcs, fc, fmt, ref pendingTable, ref pendingRow);
                    break;
                case '\v':  // soft line break
                    paraChars.Add('\n'); paraFcs.Add(fc);
                    break;
                case '':  // cell mark
                    paraChars.Add('\u0007'); paraFcs.Add(fc);
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
        FlushParagraph(section, paraChars, paraFcs, lastFc, fmt, ref pendingTable, ref pendingRow);
        // 본문 끝에서 표가 미완 상태이면 마감.
        if (pendingTable is not null)
        {
            if (pendingRow is { Cells.Count: > 0 }) pendingTable.Rows.Add(pendingRow);
            if (pendingTable.Rows.Count > 0) section.Blocks.Add(pendingTable);
        }

        if (section.Blocks.Count == 0) section.Blocks.Add(new Paragraph());
        return doc;
    }

    // Phase 1b — 누적된 (char, fc) 쌍을 한 단락으로 묶어 만든다.
    // Phase 2a — InTable=true 단락은 pendingTable 에 셀별로 누적, IsTtp 단락은 행 종료,
    //          비-InTable 단락은 표를 마감하고 section.Blocks 에 직접 추가.
    private static void FlushParagraph(
        Section section, List<char> paraChars, List<int> paraFcs, int paraEndFc, FormatStyles fmt,
        ref Table? pendingTable, ref TableRow? pendingRow)
    {
        var (paraIstd, ps, inTable, isTtp) = fmt.GetParagraphInfo(paraEndFc);

        // 행 종료 단락 (TTP). 비어 있고 행을 마무리하는 신호.
        if (isTtp)
        {
            pendingTable ??= new Table();
            if (pendingRow is { Cells.Count: > 0 }) pendingTable.Rows.Add(pendingRow);
            pendingRow = null;
            paraChars.Clear();
            paraFcs.Clear();
            return;
        }

        // 표 안 단락. 0x07 셀 마커로 분리해서 셀별 단락을 누적.
        if (inTable)
        {
            pendingTable ??= new Table();
            pendingRow   ??= new TableRow();
            SplitIntoCells(paraChars, paraFcs, paraIstd, fmt, pendingRow);
            paraChars.Clear();
            paraFcs.Clear();
            return;
        }

        // 비-InTable 단락 — 표가 진행 중이면 마감.
        if (pendingTable is not null)
        {
            if (pendingRow is { Cells.Count: > 0 }) pendingTable.Rows.Add(pendingRow);
            if (pendingTable.Rows.Count > 0) section.Blocks.Add(pendingTable);
            pendingTable = null;
            pendingRow   = null;
        }

        var para = BuildParaFromChars(paraChars, paraFcs, paraIstd, ps, fmt);
        section.Blocks.Add(para);
        paraChars.Clear();
        paraFcs.Clear();
    }

    // 단일 셀 단락 만들기 — Phase 1 의 Run 분할 알고리즘 재사용.
    private static Paragraph BuildParaFromChars(
        List<char> chars, List<int> fcs, int paraIstd, ParagraphStyle? ps, FormatStyles fmt)
    {
        var para = new Paragraph();
        if (ps is not null) para.Style = ps;
        if (chars.Count == 0) return para;

        RunStyle? curStyle = null;
        var curText = new StringBuilder();
        for (int i = 0; i < chars.Count; i++)
        {
            var rs = fmt.GetRunStyle(fcs[i], paraIstd) ?? new RunStyle();
            if (curStyle is null || !RunStyleEquals(curStyle, rs))
            {
                if (curText.Length > 0 && curStyle is not null)
                    para.AddText(curText.ToString(), curStyle);
                curStyle = rs;
                curText.Clear();
            }
            curText.Append(chars[i]);
        }
        if (curText.Length > 0 && curStyle is not null)
            para.AddText(curText.ToString(), curStyle);
        return para;
    }

    // Phase 2a — paraChars 의 0x07 (cell mark) 위치마다 셀을 나눠 pendingRow 에 추가.
    // 셀 안에는 단일 단락 가정 (멀티-단락 셀은 후속 단계).
    private static void SplitIntoCells(
        List<char> paraChars, List<int> paraFcs, int paraIstd, FormatStyles fmt, TableRow pendingRow)
    {
        var curChars = new List<char>();
        var curFcs   = new List<int>();
        for (int i = 0; i < paraChars.Count; i++)
        {
            if (paraChars[i] == '')
            {
                var cell = new TableCell();
                cell.Blocks.Add(BuildParaFromChars(curChars, curFcs, paraIstd, null, fmt));
                pendingRow.Cells.Add(cell);
                curChars.Clear();
                curFcs.Clear();
            }
            else
            {
                curChars.Add(paraChars[i]);
                curFcs.Add(paraFcs[i]);
            }
        }
        // 잔여 텍스트 — 0x07 없이 단락이 끝났다면 (셀의 미완 단락) 별개 셀로 처리.
        if (curChars.Count > 0)
        {
            var cell = new TableCell();
            cell.Blocks.Add(BuildParaFromChars(curChars, curFcs, paraIstd, null, fmt));
            pendingRow.Cells.Add(cell);
        }
    }

    private static bool RunStyleEquals(RunStyle a, RunStyle b)
        => a.Bold == b.Bold
        && a.Italic == b.Italic
        && a.Underline == b.Underline
        && a.Strikethrough == b.Strikethrough
        && a.FontSizePt == b.FontSizePt
        && string.Equals(a.FontFamily, b.FontFamily, StringComparison.Ordinal)
        && Nullable.Equals(a.Foreground, b.Foreground)
        && Nullable.Equals(a.Background, b.Background);

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
    //   PAPX:  sprmPJc80     (0x2461)  — 단락 정렬 1 byte (0=L, 1=C, 2=R, 3=J)
    //          sprmPDxaLeft  (0x845D)  — 왼쪽 들여쓰기 (signed twips, 2-byte)
    //          sprmPDxaRight (0x845E)  — 오른쪽 들여쓰기 (signed twips, 2-byte)
    //          sprmPDxaLeft1 (0x8460)  — 첫 줄 들여쓰기 (signed twips, 2-byte)
    //          sprmPDyaBefore(0xA413)  — 단락 앞 여백 (unsigned twips, 2-byte)
    //          sprmPDyaAfter (0xA415)  — 단락 뒤 여백 (unsigned twips, 2-byte)
    //          sprmPDyaLine  (0x6412)  — 줄 간격 LSPD (4-byte: dyaLine + fMultLinespace)
    //   CHPX:  sprmCFBold (0x0835), sprmCFItalic (0x0836)        — 1 byte toggle/on/off
    //          sprmCFStrike    (0x0837)                           — 1 byte
    //          sprmCKul        (0x2A3E)                           — 1 byte (0=none, 그 외=ON)
    //          sprmCHps        (0x4A43)                           — 2 byte 글자크기 (half-points)
    //          sprmCIco        (0x2A42)                           — 1 byte Word16 팔레트 전경색
    //          sprmCCv         (0x6870)                           — 4 byte RGB+1 (Word 2002+) 전경색
    //          sprmCHighlight  (0x2A0C)                           — 1 byte Word16 팔레트 하이라이트
    // 그 외 sprm 은 본 단계에서 무시 — Phase 1d/2 에서 점진적으로 추가한다 (폰트 패밀리는 STSH/STTB FFN 필요).

    private sealed class FormatStyles
    {
        private readonly byte[] _wd;
        private readonly List<BteEntry> _papxBte;
        private readonly List<BteEntry> _chpxBte;
        // Phase 1d — STTB FFN 에서 추출한 폰트명. sprmCRgFtc{0,1,2} 의 operand 인 ftc 가 이 인덱스.
        private readonly IReadOnlyList<string> _fonts;
        // Phase 1e — STSH (Style Sheet) 의 STD 배열. istd → sti/stk/name. null = 빈 슬롯.
        private readonly IReadOnlyList<StyleDef?> _styles;

        // FKP 페이지(512 byte) 파싱은 매번 동일 데이터를 다시 만지지 않도록 page → grpprl 캐시.
        // PAPX 는 (istd, sprms) 가 페어로 필요하므로 별도 캐시.
        private readonly Dictionary<(int Pn, int RgIdx), (int Istd, byte[] Sprms)?> _papxCache = new();
        private readonly Dictionary<(int Pn, int RgIdx), byte[]?> _chpxCache = new();

        private FormatStyles(byte[] wd, List<BteEntry> papx, List<BteEntry> chpx,
                             IReadOnlyList<string> fonts, IReadOnlyList<StyleDef?> styles)
        {
            _wd = wd; _papxBte = papx; _chpxBte = chpx; _fonts = fonts; _styles = styles;
        }

        public static FormatStyles Build(byte[] wd, byte[] table, Fib fib)
        {
            var papx   = ReadBte(table, (int)fib.FcPlcfBtePapx, (int)fib.LcbPlcfBtePapx);
            var chpx   = ReadBte(table, (int)fib.FcPlcfBteChpx, (int)fib.LcbPlcfBteChpx);
            var fonts  = ReadSttbfFfn(table, (int)fib.FcSttbfFfn, (int)fib.LcbSttbfFfn);
            var styles = ReadStsh(table, (int)fib.FcStshf, (int)fib.LcbStshf);
            return new FormatStyles(wd, papx, chpx, fonts, styles);
        }

        // Phase 1f — 단락 정보를 (istd, ParagraphStyle?, InTable, IsTtp) 4-tuple 로 반환.
        // Phase 2a — InTable / IsTtp 플래그가 추가됨 (sprmPFInTable / sprmPFTtp).
        public (int Istd, ParagraphStyle? Style, bool InTable, bool IsTtp) GetParagraphInfo(int paraEndFc)
        {
            var papx = LoadPapx(paraEndFc);
            if (papx is null) return (-1, null, false, false);

            var (istd, directSprms) = papx.Value;
            var style = new ParagraphStyle();
            bool touched = false;
            bool inTable = false;
            bool isTtp   = false;

            // 1. STSH built-in sti → Outline (Heading N → HN).
            if (istd >= 0 && istd < _styles.Count && _styles[istd] is { } sd && sd.Sti >= 1 && sd.Sti <= 9)
            {
                int level = Math.Min(sd.Sti, 6);
                style.Outline = (OutlineLevel)level;
                touched = true;
            }

            // 2. Phase 1g — istdBase 체인을 따라 root → leaf 순으로 STD PAPX sprms + 표 플래그 적용.
            foreach (int chainIstd in ResolveStyleChain(istd))
            {
                if (_styles[chainIstd]?.PapxSprms is { Length: > 0 } chainSprms)
                {
                    touched |= ApplyParagraphSprms(chainSprms, style);
                    ScanTableFlags(chainSprms, ref inTable, ref isTtp);
                }
            }

            // 3. 직접 PAPX sprms — 스타일 상속값을 덮어쓴다.
            touched |= ApplyParagraphSprms(directSprms, style);
            ScanTableFlags(directSprms, ref inTable, ref isTtp);

            return (istd, touched ? style : null, inTable, isTtp);
        }

        // [MS-DOC] sprmPFInTable (0x2416, 1-byte): 단락이 표 안에 있는지.
        //          sprmPFTtp    (0x2417, 1-byte): 단락이 행 종료 단락(TTP)인지.
        // ref 매개변수는 람다 캡처 불가라 로컬에 모은 뒤 호출부에서 합친다.
        private static void ScanTableFlags(byte[] grpprl, ref bool inTable, ref bool isTtp)
        {
            bool localIn  = inTable;
            bool localTtp = isTtp;
            WalkSprms(grpprl, (sprm, operand) =>
            {
                if (sprm == 0x2416 && operand.Length >= 1) localIn = operand[0] != 0;
                else if (sprm == 0x2417 && operand.Length >= 1) localTtp = operand[0] != 0;
            });
            inTable = localIn;
            isTtp   = localTtp;
        }

        // Phase 1g — istd 부터 istdBase 를 따라가 root 까지 chain 을 모은 뒤 root 부터 순회.
        // 부모가 먼저 적용되고 자식이 덮어쓰는 순서. 순환 참조와 nil(0xFFF) 종료를 모두 처리.
        private IEnumerable<int> ResolveStyleChain(int startIstd)
        {
            if (startIstd < 0 || startIstd >= _styles.Count) yield break;
            var chain  = new List<int>();
            var seen   = new HashSet<int>();
            int cur    = startIstd;
            while (cur >= 0 && cur < _styles.Count && cur != IstdNil && seen.Add(cur))
            {
                var sd = _styles[cur];
                if (sd is null) break;
                chain.Add(cur);
                cur = sd.IstdBase;
            }
            // root → leaf 순서로 yield (부모 먼저 → 자식 덮어쓰기).
            for (int i = chain.Count - 1; i >= 0; i--)
                yield return chain[i];
        }

        private static bool ApplyParagraphSprms(byte[] grpprl, ParagraphStyle style)
        {
            bool touched = false;
            WalkSprms(grpprl, (sprm, operand) =>
            {
                if (ApplyParagraphSprm(sprm, operand, style)) touched = true;
            });
            return touched;
        }

        private static bool ApplyParagraphSprm(ushort sprm, byte[] operand, ParagraphStyle style)
        {
            switch (sprm)
            {
                case 0x2461:  // sprmPJc80 — alignment
                    if (operand.Length >= 1)
                    {
                        style.Alignment = operand[0] switch
                        {
                            1 => Alignment.Center,
                            2 => Alignment.Right,
                            3 => Alignment.Justify,
                            _ => Alignment.Left,
                        };
                        return true;
                    }
                    return false;
                case 0x845D:
                    if (operand.Length >= 2)
                    { style.IndentLeftMm  = BitConverter.ToInt16(operand, 0) * TwipsToMm; return true; }
                    return false;
                case 0x845E:
                    if (operand.Length >= 2)
                    { style.IndentRightMm = BitConverter.ToInt16(operand, 0) * TwipsToMm; return true; }
                    return false;
                case 0x8460:
                    if (operand.Length >= 2)
                    { style.IndentFirstLineMm = BitConverter.ToInt16(operand, 0) * TwipsToMm; return true; }
                    return false;
                case 0xA413:
                    if (operand.Length >= 2)
                    { style.SpaceBeforePt = BitConverter.ToUInt16(operand, 0) / 20.0; return true; }
                    return false;
                case 0xA415:
                    if (operand.Length >= 2)
                    { style.SpaceAfterPt  = BitConverter.ToUInt16(operand, 0) / 20.0; return true; }
                    return false;
                case 0x6412:
                    if (operand.Length >= 4)
                    {
                        short  dyaLine = BitConverter.ToInt16(operand, 0);
                        ushort fMult   = BitConverter.ToUInt16(operand, 2);
                        if (fMult == 1 && dyaLine > 0)
                        { style.LineHeightFactor = dyaLine / 240.0; return true; }
                        if (fMult == 0 && dyaLine != 0)
                        {
                            double abs = Math.Abs(dyaLine) / 240.0;
                            if (abs > 0.5 && abs < 5.0)
                            { style.LineHeightFactor = abs; return true; }
                        }
                    }
                    return false;
            }
            return false;
        }

        // Twips → mm 변환 — 1 mm = 56.692 twips (1440/25.4).
        private const double TwipsToMm = 1.0 / 56.692;

        // Phase 1f — paraIstd 가 -1 아니면 단락 스타일의 STD chpxSprms 를 먼저 적용한 뒤
        // 직접 CHPX 로 override. Heading 의 폰트/크기/굵게 등이 자동 상속된다.
        public RunStyle? GetRunStyle(int charFc, int paraIstd)
        {
            var rs = new RunStyle();
            bool touched = false;

            // 1. Phase 1g — 단락 스타일의 istdBase 체인을 따라 root → leaf 순으로 STD CHPX sprms 적용.
            //    예: Heading 1 의 폰트가 Normal 의 폰트 패밀리를 상속.
            foreach (int chainIstd in ResolveStyleChain(paraIstd))
            {
                if (_styles[chainIstd]?.ChpxSprms is { Length: > 0 } chainChpx)
                    touched |= ApplyRunSprms(chainChpx, rs);
            }

            // 2. 직접 CHPX FKP sprms (override).
            var direct = LoadChpx(charFc);
            if (direct is { Length: > 0 })
                touched |= ApplyRunSprms(direct, rs);

            return touched ? rs : null;
        }

        private bool ApplyRunSprms(byte[] grpprl, RunStyle rs)
        {
            bool touched = false;
            WalkSprms(grpprl, (sprm, operand) =>
            {
                if (ApplyRunSprm(sprm, operand, rs)) touched = true;
            });
            return touched;
        }

        private bool ApplyRunSprm(ushort sprm, byte[] operand, RunStyle rs)
        {
            switch (sprm)
            {
                case 0x0835:  // sprmCFBold
                    if (operand.Length >= 1) { rs.Bold = operand[0] != 0; return true; }
                    return false;
                case 0x0836:  // sprmCFItalic
                    if (operand.Length >= 1) { rs.Italic = operand[0] != 0; return true; }
                    return false;
                case 0x0837:  // sprmCFStrike
                    if (operand.Length >= 1) { rs.Strikethrough = operand[0] != 0; return true; }
                    return false;
                case 0x2A3E:  // sprmCKul
                    if (operand.Length >= 1) { rs.Underline = operand[0] != 0; return true; }
                    return false;
                case 0x4A43:  // sprmCHps (font size in half-points)
                    if (operand.Length >= 2)
                    {
                        ushort halfPt = BitConverter.ToUInt16(operand, 0);
                        if (halfPt > 0 && halfPt < 1000) { rs.FontSizePt = halfPt / 2.0; return true; }
                    }
                    return false;
                case 0x2A42:  // sprmCIco
                    if (operand.Length >= 1)
                    {
                        var c = WordPaletteColor(operand[0]);
                        if (c.HasValue) { rs.Foreground = c.Value; return true; }
                    }
                    return false;
                case 0x6870:  // sprmCCv
                    if (operand.Length >= 4 && operand[3] != 0xFF)
                    {
                        rs.Foreground = new Color(operand[0], operand[1], operand[2]);
                        return true;
                    }
                    return false;
                case 0x2A0C:  // sprmCHighlight
                    if (operand.Length >= 1)
                    {
                        var c = WordPaletteColor(operand[0]);
                        if (c.HasValue) { rs.Background = c.Value; return true; }
                    }
                    return false;
                case 0x4A4F:  // sprmCRgFtc0
                case 0x4A50:  // sprmCRgFtc1
                case 0x4A51:  // sprmCRgFtc2
                    if (operand.Length >= 2)
                    {
                        int ftc = BitConverter.ToUInt16(operand, 0);
                        if (ftc >= 0 && ftc < _fonts.Count)
                        {
                            var name = _fonts[ftc];
                            if (!string.IsNullOrEmpty(name))
                            { rs.FontFamily = name; return true; }
                        }
                    }
                    return false;
            }
            return false;
        }

        // [MS-DOC] Word 16-color palette (sprmCIco, sprmCHighlight 의 1-byte 인덱스).
        // 0 = auto (부모 색 상속 — null 반환).
        private static Color? WordPaletteColor(byte ico) => ico switch
        {
            1  => new Color(0,   0,   0),     //  1 Black
            2  => new Color(0,   0,   255),   //  2 Blue
            3  => new Color(0,   255, 255),   //  3 Cyan
            4  => new Color(0,   255, 0),     //  4 Green
            5  => new Color(255, 0,   255),   //  5 Magenta
            6  => new Color(255, 0,   0),     //  6 Red
            7  => new Color(255, 255, 0),     //  7 Yellow
            8  => new Color(255, 255, 255),   //  8 White
            9  => new Color(0,   0,   128),   //  9 Dark Blue
            10 => new Color(0,   128, 128),   // 10 Dark Cyan
            11 => new Color(0,   128, 0),     // 11 Dark Green
            12 => new Color(128, 0,   128),   // 12 Dark Magenta
            13 => new Color(128, 0,   0),     // 13 Dark Red
            14 => new Color(128, 128, 0),     // 14 Dark Yellow
            15 => new Color(128, 128, 128),   // 15 Dark Gray
            16 => new Color(192, 192, 192),   // 16 Light Gray
            _  => null,                       //  0 auto / 그 외 알 수 없음
        };

        // PAPX 로드 — istd 와 sprm bytes 를 한 쌍으로 반환. STSH lookup 에 istd 가 필요해 분리.
        private (int Istd, byte[] Sprms)? LoadPapx(int fc)
        {
            var loc = LocateFkpEntry(_papxBte, fc);
            if (loc is null) return null;
            var (pn, rgIdx) = loc.Value;
            if (_papxCache.TryGetValue((pn, rgIdx), out var cached)) return cached;

            int fkpOff = pn * 512;
            int cRun   = _wd[fkpOff + 511];
            int rgbxBase = fkpOff + 4 * (cRun + 1);
            // BXPap = 13 byte; first byte = bOffset (2-byte units from FKP start).
            int bOffset = _wd[rgbxBase + rgIdx * 13];
            var result  = bOffset == 0 ? null : ReadPapxInFkp(_wd, fkpOff + bOffset * 2);
            _papxCache[(pn, rgIdx)] = result;
            return result;
        }

        // CHPX 로드 — sprm bytes 만 (istd 없음).
        private byte[]? LoadChpx(int fc)
        {
            var loc = LocateFkpEntry(_chpxBte, fc);
            if (loc is null) return null;
            var (pn, rgIdx) = loc.Value;
            if (_chpxCache.TryGetValue((pn, rgIdx), out var cached)) return cached;

            int fkpOff = pn * 512;
            int cRun   = _wd[fkpOff + 511];
            int rgbxBase = fkpOff + 4 * (cRun + 1);
            // ChpxFkp rgb = 1 byte each.
            int bOffset = _wd[rgbxBase + rgIdx];
            var result  = bOffset == 0 ? null : ReadChpxInFkp(_wd, fkpOff + bOffset * 2);
            _chpxCache[(pn, rgIdx)] = result;
            return result;
        }

        private (int Pn, int RgIdx)? LocateFkpEntry(List<BteEntry> bte, int fc)
        {
            if (bte.Count == 0) return null;
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
            for (int i = 0; i < cRun; i++)
            {
                int fc0 = BitConverter.ToInt32(_wd, fkpOff + i * 4);
                int fc1 = BitConverter.ToInt32(_wd, fkpOff + (i + 1) * 4);
                if (fc >= fc0 && fc < fc1) return (pn, i);
            }
            return null;
        }

        private static (int Istd, byte[] Sprms)? ReadPapxInFkp(byte[] data, int off)
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
            // GrpPrlAndIstd: 처음 2 byte = istd (style identifier), 그 뒤가 sprm 배열.
            int istd = BitConverter.ToUInt16(data, grpprlStart);
            int sprmsLen = grpprlSize - 2;
            var sprms = new byte[sprmsLen];
            Buffer.BlockCopy(data, grpprlStart + 2, sprms, 0, sprmsLen);
            return (istd, sprms);
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

        // [MS-DOC] §2.9.262 SttbfFfn — Word 97+ 의 폰트 이름 STTB. 각 원소는 §2.9.85 FFN.
        // Extended STTB header (6 byte): 0xFFFF marker + cData(2) + cbExtra(2).
        // 각 entry: cchData(2 byte, FFN 의 wide-char 수) + FFN(cchData*2 byte) + cbExtra byte.
        // FFN 내부에서 폰트명(xszFfn) 은 offset 40 부터 null-terminated UTF-16LE.
        private static IReadOnlyList<string> ReadSttbfFfn(byte[] table, int fc, int lcb)
        {
            var fonts = new List<string>();
            if (lcb <= 0 || fc < 0 || fc + lcb > table.Length) return fonts;

            int pos = fc;
            int end = fc + lcb;
            bool extended;
            int  cbExtra;
            if (end - pos >= 6 && BitConverter.ToUInt16(table, pos) == 0xFFFF)
            {
                extended = true;
                cbExtra  = BitConverter.ToUInt16(table, pos + 4);
                pos += 6;
            }
            else if (end - pos >= 4)
            {
                // 비-extended (legacy) — Word 97+ 에서는 거의 없지만 fallback.
                extended = false;
                cbExtra  = BitConverter.ToUInt16(table, pos + 2);
                pos += 4;
            }
            else
            {
                return fonts;
            }

            while (pos < end)
            {
                int hdrLen = extended ? 2 : 1;
                if (pos + hdrLen > end) break;
                int cchData = extended ? BitConverter.ToUInt16(table, pos) : table[pos];
                int ffnByteLen = cchData * 2;
                if (pos + hdrLen + ffnByteLen + cbExtra > end) break;

                int ffnStart = pos + hdrLen;
                string name = string.Empty;
                if (ffnByteLen > 40)
                {
                    int nameStart = ffnStart + 40;
                    int nameMax   = ffnStart + ffnByteLen;
                    int nameEnd   = nameStart;
                    while (nameEnd + 1 < nameMax && BitConverter.ToUInt16(table, nameEnd) != 0)
                        nameEnd += 2;
                    if (nameEnd > nameStart)
                        name = Encoding.Unicode.GetString(table, nameStart, nameEnd - nameStart);
                }
                fonts.Add(name);

                pos += hdrLen + ffnByteLen + cbExtra;
            }
            return fonts;
        }

        // [MS-DOC] §2.9.270 STD 의 핵심 필드 — sti(12 bit): built-in style identifier
        // (0=Normal, 1..9=Heading 1..9), stk(4 bit): 1=paragraph, 2=character, 3=table, 4=numbering.
        // Phase 1f — STD 의 grLPUpxSw 에서 추출한 단락/문자 sprm 기본값.
        // Phase 1g — IstdBase(12 bit): 상속 부모 styleId. 0xFFF = nil (체인 종료).
        private sealed record StyleDef(int Sti, int Stk, int IstdBase, string? Name,
                                       byte[]? PapxSprms, byte[]? ChpxSprms);

        private const int IstdNil = 0xFFF;

        // [MS-DOC] §2.9.271 STSH 파서.
        // STSH = LPStshi (2 byte cbStshi + Stshi) + rgLPStd[cstd] (각 LPStd: 2 byte cbStd + STD).
        // 본 단계에서는 STD 의 sti/stk 만 추출 — STD grpprl 상속(default 서식)은 후속 단계.
        private static IReadOnlyList<StyleDef?> ReadStsh(byte[] table, int fc, int lcb)
        {
            var empty = Array.Empty<StyleDef?>();
            if (lcb < 4 || fc < 0 || fc + lcb > table.Length) return empty;

            int pos = fc;
            int end = fc + lcb;

            // LPStshi: cbStshi(2) + Stshi(cbStshi)
            int cbStshi = BitConverter.ToUInt16(table, pos);
            pos += 2;
            if (cbStshi < 4 || pos + cbStshi > end) return empty;

            // Stshi 의 첫 4 byte: cstd(2) + cbSTDBaseInFile(2)
            int cstd            = BitConverter.ToUInt16(table, pos);
            int cbSTDBaseInFile = BitConverter.ToUInt16(table, pos + 2);
            pos += cbStshi;

            if (cstd <= 0 || cstd > 8192) return empty;
            if (cbSTDBaseInFile is not (10 or 18)) cbSTDBaseInFile = 10;

            var styles = new StyleDef?[cstd];
            for (int i = 0; i < cstd && pos + 2 <= end; i++)
            {
                int cbStd = BitConverter.ToUInt16(table, pos);
                pos += 2;
                if (cbStd == 0) { styles[i] = null; continue; }
                if (pos + cbStd > end) break;

                int stdStart = pos;
                int stdEnd   = stdStart + cbStd;
                if (cbStd >= 4)
                {
                    ushort word0 = BitConverter.ToUInt16(table, stdStart);
                    ushort word1 = BitConverter.ToUInt16(table, stdStart + 2);
                    ushort word2 = cbStd >= 6 ? BitConverter.ToUInt16(table, stdStart + 4) : (ushort)0;
                    int sti       = word0 & 0x0FFF;
                    int stk       = word1 & 0x000F;
                    int istdBase  = (word1 >> 4) & 0x0FFF;  // 부모 styleId — 0xFFF 면 nil.
                    int cupx      = word2 & 0x000F;

                    // xstzName — stdfBase(+stdfPost2000) 뒤. Xstz: cchData(2) + chars(2*cchData) + null(2).
                    string? name = null;
                    int nameOff  = stdStart + cbSTDBaseInFile;
                    int afterName = nameOff;
                    if (nameOff + 2 <= stdEnd)
                    {
                        int cch = BitConverter.ToUInt16(table, nameOff);
                        int nameStart = nameOff + 2;
                        int nameByteLen = cch * 2;
                        if (nameStart + nameByteLen + 2 <= stdEnd)
                        {
                            if (cch > 0)
                                name = Encoding.Unicode.GetString(table, nameStart, nameByteLen);
                            afterName = nameStart + nameByteLen + 2;  // +null terminator
                        }
                        else
                        {
                            afterName = stdEnd;  // 손상된 STD — grLPUpxSw 진입 안 함
                        }
                    }
                    // 2-byte 정렬은 STD 내부 offset 기준이라 stdStart 의 parity 와 비교.
                    if (((afterName - stdStart) & 1) != 0) afterName++;

                    // [MS-DOC] §2.9.135 grLPUpxSw — cupx 개의 LPUpx (각각 cbUpx(2) + UPX(cbUpx)).
                    //   paragraph style(stk=1): LPUpx[0]=PAPX (UPX = istd(2)+grpprl), LPUpx[1]=CHPX (UPX=grpprl)
                    //   character style(stk=2): LPUpx[0]=CHPX
                    var (papx, chpx) = ParseGrLPUpxSw(table, stdStart, afterName, stdEnd, cupx, stk);
                    styles[i] = new StyleDef(sti, stk, istdBase, name, papx, chpx);
                }
                pos += cbStd;
            }
            return styles;
        }

        // [MS-DOC] §2.9.135 grLPUpxSw — 한 STD 안의 cupx 개 LPUpx 를 순서대로 파싱.
        // cupx=2 paragraph style: LPUpx[0]=PAPX (UPX는 istd(2)+grpprl), LPUpx[1]=CHPX (UPX=grpprl)
        // cupx=1 character style: LPUpx[0]=CHPX
        // (table/numbering style 은 본 단계에서 무시.)
        private static (byte[]? Papx, byte[]? Chpx) ParseGrLPUpxSw(
            byte[] table, int stdStart, int start, int end, int cupx, int stk)
        {
            byte[]? papx = null;
            byte[]? chpx = null;
            int pos = start;
            for (int idx = 0; idx < cupx && pos + 2 <= end; idx++)
            {
                int cbUpx = BitConverter.ToUInt16(table, pos);
                pos += 2;
                if (pos + cbUpx > end) break;

                if (stk == 1 && idx == 0)
                {
                    // Paragraph style 의 PAPX UPX: istd(2 byte) + grpprl. istd 는 styleId 라 무시.
                    if (cbUpx > 2)
                    {
                        int sprmLen = cbUpx - 2;
                        papx = new byte[sprmLen];
                        Buffer.BlockCopy(table, pos + 2, papx, 0, sprmLen);
                    }
                }
                else if ((stk == 1 && idx == 1) || (stk == 2 && idx == 0))
                {
                    // CHPX UPX = grpprl 그대로.
                    if (cbUpx > 0)
                    {
                        chpx = new byte[cbUpx];
                        Buffer.BlockCopy(table, pos, chpx, 0, cbUpx);
                    }
                }
                pos += cbUpx;
                // LPUpx 는 2-byte 정렬 — STD 내부 offset 기준.
                if (((pos - stdStart) & 1) != 0) pos++;
            }
            return (papx, chpx);
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
