using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using OpenMcdf;
using PolyDonky.Core;

namespace PolyDonky.Convert.Doc;

/// <summary>
/// IWPF → Word 97-2003 (.doc) OLE2 바이너리 writer.
///
/// 현재 단계 — Phase F1-W2e (FidelityCapsules round-trip):
///   * CFB 컨테이너 + FIB + CLX (단일 piece) + 본문 텍스트
///   * SttbfFfn (폰트 테이블)
///   * PAPX FKP 페이지 — 단락별 정렬·들여쓰기·간격·줄높이
///   * CHPX FKP 페이지 — Run 별 굵게/이탤릭/밑줄/취소선·글자크기·전경색·폰트 인덱스
///   * PlcfBtePapx / PlcfBteChpx — FKP 페이지 인덱스
///   * SttbfBkmk + PlcfBkf + PlcfBkl — Run.BookmarkStart/End 책갈피
///   * ImageBlock / ShapeObject / TextBoxObject / Table → 가시 placeholder (텍스트 정보 보존)
///   * FidelityCapsules (msdoc/macros / msdoc/signatures / msdoc/preserved) → 원래 OLE2 storage 로 복원.
///     VBA 매크로 / 디지털 서명 / 미인식 root storage 가 .doc → .iwpf → .doc 라운드트립을 통과.
///
/// 후속 단계 (v1.0.0+):
///   헤더/푸터, 섹션 SEPX, 각주/미주, 주석, 필드, OfficeArt 이미지/도형 binary 임베드
///
/// 참고 사양:
///   [MS-CFB]  Compound File Binary File Format        (OpenMcdf 가 처리)
///   [MS-DOC]  Word (.doc) Binary File Format §2.5 FIB, §2.6 Sprm, §2.7 FKP,
///             §2.8 PlcBteFcN, §2.9.85 FFN / §2.9.262 SttbfFfn, §2.9.177 Pcd
/// </summary>
public sealed class DocBinaryWriter
{
    // ── FIB 레이아웃 상수 ───────────────────────────────────────────────────────

    private const ushort Magic     = 0xA5EC;
    private const ushort NFib      = 0x00C1;          // Word 97-2003 표준 (193)
    private const ushort NFibBack  = 0x00BF;
    /// <summary>FibBase flags @ 0x000A — [MS-DOC] §2.5.1 ABCDEFGH/IJKLMNOP:
    ///   bit 2  (0x0004) C = fComplex            (CLX piece table 사용)
    ///   bits 4-7 (0x00F0) EFGH = cQuickSaves    ([MS-DOC] MUST be 0xF when nFib &lt; 0x00D9)
    ///   bit 9  (0x0200) J = fWhichTblStm        (1 → "1Table" 사용)
    /// → 0x0004 | 0x00F0 | 0x0200 = 0x02F4.
    /// 누락 시 한글/Word Doc Viewer 가 "지원하지 않는 파일 형식" 으로 거부.</summary>
    private const ushort FibFlags  = 0x02F4;
    private const ushort Csw       = 0x000E;          // FibRgW97 size (14 shorts)
    private const ushort Cslw      = 0x0016;          // FibRgLw97 size (22 longs)
    private const ushort CbRgFcLcb = 0x005D;          // FibRgFcLcbBlob pair count (93)
    private const ushort CswNew    = 0x0002;
    /// <summary>FIB 영역 패딩 — Reader 가드(최대 0x0316) + cswNew(0x0382) 까지 충분.</summary>
    private const int    FibPadSize = 0x400;

    /// <summary>Word.Document.8 CLSID (Word 97-2003 binary). 한글/Word 등 OLE consumer 가 이 CLSID 로
    /// 파일 종류를 판단하므로 RootStorage 에 반드시 박혀 있어야 한다.</summary>
    private static readonly Guid WordDocument8ClsId = new("00020906-0000-0000-C000-000000000046");

    // FibRgFcLcbBlob 안의 (fc, lcb) pair 시작 — FIB offset 0x009A. pair index = (offset - 0x9A) / 8.
    private const int RgFcLcbBase = 0x009A;
    private const int PairFcStshf       = 1;    // FIB 0x00A2 / 0x00A6 — STSH (스타일시트)
    private const int PairFcPlcfSed     = 6;    // FIB 0x00CA / 0x00CE — PlcfSed (섹션 테이블)
    private const int PairFcPlcfBteChpx = 12;   // FIB 0x00FA / 0x00FE
    private const int PairFcPlcfBtePapx = 13;   // FIB 0x0102 / 0x0106
    private const int PairFcSttbfFfn    = 15;   // FIB 0x0112 / 0x0116  ← 14 가 아님 (0x10A = pair 14 는 fcPlcfFldMom)
    private const int PairFcClx         = 33;   // FIB 0x01A2 / 0x01A6
    // Phase F1-W2c — bookmarks. FibRgFcLcb97 pair 29 / 30 / 31.
    private const int PairFcSttbfBkmk   = 29;   // FIB 0x0182 / 0x0186
    private const int PairFcPlcfBkf     = 30;   // FIB 0x018A / 0x018E
    private const int PairFcPlcfBkl     = 31;   // FIB 0x0192 / 0x0196

    // FKP 페이지 크기 — [MS-DOC] §2.7 FKP 는 항상 512 byte.
    private const int FkpPageSize = 512;

    // ── 진입점 ──────────────────────────────────────────────────────────────────

    public void Write(PolyDonkyument doc, Stream output)
    {
        var ctx = new BuildContext();
        ctx.Collect(doc);

        byte[] wd    = BuildWordDocument(ctx);
        byte[] table = BuildTable(ctx);

        WriteFib(wd, ctx, (uint)table.Length);

        // OpenMcdf v3 — 비-Transacted 모드는 stream Dispose 시점에 flush. LeaveOpen 으로 호출 측 stream 보존.
        using var root = RootStorage.Create(output, OpenMcdf.Version.V3, StorageModeFlags.LeaveOpen);

        // OLE consumer (한글/Word/LibreOffice 등) 가 root CLSID 로 파일 종류를 식별.
        // 누락 시 한글이 "지원하지 않는 파일 형식" 오류를 내며 거부.
        root.CLSID = WordDocument8ClsId;

        WriteStream(root, "WordDocument", wd);
        WriteStream(root, "1Table",       table);
        // OLE compatibility — 한글/Word 가 root CLSID 와 별도로 \1CompObj / \5SummaryInformation 도
        // 점검할 수 있어 표준 형식의 빈 스트림이라도 박아 둔다.
        WriteStream(root, "CompObj",            BuildCompObjStream());
        WriteStream(root, "SummaryInformation", BuildSummaryInformationStream(doc.Metadata));
        WriteStream(root, "DocumentSummaryInformation", BuildDocumentSummaryInformationStream());

        // Phase F1-W2e — FidelityCapsules 복원 (VBA 매크로 / 디지털 서명 / 미인식 root storage).
        //   Reader 가 .doc → .iwpf 할 때 fidelity/capsules/msdoc/... 로 저장한 raw bytes 를
        //   원래 OLE2 storage 위치로 복원해 .doc → .iwpf → .doc 라운드트립을 닫는다.
        WriteFidelityCapsules(root, doc.FidelityCapsules);
    }

    // ── OLE Compatibility 스트림 ────────────────────────────────────────────────

    /// <summary>\1CompObj — OLE 가 문서 클래스(Word.Document.8) 와 user-friendly 이름을 식별하는 데 사용.
    /// [MS-OLEDS] §2.3.6 CompObjStream 의 최소 변형.
    /// 28-byte 헤더 + LengthPrefixedAnsiString (AnsiUserType) + ClipboardFormatOrAnsiString 만 작성 — 한글이
    /// 요구하는 최소 필드만 채우고 Unicode 섹션은 생략 (UnicodeMarker 없이 종료).</summary>
    private static byte[] BuildCompObjStream()
    {
        using var ms = new MemoryStream();
        var bw = new BinaryWriter(ms);

        // Header (28 byte) — [MS-OLEDS] §2.3.7
        bw.Write((uint)0xFFFE0001);                // Reserved1 magic
        bw.Write((uint)0x00000A03);                // Version (Office 2.0 binary)
        // Reserved2 (20 byte): 16 byte CLSID (Word.Document.8) + 4 byte 0
        bw.Write(WordDocument8ClsId.ToByteArray());
        bw.Write((uint)0);

        // AnsiUserType: LengthPrefixedAnsiString — length(4) + bytes + null
        var userType = "Microsoft Word Document\0";
        bw.Write((uint)userType.Length);
        bw.Write(System.Text.Encoding.ASCII.GetBytes(userType));

        // AnsiClipboardFormat: ClipboardFormatOrAnsiString — length(4) + bytes + null
        var clipFmt = "MSWordDoc\0";
        bw.Write((uint)clipFmt.Length);
        bw.Write(System.Text.Encoding.ASCII.GetBytes(clipFmt));

        // Reserved1: LengthPrefixedAnsiString — Word.Document.8 ProgID (null-terminated)
        var progId = "Word.Document.8\0";
        bw.Write((uint)progId.Length);
        bw.Write(System.Text.Encoding.ASCII.GetBytes(progId));

        // UnicodeMarker + Unicode 섹션은 생략. 한글/Word 는 Ansi 섹션만으로도 동작.
        return ms.ToArray();
    }

    /// <summary>\5SummaryInformation — Document metadata. [MS-OLEPS] §2.18 PropertySetStream.
    /// Title/Author/Application 등 최소 정보만 박는 4KB 미만 스트림.</summary>
    private static byte[] BuildSummaryInformationStream(DocumentMetadata meta)
    {
        // SummaryInformation FMTID = F29F85E0-4FF9-1068-AB91-08002B27B3D9
        // PID 값: 0x02=Title, 0x04=Author, 0x12=AppName ([MS-OLEPS] §2.18)
        var fmtId = new Guid("F29F85E0-4FF9-1068-AB91-08002B27B3D9");
        return BuildPropertySetStream(fmtId, new (int Pid, object? Value)[]
        {
            (0x02, meta.Title),
            (0x04, meta.Author),
            (0x12, meta.Application ?? "PolyDonky"),
        });
    }

    /// <summary>\5DocumentSummaryInformation — 보조 metadata. 한글이 SummaryInformation 만 있어도 동작하지만,
    /// 둘 다 있어야 "정상 Word 문서" 로 인식하는 경우가 있어 빈 PropertySet 으로 채워둠.</summary>
    private static byte[] BuildDocumentSummaryInformationStream()
    {
        // DocumentSummaryInformation FMTID = D5CDD502-2E9C-101B-9397-08002B2CF9AE
        var fmtId = new Guid("D5CDD502-2E9C-101B-9397-08002B2CF9AE");
        return BuildPropertySetStream(fmtId, Array.Empty<(int, object?)>());
    }

    /// <summary>최소 PropertySet 직렬화 — [MS-OLEPS] §2.21 PropertySetStream 의 1-PropertySet 변형.
    /// Header(28) + FMTID/Offset(20) + PropertySet (size + count + idAndOffset[count] + properties).
    /// 지원 Variant: VT_LPSTR (0x001E, ANSI string). 값이 null/빈 문자열이면 property 자체 생략.</summary>
    private static byte[] BuildPropertySetStream(Guid fmtId, IReadOnlyList<(int Pid, object? Value)> entries)
    {
        // PropertySet 본체 먼저 빌드 (size 계산 위해 두 패스 필요).
        var nonNull = entries.Where(e => e.Value is string s && s.Length > 0).ToList();
        int count = nonNull.Count;

        using var psBody = new MemoryStream();
        var pw = new BinaryWriter(psBody);

        // PropertySet 헤더: size(4) + count(4) + idAndOffset[count] (각 8 byte: pid + offset)
        int propsArrayStart = 8 + count * 8;
        var propBytes = new List<byte[]>();
        var offsets   = new List<int>();

        foreach (var (pid, value) in nonNull)
        {
            var s = (string)value!;
            offsets.Add(propsArrayStart + propBytes.Sum(b => b.Length));
            using var pms = new MemoryStream();
            var pbw = new BinaryWriter(pms);
            pbw.Write((uint)0x001E);   // VT_LPSTR
            var ansi = System.Text.Encoding.UTF8.GetBytes(s + "\0");
            pbw.Write((uint)ansi.Length);
            pbw.Write(ansi);
            // 4-byte align
            while (pms.Length % 4 != 0) pbw.Write((byte)0);
            propBytes.Add(pms.ToArray());
        }

        int psSize = propsArrayStart + propBytes.Sum(b => b.Length);
        pw.Write((uint)psSize);
        pw.Write((uint)count);
        for (int i = 0; i < count; i++)
        {
            pw.Write((uint)nonNull[i].Pid);
            pw.Write((uint)offsets[i]);
        }
        foreach (var b in propBytes) pw.Write(b);

        // Outer PropertySetStream
        using var ms = new MemoryStream();
        var bw = new BinaryWriter(ms);
        bw.Write((ushort)0xFFFE);                  // ByteOrder (LE marker)
        bw.Write((ushort)0x0000);                  // Version 0
        bw.Write((uint)0x00020105);                // SystemIdentifier (Windows 2.0+)
        bw.Write(Guid.Empty.ToByteArray());        // CLSID (16 byte, 0 OK)
        bw.Write((uint)1);                         // cSections = 1
        bw.Write(fmtId.ToByteArray());             // FMTID
        bw.Write((uint)(28 + 20));                 // Offset to PropertySet
        bw.Write(psBody.ToArray());
        return ms.ToArray();
    }

    // ─────────────────────────────── BuildContext ──────────────────────────────

    /// <summary>1패스 수집 결과 — 텍스트 byte, 단락/Run FC 범위, sprm 바이트, 폰트 테이블.</summary>
    private sealed class BuildContext
    {
        public byte[] TextBytes = Array.Empty<byte>();
        public uint   CcpText;
        public uint   FcMin = FibPadSize;
        public uint   FcMac;
        public readonly List<ParaInfo> Paragraphs = new();
        public readonly List<string>   Fonts      = new();   // index 0 = default ("Times New Roman")

        // Layout (BuildTable 가 채움)
        public uint FcSttbfFfn,    LcbSttbfFfn;
        public uint FcPlcfBteChpx, LcbPlcfBteChpx;
        public uint FcPlcfBtePapx, LcbPlcfBtePapx;
        public uint FcClx,         LcbClx;
        // Phase F1-W2c — bookmarks
        public uint FcSttbfBkmk,   LcbSttbfBkmk;
        public uint FcPlcfBkf,     LcbPlcfBkf;
        public uint FcPlcfBkl,     LcbPlcfBkl;
        // 한글/Word strict parser 호환용 skeleton 구조 — STSH(스타일시트) + PlcfSed(섹션테이블).
        public uint FcStshf,       LcbStshf;
        public uint FcPlcfSed,     LcbPlcfSed;
        /// <summary>SEPX 의 WordDocument stream 내 위치 — BuildWordDocument 가 append 후 채움.</summary>
        public uint SepxFc;
        /// <summary>첫 번째 섹션의 페이지 설정 — SEPX 에 page width/height/margin sprm 으로 기록.</summary>
        public PageSettings Page = new();

        // BuildWordDocument 가 채우는 BTE 목록 — 절대 페이지 번호로 치환된 상태로 BuildTable 에 전달.
        public List<BteEntry> PapxBtes = new();
        public List<BteEntry> ChpxBtes = new();

        // Phase F1-W2c — 책갈피 (name, startCp, endCp). startCp / endCp 는 main text CP 단위.
        //   Run.BookmarkStart 이 발견되면 stack 에 push, BookmarkEnd 시 pop 해서 완성된 항목으로 등록.
        public readonly List<BookmarkRecord> Bookmarks = new();

        public void Collect(PolyDonkyument doc)
        {
            Fonts.Add("Times New Roman");
            if (doc.Sections.Count > 0) Page = doc.Sections[0].Page;

            var sb = new StringBuilder();
            foreach (var sec in doc.Sections)
            foreach (var blk in sec.Blocks)
                CollectBlock(blk, sb);

            FinalizeBody(sb);
        }

        /// <summary>임의의 Block 을 처리 — Paragraph 는 본문으로, 나머지는 Phase F1-W2d 단계에서는
        /// 가시 placeholder 단락으로 변환해 본문에 emit.</summary>
        private void CollectBlock(Block blk, StringBuilder sb)
        {
            switch (blk)
            {
                case Paragraph p:
                    CollectParagraph(p, sb);
                    break;
                case ImageBlock img:
                    CollectParagraph(MakePlaceholderPara(BuildImagePlaceholder(img)), sb);
                    break;
                case ShapeObject shape:
                    CollectParagraph(MakePlaceholderPara($"[Shape: {shape.Kind}, {shape.WidthMm:0.#}×{shape.HeightMm:0.#}mm]"), sb);
                    break;
                case TextBoxObject tb:
                    {
                        var label = $"[TextBox {tb.WidthMm:0.#}×{tb.HeightMm:0.#}mm]";
                        CollectParagraph(MakePlaceholderPara(label), sb);
                        // 글상자 안 단락도 본문으로 함께 emit — 텍스트 정보가 살아남도록.
                        foreach (var inner in tb.Content)
                            CollectBlock(inner, sb);
                        CollectParagraph(MakePlaceholderPara("[/TextBox]"), sb);
                    }
                    break;
                case Table table:
                    {
                        CollectParagraph(MakePlaceholderPara($"[Table {table.Rows.Count}×{(table.Rows.FirstOrDefault()?.Cells.Count ?? 0)}]"), sb);
                        foreach (var row in table.Rows)
                            foreach (var cell in row.Cells)
                                foreach (var inner in cell.Blocks) CollectBlock(inner, sb);
                        CollectParagraph(MakePlaceholderPara("[/Table]"), sb);
                    }
                    break;
                case ContainerBlock c:
                    foreach (var inner in c.Children) CollectBlock(inner, sb);
                    break;
                case ThematicBreakBlock:
                    CollectParagraph(MakePlaceholderPara("────────"), sb);
                    break;
                case TocBlock toc:
                    CollectParagraph(MakePlaceholderPara($"[TOC: levels {toc.MinLevel}..{toc.MaxLevel}, {toc.Entries.Count} entries]"), sb);
                    // 본문에 entry 들의 텍스트도 emit — 사용자가 목차 항목을 잃지 않도록.
                    foreach (var e in toc.Entries)
                        CollectParagraph(MakePlaceholderPara($"  {new string(' ', Math.Max(0, e.Level - 1) * 2)}{e.Text}"), sb);
                    break;
                case OpaqueBlock opaque:
                    CollectParagraph(MakePlaceholderPara($"[Opaque: {opaque.DisplayLabel}]"), sb);
                    break;
            }
        }

        private static Paragraph MakePlaceholderPara(string text)
        {
            var p = new Paragraph();
            // 이탤릭 + 회색 으로 placeholder 임을 시각적으로 표시.
            p.AddText(text, new RunStyle { Italic = true, Foreground = new Color(0x80, 0x80, 0x80) });
            return p;
        }

        private static string BuildImagePlaceholder(ImageBlock img)
        {
            string type = string.IsNullOrEmpty(img.MediaType) ? "image" : img.MediaType!;
            string size = img.WidthMm > 0 && img.HeightMm > 0
                ? $", {img.WidthMm:0.#}×{img.HeightMm:0.#}mm" : string.Empty;
            int bytes = img.Data?.Length ?? 0;
            string byteStr = bytes > 1024 ? $", {bytes / 1024} KB" : $", {bytes} B";
            return $"[Image: {type}{size}{byteStr}]";
        }

        /// <summary>본문 마지막에 단락 마크 보장 + bytes/CCP 확정. Collect() 진입 마지막 단계.</summary>
        private void FinalizeBody(StringBuilder sb)
        {
            if (sb.Length == 0 || sb[^1] != '\r')
            {
                int fcStart = sb.Length;
                sb.Append('\r');
                Paragraphs.Add(new ParaInfo
                {
                    CpStart  = fcStart,
                    CpEnd    = fcStart + 1,
                    PapxBody = BuildPapxBody(new ParagraphStyle()),
                    Runs     = new List<RunInfo>(),
                });
            }

            string text = sb.ToString();
            TextBytes   = Encoding.Unicode.GetBytes(text);
            CcpText     = (uint)text.Length;
            FcMac       = FcMin + (uint)TextBytes.Length;
        }

        private void CollectParagraph(Paragraph para, StringBuilder sb)
        {
            int paraCpStart = sb.Length;
            var runs = new List<RunInfo>();
            var style = para.Style ?? new ParagraphStyle();

            // 책갈피 시작/끝 marker 누적 — Run 의 양 끝에 가상 marker 적용.
            //   Run.BookmarkStart 가 있으면 현재 CP 를 시작점으로 stack push,
            //   BookmarkEnd 가 있으면 stack 에서 같은 이름 pop 후 항목 등록.
            var openBookmarks = new Dictionary<string, int>(StringComparer.Ordinal);

            foreach (var run in para.Runs)
            {
                int curCp = sb.Length;

                if (!string.IsNullOrEmpty(run.BookmarkStart))
                    openBookmarks[run.BookmarkStart!] = curCp;

                if (!string.IsNullOrEmpty(run.Text))
                {
                    string sanitized = SanitizeForBody(run.Text);
                    if (sanitized.Length > 0)
                    {
                        int runCpStart = sb.Length;
                        sb.Append(sanitized);
                        int runCpEnd = sb.Length;
                        runs.Add(new RunInfo
                        {
                            CpStart  = runCpStart,
                            CpEnd    = runCpEnd,
                            ChpxBody = BuildChpxBody(run.Style ?? new RunStyle(), this),
                        });
                    }
                }

                if (!string.IsNullOrEmpty(run.BookmarkEnd) &&
                    openBookmarks.Remove(run.BookmarkEnd!, out int startCp))
                {
                    Bookmarks.Add(new BookmarkRecord(run.BookmarkEnd!, startCp, sb.Length));
                }
            }

            // 단락 끝까지 열려 있는 책갈피는 단락 마크 직전(\r 직전) 까지 확장.
            foreach (var kv in openBookmarks)
                Bookmarks.Add(new BookmarkRecord(kv.Key, kv.Value, sb.Length));

            // 단락 마크 \r 자체도 본문 character — CHPX 가 적용되어야 단락 끝의 폰트도 보존됨.
            int markCp = sb.Length;
            sb.Append('\r');

            Paragraphs.Add(new ParaInfo
            {
                CpStart  = paraCpStart,
                CpEnd    = sb.Length,         // includes \r
                PapxBody = BuildPapxBody(style),
                Runs     = runs,
            });
            // 단락 마크에 대해 별도 CHPX 가 필요한 경우는 향후 (Phase 3h-3 paragraph revision 등) — 현재 생략.
        }

        public int RegisterFont(string family)
        {
            if (string.IsNullOrEmpty(family)) return 0;
            int i = Fonts.IndexOf(family);
            if (i >= 0) return i;
            Fonts.Add(family);
            return Fonts.Count - 1;
        }
    }

    /// <summary>한 단락의 메타 — CP 범위 (텍스트 시작~\r 포함 끝), PAPX grpprl 본문, Run 목록.</summary>
    private sealed class ParaInfo
    {
        public int    CpStart;
        public int    CpEnd;
        public byte[] PapxBody = Array.Empty<byte>();    // istd 뒤에 오는 sprm 들만 — istd 자체는 별도
        public List<RunInfo> Runs = new();
        public ushort Istd = 0;                          // 0 = Normal style
    }

    private sealed class RunInfo
    {
        public int    CpStart;
        public int    CpEnd;
        public byte[] ChpxBody = Array.Empty<byte>();    // CHPX sprm 본문 (cb 헤더 없음)
    }

    // ─────────────────────────────── sprm 작성 ─────────────────────────────────

    /// <summary>ParagraphStyle → PAPX grpprl sprm 본문 (istd 제외).
    /// Reader 가 인식하는 sprms: 0x2461 정렬 / 0x845D 좌측 / 0x845E 우측 /
    /// 0x8460 첫줄 들여쓰기 / 0xA413 앞 간격 / 0xA415 뒤 간격 / 0x6412 줄 간격.</summary>
    private static byte[] BuildPapxBody(ParagraphStyle ps)
    {
        var ms = new MemoryStream(32);
        var bw = new BinaryWriter(ms);

        // sprmPJc80 (0x2461, spra=1, 1-byte): 정렬
        byte align = ps.Alignment switch
        {
            Alignment.Center  => (byte)1,
            Alignment.Right   => (byte)2,
            Alignment.Justify => (byte)3,
            _                 => (byte)0,
        };
        if (align != 0) WriteSprm1(bw, 0x2461, align);

        // 들여쓰기·간격 (mm → twips, signed/unsigned 2-byte)
        int liT  = MmToTwips(ps.IndentLeftMm);
        int riT  = MmToTwips(ps.IndentRightMm);
        int fiT  = MmToTwips(ps.IndentFirstLineMm);
        int sbT  = (int)(ps.SpaceBeforePt * 20);
        int saT  = (int)(ps.SpaceAfterPt  * 20);

        if (liT != 0) WriteSprm2Signed(bw, 0x845D, liT);
        if (riT != 0) WriteSprm2Signed(bw, 0x845E, riT);
        if (fiT != 0) WriteSprm2Signed(bw, 0x8460, fiT);
        if (sbT > 0)  WriteSprm2Unsigned(bw, 0xA413, sbT);
        if (saT > 0)  WriteSprm2Unsigned(bw, 0xA415, saT);

        // sprmPDyaLine (0x6412, spra=4, 4-byte LSPD: dyaLine + fMultLinespace)
        // LineHeightFactor != 1.2 (default) 일 때만 multi-mode 로 기록.
        if (ps.LineHeightFactor > 0 && Math.Abs(ps.LineHeightFactor - 1.2) > 0.01)
        {
            short  dyaLine = (short)Math.Round(ps.LineHeightFactor * 240.0);
            WriteSprm4LSPD(bw, 0x6412, dyaLine, fMult: 1);
        }

        return ms.ToArray();
    }

    /// <summary>RunStyle → CHPX grpprl sprm 본문.
    /// Reader 인식: 0x0835 굵게 / 0x0836 이탤릭 / 0x0837 취소선 / 0x2A3E 밑줄 /
    /// 0x4A43 글자 크기 / 0x6870 RGB 전경색 / 0x2A0C 하이라이트(팔레트) /
    /// 0x4A4F/0x4A50/0x4A51 폰트 인덱스 (ascii/far-east/other).</summary>
    private static byte[] BuildChpxBody(RunStyle rs, BuildContext ctx)
    {
        var ms = new MemoryStream(32);
        var bw = new BinaryWriter(ms);

        if (rs.Bold)          WriteSprm1(bw, 0x0835, 1);
        if (rs.Italic)        WriteSprm1(bw, 0x0836, 1);
        if (rs.Strikethrough) WriteSprm1(bw, 0x0837, 1);
        if (rs.Underline)     WriteSprm1(bw, 0x2A3E, 1);

        // sprmCHps (0x4A43, spra=2, 2-byte): half-points
        if (rs.FontSizePt > 0)
        {
            ushort halfPt = (ushort)Math.Round(rs.FontSizePt * 2);
            if (halfPt > 0 && halfPt < 1000) WriteSprm2Unsigned(bw, 0x4A43, halfPt);
        }

        // sprmCCv (0x6870, spra=3, 4-byte): RGB+1 (R, G, B, 0). cvAuto 면 0xFF.
        if (rs.Foreground.HasValue)
        {
            var c = rs.Foreground.Value;
            byte[] op = { c.R, c.G, c.B, 0 };
            WriteSprmHeader(bw, 0x6870);
            bw.Write(op);
        }

        // 폰트 인덱스 — ascii / far-east / other 세 슬롯 모두 동일하게 설정 (Word 의 표준 동작)
        if (!string.IsNullOrEmpty(rs.FontFamily))
        {
            ushort ftc = (ushort)ctx.RegisterFont(rs.FontFamily!);
            if (ftc > 0)   // 0 = default, 굳이 sprm 으로 명시 안 함
            {
                WriteSprm2Unsigned(bw, 0x4A4F, ftc);
                WriteSprm2Unsigned(bw, 0x4A50, ftc);
                WriteSprm2Unsigned(bw, 0x4A51, ftc);
            }
        }

        return ms.ToArray();
    }

    // sprm header + 1-byte operand (spra=0 → 1-byte). 0x2461 / 0x0835 / 0x0836 / 0x0837 / 0x2A3E.
    private static void WriteSprm1(BinaryWriter bw, ushort sprm, byte operand)
    {
        WriteSprmHeader(bw, sprm);
        bw.Write(operand);
    }

    // sprm header + 2-byte unsigned (spra=4 / 5). 0xA413 / 0xA415 / 0x4A43 / 0x4A4F-51.
    private static void WriteSprm2Unsigned(BinaryWriter bw, ushort sprm, int operand)
    {
        WriteSprmHeader(bw, sprm);
        bw.Write((ushort)operand);
    }

    // sprm header + 2-byte signed (spra=2). 0x845D / 0x845E / 0x8460.
    private static void WriteSprm2Signed(BinaryWriter bw, ushort sprm, int operand)
    {
        WriteSprmHeader(bw, sprm);
        bw.Write((short)Math.Clamp(operand, short.MinValue, short.MaxValue));
    }

    // sprm header + LSPD (2-byte dyaLine + 2-byte fMultLinespace).
    private static void WriteSprm4LSPD(BinaryWriter bw, ushort sprm, short dyaLine, ushort fMult)
    {
        WriteSprmHeader(bw, sprm);
        bw.Write(dyaLine);
        bw.Write(fMult);
    }

    private static void WriteSprmHeader(BinaryWriter bw, ushort sprm) => bw.Write(sprm);

    // ───────────────────────── WordDocument 조립 ───────────────────────────────

    /// <summary>WordDocument stream 조립 — FIB(0x400 pad) + 텍스트 + (512-align) + PAPX FKP* + CHPX FKP*.</summary>
    private static byte[] BuildWordDocument(BuildContext ctx)
    {
        // 1. 텍스트까지 layout
        var ms = new MemoryStream();
        ms.SetLength(FibPadSize);                 // FIB 영역 zero-pad
        ms.Position = FibPadSize;
        ms.Write(ctx.TextBytes, 0, ctx.TextBytes.Length);

        // 2. 512-byte 페이지 경계로 정렬
        PadTo512(ms);

        // 3. PAPX FKP 페이지들 생성 — Body 는 sprm grpprl 본문만 (istd / cb 헤더는 PackFkpPages 가 감쌈).
        var papxItems = ctx.Paragraphs.Select(p => new FkpItem
        {
            FcStart = CpToFc(p.CpStart, ctx),
            FcEnd   = CpToFc(p.CpEnd,   ctx),
            Body    = p.PapxBody,
        }).ToList();

        var (papxPages, papxBtes) = PackFkpPages(papxItems, isPapx: true);
        foreach (var page in papxPages)
        {
            int pageIndex = (int)(ms.Position / FkpPageSize);
            // BTE 의 PnFkp 를 절대 페이지 번호로 갱신 — Pack 단계에서 임시 인덱스만 채워 둠.
            for (int i = 0; i < papxBtes.Count; i++)
                if (papxBtes[i].PnFkp == -(papxPages.IndexOf(page) + 1))   // sentinel
                    papxBtes[i] = papxBtes[i] with { PnFkp = pageIndex };
            ms.Write(page, 0, page.Length);
        }

        // 4. CHPX FKP 페이지들 — Run 단위
        var chpxItems = new List<FkpItem>();
        foreach (var p in ctx.Paragraphs)
        {
            foreach (var r in p.Runs)
            {
                chpxItems.Add(new FkpItem
                {
                    FcStart = CpToFc(r.CpStart, ctx),
                    FcEnd   = CpToFc(r.CpEnd,   ctx),
                    Body    = r.ChpxBody,
                });
            }
            // 단락 마크 (\r) 의 CHPX 도 빈 본문으로 등록 — Reader 의 FC 범위 누락 방지.
            chpxItems.Add(new FkpItem
            {
                FcStart = CpToFc(p.CpEnd - 1, ctx),
                FcEnd   = CpToFc(p.CpEnd,     ctx),
                Body    = Array.Empty<byte>(),
            });
        }

        var (chpxPages, chpxBtes) = PackFkpPages(chpxItems, isPapx: false);
        foreach (var page in chpxPages)
        {
            int pageIndex = (int)(ms.Position / FkpPageSize);
            for (int i = 0; i < chpxBtes.Count; i++)
                if (chpxBtes[i].PnFkp == -(chpxPages.IndexOf(page) + 1))
                    chpxBtes[i] = chpxBtes[i] with { PnFkp = pageIndex };
            ms.Write(page, 0, page.Length);
        }

        // 5. BTE 데이터를 context 에 임시 저장 — BuildTable 에서 직렬화.
        //    SEPX 는 [MS-DOC] §2.9.241 상 Table stream 에 위치하므로 BuildTable 이 작성.
        ctx.PapxBtes = papxBtes;
        ctx.ChpxBtes = chpxBtes;

        return ms.ToArray();
    }

    private static int CpToFc(int cp, BuildContext ctx) => (int)ctx.FcMin + cp * 2;

    private static void PadTo512(MemoryStream ms)
    {
        int rem = (int)(ms.Length % FkpPageSize);
        if (rem == 0) return;
        int pad = FkpPageSize - rem;
        ms.SetLength(ms.Length + pad);
        ms.Position = ms.Length;
    }

    /// <summary>PAPX record: cb=0 form. grpprl = istd(2 byte) + sprms (even-padded).
    /// 입력 <paramref name="papxBody"/> 는 sprm grpprl 본문만 (istd 제외).</summary>
    private static byte[] BuildPapxRecord(byte[] papxBody, ushort istd)
    {
        int sprmsLen = papxBody.Length;
        if ((sprmsLen & 1) == 1) sprmsLen++;
        int grpprlSize = 2 + sprmsLen;              // istd(2) + padded sprms
        if ((grpprlSize & 1) == 1) grpprlSize++;

        // cb=0 form: 1 byte cb=0 + 1 byte cb2=grpprlSize/2 + grpprl. record size = 2 + grpprlSize.
        var rec = new byte[2 + grpprlSize];
        rec[0] = 0;
        rec[1] = (byte)(grpprlSize / 2);
        rec[2] = (byte)(istd & 0xFF);
        rec[3] = (byte)((istd >> 8) & 0xFF);
        Buffer.BlockCopy(papxBody, 0, rec, 4, papxBody.Length);
        // 잉여 trailing bytes 는 0 — Reader 의 WalkSprms 가 sprm 0x0000 (spra=0 1-byte operand) 로
        //   안전하게 통과시키고 ApplyParagraphSprm switch 가 default false 반환.
        return rec;
    }

    /// <summary>CHPX record: cb(1 byte) + sprms[cb]. bOffset 이 /2 로 인코딩되므로 record offset 은 even.
    /// → sprms 길이를 even 으로 padding (cb 도 even 이어야 record 전체가 even).</summary>
    private static byte[] BuildChpxRecord(byte[] body)
    {
        int len = body.Length;
        if ((len & 1) == 1) len++;          // even sprms
        // cb = len. record size = 1 + len. odd → pad 1 byte for even-alignment.
        int recSize = 1 + len;
        if ((recSize & 1) == 1) recSize++;
        var rec = new byte[recSize];
        rec[0] = (byte)len;
        Buffer.BlockCopy(body, 0, rec, 1, body.Length);
        return rec;
    }

    // ── FKP 페이지 패킹 ────────────────────────────────────────────────────────

    private record struct FkpItem
    {
        public int    FcStart;
        public int    FcEnd;
        public byte[] Body;
    }

    /// <summary>한 페이지에 PAPX/CHPX 항목들을 가능한 만큼 패킹. 페이지를 넘으면 새 페이지로.
    /// 반환되는 BTE 의 <c>PnFkp</c> 는 sentinel(-(pageIndex+1)) — <see cref="BuildWordDocument"/>
    /// 에서 실제 absolute page number 로 치환.</summary>
    private static (List<byte[]> Pages, List<BteEntry> Btes) PackFkpPages(List<FkpItem> items, bool isPapx)
    {
        var pages = new List<byte[]>();
        var btes  = new List<BteEntry>();
        if (items.Count == 0) return (pages, btes);

        int idx = 0;
        while (idx < items.Count)
        {
            // 한 페이지에 들어갈 항목 수 결정 — back-to-front 로 record 누적, front 에서 rgfc+rgbx 누적.
            int rgbxEntrySize = isPapx ? 13 : 1;
            int recordsBytes  = 0;
            int count         = 0;
            for (int i = idx; i < items.Count; i++)
            {
                var rec = isPapx ? BuildPapxRecord(items[i].Body, istd: 0) : BuildChpxRecord(items[i].Body);
                int neededFront = 4 * (count + 2) + rgbxEntrySize * (count + 1);
                int neededBack  = recordsBytes + rec.Length;
                int cparaByte   = 1;
                if (neededFront + neededBack + cparaByte > FkpPageSize) break;
                recordsBytes += rec.Length;
                count++;
            }
            if (count == 0)
                throw new InvalidOperationException(
                    "단일 PAPX/CHPX record 가 FKP 페이지(512 byte) 보다 큽니다 — 단락/Run 의 sprm 본문이 비정상적으로 큼.");

            // 페이지 채우기
            var page = new byte[FkpPageSize];
            int writePos = FkpPageSize - 1;                // crun byte
            page[writePos] = (byte)count;

            // record 들 — 페이지 끝부터 역방향으로 배치, 각 record 의 offset 을 rgbx 에 기록
            // bOffset 은 /2 로 인코딩됨 (PAPX 와 CHPX 모두) → record offset 은 even.
            int recCursor = FkpPageSize - 1;               // crun 위
            var offsets   = new int[count];
            for (int i = count - 1; i >= 0; i--)
            {
                var rec = isPapx
                    ? BuildPapxRecord(items[idx + i].Body, istd: 0)
                    : BuildChpxRecord(items[idx + i].Body);
                recCursor -= rec.Length;
                if ((recCursor & 1) == 1) recCursor--;     // even alignment
                Buffer.BlockCopy(rec, 0, page, recCursor, rec.Length);
                offsets[i] = recCursor;
            }

            // rgfc[count+1] @ offset 0 — FC array
            for (int i = 0; i <= count; i++)
            {
                int fc = i == count ? items[idx + count - 1].FcEnd : items[idx + i].FcStart;
                WriteInt32(page, i * 4, fc);
            }

            // rgbx[count] — 각 항목별 (PAPX: bOffset(1)+PHE(12)=13 byte, CHPX: bOffset(1)=1 byte)
            int rgbxBase = 4 * (count + 1);
            for (int i = 0; i < count; i++)
            {
                page[rgbxBase + i * rgbxEntrySize] = (byte)(offsets[i] / 2);
                // PHE 12 byte 는 0 으로 둠 (paragraph height — Reader 가 안 읽음)
            }

            int pageSentinel = -(pages.Count + 1);
            btes.Add(new BteEntry(items[idx].FcStart, items[idx + count - 1].FcEnd, pageSentinel));
            pages.Add(page);
            idx += count;
        }

        return (pages, btes);
    }

    // ───────────────────────── 1Table 조립 ────────────────────────────────────

    private static byte[] BuildTable(BuildContext ctx)
    {
        var ms = new MemoryStream();

        // 1. CLX (단일 piece, Unicode)
        ctx.FcClx = (uint)ms.Position;
        byte[] clx = BuildClx(ctx.FcMin, ctx.CcpText);
        ms.Write(clx, 0, clx.Length);
        ctx.LcbClx = (uint)clx.Length;

        // 2. SttbfFfn
        ctx.FcSttbfFfn = (uint)ms.Position;
        byte[] sttbFfn = BuildSttbfFfn(ctx.Fonts);
        ms.Write(sttbFfn, 0, sttbFfn.Length);
        ctx.LcbSttbfFfn = (uint)sttbFfn.Length;

        // 3. PlcfBteChpx — aFC[n+1] + aPnFkp[n]
        ctx.FcPlcfBteChpx = (uint)ms.Position;
        byte[] plcChpx = BuildPlcBte(ctx.ChpxBtes);
        ms.Write(plcChpx, 0, plcChpx.Length);
        ctx.LcbPlcfBteChpx = (uint)plcChpx.Length;

        // 4. PlcfBtePapx
        ctx.FcPlcfBtePapx = (uint)ms.Position;
        byte[] plcPapx = BuildPlcBte(ctx.PapxBtes);
        ms.Write(plcPapx, 0, plcPapx.Length);
        ctx.LcbPlcfBtePapx = (uint)plcPapx.Length;

        // 5. Phase F1-W2c — bookmarks (SttbfBkmk + PlcfBkf + PlcfBkl)
        if (ctx.Bookmarks.Count > 0)
        {
            // (a) SttbfBkmk — bookmark names (extended SttbExtend)
            ctx.FcSttbfBkmk = (uint)ms.Position;
            byte[] sttbBkmk = BuildSttbExtend(ctx.Bookmarks.Select(b => b.Name).ToList());
            ms.Write(sttbBkmk, 0, sttbBkmk.Length);
            ctx.LcbSttbfBkmk = (uint)sttbBkmk.Length;

            // (b) PlcfBkf — aCP[n+1] + BKF[n] (4 byte each: ibkl(2) + flags(2))
            ctx.FcPlcfBkf = (uint)ms.Position;
            byte[] plcBkf = BuildPlcfBkf(ctx.Bookmarks);
            ms.Write(plcBkf, 0, plcBkf.Length);
            ctx.LcbPlcfBkf = (uint)plcBkf.Length;

            // (c) PlcfBkl — aCP[n+1] only (lcb = 4(n+1)).
            ctx.FcPlcfBkl = (uint)ms.Position;
            byte[] plcBkl = BuildPlcfBkl(ctx.Bookmarks);
            ms.Write(plcBkl, 0, plcBkl.Length);
            ctx.LcbPlcfBkl = (uint)plcBkl.Length;
        }

        // 6. STSH (스타일시트) — 한글/Word strict parser 가 istd(=0 Normal) 참조를 해석하려면 필수.
        ctx.FcStshf = (uint)ms.Position;
        byte[] stsh = BuildStsh();
        ms.Write(stsh, 0, stsh.Length);
        ctx.LcbStshf = (uint)stsh.Length;

        // 7. SEPX (Section Properties) — [MS-DOC] §2.9.241 상 Table stream 에 위치.
        //    PlcfSed 의 Sed.fcSepx 가 이 위치를 가리킨다.
        ctx.SepxFc = (uint)ms.Position;
        byte[] sepx = BuildSepx(ctx.Page);
        ms.Write(sepx, 0, sepx.Length);

        // 8. PlcfSed (섹션 테이블) — 1개 섹션. Sed.fcSepx → 위 SEPX 의 Table stream 위치.
        ctx.FcPlcfSed = (uint)ms.Position;
        byte[] plcfSed = BuildPlcfSed(ctx.CcpText, ctx.SepxFc);
        ms.Write(plcfSed, 0, plcfSed.Length);
        ctx.LcbPlcfSed = (uint)plcfSed.Length;

        return ms.ToArray();
    }

    // ── 한글/Word strict parser 호환 skeleton 구조 ──────────────────────────────

    /// <summary>SEPX (Section Properties) — [MS-DOC] §2.8.26. cb(2) + grpprl(section sprms).
    /// 첫 섹션의 페이지 크기·여백을 sprmSDxaPage/SDyaPage/SDxaLeft/SDxaRight/SDyaTop/SDyaBottom 으로 기록.</summary>
    private static byte[] BuildSepx(PageSettings page)
    {
        using var body = new MemoryStream();
        var bw = new BinaryWriter(body);

        // sprmSBkc (0x3009, spra=0, 1-byte): break code = 2 (new page) — 섹션 시작 방식
        WriteSprm1(bw, 0x3009, 2);
        // 페이지 크기 (twips). EffectiveWidth/Height 로 가로/세로 반영.
        WriteSprm2Unsigned(bw, 0xB016, MmToTwips(page.EffectiveWidthMm));   // sprmSDxaPage
        WriteSprm2Unsigned(bw, 0xB017, MmToTwips(page.EffectiveHeightMm));  // sprmSDyaPage
        WriteSprm2Unsigned(bw, 0xB02C, MmToTwips(page.MarginLeftMm));       // sprmSDxaLeft
        WriteSprm2Unsigned(bw, 0xB02D, MmToTwips(page.MarginRightMm));      // sprmSDxaRight
        WriteSprm2Signed  (bw, 0xB02E, MmToTwips(page.MarginTopMm));        // sprmSDyaTop
        WriteSprm2Signed  (bw, 0xB02F, MmToTwips(page.MarginBottomMm));     // sprmSDyaBottom

        byte[] grpprl = body.ToArray();
        using var ms = new MemoryStream();
        var bw2 = new BinaryWriter(ms);
        bw2.Write((ushort)grpprl.Length);   // cb
        bw2.Write(grpprl);
        return ms.ToArray();
    }

    /// <summary>PlcfSed — aCP[n+1] (섹션 경계 CP) + Sed[n] (각 12 byte).
    /// 1개 섹션: aCP=[0, ccpText], Sed[0] = fn(2)=0 + fcSepx(4) + fnMpr(2)=0 + fcMpr(4)=0xFFFFFFFF.</summary>
    private static byte[] BuildPlcfSed(uint ccpText, uint sepxFc)
    {
        using var ms = new MemoryStream();
        var bw = new BinaryWriter(ms);
        // aCP[2]
        bw.Write((uint)0);
        bw.Write(ccpText);
        // Sed[0] (12 byte)
        bw.Write((ushort)0);            // fn
        bw.Write(sepxFc);               // fcSepx → WordDocument stream
        bw.Write((ushort)0);            // fnMpr
        bw.Write((uint)0xFFFFFFFF);     // fcMpr (none)
        return ms.ToArray();
    }

    /// <summary>STSH (Style Sheet) — [MS-DOC] §2.9.271. LPStshi + rgLPStd.
    /// 최소: Stshif(18 byte) + Normal 스타일 1개 (sti=0, stk=1 paragraph). istd=0 참조 해석용.</summary>
    private static byte[] BuildStsh()
    {
        using var ms = new MemoryStream();
        var bw = new BinaryWriter(ms);

        // ── LPStshi: cbStshi(2) + Stshif(18) ──
        using var stshi = new MemoryStream();
        var sw = new BinaryWriter(stshi);
        sw.Write((ushort)1);        // cstd = 1 (Normal only)
        sw.Write((ushort)0x000A);   // cbSTDBaseInFile = 10 (Word 97 STD base)
        sw.Write((ushort)0x0000);   // fStdStylenamesWritten + grfReserved
        sw.Write((ushort)0);        // stiMaxWhenSaved
        sw.Write((ushort)1);        // istdMaxFixedWhenSaved
        sw.Write((ushort)0);        // nVerBuiltInNamesWhenSaved
        sw.Write((ushort)0);        // rgftcStandardChpStsh[0] (ftcAsci)
        sw.Write((ushort)0);        // [1] ftcFE
        sw.Write((ushort)0);        // [2] ftcOther
        byte[] stshiBytes = stshi.ToArray();   // 18 byte
        bw.Write((ushort)stshiBytes.Length);
        bw.Write(stshiBytes);

        // ── rgLPStd[0] = Normal: cbStd(2) + STD ──
        using var std = new MemoryStream();
        var dw = new BinaryWriter(std);
        // StdfBase (10 byte)
        dw.Write((ushort)0x0000);   // sti=0 + flags
        dw.Write((ushort)0xFFF1);   // stk=1(para) | istdBase=0xFFF<<4
        dw.Write((ushort)0x0002);   // cupx=2 | istdNext=0
        dw.Write((ushort)0);        // bchUpe
        dw.Write((ushort)0);        // grfstd
        // xstzName "Normal": cchData(2) + UTF-16 + null(2)
        const string name = "Normal";
        dw.Write((ushort)name.Length);
        dw.Write(System.Text.Encoding.Unicode.GetBytes(name));
        dw.Write((ushort)0);
        // grLPUpxSw (para style, cupx=2): LPUpxPapx + LPUpxChpx
        //   LPUpxPapx: cbUPX(2)=2 + UpxPapx(istd 2 byte) → 그 뒤 2-byte 정렬 (이미 even)
        dw.Write((ushort)2);        // cbUPX (PAPX) = 2
        dw.Write((ushort)0);        // istd = 0 (UpxPapx)
        //   LPUpxChpx: cbUPX(2)=0 (빈 grpprl)
        dw.Write((ushort)0);        // cbUPX (CHPX) = 0
        byte[] stdBytes = std.ToArray();
        bw.Write((ushort)stdBytes.Length);
        bw.Write(stdBytes);

        return ms.ToArray();
    }

    /// <summary>SttbExtend (extended SttbF) — 0xFFFF + cData + cbExtra=0 + entries[cchData(2) + UTF-16 chars].
    /// Reader 의 SttbfBkmk / SttbfRMark / SttbfFnm 등이 모두 이 포맷.</summary>
    private static byte[] BuildSttbExtend(IReadOnlyList<string> strings)
    {
        var ms = new MemoryStream();
        var bw = new BinaryWriter(ms);
        bw.Write((ushort)0xFFFF);
        bw.Write((ushort)strings.Count);
        bw.Write((ushort)0);                              // cbExtra
        foreach (var s in strings)
        {
            byte[] utf16 = Encoding.Unicode.GetBytes(s ?? string.Empty);
            ushort cch = (ushort)((s ?? string.Empty).Length);
            bw.Write(cch);
            bw.Write(utf16);
        }
        return ms.ToArray();
    }

    /// <summary>PlcfBkf — aCP[n+1] (시작 CP) + BKF[n] (각 4 byte: ibkl(2) + flags(2)).
    /// ibkl 은 PlcfBkl 안에서 매칭되는 끝 CP 의 인덱스. 우리는 1:1 매핑이라 ibkl[i] = i.</summary>
    private static byte[] BuildPlcfBkf(IReadOnlyList<BookmarkRecord> bookmarks)
    {
        int n = bookmarks.Count;
        var ms = new MemoryStream();
        var bw = new BinaryWriter(ms);
        // aCP[n+1] — 시작 CP 들 + sentinel (= 마지막 끝 CP 또는 0). Reader 는 마지막을 무시하므로 0 충분.
        for (int i = 0; i < n; i++) bw.Write(bookmarks[i].StartCp);
        bw.Write(0);                                      // sentinel
        // BKF[n] — 각 4 byte
        for (int i = 0; i < n; i++)
        {
            bw.Write((ushort)i);                          // ibkl
            bw.Write((ushort)0);                          // flags = 0 (col-first 등 미사용)
        }
        return ms.ToArray();
    }

    /// <summary>PlcfBkl — aCP[n+1] (끝 CP 들 + sentinel). 데이터 없음 (frdSize=0).</summary>
    private static byte[] BuildPlcfBkl(IReadOnlyList<BookmarkRecord> bookmarks)
    {
        int n = bookmarks.Count;
        var ms = new MemoryStream();
        var bw = new BinaryWriter(ms);
        for (int i = 0; i < n; i++) bw.Write(bookmarks[i].EndCp);
        bw.Write(0);                                      // sentinel
        return ms.ToArray();
    }

    /// <summary>CLX = clxtPcdt(0x02) + lcb(4) + plcfpcd. 단일 piece, Unicode.</summary>
    private static byte[] BuildClx(uint fcText, uint ccpText)
    {
        const int n = 1;
        int cpsSize = (n + 1) * 4;
        int pcdSize = n * 8;
        int plcSize = cpsSize + pcdSize;

        var ms = new MemoryStream(1 + 4 + plcSize);
        var bw = new BinaryWriter(ms);
        bw.Write((byte)0x02);
        bw.Write((uint)plcSize);
        bw.Write((uint)0);             // cp[0] = 0
        bw.Write(ccpText);             // cp[1] = ccpText

        // Pcd: A/B/C/D flags(2) + fc(4) Unicode + prm(2)
        bw.Write((ushort)0);
        bw.Write(fcText);              // bit 30 = 0 → Unicode
        bw.Write((ushort)0);
        return ms.ToArray();
    }

    /// <summary>SttbfFfn extended format. FFN 내용은 40 byte zero header + UTF-16LE 이름 + null terminator.</summary>
    private static byte[] BuildSttbfFfn(List<string> fonts)
    {
        var ms = new MemoryStream();
        var bw = new BinaryWriter(ms);
        bw.Write((ushort)0xFFFF);                // extended marker
        bw.Write((ushort)fonts.Count);           // cData
        bw.Write((ushort)0);                     // cbExtra = 0

        foreach (var name in fonts)
        {
            // FFN byte length = 40 (header) + (name + null) * 2
            byte[] nameUtf16 = Encoding.Unicode.GetBytes(name);
            int ffnBytes = 40 + nameUtf16.Length + 2;   // +2 = null terminator (UTF-16 0x0000)
            ushort cchData = (ushort)(ffnBytes / 2);

            bw.Write(cchData);
            for (int k = 0; k < 40; k++) bw.Write((byte)0);   // FFN header — Reader 가 무시
            bw.Write(nameUtf16);
            bw.Write((ushort)0);                              // null terminator
        }
        return ms.ToArray();
    }

    /// <summary>PlcBteFcN: aFC[n+1] (4 byte) + aPnFkp[n] (4 byte, lower 22 bits = page number).</summary>
    private static byte[] BuildPlcBte(List<BteEntry> btes)
    {
        if (btes.Count == 0) return Array.Empty<byte>();
        int n = btes.Count;
        var ms = new MemoryStream(4 * (n + 1) + 4 * n);
        var bw = new BinaryWriter(ms);
        for (int i = 0; i < n; i++) bw.Write(btes[i].FcStart);
        bw.Write(btes[n - 1].FcEnd);                  // closing FC
        for (int i = 0; i < n; i++) bw.Write((uint)(btes[i].PnFkp & 0x003FFFFF));
        return ms.ToArray();
    }

    // ──────────────────────────── FIB 작성 ─────────────────────────────────────

    private static void WriteFib(byte[] wd, BuildContext ctx, uint tableLength)
    {
        WriteUInt16(wd, 0x0000, Magic);
        WriteUInt16(wd, 0x0002, NFib);
        WriteUInt16(wd, 0x000A, FibFlags);
        WriteUInt16(wd, 0x000C, NFibBack);
        // [MS-DOC] §2.5.1 fcMin / fcMac — Word 97+ 에서는 본문 위치를 CLX piece table 로만 표현하고
        // 이 두 필드는 SHOULD be 0. 0이 아닌 값이면 한글/Word Doc Viewer 같은 strict parser 가 거부.
        // (CLX 의 Pcd.fc 는 여전히 ctx.FcMin = 0x400 을 가리켜 실제 텍스트 위치는 정상.)
        WriteUInt32(wd, 0x0018, 0);
        WriteUInt32(wd, 0x001C, 0);

        WriteUInt16(wd, 0x0020, Csw);
        WriteUInt16(wd, 0x003E, Cslw);
        WriteUInt32(wd, 0x004C, ctx.CcpText);

        WriteUInt16(wd, 0x0098, CbRgFcLcb);
        WritePair(wd, PairFcPlcfBteChpx, ctx.FcPlcfBteChpx, ctx.LcbPlcfBteChpx);
        WritePair(wd, PairFcPlcfBtePapx, ctx.FcPlcfBtePapx, ctx.LcbPlcfBtePapx);
        WritePair(wd, PairFcSttbfFfn,    ctx.FcSttbfFfn,    ctx.LcbSttbfFfn);
        WritePair(wd, PairFcClx,         ctx.FcClx,         ctx.LcbClx);
        WritePair(wd, PairFcSttbfBkmk,   ctx.FcSttbfBkmk,   ctx.LcbSttbfBkmk);
        WritePair(wd, PairFcPlcfBkf,     ctx.FcPlcfBkf,     ctx.LcbPlcfBkf);
        WritePair(wd, PairFcPlcfBkl,     ctx.FcPlcfBkl,     ctx.LcbPlcfBkl);
        WritePair(wd, PairFcStshf,       ctx.FcStshf,       ctx.LcbStshf);
        WritePair(wd, PairFcPlcfSed,     ctx.FcPlcfSed,     ctx.LcbPlcfSed);

        int cswNewOffset = RgFcLcbBase + CbRgFcLcb * 8;
        WriteUInt16(wd, cswNewOffset,     CswNew);
        WriteUInt16(wd, cswNewOffset + 2, NFib);

        _ = tableLength;   // 현재 FIB 에 직접 기록하는 필드는 없음 — 모든 lcb 는 위에서 처리.
    }

    // ── 헬퍼 ────────────────────────────────────────────────────────────────────

    private static string SanitizeForBody(string s)
    {
        var sb = new StringBuilder(s.Length);
        foreach (char c in s)
        {
            if (c == '\r' || c == '\n') sb.Append(' ');
            else if (c == '\t')         sb.Append('\t');
            else if (c < 0x20)          sb.Append(' ');
            else                        sb.Append(c);
        }
        return sb.ToString();
    }

    private static int MmToTwips(double mm) => (int)Math.Round(mm * 56.692);

    private static void WritePair(byte[] wd, int pairIndex, uint fc, uint lcb)
    {
        int fcOff = RgFcLcbBase + pairIndex * 8;
        WriteUInt32(wd, fcOff,     fc);
        WriteUInt32(wd, fcOff + 4, lcb);
    }

    private static void WriteUInt16(byte[] buf, int offset, ushort value)
    {
        buf[offset    ] = (byte) (value       & 0xFF);
        buf[offset + 1] = (byte)((value >> 8) & 0xFF);
    }

    private static void WriteUInt32(byte[] buf, int offset, uint value)
    {
        buf[offset    ] = (byte) (value        & 0xFF);
        buf[offset + 1] = (byte)((value >>  8) & 0xFF);
        buf[offset + 2] = (byte)((value >> 16) & 0xFF);
        buf[offset + 3] = (byte)((value >> 24) & 0xFF);
    }

    private static void WriteInt32(byte[] buf, int offset, int value) => WriteUInt32(buf, offset, unchecked((uint)value));

    private static void WriteStream(RootStorage root, string name, byte[] data)
    {
        using var s = root.CreateStream(name);
        s.Write(data, 0, data.Length);
    }

    // ── Phase F1-W2e: FidelityCapsules → OLE2 storage 복원 ──────────────────────

    /// <summary>이 writer 가 직접 출력하는 root entry — capsule 이 같은 이름을 가져도 덮어쓰지 않음.</summary>
    private static readonly HashSet<string> ReservedRootEntries = new(StringComparer.OrdinalIgnoreCase)
    {
        "WordDocument", "1Table", "0Table", "Data",
    };

    /// <summary>capsule 키 prefix → 첫 번째 path 컴포넌트가 root storage 이름.
    /// 키 형식: <c>msdoc/&lt;category&gt;/&lt;rootStorageName&gt;/&lt;inner/path&gt;</c>.
    /// 예: <c>msdoc/macros/Macros/VBA/Module1</c> → root="Macros", path="VBA/Module1".</summary>
    private static void WriteFidelityCapsules(RootStorage root, IDictionary<string, byte[]>? capsules)
    {
        if (capsules is null || capsules.Count == 0) return;

        foreach (var kv in capsules)
        {
            string key = kv.Key;
            if (string.IsNullOrEmpty(key) || !key.StartsWith("msdoc/", StringComparison.Ordinal))
                continue;   // 다른 prefix (ooxml/, hwpx/, hancom/) 는 DOC writer 가 처리 안 함

            // msdoc/ 다음의 한 segment 는 category (macros/signatures/preserved). 그 뒤가 OLE2 path.
            string afterMsdoc = key.Substring("msdoc/".Length);
            int catSlash = afterMsdoc.IndexOf('/');
            if (catSlash <= 0) continue;
            string olePath = afterMsdoc.Substring(catSlash + 1);
            if (string.IsNullOrEmpty(olePath)) continue;

            string[] parts = olePath.Split('/');
            if (parts.Length == 0 || string.IsNullOrEmpty(parts[0])) continue;

            // Root entry 가 우리가 이미 쓴 stream 과 충돌하면 skip (안전망 — Reader 는 자기 stream 을 capsule 로 안 만든다).
            if (ReservedRootEntries.Contains(parts[0])) continue;

            try
            {
                WriteCapsuleAtPath(root, parts, kv.Value);
            }
            catch
            {
                // 같은 stream 이 두 번 쓰여지면 OpenMcdf 가 예외 — 무시하고 다음 capsule 진행.
            }
        }
    }

    private static void WriteCapsuleAtPath(RootStorage root, string[] parts, byte[] data)
    {
        Storage current = root;
        for (int i = 0; i < parts.Length - 1; i++)
        {
            string seg = parts[i];
            if (string.IsNullOrEmpty(seg)) continue;
            current = current.TryOpenStorage(seg, out var existing)
                ? existing
                : current.CreateStorage(seg);
        }
        string streamName = parts[^1];
        if (string.IsNullOrEmpty(streamName)) return;
        using var s = current.CreateStream(streamName);
        s.Write(data, 0, data.Length);
    }
}

/// <summary>PlcfBteChpx / PlcfBtePapx 의 한 엔트리 — FC 범위와 FKP 페이지 번호.</summary>
internal record struct BteEntry(int FcStart, int FcEnd, int PnFkp);

/// <summary>Phase F1-W2c — 책갈피 한 항목 (이름 + main text CP 범위).</summary>
internal sealed record BookmarkRecord(string Name, int StartCp, int EndCp);
