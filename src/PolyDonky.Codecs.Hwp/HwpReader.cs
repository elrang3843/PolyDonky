using System.IO.Compression;
using System.Text;
using OpenMcdf;
using PolyDonky.Core;

namespace PolyDonky.Codecs.Hwp;

/// <summary>
/// HwpReader 전용 파일 로거. d:\Temp\PolyDonky-HwpReader.log 에 기록한다.
/// Debug.WriteLine 은 외부 CLI 프로세스에서 VS 출력 창에 표시되지 않으므로
/// 파일 로그로 대체해 진단한다.
/// </summary>
internal static class HwpLog
{
    private static readonly string LogPath = Path.Combine(
        Environment.OSVersion.Platform == PlatformID.Win32NT ? @"d:\Temp" : "/tmp",
        "PolyDonky-HwpReader.log");

    static HwpLog()
    {
        try
        {
            Directory.CreateDirectory(Path.GetDirectoryName(LogPath)!);
            File.AppendAllText(LogPath, $"\n=== HwpReader session {DateTime.Now:yyyy-MM-dd HH:mm:ss} ===\n");
        }
        catch { }
    }

    public static void Write(string message)
    {
        var line = $"[{DateTime.Now:HH:mm:ss.fff}] {message}";
        System.Diagnostics.Debug.WriteLine(line);
        try { File.AppendAllText(LogPath, line + "\n"); } catch { }
    }
}

/// <summary>
/// HWP 5.x (KS X 5700) → PolyDonkyument 리더.
///
/// OLE2 Compound File Binary 컨테이너 구조:
///   FileHeader  — HWP 서명·버전·압축 플래그
///   DocInfo     — 폰트·글자서식·단락서식·스타일 (레코드 스트림, zlib 압축 가능)
///   BodyText/SectionN — 본문 (레코드 스트림, zlib 압축 가능)
///   PrvText     — 미리보기 텍스트 (UTF-16 LE)
///
/// 레코드 헤더 (4바이트 DWORD, 공식 KS X 5700 스펙):
///   bit  9~ 0: Tag ID  (10비트)
///   bit 19~10: Level   (10비트)
///   bit 31~20: Size    (12비트)
///   Size == 0xFFF → 다음 4바이트 uint32 가 실제 크기(확장).
///
/// Tag ID 베이스 (KS X 5700 기준):
///   HWPTAG_BEGIN (=0x010) → DocInfo 태그 시작
///   BodyText 태그: PARA_HEADER=0x042, PARA_TEXT=0x043, CTRL_HEADER=0x047,
///                  LIST_HEADER=0x048, PAGE_DEF=0x049, SHAPE_COMPONENT=0x04C, …
///
/// 지원 범위:
///   텍스트/단락, 용지 설정(PAGE_DEF), 글상자(CTRL_HEADER + LIST_HEADER),
///   도형(SHAPE_COMPONENT + 서브태그), 이미지(PICTURE_COMPONENT + BinData).
/// 미지원: 암호화, 변경추적, 수식.
/// </summary>
public sealed class HwpReader : IDocumentReader
{
    public string FormatId => "hwp";

    // HWPUNIT: 1/7200 inch = 25.4/7200 mm ≈ 0.003528 mm/unit
    private const double HwpUnitToMm = 25.4 / 7200.0;

    // ── DocInfo Tag ID (HWPTAG_BEGIN = 0x010) ──────────────────────────────
    private const uint TAG_DOCUMENT_PROPERTIES = 0x010;
    private const uint TAG_BIN_DATA            = 0x012;
    private const uint TAG_FACE_NAME           = 0x013;
    private const uint TAG_BORDER_FILL         = 0x014;
    private const uint TAG_CHAR_SHAPE          = 0x015;
    private const uint TAG_PARA_SHAPE          = 0x019;
    private const uint TAG_STYLE               = 0x01A;

    // ── BodyText Tag ID (KS X 5700, 실제 값은 0x042 부터 시작) ────────────────
    private const uint TAG_PARA_HEADER         = 0x042;  // HWPTAG_PARA_HEADER
    private const uint TAG_PARA_TEXT           = 0x043;  // HWPTAG_PARA_TEXT
    private const uint TAG_PARA_CHAR_SHAPE     = 0x044;  // HWPTAG_PARA_CHAR_SHAPE
    private const uint TAG_PARA_LINE_SEG       = 0x045;  // HWPTAG_PARA_LINE_SEG
    private const uint TAG_CTRL_HEADER         = 0x047;  // HWPTAG_CTRL_HEADER
    private const uint TAG_LIST_HEADER         = 0x048;  // HWPTAG_LIST_HEADER
    private const uint TAG_PAGE_DEF            = 0x049;  // HWPTAG_PAGE_DEF
    private const uint TAG_SHAPE_COMPONENT     = 0x04C;  // HWPTAG_SHAPE_COMPONENT
    private const uint TAG_TABLE               = 0x04D;  // HWPTAG_TABLE
    private const uint TAG_LINE_COMPONENT      = 0x04E;  // HWPTAG_SHAPE_COMPONENT_LINE
    private const uint TAG_RECT_COMPONENT      = 0x04F;  // HWPTAG_SHAPE_COMPONENT_RECTANGLE
    private const uint TAG_ELLIPSE_COMPONENT   = 0x050;  // HWPTAG_SHAPE_COMPONENT_ELLIPSE
    private const uint TAG_ARC_COMPONENT       = 0x051;  // HWPTAG_SHAPE_COMPONENT_ARC
    private const uint TAG_POLYGON_COMPONENT   = 0x052;  // HWPTAG_SHAPE_COMPONENT_POLYGON
    private const uint TAG_CURVE_COMPONENT     = 0x053;  // HWPTAG_SHAPE_COMPONENT_CURVE
    private const uint TAG_OLE_COMPONENT       = 0x054;  // HWPTAG_SHAPE_COMPONENT_OLE
    private const uint TAG_PICTURE_COMPONENT   = 0x055;  // HWPTAG_SHAPE_COMPONENT_PICTURE
    private const uint TAG_CONTAINER_COMPONENT = 0x056;  // HWPTAG_SHAPE_COMPONENT_CONTAINER
    private const uint TAG_TEXTBOX_COMPONENT   = 0x059;  // HWPTAG_SHAPE_COMPONENT_TEXTBOX

    // CTRL_ID_GSO: 그리기 개체 (도형/글상자/이미지 공통). 정확히 "gso " (space at end).
    // 다른 컨트롤과 마찬가지로 빅엔디언 표기법: ('g'<<24)|('s'<<16)|('o'<<8)|' '
    private const uint CTRL_ID_GSO = ('g' << 24) | ('s' << 16) | ('o' << 8) | ' '; // 0x67736F20

    // 비-GSO 컨트롤 ID (LE uint32). 메모리상 바이트 순서가 역순이므로,
    // 'a','b','c','d' 4글자 컨트롤 이름은 LE uint32 = ('d'<<24)|('c'<<16)|('b'<<8)|'a'.
    // 다만 HWP는 빅엔디언 표기 컨벤션 ("head", "foot")으로 문서에 정의되므로
    // 그에 맞춰 ('h'<<24)|('e'<<16)|('a'<<8)|'d' 식으로 정의.
    private const uint CTRL_ID_HEADER = ('h' << 24) | ('e' << 16) | ('a' << 8) | 'd'; // 0x68656164
    private const uint CTRL_ID_FOOTER = ('f' << 24) | ('o' << 16) | ('o' << 8) | 't'; // 0x666F6F74
    private const uint CTRL_ID_TABLE  = ('t' << 24) | ('b' << 16) | ('l' << 8) | ' '; // 0x74626C20
    private const uint CTRL_ID_SECD   = ('s' << 24) | ('e' << 16) | ('c' << 8) | 'd'; // 0x73656364
    private const uint CTRL_ID_COLD   = ('c' << 24) | ('o' << 16) | ('l' << 8) | 'd'; // 0x636F6C64
    private const uint CTRL_ID_FN     = ('f' << 24) | ('n' << 16) | (' ' << 8) | ' '; // footnote
    private const uint CTRL_ID_EN     = ('e' << 24) | ('n' << 16) | (' ' << 8) | ' '; // endnote

    // ──────────────────────────────────────────────────────────────────────
    public PolyDonkyument Read(Stream input)
    {
        ArgumentNullException.ThrowIfNull(input);

        HwpLog.Write("[HwpReader.Read] 시작");
        var tmpPath = Path.GetTempFileName();
        try
        {
            using (var fs = File.Create(tmpPath))
                input.CopyTo(fs);

            using var root = RootStorage.OpenRead(tmpPath);

            var header  = ParseFileHeader(root);
            var docInfo = ParseDocInfo(root, header.IsCompressed);
            var body    = ParseBodyText(root, header.IsCompressed);
            return BuildDocument(docInfo, body, root);
        }
        finally
        {
            try { File.Delete(tmpPath); } catch { }
        }
    }

    // ── FileHeader ─────────────────────────────────────────────────────────

    private static HwpFileHeader ParseFileHeader(RootStorage root)
    {
        using var stream = root.OpenStream("FileHeader");
        Span<byte> buf = stackalloc byte[256];
        int read = stream.Read(buf);
        if (read < 40)
            throw new InvalidOperationException("FileHeader too short");

        var sig = Encoding.ASCII.GetString(buf[..18]).TrimEnd('\0');
        if (sig != "HWP Document File")
            throw new InvalidOperationException($"Invalid HWP signature: {sig}");

        uint flags = BitConverter.ToUInt32(buf[36..]);
        if ((flags & 0x02) != 0)
            throw new InvalidOperationException("Encrypted HWP files are not supported");

        bool compressed = (flags & 0x01) != 0;
        HwpLog.Write($"[ParseFileHeader] flags=0x{flags:X8}, IsCompressed={compressed}");
        return new HwpFileHeader { IsCompressed = compressed };
    }

    // ── DocInfo ────────────────────────────────────────────────────────────

    private static HwpDocInfo ParseDocInfo(RootStorage root, bool isCompressed)
    {
        using var stream = root.OpenStream("DocInfo");
        var data = ReadAllBytes(stream);
        if (isCompressed) data = Decompress(data);

        var info = new HwpDocInfo();
        int binSeq = 0; // BIN_DATA records are 1-based sequential in DocInfo

        // DocInfo 태그 요약
        var docInfoTags = new Dictionary<uint, int>();
        ForEachRecord(data, (tagId, level, payload) =>
        {
            if (!docInfoTags.ContainsKey(tagId))
                docInfoTags[tagId] = 0;
            docInfoTags[tagId]++;
        });
        HwpLog.Write($"[HwpReader.ParseDocInfo] DocInfo tags: {string.Join(", ", docInfoTags.OrderBy(x => x.Key).Select(x => $"0x{x.Key:X3}(×{x.Value})"))}");

        ForEachRecord(data, (tagId, level, payload) =>
        {
            switch (tagId)
            {
                case TAG_DOCUMENT_PROPERTIES when payload.Length >= 2:
                    info.SectionCount = BitConverter.ToUInt16(payload, 0);
                    HwpLog.Write($"[HwpReader] TAG_DOCUMENT_PROPERTIES: SectionCount={info.SectionCount}");
                    break;

                case TAG_FACE_NAME:
                    // HWPTAG_FACE_NAME 레이아웃:
                    //   byte 0     : properties (hasAlt/hasSubst/hasTypeInfo 비트)
                    //   bytes 1-2  : faceNameLen (uint16, in chars)
                    //   bytes 3..  : faceName (UTF-16 LE, len*2 bytes)
                    //   …(타입정보·대체폰트·기본폰트 추가 필드)
                    try
                    {
                        if (payload.Length >= 3)
                        {
                            int nameLen = BitConverter.ToUInt16(payload, 1);
                            int nameBytes = nameLen * 2;
                            if (3 + nameBytes <= payload.Length)
                            {
                                var name = Encoding.Unicode.GetString(payload, 3, nameBytes);
                                info.FontNames.Add(name);
                            }
                            else
                            {
                                info.FontNames.Add("");
                            }
                        }
                        else
                        {
                            info.FontNames.Add("");
                        }
                    }
                    catch { info.FontNames.Add(""); }
                    break;

                case TAG_BORDER_FILL:
                    info.BorderFills.Add(ParseBorderFill(payload));
                    break;

                case TAG_BIN_DATA when payload.Length >= 4:
                    {
                        binSeq++;
                        ushort binId   = BitConverter.ToUInt16(payload, 0);
                        ushort binType = BitConverter.ToUInt16(payload, 2);
                        // binType: 0=link, 1=embedded, 2=stored
                        var binfo = new HwpBinInfo
                        {
                            Id         = binId > 0 ? binId : binSeq,
                            IsEmbedded = binType == 1 || binType == 2,
                        };

                        // For link type (binType=0), payload[4..] may contain a filename
                        if (binType == 0 && payload.Length > 4)
                        {
                            try
                            {
                                ushort nameLen = BitConverter.ToUInt16(payload, 4);
                                if (payload.Length >= 6 + nameLen * 2)
                                    binfo.LinkPath = Encoding.Unicode.GetString(payload, 6, nameLen * 2);
                            }
                            catch { }
                        }

                        // For embedded/stored, try to detect extension from payload[6..7] (format code)
                        if (payload.Length >= 8)
                        {
                            ushort fmt = BitConverter.ToUInt16(payload, 6);
                            binfo.Format = fmt switch
                            {
                                1  => "bmp",
                                2  => "gif",
                                3  => "jpg",
                                4  => "png",
                                5  => "wmf",
                                6  => "ole",
                                _  => ""
                            };
                        }

                        info.BinInfos.Add(binfo);
                    }
                    break;

                case TAG_CHAR_SHAPE:
                    info.CharShapes.Add(ParseCharShape(payload, info.FontNames));
                    break;

                case TAG_PARA_SHAPE:
                    info.ParaShapes.Add(ParseParaShape(payload));
                    break;
            }
        });

        return info;
    }

    // ── BodyText ───────────────────────────────────────────────────────────

    private static HwpBodyText ParseBodyText(RootStorage root, bool isCompressed)
    {
        var body = new HwpBodyText();

        if (!root.TryOpenStorage("BodyText", out var bodyDir))
            return body;

        for (int i = 0; i < 512; i++)
        {
            if (!bodyDir.TryOpenStream($"Section{i}", out var sectionStream))
                break;

            using (sectionStream)
            {
                var data = ReadAllBytes(sectionStream);
                if (isCompressed) data = Decompress(data);
                var recs = CollectRecords(data);
                ParseSectionRecords(recs, body);
            }
        }

        return body;
    }

    // ── Section record parsing (level-aware state machine) ─────────────────

    private static void ParseSectionRecords(List<HwpRecord> recs, HwpBodyText body)
    {
        HwpParagraph? current = null;
        int i = 0;
        // 페이지 인덱스 추적: 페이지 나누기 플래그(columnType bit 2)가 등장할 때마다 증가.
        // 단, 첫 단락의 페이지 나누기 플래그는 무시(문서 시작이므로 page 0 에 머무름).
        int currentPageIndex = 0;
        bool sawFirstParagraph = false;

        HwpLog.Write(
            $"[HwpReader.ParseSectionRecords] Total records: {recs.Count}");

        // 진단: CTRL_HEADER 레코드 구조 미리 스캔
        var ctrlHeaders = new Dictionary<string, (int count, uint level)>();
        for (int j = 0; j < recs.Count; j++)
        {
            if (recs[j].TagId == TAG_CTRL_HEADER && recs[j].Payload.Length >= 4)
            {
                uint ctrlId = BitConverter.ToUInt32(recs[j].Payload, 0);
                string ctrlName = ctrlId switch
                {
                    CTRL_ID_GSO => "gso",
                    CTRL_ID_HEADER => "head",
                    CTRL_ID_FOOTER => "foot",
                    CTRL_ID_TABLE => "tbl ",
                    CTRL_ID_SECD => "secd",
                    CTRL_ID_COLD => "cold",
                    _ => $"0x{ctrlId:X8}"
                };
                if (!ctrlHeaders.ContainsKey(ctrlName))
                    ctrlHeaders[ctrlName] = (0, recs[j].Level);
                var (cnt, _) = ctrlHeaders[ctrlName];
                ctrlHeaders[ctrlName] = (cnt + 1, recs[j].Level);
            }
        }
        if (ctrlHeaders.Count > 0)
            HwpLog.Write($"[ParseSectionRecords] CTRL_HEADER summary: {string.Join(", ", ctrlHeaders.OrderBy(x => x.Key).Select(x => $"{x.Key}={x.Value.count}@L{x.Value.level}"))}");
        else
            HwpLog.Write($"[ParseSectionRecords] ⚠ No CTRL_HEADER records found at any level!");

        while (i < recs.Count)
        {
            var rec = recs[i];

            switch (rec.TagId)
            {
                case TAG_PAGE_DEF:
                    if (body.PageDef == null && rec.Payload.Length >= 32)
                        body.PageDef = ParsePageDef(rec.Payload);
                    break;

                // 진단: 모든 CTRL_HEADER 레코드 로깅 (레벨 무관)
                case TAG_CTRL_HEADER when rec.Level != 1 && rec.Payload.Length >= 4:
                    {
                        uint ctrlId = BitConverter.ToUInt32(rec.Payload, 0);
                        string ctrlName = ctrlId switch
                        {
                            CTRL_ID_GSO => "gso",
                            CTRL_ID_HEADER => "head",
                            CTRL_ID_FOOTER => "foot",
                            CTRL_ID_TABLE => "tbl ",
                            CTRL_ID_SECD => "secd",
                            CTRL_ID_COLD => "cold",
                            _ => $"0x{ctrlId:X8}"
                        };
                        HwpLog.Write($"[ParseSectionRecords] ⚠ Found CTRL_HEADER at unexpected level {rec.Level} (index {i}): ctrlId={ctrlName}");
                    }
                    break;

                // HWP 레벨 구조: PARA_HEADER(N) → 자식 PARA_TEXT/CHAR_SHAPE/LINE_SEG/CTRL_HEADER(N+1).
                // 본문 단락은 PARA_HEADER 가 레벨 0 에 있고, 그 PARA_TEXT 는 레벨 1.
                case TAG_PARA_HEADER when rec.Level == 0:
                    if (current != null)
                        body.Blocks.Add(new HwpParagraphBlock { Paragraph = current });
                    current = new HwpParagraph();
                    // PARA_HEADER payload layout (KS X 5700):
                    //   offset 0-3:  nChars (uint32)
                    //   offset 4-7:  nCtrlMask (uint32)
                    //   offset 8-9:  paraShapeId (uint16) — 단락 모양 참조
                    //   offset 10:   styleId (uint8)
                    //   offset 11:   columnType (uint8) — 단/페이지 나누기 종류
                    //                  bit 2 (0x04): page break before (페이지 나누기)
                    //   offset 12-13: nCharShapes (uint16)
                    //   offset 14-15: nRangeTags (uint16)
                    if (rec.Payload.Length >= 12)
                    {
                        // paraShapeId 는 uint16 으로 정확히 읽음 (기존 코드는 uint32 였으나
                        // 상위 16비트는 styleId/columnType 이므로 결과적으로 잘못된 값).
                        current.ParaShapeId = BitConverter.ToUInt16(rec.Payload, 8);
                        byte columnType = rec.Payload[11];
                        // bit 2 = page break before. 단, 첫 단락은 문서 시작이므로 무시.
                        current.PageBreakBefore = (columnType & 0x04) != 0;
                        if (current.PageBreakBefore && sawFirstParagraph)
                        {
                            currentPageIndex++;
                            HwpLog.Write($"[ParseSectionRecords] Page break before PARA_HEADER@{i} → page {currentPageIndex}");
                        }
                    }
                    else if (rec.Payload.Length >= 10)
                    {
                        current.ParaShapeId = BitConverter.ToUInt16(rec.Payload, 8);
                    }
                    sawFirstParagraph = true;
                    break;

                case TAG_PARA_TEXT when rec.Level == 1:
                    if (current == null) current = new HwpParagraph();
                    try
                    {
                        var text = ExtractHwpText(rec.Payload);
                        current.Text += text;
                    }
                    catch { }
                    break;

                case TAG_PARA_CHAR_SHAPE when rec.Level == 1 && current != null && current.CharShapeId < 0:
                    // PARA_CHAR_SHAPE payload: 반복 (position uint32, charShapeId uint32).
                    // 단순화: 첫 번째 pair 의 charShapeId 사용 (단락 전체 동일 가정).
                    if (rec.Payload.Length >= 8)
                        current.CharShapeId = (int)BitConverter.ToUInt32(rec.Payload, 4);
                    break;

                // CTRL_HEADER 는 본문 단락(레벨 0)의 자식 인라인 컨트롤이므로 레벨 1 에서 처리.
                case TAG_CTRL_HEADER when rec.Level == 1 && rec.Payload.Length >= 4:
                    {
                        uint ctrlId = BitConverter.ToUInt32(rec.Payload, 0);

                        // 진단: 모든 CTRL_HEADER 의 ctrlId 로깅
                        string ctrlName = "unknown";
                        if (ctrlId == CTRL_ID_GSO) ctrlName = "gso";
                        else if (ctrlId == CTRL_ID_HEADER) ctrlName = "head";
                        else if (ctrlId == CTRL_ID_FOOTER) ctrlName = "foot";
                        else if (ctrlId == CTRL_ID_TABLE) ctrlName = "tbl ";
                        else if (ctrlId == CTRL_ID_SECD) ctrlName = "secd";
                        else if (ctrlId == CTRL_ID_COLD) ctrlName = "cold";
                        else ctrlName = $"0x{ctrlId:X8}";
                        HwpLog.Write($"[ParseSectionRecords] Found CTRL_HEADER at index {i}: ctrlId={ctrlName}");

                        if (ctrlId == CTRL_ID_GSO)
                        {
                            // GSO CTRL_HEADER 페이로드 레이아웃:
                            //   0-3: ctrlId
                            //   4-7: flags (비트 필드):
                            //        bits 0-1: wrapType (0=square, 1=tight, 2=through, 3=none)
                            //        bits 2-3: wrapSide (0=both, 1=left, 2=right, 3=bigger)
                            //        bits 4-5: anchorType (0=page, 1=paragraph, 2=character/inline)
                            //        ... (그 외 수직/수평 기준 등)
                            //   8-11: xOffset (int32, HWPUNIT) — anchor 기준 X
                            //  12-15: yOffset (int32, HWPUNIT) — anchor 기준 Y
                            //  16-19: width (uint32, HWPUNIT)
                            //  20-23: height (uint32, HWPUNIT)
                            double gsoXMm = 0, gsoYMm = 0, gsoWMm = 0, gsoHMm = 0;
                            uint gsoFlags = 0;
                            var p = rec.Payload;
                            if (p.Length >= 8)  gsoFlags = BitConverter.ToUInt32(p, 4);
                            if (p.Length >= 24)
                            {
                                gsoXMm = BitConverter.ToInt32(p, 8)  * HwpUnitToMm;
                                gsoYMm = BitConverter.ToInt32(p, 12) * HwpUnitToMm;
                                gsoWMm = BitConverter.ToUInt32(p, 16) * HwpUnitToMm;
                                gsoHMm = BitConverter.ToUInt32(p, 20) * HwpUnitToMm;
                            }
                            HwpLog.Write($"[ParseSectionRecords] GSO flags=0x{gsoFlags:X8}, anchorType={(gsoFlags >> 4) & 0x3}");
                            // 그리기 개체(GSO): 도형/글상자/이미지
                            i = ParseGsoControl(recs, i + 1, rec.Level + 1, body, gsoXMm, gsoYMm, gsoWMm, gsoHMm, currentPageIndex, gsoFlags);
                            continue;
                        }
                        if (ctrlId == CTRL_ID_HEADER)
                        {
                            HwpLog.Write($"[ParseSectionRecords] → Parsing header at index {i}");
                            var hf = ParseHeaderFooter(recs, ref i, rec.Level + 1);
                            body.Headers.Add(hf);
                            HwpLog.Write($"[ParseSectionRecords] → Header complete, {hf.Paragraphs.Count} paragraphs");
                            continue;
                        }
                        if (ctrlId == CTRL_ID_FOOTER)
                        {
                            HwpLog.Write($"[ParseSectionRecords] → Parsing footer at index {i}");
                            var hf = ParseHeaderFooter(recs, ref i, rec.Level + 1);
                            body.Footers.Add(hf);
                            HwpLog.Write($"[ParseSectionRecords] → Footer complete, {hf.Paragraphs.Count} paragraphs");
                            continue;
                        }
                        if (ctrlId == CTRL_ID_TABLE)
                        {
                            var tbl = ParseTable(recs, ref i, rec.Level + 1);
                            if (tbl != null) body.Blocks.Add(tbl);
                            continue;
                        }
                        // 그 외(secd/cold/fn/en 등): 자식 PARA_HEADER 들이 본문에 섞이지 않도록
                        // 자식 레코드 스킵 (단, PAGE_DEF 는 위 case 에서 별도 수집됨).
                        // 본문에 섞이는 걸 막기 위해 같은 레벨로 돌아갈 때까지 nested 레코드 패스만 PAGE_DEF만 수집.

                        // 진단: secd 컨트롤의 자식 레코드 구조 분석
                        if (ctrlId == CTRL_ID_SECD)
                        {
                            HwpLog.Write($"[ParseSectionRecords] → Analyzing SECD control at index {i}");
                            int secdEndIdx = i + 1;
                            while (secdEndIdx < recs.Count && recs[secdEndIdx].Level > rec.Level + 1)
                                secdEndIdx++;

                            var secdTags = new Dictionary<uint, int>();
                            for (int j = i + 1; j < secdEndIdx && j < recs.Count; j++)
                            {
                                var child = recs[j];
                                if (!secdTags.ContainsKey(child.TagId))
                                    secdTags[child.TagId] = 0;
                                secdTags[child.TagId]++;
                            }
                            if (secdTags.Count > 0)
                                HwpLog.Write($"[ParseSectionRecords]    SECD children: {string.Join(", ", secdTags.OrderBy(x => x.Key).Select(x => $"0x{x.Key:X3}(×{x.Value})"))}");
                        }

                        HwpLog.Write($"[ParseSectionRecords] → Skipping non-GSO/head/foot/tbl control: {ctrlName}");
                        i = SkipControlChildrenButKeepPageDef(recs, i + 1, rec.Level + 1, body);
                        continue;
                    }
            }

            i++;
        }

        if (current != null)
            body.Blocks.Add(new HwpParagraphBlock { Paragraph = current });

        HwpLog.Write($"[HwpReader] ParseSectionRecords complete: " +
            $"{body.Blocks.Count} blocks ({body.Paragraphs.Count()} paragraphs, " +
            $"{body.Blocks.OfType<HwpTableBlock>().Count()} tables), " +
            $"{body.Headers.Count} headers, {body.Footers.Count} footers, " +
            $"{body.Images.Count} images, {body.TextBoxes.Count} textboxes, " +
            $"{body.Shapes.Count} shapes, PageDef={body.PageDef != null}");
    }

    // ── GSO (General Shape Object) control handler ─────────────────────────

    private static int ParseGsoControl(List<HwpRecord> recs, int startIdx, uint minLevel, HwpBodyText body,
        double ctrlXMm = 0, double ctrlYMm = 0, double ctrlWMm = 0, double ctrlHMm = 0,
        int anchorPageIndex = 0, uint ctrlFlags = 0)
    {
        // 기본값: CTRL_HEADER 에서 가져온 위치/크기 (SHAPE_COMPONENT 가 덮어쓸 수 있음).
        double xMm = ctrlXMm, yMm = ctrlYMm, wMm = ctrlWMm, hMm = ctrlHMm;
        bool hasShape = ctrlWMm > 0 || ctrlHMm > 0;
        HwpShapeKind kind = HwpShapeKind.Rectangle;
        int binDataId = 0;
        List<HwpParagraph>? tbContent = null;

        // 진단: GSO 컨트롤 자식 레코드 추적
        int gsoStartIdx = startIdx;
        var gsoChildTags = new List<string>();

        int i = startIdx;
        while (i < recs.Count)
        {
            var rec = recs[i];
            if (rec.Level < minLevel) break;

            // 진단: 모든 GSO 자식 레코드 기록
            gsoChildTags.Add($"0x{rec.TagId:X3}@L{rec.Level}");

            switch (rec.TagId)
            {
                case TAG_SHAPE_COMPONENT when rec.Payload.Length >= 32:
                    {
                        // SHAPE_COMPONENT (HWPTAG_SHAPE_COMPONENT) 레이아웃 (KS X 5700):
                        //   0-3   : childCtrlId (4 ASCII chars: "rect","elli","spol","pic ","ole ","txt ","cont")
                        //   4-7   : groupLevel (uint16) + localFileVersion (uint16)
                        //   8-11  : xPosShape (int32, HWPUNIT) — 그룹 내부면 부모 기준 상대 위치
                        //   12-15 : yPosShape (int32, HWPUNIT)
                        //   16-19 : groupingLevel / ngrp (uint32)
                        //   20-23 : nlevel (uint32)
                        //   24-27 : objW (uint32, HWPUNIT) — 초기 너비
                        //   28-31 : objH (uint32, HWPUNIT) — 초기 높이
                        // CTRL_HEADER 의 위치를 우선 사용 (절대 좌표). SHAPE_COMPONENT 의 xPos/yPos 는 그룹 내 상대.
                        var p = rec.Payload;
                        if (wMm <= 0)
                            wMm = BitConverter.ToUInt32(p, 24) * HwpUnitToMm;
                        if (hMm <= 0)
                            hMm = BitConverter.ToUInt32(p, 28) * HwpUnitToMm;
                        hasShape = true;
                    }
                    break;

                case TAG_LINE_COMPONENT:
                    kind = HwpShapeKind.Line;
                    break;
                case TAG_RECT_COMPONENT:
                    kind = HwpShapeKind.Rectangle;
                    break;
                case TAG_ELLIPSE_COMPONENT:
                    kind = HwpShapeKind.Ellipse;
                    break;
                case TAG_ARC_COMPONENT:
                    kind = HwpShapeKind.Arc;
                    break;
                case TAG_POLYGON_COMPONENT:
                    kind = HwpShapeKind.Polygon;
                    break;
                case TAG_CURVE_COMPONENT:
                    kind = HwpShapeKind.Curve;
                    break;
                case TAG_OLE_COMPONENT:
                    kind = HwpShapeKind.Ole;
                    // OLE 도 binDataId 가 포함되는 경우가 많아 동일한 휴리스틱 사용.
                    if (rec.Payload.Length >= 4 && binDataId == 0)
                        binDataId = TryReadBinDataId(rec.Payload);
                    break;

                case TAG_PICTURE_COMPONENT when rec.Payload.Length >= 4:
                    kind = HwpShapeKind.Picture;
                    // Try to read binDataId:
                    //   offset 0-1: border fill flags
                    //   offset 2-3: picture type / attrs
                    //   Further offsets hold the actual binDataId; try offset 50 (after border data),
                    //   falling back to a sequential counter managed by body.
                    binDataId = TryReadBinDataId(rec.Payload);
                    break;

                case TAG_CONTAINER_COMPONENT:
                    kind = HwpShapeKind.Container;
                    break;

                case TAG_LIST_HEADER when hasShape && kind != HwpShapeKind.Picture:
                    // Textbox: nested paragraphs follow.
                    // HWP 구조: LIST_HEADER@L_n → PARA_HEADER@L_n → PARA_TEXT@L_(n+1).
                    // 즉, PARA_HEADER 는 LIST_HEADER 와 같은 레벨, PARA_TEXT 는 한 레벨 더 깊다.
                    kind = HwpShapeKind.TextBox;
                    var paraLevel = rec.Level;       // PARA_HEADER 레벨 (LIST_HEADER 와 동일)
                    var textLevel = rec.Level + 1;   // PARA_TEXT/CHAR_SHAPE 레벨
                    i++;
                    var tbParas = new List<HwpParagraph>();
                    HwpParagraph? tbCur = null;
                    while (i < recs.Count)
                    {
                        var ir = recs[i];
                        if (ir.Level < paraLevel) break;
                        // 같은 레벨에 paragraph 가 아닌 다른 컨트롤(예: RECT_COMPONENT) 이 나오면 종료
                        if (ir.Level == paraLevel && ir.TagId != TAG_PARA_HEADER)
                            break;
                        switch (ir.TagId)
                        {
                            case TAG_PARA_HEADER when ir.Level == paraLevel:
                                if (tbCur != null) tbParas.Add(tbCur);
                                tbCur = new HwpParagraph();
                                if (ir.Payload.Length >= 10)
                                    tbCur.ParaShapeId = BitConverter.ToUInt16(ir.Payload, 8);
                                break;
                            case TAG_PARA_TEXT when ir.Level == textLevel:
                                if (tbCur == null) tbCur = new HwpParagraph();
                                try { tbCur.Text += ExtractHwpText(ir.Payload); }
                                catch { }
                                break;
                            case TAG_PARA_CHAR_SHAPE when ir.Level == textLevel && tbCur != null && tbCur.CharShapeId < 0:
                                if (ir.Payload.Length >= 8)
                                    tbCur.CharShapeId = (int)BitConverter.ToUInt32(ir.Payload, 4);
                                break;
                        }
                        i++;
                    }
                    if (tbCur != null) tbParas.Add(tbCur);
                    tbContent = tbParas;
                    continue;
            }

            i++;
        }

        // LIST_HEADER 가 존재하고 텍스트 단락이 수집되었으면 무조건 TextBox 로 판정.
        // (RECT_COMPONENT 등 후속 shape component 가 kind 를 덮어쓰는 것을 방지)
        if (tbContent != null && tbContent.Count > 0)
            kind = HwpShapeKind.TextBox;

        HwpLog.Write($"[ParseGsoControl] GSO at idx={gsoStartIdx}: kind={kind}, hasShape={hasShape}, " +
            $"size={wMm:F1}x{hMm:F1}mm, pos=({xMm:F1},{yMm:F1}), " +
            $"tbContent={(tbContent != null ? tbContent.Count.ToString() : "null")}, " +
            $"children=[{string.Join(",", gsoChildTags.Take(20))}{(gsoChildTags.Count > 20 ? "..." : "")}]");

        if (!hasShape) return i;

        // ── 인라인/단락 앵커 LINE → ThematicBreakBlock ─────────────────────────
        // CTRL_HEADER flags (offset 4) 의 bits 4-5: anchorType
        //   0 = page (절대 좌표) | 1 = paragraph (단락 기준) | 2 = character (인라인)
        // 단락·문자 앵커된 수평 선(LINE)은 본문 흐름 안에 구분선으로 삽입해야 한다.
        // 이러한 경우 overlay ShapeObject 로 렌더하면 위치가 완전히 틀어지므로
        // ThematicBreakBlock 으로 변환해 인라인 흐름에 끼워 넣는다.
        if (kind == HwpShapeKind.Line)
        {
            uint anchorType = (ctrlFlags >> 4) & 0x3;
            bool isNonPageAnchor = anchorType != 0;
            // 폴백: anchorType 비트를 읽지 못한 경우에도, 넓은 수평선(넓이>100mm, 높이≈0)이
            // 페이지 절대좌표로 배치 불가능한 좌표(너비가 페이지 밖)를 가지면 인라인으로 판단.
            if (!isNonPageAnchor && wMm > 100 && Math.Abs(hMm) < 5 && xMm + wMm > 220)
                isNonPageAnchor = true;

            HwpLog.Write($"[ParseGsoControl] LINE anchorType={anchorType} ({(isNonPageAnchor ? "non-page → ThematicBreak" : "page → overlay")})");
            if (isNonPageAnchor)
            {
                body.Blocks.Add(new HwpThematicBreakBlock());
                return i;
            }
        }

        // 좌표·크기 sanity 검증. SHAPE_COMPONENT 오프셋이 부정확한 경우 비현실적인 값(수천 mm)이 나옴.
        const double MaxReasonableMm = 1000.0;  // 가장 큰 용지(A0) 도 1000mm 미만
        if (Math.Abs(xMm) > MaxReasonableMm || Math.Abs(yMm) > MaxReasonableMm)
        {
            HwpLog.Write($"[ParseGsoControl] ⚠ Position out of range, clamping: ({xMm:F1},{yMm:F1}) → (0,0)");
            xMm = 0; yMm = 0;
        }
        if (wMm > MaxReasonableMm || wMm < 0)
        {
            HwpLog.Write($"[ParseGsoControl] ⚠ Width out of range, clamping: {wMm:F1} → 100");
            wMm = 100;
        }
        if (hMm > MaxReasonableMm || hMm < 0)
        {
            HwpLog.Write($"[ParseGsoControl] ⚠ Height out of range, clamping: {hMm:F1} → 30");
            hMm = 30;
        }
        xMm = Math.Max(0, xMm);
        yMm = Math.Max(0, yMm);

        switch (kind)
        {
            case HwpShapeKind.Picture:
                body.Images.Add(new HwpImage
                {
                    XMm = xMm, YMm = yMm, WidthMm = wMm, HeightMm = hMm,
                    BinDataId = binDataId,
                    AnchorPageIndex = anchorPageIndex,
                });
                break;

            case HwpShapeKind.TextBox:
                body.TextBoxes.Add(new HwpTextBox
                {
                    XMm = xMm, YMm = yMm, WidthMm = wMm, HeightMm = hMm,
                    Paragraphs = tbContent ?? new List<HwpParagraph>(),
                    AnchorPageIndex = anchorPageIndex,
                });
                break;

            default:
                body.Shapes.Add(new HwpShape
                {
                    XMm = xMm, YMm = yMm, WidthMm = wMm, HeightMm = hMm, Kind = kind,
                    BinDataId = binDataId,
                    AnchorPageIndex = anchorPageIndex,
                });
                break;
        }

        return i;
    }

    // ── CHAR_SHAPE / PARA_SHAPE 파싱 (DocInfo) ─────────────────────────────

    /// <summary>
    /// HWPTAG_CHAR_SHAPE 페이로드 → HwpCharShape.
    /// 레이아웃 (KS X 5700 §5.4.4):
    ///   0-13 : faceNameId (uint16 × 7) — Hangul/Latin/Hanja/Japanese/Other/Symbol/User
    ///   14-20: ratio       (uint8 × 7, 50-200%)
    ///   21-27: charSpacing (int8  × 7, -50~50)
    ///   28-34: relSize     (uint8 × 7, 10-250%)
    ///   35-41: charOffset  (int8  × 7, -50~50)
    ///   42-45: baseSize    (int32, 1/100 pt)
    ///   46-49: properties  (uint32 비트플래그)
    ///   50   : shadowOffsetX (int8)
    ///   51   : shadowOffsetY (int8)
    ///   52-55: color           (uint32 RGB)
    ///   56-59: underlineColor  (uint32 RGB)
    ///   60-63: shadeColor      (uint32 RGB)
    ///   64-67: shadowColor     (uint32 RGB)
    /// </summary>
    private static HwpCharShape ParseCharShape(byte[] p, List<string> fontNames)
    {
        var cs = new HwpCharShape();

        if (p.Length >= 2)
        {
            ushort hangulFaceId = BitConverter.ToUInt16(p, 0);
            if (hangulFaceId < fontNames.Count)
                cs.FontFamily = fontNames[hangulFaceId];
        }

        if (p.Length >= 16) cs.WidthPercent  = p[14];                  // ratio (장평)
        if (p.Length >= 22) cs.LetterSpacingPx = (sbyte)p[21] * 0.5;   // charSpacing (자간) - 대략

        if (p.Length >= 46)
        {
            int baseSize100 = BitConverter.ToInt32(p, 42);
            if (baseSize100 > 0 && baseSize100 < 100000)  // sanity
                cs.FontSizePt = baseSize100 / 100.0;
        }

        if (p.Length >= 50)
        {
            uint props = BitConverter.ToUInt32(p, 46);
            cs.Italic        = (props & 0x00000001u) != 0;
            cs.Bold          = (props & 0x00000002u) != 0;
            // bits 2-4: underline kind (0=none, 1=under, 2=over, 3=through)
            uint ulKind      = (props >> 2) & 0x07u;
            cs.Underline     = ulKind == 1;
            cs.Strikethrough = ulKind == 3;
            // bit 18: strike
            if ((props & (1u << 18)) != 0) cs.Strikethrough = true;
            // bits 15-16: super/subscript
            cs.Superscript   = (props & (1u << 15)) != 0;
            cs.Subscript     = (props & (1u << 16)) != 0;
        }

        if (p.Length >= 56)
        {
            uint rgb = BitConverter.ToUInt32(p, 52);
            cs.Color = FormatRgb(rgb);
        }

        return cs;
    }

    /// <summary>
    /// HWPTAG_PARA_SHAPE 페이로드 → HwpParaShape.
    /// 레이아웃 (요약):
    ///   0-3  : properties (정렬·줄나눔·들여쓰기 종류 비트플래그)
    ///   4-7  : marginLeft  (int32, HWPUNIT)
    ///   8-11 : marginRight
    ///   12-15: indent (int32, HWPUNIT, 첫줄 들여쓰기)
    ///   16-19: marginPrev (단락 위 여백)
    ///   20-23: marginNext (단락 아래 여백)
    ///   24-27: lineSpacing (uint32)
    ///   28-31: tabDefId
    ///   32-35: numberingId
    ///   ...
    /// properties 비트필드 (HWP 5.x):
    ///   bits 2-4: alignment (0=both, 1=left, 2=right, 3=center, 4=distribute, 5=division)
    /// </summary>
    private static HwpParaShape ParseParaShape(byte[] p)
    {
        var ps = new HwpParaShape();

        if (p.Length >= 4)
        {
            uint props = BitConverter.ToUInt32(p, 0);
            uint align = (props >> 2) & 0x7u;
            ps.Alignment = align switch
            {
                1 => Alignment.Left,
                2 => Alignment.Right,
                3 => Alignment.Center,
                4 => Alignment.Distributed, // distribute (균등 분배) — "공 문" 스타일 넓은 자간
                5 => Alignment.Distributed, // division
                _ => Alignment.Justify,  // 0 = both (양쪽 혼합)
            };
        }

        if (p.Length >= 8)  ps.IndentLeftMm    = BitConverter.ToInt32(p, 4) * HwpUnitToMm;
        if (p.Length >= 12) ps.IndentRightMm   = BitConverter.ToInt32(p, 8) * HwpUnitToMm;
        if (p.Length >= 16) ps.IndentFirstLineMm = BitConverter.ToInt32(p, 12) * HwpUnitToMm;
        if (p.Length >= 20) ps.SpaceBeforePt   = BitConverter.ToInt32(p, 16) * HwpUnitToMm * (72.0 / 25.4);
        if (p.Length >= 24) ps.SpaceAfterPt    = BitConverter.ToInt32(p, 20) * HwpUnitToMm * (72.0 / 25.4);

        // 줄간격: HWP 의 lineSpacing 은 100 = 100% 인 비율 → factor.
        if (p.Length >= 28)
        {
            uint ls = BitConverter.ToUInt32(p, 24);
            if (ls > 0 && ls < 1000)
                ps.LineHeightFactor = ls / 100.0;
        }

        // 음수·과대 마진은 0 으로 클램프 (스펙 외 값 방지)
        if (ps.IndentLeftMm  < 0) ps.IndentLeftMm  = 0;
        if (ps.IndentRightMm < 0) ps.IndentRightMm = 0;
        if (ps.IndentFirstLineMm < -50) ps.IndentFirstLineMm = -50;
        if (ps.IndentFirstLineMm >  50) ps.IndentFirstLineMm =  50;
        // 단락 위·아래 여백 상한선: 18pt 초과는 스펙 오독 가능성 (일반 문서는 12pt 이하가 표준)
        if (ps.SpaceBeforePt > 18) ps.SpaceBeforePt = 0;
        if (ps.SpaceAfterPt  > 18) ps.SpaceAfterPt  = 0;

        return ps;
    }

    private static string FormatRgb(uint rgb)
    {
        // HWP stores color as 0x00BBGGRR (B in highest byte after alpha 0).
        // Common interpretation: low 3 bytes are R, G, B in that order.
        byte r = (byte)(rgb & 0xFF);
        byte g = (byte)((rgb >> 8) & 0xFF);
        byte b = (byte)((rgb >> 16) & 0xFF);
        return $"#{r:X2}{g:X2}{b:X2}";
    }

    private static string FormatRgbAbgr(uint abgr)
    {
        // ABGR format: 0xAABBGGRR (alpha, blue, green, red)
        // Extract BGR and format as RGB hex
        byte r = (byte)(abgr & 0xFF);
        byte g = (byte)((abgr >> 8) & 0xFF);
        byte b = (byte)((abgr >> 16) & 0xFF);
        return $"#{r:X2}{g:X2}{b:X2}";
    }

    /// <summary>
    /// HWPTAG_BORDER_FILL (KS X 5700) 페이로드 파싱.
    /// Layout:
    ///   0-1  : attr (WORD)
    ///   2-7  : top border (HWPLINE = type 1B, width 1B, color BBGGRR 4B)
    ///   8-13 : left border
    ///   14-19: bottom border
    ///   20-25: right border
    ///   26-31: diagonal border
    ///   32-35: fill kind (0=none, 1=color, 2=image, 4=gradient)
    ///   36-39: background color (BBGGRR, fill kind=1 only)
    ///   40-43: pattern color (BBGGRR, used if background is 0xFFFFFFFF)
    /// </summary>
    private static HwpBorderFill ParseBorderFill(byte[] p)
    {
        var bf = new HwpBorderFill();
        // HWPLINE 파서: type(1B), width(1B), color BBGGRR(4B) = 6B
        static (byte kind, byte width, string? color) ReadLine(byte[] b, int off)
        {
            if (off + 6 > b.Length) return (0, 0, null);
            byte kind  = b[off];
            byte width = b[off + 1];
            uint bbggrr = BitConverter.ToUInt32(b, off + 2);
            string? color = (kind == 0) ? null : FormatRgb(bbggrr);
            return (kind, width, color);
        }

        if (p.Length >= 8)  { var (k, w, c) = ReadLine(p, 2);  bf.TopKind = k; bf.TopWidth = w; bf.TopColor = c; }
        if (p.Length >= 14) { var (k, w, c) = ReadLine(p, 8);  bf.LeftKind = k; bf.LeftWidth = w; bf.LeftColor = c; }
        if (p.Length >= 20) { var (k, w, c) = ReadLine(p, 14); bf.BottomKind = k; bf.BottomWidth = w; bf.BottomColor = c; }
        if (p.Length >= 26) { var (k, w, c) = ReadLine(p, 20); bf.RightKind = k; bf.RightWidth = w; bf.RightColor = c; }

        // fill (offset 32+)
        if (p.Length >= 40)
        {
            uint fillKind = BitConverter.ToUInt32(p, 32);
            if (fillKind == 1)  // color fill
            {
                uint bgColor  = BitConverter.ToUInt32(p, 36);  // BBGGRR
                uint patColor = p.Length >= 44 ? BitConverter.ToUInt32(p, 40) : 0;
                // 0xFFFFFFFF = transparent/no fill → fall through to pattern color
                if (bgColor != 0xFFFFFFFF && bgColor != 0)
                    bf.BackgroundColor = FormatRgb(bgColor);
                else if (patColor != 0xFFFFFFFF && patColor != 0)
                    bf.BackgroundColor = FormatRgb(patColor);
            }
        }
        return bf;
    }

    // HWP 선 너비 인덱스(0-15) → pt 변환 (KS X 5700 §4.2.1)
    private static readonly double[] HwpLineWidthTable =
        [0.284, 0.340, 0.425, 0.567, 0.709, 0.850, 1.134, 1.417,
         1.701, 1.984, 2.835, 4.252, 5.669, 8.504, 11.339, 14.173];

    private static double HwpLineWidthToPt(byte widthIdx) =>
        widthIdx < HwpLineWidthTable.Length ? HwpLineWidthTable[widthIdx] : 0.75;

    private static BorderLineStyle HwpLineKindToStyle(byte kind) => kind switch
    {
        1 => BorderLineStyle.Solid,
        2 => BorderLineStyle.Dashed,
        3 => BorderLineStyle.Dotted,
        4 => BorderLineStyle.DashDot,
        5 => BorderLineStyle.Double,
        _ => BorderLineStyle.Solid,
    };

    // ── 머리말/꼬리말 파싱 ───────────────────────────────────────────────────

    /// <summary>
    /// CTRL_HEADER "head"/"foot" 뒤의 nested 레코드를 파싱해 단락 목록을 추출.
    /// 레코드 구조: LIST_HEADER(level=ctrlLevel+1) → PARA_HEADER(같은 레벨) → PARA_TEXT(level+1) ...
    /// </summary>
    private static HwpHeaderFooter ParseHeaderFooter(List<HwpRecord> recs, ref int i, uint minLevel)
    {
        var hf = new HwpHeaderFooter();
        HwpParagraph? cur = null;

        // 첫 번째 CTRL_HEADER 자체는 이미 처리되었으므로 다음 인덱스부터 시작
        i++;

        while (i < recs.Count)
        {
            var rec = recs[i];
            if (rec.Level < minLevel) break;

            switch (rec.TagId)
            {
                case TAG_PARA_HEADER when rec.Level == minLevel:
                    if (cur != null) hf.Paragraphs.Add(cur);
                    cur = new HwpParagraph();
                    if (rec.Payload.Length >= 10)
                        cur.ParaShapeId = BitConverter.ToUInt16(rec.Payload, 8);
                    break;

                case TAG_PARA_TEXT when rec.Level == minLevel + 1:
                    if (cur == null) cur = new HwpParagraph();
                    try { cur.Text += ExtractHwpText(rec.Payload); }
                    catch { }
                    break;

                case TAG_PARA_CHAR_SHAPE when rec.Level == minLevel + 1 && cur != null && cur.CharShapeId < 0:
                    if (rec.Payload.Length >= 8)
                        cur.CharShapeId = (int)BitConverter.ToUInt32(rec.Payload, 4);
                    break;
            }
            i++;
        }

        if (cur != null) hf.Paragraphs.Add(cur);
        return hf;
    }

    // ── 표 파싱 ──────────────────────────────────────────────────────────────

    /// <summary>
    /// CTRL_HEADER "tbl " 뒤의 nested 레코드를 파싱해 표 정보 + 셀별 단락 추출.
    /// 레코드 구조:
    ///   TAG_TABLE(level=ctrlLevel+1)  ← 행/열 정보
    ///   LIST_HEADER(level=ctrlLevel+1) ← 각 셀 헤더 (rowCnt × colCnt 개)
    ///     PARA_HEADER(level=ctrlLevel+2)
    ///       PARA_TEXT(level=ctrlLevel+3)
    /// </summary>
    private static HwpTableBlock? ParseTable(List<HwpRecord> recs, ref int i, uint minLevel)
    {
        var tbl = new HwpTableBlock();
        HwpTableCell? curCell = null;
        HwpParagraph? curPara = null;

        int tblStartIdx = i;
        var tblChildTags = new List<string>();
        i++; // CTRL_HEADER 다음부터 처리

        while (i < recs.Count)
        {
            var rec = recs[i];
            if (rec.Level < minLevel) break;

            tblChildTags.Add($"0x{rec.TagId:X3}@L{rec.Level}");

            switch (rec.TagId)
            {
                case TAG_TABLE when rec.Level == minLevel && rec.Payload.Length >= 16:
                    {
                        // TAG_TABLE layout:
                        //   0-3: properties (flags)
                        //   4-5: rowCnt
                        //   6-7: colCnt
                        //   8-9: cellSpacing
                        //   10-13: margin left/right (각 2바이트)
                        //   14-17: margin top/bottom
                        //   ...
                        var p = rec.Payload;
                        tbl.RowCount = BitConverter.ToUInt16(p, 4);
                        tbl.ColCount = BitConverter.ToUInt16(p, 6);
                        HwpLog.Write($"[ParseTable] TAG_TABLE: {tbl.RowCount} rows × {tbl.ColCount} cols (payload len={rec.Payload.Length})");
                    }
                    break;

                case TAG_LIST_HEADER when rec.Level == minLevel && rec.Payload.Length >= 28:
                    {
                        // 셀의 LIST_HEADER (표 셀 정보 포함)
                        // payload 첫 4바이트: paraCount
                        // 그 뒤로 셀 좌표/크기 정보 (col, row, colSpan, rowSpan)
                        if (curCell != null && curPara != null)
                        {
                            curCell.Paragraphs.Add(curPara);
                            curCell.Blocks.Add(curPara);
                            curPara = null;
                        }
                        if (curCell != null)
                            tbl.Cells.Add(curCell);

                        var p = rec.Payload;
                        curCell = new HwpTableCell();
                        // LIST_HEADER payload (KS X 5700 §4.2.10):
                        //   0-1: nParagraphs, 2-7: flags/direction
                        //   8-9: col, 10-11: row (uint16 each)
                        //   12-13: colSpan, 14-15: rowSpan (uint16 each)
                        //   16-19: width (int32, HWPUNIT), 20-23: height (int32, HWPUNIT)
                        //   24-25: cellPaddingTop, 26-27: cellPaddingBottom (uint16 HWPUNIT)
                        //   28-29: cellPaddingLeft, 30-31: cellPaddingRight (uint16 HWPUNIT)
                        //   32-33: borderFillId (uint16, 1-based index into DocInfo BorderFills)
                        if (p.Length >= 16)
                        {
                            curCell.Col     = BitConverter.ToUInt16(p, 8);
                            curCell.Row     = BitConverter.ToUInt16(p, 10);
                            curCell.ColSpan = Math.Max(1, (int)BitConverter.ToUInt16(p, 12));
                            curCell.RowSpan = Math.Max(1, (int)BitConverter.ToUInt16(p, 14));
                        }
                        // 셀 너비 (offset 16-19, int32 HWPUNIT)
                        if (p.Length >= 20)
                        {
                            int widthUnit = BitConverter.ToInt32(p, 16);
                            if (widthUnit > 0)
                                curCell.WidthMm = widthUnit * HwpUnitToMm;
                        }
                        // 셀 높이 (offset 20-23, int32 HWPUNIT)
                        if (p.Length >= 24)
                        {
                            int heightUnit = BitConverter.ToInt32(p, 20);
                            if (heightUnit > 0)
                                curCell.HeightMm = heightUnit * HwpUnitToMm;
                        }
                        // 셀 내부 여백 (offset 24-31, uint16 HWPUNIT 각)
                        if (p.Length >= 32)
                        {
                            curCell.PaddingTopMm    = BitConverter.ToUInt16(p, 24) * HwpUnitToMm;
                            curCell.PaddingBottomMm = BitConverter.ToUInt16(p, 26) * HwpUnitToMm;
                            curCell.PaddingLeftMm   = BitConverter.ToUInt16(p, 28) * HwpUnitToMm;
                            curCell.PaddingRightMm  = BitConverter.ToUInt16(p, 30) * HwpUnitToMm;
                        }
                        // borderFillId (offset 32-33, uint16, 1-based)
                        if (p.Length >= 34)
                            curCell.BorderFillId = BitConverter.ToUInt16(p, 32);
                    }
                    break;

                case TAG_CTRL_HEADER when rec.Level == minLevel && curCell != null && rec.Payload.Length >= 4:
                    {
                        // 셀 내 중첩 표 처리
                        uint ctrlId = BitConverter.ToUInt32(rec.Payload, 0);
                        if (ctrlId == CTRL_ID_TABLE)
                        {
                            // 현재 단락을 저장
                            if (curPara != null)
                            {
                                curCell.Paragraphs.Add(curPara);
                                curCell.Blocks.Add(curPara);
                                curPara = null;
                            }

                            // 중첩 표 파싱
                            var nestedTbl = ParseTable(recs, ref i, minLevel + 1);
                            if (nestedTbl != null)
                            {
                                curCell.Blocks.Add(nestedTbl);
                                HwpLog.Write($"[ParseTable] Nested table found in cell ({curCell.Row},{curCell.Col})");
                            }
                            continue;  // i 가 이미 증가했으므로 루프 끝 i++ 스킵
                        }
                    }
                    break;

                case TAG_PARA_HEADER when rec.Level == minLevel && curCell != null:
                    if (curPara != null)
                    {
                        curCell.Paragraphs.Add(curPara);
                        curCell.Blocks.Add(curPara);
                    }
                    curPara = new HwpParagraph();
                    if (rec.Payload.Length >= 10)
                        curPara.ParaShapeId = BitConverter.ToUInt16(rec.Payload, 8);
                    break;

                case TAG_PARA_TEXT when rec.Level == minLevel + 1 && curPara != null:
                    try { curPara.Text += ExtractHwpText(rec.Payload); }
                    catch { }
                    break;

                case TAG_PARA_CHAR_SHAPE when rec.Level == minLevel + 1
                                          && curPara != null && curPara.CharShapeId < 0:
                    if (rec.Payload.Length >= 8)
                        curPara.CharShapeId = (int)BitConverter.ToUInt32(rec.Payload, 4);
                    break;
            }
            i++;
        }

        if (curCell != null && curPara != null)
        {
            curCell.Paragraphs.Add(curPara);
            curCell.Blocks.Add(curPara);
        }
        if (curCell != null) tbl.Cells.Add(curCell);

        HwpLog.Write($"[ParseTable] Complete: {tbl.RowCount}×{tbl.ColCount} table, {tbl.Cells.Count} cells, " +
            $"children=[{string.Join(",", tblChildTags.Take(30))}{(tblChildTags.Count > 30 ? "..." : "")}]");

        if (tbl.RowCount == 0 || tbl.ColCount == 0) return null;
        return tbl;
    }
    /// <summary>
    /// 비-GSO·비-head/foot/tbl 컨트롤(secd/cold/fn/en 등)의 자식 레코드를 건너뛰되,
    /// PAGE_DEF 만은 body.PageDef 에 수집한다.
    /// </summary>
    private static int SkipControlChildrenButKeepPageDef(
        List<HwpRecord> recs, int startIdx, uint minLevel, HwpBodyText body)
    {
        int i = startIdx;
        while (i < recs.Count && recs[i].Level >= minLevel)
        {
            var rec = recs[i];
            if (rec.TagId == TAG_PAGE_DEF && body.PageDef == null && rec.Payload.Length >= 32)
            {
                body.PageDef = ParsePageDef(rec.Payload);
            }
            i++;
        }
        return i;
    }

    // ── PAGE_DEF parsing ───────────────────────────────────────────────────

    private static HwpPageDef ParsePageDef(byte[] p)
    {
        // PAGE_DEF layout (all fields uint32, HWPUNIT):
        //   0: paperWidth
        //   4: paperHeight
        //   8: marginLeft
        //  12: marginRight
        //  16: marginTop
        //  20: marginBottom
        //  24: headerMargin (distance from paper top to header baseline)
        //  28: footerMargin
        //  32: gutterMargin (제본 여백)
        //  36: flags / textDirection (bit0: landscape flag in some versions)
        double pw = BitConverter.ToUInt32(p, 0)  * HwpUnitToMm;
        double ph = BitConverter.ToUInt32(p, 4)  * HwpUnitToMm;
        double ml = BitConverter.ToUInt32(p, 8)  * HwpUnitToMm;
        double mr = BitConverter.ToUInt32(p, 12) * HwpUnitToMm;
        double mt = BitConverter.ToUInt32(p, 16) * HwpUnitToMm;
        double mb = BitConverter.ToUInt32(p, 20) * HwpUnitToMm;
        double mh = p.Length >= 28 ? BitConverter.ToUInt32(p, 24) * HwpUnitToMm : 10;
        double mf = p.Length >= 32 ? BitConverter.ToUInt32(p, 28) * HwpUnitToMm : 10;

        return new HwpPageDef
        {
            PaperWidthMm   = pw,
            PaperHeightMm  = ph,
            MarginLeftMm   = ml,
            MarginRightMm  = mr,
            MarginTopMm    = mt,
            MarginBottomMm = mb,
            MarginHeaderMm = mh,
            MarginFooterMm = mf,
        };
    }

    // ── BinData ID extraction ──────────────────────────────────────────────

    // Try to extract the BinData reference ID from a PICTURE_COMPONENT payload.
    // The exact offset is version-dependent; scan for a small plausible uint16.
    private static int TryReadBinDataId(byte[] p)
    {
        // The BinData ID is a 1-based uint16 in the payload.
        // Common positions: 50 (after border/fill/frame data), then 4, then 2.
        // Plausible range: 1 to 256.
        foreach (int off in new[] { 50, 4, 2, 52, 6 })
        {
            if (off + 2 > p.Length) continue;
            int candidate = BitConverter.ToUInt16(p, off);
            if (candidate is >= 1 and <= 256)
                return candidate;
        }
        return 0;
    }

    private static byte[]? ReadBinData(RootStorage root, int binId)
    {
        if (binId <= 0) return null;
        if (!root.TryOpenStorage("BinData", out var binDir)) return null;
        var streamName = $"BIN{binId:X4}";
        if (!binDir.TryOpenStream(streamName, out var binStream)) return null;
        using (binStream)
            return ReadAllBytes(binStream);
    }

    private static string DetectMediaType(byte[] data)
    {
        if (data.Length < 4) return "image/png";
        if (data[0] == 0x89 && data[1] == 0x50) return "image/png";
        if (data[0] == 0xFF && data[1] == 0xD8) return "image/jpeg";
        if (data[0] == 0x47 && data[1] == 0x49) return "image/gif";
        if (data[0] == 0x42 && data[1] == 0x4D) return "image/bmp";
        return "image/png";
    }

    // ── Document model construction ────────────────────────────────────────

    private static PolyDonkyument BuildDocument(HwpDocInfo docInfo, HwpBodyText body, RootStorage root)
    {
        var doc     = new PolyDonkyument();
        var section = new Section();
        doc.Sections.Add(section);

        // ── Page settings ──────────────────────────────────────────────────
        if (body.PageDef is { } pd)
        {
            var ps = section.Page;
            double pw = pd.PaperWidthMm;
            double ph = pd.PaperHeightMm;

            HwpLog.Write(
                $"[HwpReader] PAGE_DEF found: width={pw:F1}mm, height={ph:F1}mm");

            // HWP stores actual paper dimensions: landscape → width > height
            if (pw > ph && pw > 10 && ph > 10)
            {
                // Landscape: normalize to portrait (short=width, long=height)
                ps.WidthMm      = ph;
                ps.HeightMm     = pw;
                ps.Orientation  = PageOrientation.Landscape;
                HwpLog.Write(
                    $"[HwpReader] → Landscape detected: {ps.WidthMm:F0}x{ps.HeightMm:F0}mm, Orientation={ps.Orientation}");
            }
            else if (pw > 10 && ph > 10)
            {
                ps.WidthMm      = pw;
                ps.HeightMm     = ph;
                ps.Orientation  = PageOrientation.Portrait;
                HwpLog.Write(
                    $"[HwpReader] → Portrait detected: {ps.WidthMm:F0}x{ps.HeightMm:F0}mm, Orientation={ps.Orientation}");
            }

            ps.SizeKind = MatchPaperSize(ps.WidthMm, ps.HeightMm);

            if (pd.MarginLeftMm   > 0) ps.MarginLeftMm   = pd.MarginLeftMm;
            if (pd.MarginRightMm  > 0) ps.MarginRightMm  = pd.MarginRightMm;
            if (pd.MarginTopMm    > 0) ps.MarginTopMm    = pd.MarginTopMm;
            if (pd.MarginBottomMm > 0) ps.MarginBottomMm = pd.MarginBottomMm;
            if (pd.MarginHeaderMm > 0) ps.MarginHeaderMm = pd.MarginHeaderMm;
            if (pd.MarginFooterMm > 0) ps.MarginFooterMm = pd.MarginFooterMm;
        }

        // ── Headers / Footers ──────────────────────────────────────────────
        // 처음 발견된 머리말/꼬리말을 Center 슬롯에 배치 (HWP는 Left/Center/Right 슬롯 구조가 없음).
        if (body.Headers.Count > 0)
        {
            var slot = section.Page.Header.Center;
            foreach (var hp in body.Headers[0].Paragraphs)
                foreach (var p in ConvertHwpParagraphMulti(hp, docInfo))
                    slot.Paragraphs.Add(p);
        }
        if (body.Footers.Count > 0)
        {
            var slot = section.Page.Footer.Center;
            foreach (var fp in body.Footers[0].Paragraphs)
                foreach (var p in ConvertHwpParagraphMulti(fp, docInfo))
                    slot.Paragraphs.Add(p);
        }

        // ── Body blocks (paragraphs + tables + thematic breaks in order) ──────
        foreach (var block in body.Blocks)
        {
            switch (block)
            {
                case HwpParagraphBlock pb:
                    foreach (var para in ConvertHwpParagraphMulti(pb.Paragraph, docInfo))
                        section.Blocks.Add(para);
                    break;
                case HwpTableBlock tb:
                    var table = ConvertHwpTable(tb, docInfo);
                    if (table != null) section.Blocks.Add(table);
                    break;
                case HwpThematicBreakBlock:
                    section.Blocks.Add(new ThematicBreakBlock { LineColor = "#000000" });
                    break;
            }
        }

        // ── Text boxes ─────────────────────────────────────────────────────
        foreach (var tb in body.TextBoxes)
        {
            var tbo = new TextBoxObject
            {
                WrapMode     = ImageWrapMode.InFrontOfText,
                OverlayXMm   = tb.XMm,
                OverlayYMm   = tb.YMm,
                WidthMm      = tb.WidthMm  > 1 ? tb.WidthMm  : 60,
                HeightMm     = tb.HeightMm > 1 ? tb.HeightMm : 30,
                AnchorPageIndex = tb.AnchorPageIndex,
            };
            foreach (var tp in tb.Paragraphs)
                foreach (var para in ConvertHwpParagraphMulti(tp, docInfo))
                    tbo.Content.Add(para);
            section.Blocks.Add(tbo);
        }

        // ── Images ─────────────────────────────────────────────────────────
        // 사용한 BinDataId 추적 — 동일 ID 가 여러 이미지에 매핑되는 폴백 버그 방지.
        var usedBinIds = new HashSet<int>();
        int nextSeqBinId = 1;
        foreach (var img in body.Images)
        {
            byte[]? imgData = null;
            int effectiveId = 0;

            // Try the declared BinDataId first
            if (img.BinDataId > 0)
            {
                imgData = ReadBinData(root, img.BinDataId);
                if (imgData != null) effectiveId = img.BinDataId;
            }

            // Fallback: 다음 사용되지 않은 시퀀셜 BIN#### 사용
            if (imgData == null)
            {
                while (nextSeqBinId <= 256)
                {
                    int candidate = nextSeqBinId++;
                    if (usedBinIds.Contains(candidate)) continue;
                    var tryData = ReadBinData(root, candidate);
                    if (tryData != null && tryData.Length > 0)
                    {
                        imgData = tryData;
                        effectiveId = candidate;
                        break;
                    }
                }
            }

            if (imgData == null || imgData.Length == 0) continue;
            usedBinIds.Add(effectiveId);

            var ib = new ImageBlock
            {
                Data      = imgData,
                MediaType = DetectMediaType(imgData),
                WrapMode  = ImageWrapMode.InFrontOfText,
                WidthMm   = img.WidthMm  > 1 ? img.WidthMm  : 80,
                HeightMm  = img.HeightMm > 1 ? img.HeightMm : 60,
                OverlayXMm      = img.XMm,
                OverlayYMm      = img.YMm,
                AnchorPageIndex = img.AnchorPageIndex,
            };
            section.Blocks.Add(ib);
        }

        // ── Shapes ─────────────────────────────────────────────────────────
        foreach (var sh in body.Shapes)
        {
            byte[]? oleData = null;

            // OLE 객체 바이너리 데이터 추출 (위치 정보는 유지함).
            if (sh.Kind == HwpShapeKind.Ole && sh.BinDataId > 0)
                oleData = ReadBinData(root, sh.BinDataId);

            var so = new ShapeObject
            {
                Kind         = MapShapeKind(sh.Kind),
                WrapMode     = ImageWrapMode.InFrontOfText,
                OverlayXMm   = sh.XMm,
                OverlayYMm   = sh.YMm,
                WidthMm      = sh.WidthMm  > 1 ? sh.WidthMm  : 40,
                HeightMm     = sh.HeightMm > 1 ? sh.HeightMm : 20,
                AnchorPageIndex = sh.AnchorPageIndex,
                StrokeColor  = "#000000",
                StrokeThicknessPt = 1.0,
                OleData      = oleData,  // OLE 바이너리 데이터 (Ole kind 일 때만 사용)
            };
            section.Blocks.Add(so);
        }

        return doc;
    }

    // ── Helpers ────────────────────────────────────────────────────────────

    /// <summary>
    /// HwpParagraph → Core.Paragraph 변환. 비어 있으면 null.
    /// HWP 줄바꿈 문자(\r)는 별도 Run 으로 처리하지 않고 하나의 단락으로 합침.
    /// docInfo 가 제공되면 ParaShapeId/CharShapeId 로 정렬/들여쓰기/폰트 등 속성을 적용.
    /// </summary>
    /// <summary>
    /// HwpParagraph → Core.Paragraph 다중 분할 변환.
    /// PARA_TEXT 안에 soft line break(0x000A)가 있으면 별도 단락으로 분리한다.
    /// </summary>
    private static IEnumerable<Core.Paragraph> ConvertHwpParagraphMulti(HwpParagraph hp, HwpDocInfo? docInfo = null)
    {
        if (string.IsNullOrWhiteSpace(hp.Text)) yield break;

        var fullText = hp.Text.Replace("\r", "");
        if (string.IsNullOrWhiteSpace(fullText)) yield break;

        // soft line break 단위로 분리. 첫 줄에만 PageBreakBefore 적용.
        var lines = fullText.Split('\n');
        bool isFirstLine = true;
        bool isLastLine = lines.Length <= 1;
        int lineIndex = 0;

        foreach (var line in lines)
        {
            if (string.IsNullOrWhiteSpace(line)) continue;
            isLastLine = (lineIndex == lines.Length - 1);

            var hpLine = new HwpParagraph
            {
                Text            = line,
                ParaShapeId     = hp.ParaShapeId,
                CharShapeId     = hp.CharShapeId,
                PageBreakBefore = isFirstLine && hp.PageBreakBefore,
                IsIntermediateLine = !isLastLine,  // 중간 줄 표시 (SpaceAfter 제거용)
            };
            var p = ConvertHwpParagraph(hpLine, docInfo);
            if (p != null) yield return p;
            isFirstLine = false;
            lineIndex++;
        }
    }

    private static Core.Paragraph? ConvertHwpParagraph(HwpParagraph hp, HwpDocInfo? docInfo = null)
    {
        if (string.IsNullOrWhiteSpace(hp.Text)) return null;

        var paragraph = new Core.Paragraph();
        // HWP 단락 끝 \r 은 PARA 내부에서만 의미가 있고 우리 모델은 단락 단위로 분리되어 있으므로 제거.
        var text = hp.Text.Replace("\r", "").Replace("\n", "");
        if (string.IsNullOrWhiteSpace(text)) return null;

        // 페이지 나누기 (HWP PARA_HEADER columnType bit 2)
        paragraph.Style.ForcePageBreakBefore = hp.PageBreakBefore;

        var run = new Run { Text = text };

        if (docInfo != null)
        {
            // 단락 속성 적용
            if (hp.ParaShapeId >= 0 && hp.ParaShapeId < docInfo.ParaShapes.Count)
            {
                var ps = docInfo.ParaShapes[hp.ParaShapeId];
                // 진단: 기본(Justify/Left)이 아닌 정렬값 로그
                if (ps.Alignment != Alignment.Justify && ps.Alignment != Alignment.Left)
                    HwpLog.Write($"[ConvertHwpParagraph] ParaShape[{hp.ParaShapeId}] alignment={ps.Alignment}, text='{hp.Text.Substring(0, Math.Min(15, hp.Text.Length))}'");
                paragraph.Style.Alignment        = ps.Alignment;
                paragraph.Style.IndentFirstLineMm = ps.IndentFirstLineMm;
                paragraph.Style.IndentLeftMm     = ps.IndentLeftMm;
                paragraph.Style.IndentRightMm    = ps.IndentRightMm;
                paragraph.Style.LineHeightFactor = ps.LineHeightFactor;
                // soft line break 분할 시 중간 줄들은 SpaceBeforePt/SpaceAfterPt 을 0 으로 설정해 누적 간격 방지
                paragraph.Style.SpaceBeforePt    = hp.IsIntermediateLine ? 0 : ps.SpaceBeforePt;
                paragraph.Style.SpaceAfterPt     = hp.IsIntermediateLine ? 0 : ps.SpaceAfterPt;
            }

            // 글자 속성 적용
            if (hp.CharShapeId >= 0 && hp.CharShapeId < docInfo.CharShapes.Count)
            {
                var cs = docInfo.CharShapes[hp.CharShapeId];
                if (!string.IsNullOrEmpty(cs.FontFamily)) run.Style.FontFamily = cs.FontFamily;
                if (cs.FontSizePt > 0) run.Style.FontSizePt = cs.FontSizePt;
                run.Style.Bold          = cs.Bold;
                run.Style.Italic        = cs.Italic;
                run.Style.Underline     = cs.Underline;
                run.Style.Strikethrough = cs.Strikethrough;
                run.Style.Superscript   = cs.Superscript;
                run.Style.Subscript     = cs.Subscript;
                if (!string.IsNullOrEmpty(cs.Color))
                {
                    try { run.Style.Foreground = Color.FromHex(cs.Color); } catch { }
                }
                if (cs.WidthPercent > 0) run.Style.WidthPercent = cs.WidthPercent;
                if (cs.LetterSpacingPx != 0) run.Style.LetterSpacingPx = cs.LetterSpacingPx;
            }
        }

        paragraph.Runs.Add(run);
        return paragraph;
    }

    /// <summary>
    /// HwpTableBlock → Core.Table 변환.
    /// 셀의 (row, col) 정보로 매트릭스 구성. 누락된 셀은 빈 셀로 채움.
    /// </summary>
    private static Table? ConvertHwpTable(HwpTableBlock ht, HwpDocInfo? docInfo = null)
    {
        if (ht.RowCount <= 0 || ht.ColCount <= 0) return null;
        if (ht.Cells.Count == 0) return null;

        var table = new Table();

        // 열 너비 결정: colspan=1인 셀의 너비 우선 사용,
        // 미결정 열은 colspan>1 셀의 너비를 균등 분할로 보완
        var colWidths = new double[ht.ColCount];
        // 1단계: colspan=1 셀에서 개별 열 너비 추출
        foreach (var hc in ht.Cells)
        {
            if (hc.ColSpan == 1 && hc.WidthMm > 0 && hc.Col < ht.ColCount)
                colWidths[hc.Col] = hc.WidthMm;
        }
        // 2단계: 아직 미결정 열은 병합 셀 너비를 균등 분할하여 보완
        foreach (var hc in ht.Cells)
        {
            if (hc.ColSpan <= 1 || hc.WidthMm <= 0) continue;
            double perCol = hc.WidthMm / hc.ColSpan;
            for (int c = hc.Col; c < hc.Col + hc.ColSpan && c < ht.ColCount; c++)
            {
                if (colWidths[c] == 0)
                    colWidths[c] = perCol;
            }
        }
        for (int c = 0; c < ht.ColCount; c++)
            table.Columns.Add(new TableColumn { WidthMm = colWidths[c] });

        // 셀을 (row, col) 키로 인덱싱
        var cellMap = new Dictionary<(int r, int c), HwpTableCell>();
        foreach (var hc in ht.Cells)
        {
            if (hc.Row >= 0 && hc.Row < ht.RowCount && hc.Col >= 0 && hc.Col < ht.ColCount)
                cellMap[(hc.Row, hc.Col)] = hc;
        }

        // 병합 셀이 차지하는 위치 (시작 위치 제외) — 이 위치는 WPF 셀 추가 시 건너뛴다
        var occupied = new HashSet<(int r, int c)>();
        foreach (var hc in ht.Cells)
        {
            for (int rr = hc.Row; rr < hc.Row + hc.RowSpan && rr < ht.RowCount; rr++)
                for (int cc = hc.Col; cc < hc.Col + hc.ColSpan && cc < ht.ColCount; cc++)
                {
                    if (rr == hc.Row && cc == hc.Col) continue;
                    occupied.Add((rr, cc));
                }
        }

        for (int r = 0; r < ht.RowCount; r++)
        {
            var row = new TableRow();
            double rowHeightMm = 0;  // 행의 높이 (첫 셀의 높이 사용)

            for (int c = 0; c < ht.ColCount; c++)
            {
                // 병합 셀에 의해 이미 점유된 위치는 WPF 셀을 추가하지 않는다
                if (occupied.Contains((r, c))) continue;
                var tableCell = new TableCell();
                if (cellMap.TryGetValue((r, c), out var hc))
                {
                    tableCell.ColumnSpan = hc.ColSpan;
                    tableCell.RowSpan = hc.RowSpan;

                    // 셀 너비
                    if (hc.WidthMm > 0)
                        tableCell.WidthMm = hc.WidthMm;

                    // 셀 내부 여백
                    if (hc.PaddingTopMm > 0)    tableCell.PaddingTopMm    = hc.PaddingTopMm;
                    if (hc.PaddingBottomMm > 0) tableCell.PaddingBottomMm = hc.PaddingBottomMm;
                    if (hc.PaddingLeftMm > 0)   tableCell.PaddingLeftMm   = hc.PaddingLeftMm;
                    if (hc.PaddingRightMm > 0)  tableCell.PaddingRightMm  = hc.PaddingRightMm;

                    // BORDER_FILL 테이블에서 배경색 및 테두리 조회
                    if (docInfo != null && hc.BorderFillId >= 1 &&
                        hc.BorderFillId - 1 < docInfo.BorderFills.Count)
                    {
                        var bf = docInfo.BorderFills[hc.BorderFillId - 1];
                        if (!string.IsNullOrEmpty(bf.BackgroundColor))
                            tableCell.BackgroundColor = bf.BackgroundColor;
                        if (bf.TopKind != 0 && !string.IsNullOrEmpty(bf.TopColor))
                            tableCell.BorderTop = new CellBorderSide(HwpLineWidthToPt(bf.TopWidth), bf.TopColor, HwpLineKindToStyle(bf.TopKind));
                        if (bf.LeftKind != 0 && !string.IsNullOrEmpty(bf.LeftColor))
                            tableCell.BorderLeft = new CellBorderSide(HwpLineWidthToPt(bf.LeftWidth), bf.LeftColor, HwpLineKindToStyle(bf.LeftKind));
                        if (bf.BottomKind != 0 && !string.IsNullOrEmpty(bf.BottomColor))
                            tableCell.BorderBottom = new CellBorderSide(HwpLineWidthToPt(bf.BottomWidth), bf.BottomColor, HwpLineKindToStyle(bf.BottomKind));
                        if (bf.RightKind != 0 && !string.IsNullOrEmpty(bf.RightColor))
                            tableCell.BorderRight = new CellBorderSide(HwpLineWidthToPt(bf.RightWidth), bf.RightColor, HwpLineKindToStyle(bf.RightKind));
                    }

                    // 행 높이: 첫 셀의 높이로 설정 (같은 행의 모든 셀이 같은 높이)
                    if (c == 0 && hc.HeightMm > 0)
                        rowHeightMm = hc.HeightMm;

                    // 셀의 블록들(단락 또는 중첩 표) 변환
                    foreach (var block in hc.Blocks)
                    {
                        if (block is HwpParagraph hp)
                        {
                            foreach (var para in ConvertHwpParagraphMulti(hp, docInfo))
                                tableCell.Blocks.Add(para);
                        }
                        else if (block is HwpTableBlock htb)
                        {
                            // 중첩 표 변환
                            var nestedTable = ConvertHwpTable(htb, docInfo);
                            if (nestedTable != null) tableCell.Blocks.Add(nestedTable);
                        }
                    }
                }
                // 빈 셀은 빈 단락 1개로 채움 (편집 가능 상태)
                if (tableCell.Blocks.Count == 0)
                    tableCell.Blocks.Add(new Core.Paragraph());

                row.Cells.Add(tableCell);
            }
            // 행 높이 설정 (첫 셀에서 읽은 높이)
            if (rowHeightMm > 0)
                row.HeightMm = rowHeightMm;
            table.Rows.Add(row);
        }

        return table;
    }

    private static PaperSizeKind MatchPaperSize(double wMm, double hMm)
    {
        // Match within ±3 mm tolerance
        foreach (PaperSizeKind kind in Enum.GetValues<PaperSizeKind>())
        {
            var dim = PageSettings.GetStandardDimensions(kind);
            if (dim is not { } d) continue;
            if (Math.Abs(d.W - wMm) < 3 && Math.Abs(d.H - hMm) < 3)
                return kind;
        }
        return PaperSizeKind.Custom;
    }

    private static ShapeKind MapShapeKind(HwpShapeKind k) => k switch
    {
        HwpShapeKind.Line      => ShapeKind.Line,
        HwpShapeKind.Ellipse   => ShapeKind.Ellipse,
        HwpShapeKind.Polygon   => ShapeKind.Polygon,
        HwpShapeKind.Curve     => ShapeKind.Spline,
        HwpShapeKind.Arc       => ShapeKind.HalfCircle,
        HwpShapeKind.Ole       => ShapeKind.Ole,
        _                      => ShapeKind.Rectangle,
    };

    // ── HWP 텍스트 추출 ────────────────────────────────────────────────────────

    /// <summary>
    /// PARA_TEXT 페이로드(UTF-16 LE)에서 한글/영문 텍스트 추출.
    ///
    /// HWP PARA_TEXT 구조는 복잡 (메타데이터와 텍스트 혼재).
    /// 전략: 페이로드를 2바이트 오프셋으로 슬라이딩하며 각 오프셋에서 시작하는
    /// 텍스트 스팬의 길이 측정. 가장 긴 유효 텍스트 스팬 반환.
    /// </summary>
    private static string ExtractHwpText(byte[] payload)
    {
        if (payload.Length < 2)
            return "";

        string longestText = "";

        // Try different starting offsets (2-byte aligned)
        for (int startOffset = 0; startOffset + 1 < payload.Length; startOffset += 2)
        {
            var sb = new StringBuilder();
            int consecutiveInvalid = 0;

            for (int i = startOffset; i + 1 < payload.Length; i += 2)
            {
                char c = (char)BitConverter.ToUInt16(payload, i);

                bool isValid = false;
                if (c >= 0xAC00 && c <= 0xD7AF) isValid = true;       // Korean syllables
                else if (c == 0x000A) isValid = true;                 // Soft line break (preserved as \n)
                else if (c >= 0x0020 && c < 0x0080) isValid = true;   // ASCII printable
                else if (c >= 0x00A0 && c <= 0x00FF) isValid = true;  // Latin-1 supplement (©, °, ±, ÷, ×, etc.)
                else if (c >= 0x2010 && c <= 0x206F) isValid = true;  // General punctuation (bullets •, en/em dashes, quotes)
                else if (c >= 0x2070 && c <= 0x209F) isValid = true;  // Super/Subscript
                else if (c >= 0x20A0 && c <= 0x20CF) isValid = true;  // Currency symbols (₩, $, €)
                else if (c >= 0x2100 && c <= 0x214F) isValid = true;  // Letter-like symbols (™, ©, ®)
                else if (c >= 0x2150 && c <= 0x218F) isValid = true;  // Number forms (Roman numerals ⅰ ⅱ)
                else if (c >= 0x2190 && c <= 0x21FF) isValid = true;  // Arrows (→ ← ↑ ↓)
                else if (c >= 0x2200 && c <= 0x22FF) isValid = true;  // Mathematical operators
                else if (c >= 0x2500 && c <= 0x257F) isValid = true;  // Box drawing
                else if (c >= 0x25A0 && c <= 0x25FF) isValid = true;  // Geometric shapes (■ □ ▶ ●)
                else if (c >= 0x2600 && c <= 0x26FF) isValid = true;  // Misc symbols (★ ☆ ♥)
                else if (c >= 0x3000 && c <= 0x303F) isValid = true;  // CJK symbols/punctuation (、。「」『』)
                else if (c >= 0x3130 && c <= 0x318F) isValid = true;  // Hangul compatibility jamo (ㄱ ㄴ ㅏ)
                else if (c >= 0xFF00 && c <= 0xFFEF) isValid = true;  // Halfwidth/fullwidth forms (ABC numbers)

                if (isValid)
                {
                    if (consecutiveInvalid < 2)  // Allow up to 1 invalid char in between
                    {
                        sb.Append(c);
                        consecutiveInvalid = 0;
                    }
                    else
                    {
                        break;  // Too many invalid chars, stop
                    }
                }
                else if (sb.Length > 0)
                {
                    consecutiveInvalid++;
                    if (consecutiveInvalid >= 4)
                    {
                        break;  // 4 consecutive invalid chars = end of text
                    }
                }
            }

            // Prefer text with Korean characters (more likely to be body text than ASCII metadata)
            string currentText = sb.ToString().Trim('\0');
            bool hasKorean = currentText.Any(c => c >= 0xAC00 && c <= 0xD7AF);
            bool longerHasKorean = longestText.Any(c => c >= 0xAC00 && c <= 0xD7AF);

            if ((hasKorean && !longerHasKorean) || (hasKorean == longerHasKorean && currentText.Length > longestText.Length))
            {
                longestText = currentText;
            }
        }

        if (longestText.Length > 0)
        {
            HwpLog.Write($"[ExtractHwpText] Found longest text: {longestText.Length} chars, starts with '{longestText.Substring(0, Math.Min(20, longestText.Length))}'");
        }

        return longestText;
    }

    /// <summary>
    /// 지정 레벨보다 높은 레코드들을 건너뛰어 컨트롤 그룹 밖의 첫 인덱스를 반환.
    /// </summary>
    private static int SkipToLevel(List<HwpRecord> recs, int startIdx, uint maxLevel)
    {
        int i = startIdx;
        while (i < recs.Count && recs[i].Level > maxLevel) i++;
        return i;
    }

    // ── 레코드 수집 ─────────────────────────────────────────────────────────

    private static List<HwpRecord> CollectRecords(byte[] data)
    {
        var list = new List<HwpRecord>();
        ForEachRecord(data, (tagId, level, payload) =>
            list.Add(new HwpRecord(tagId, level, payload)));
        return list;
    }

    /// <summary>
    /// HWP 레코드 스트림 순회.
    /// 헤더 DWORD:  bit 9-0 = Tag ID,  bit 19-10 = Level,  bit 31-20 = Size.
    /// Size == 0xFFF → 다음 uint32 가 실제 크기.
    /// </summary>
    private static void ForEachRecord(byte[] data, Action<uint, uint, byte[]> callback)
    {
        int offset = 0;
        while (offset + 4 <= data.Length)
        {
            uint dword = BitConverter.ToUInt32(data, offset);
            offset += 4;

            uint tagId = dword & 0x3FFu;
            uint level = (dword >> 10) & 0x3FFu;
            uint size  = dword >> 20;

            if (size == 0xFFFu)
            {
                if (offset + 4 > data.Length) break;
                size = BitConverter.ToUInt32(data, offset);
                offset += 4;
            }

            if (offset + (int)size > data.Length) break;

            var payload = new byte[size];
            Array.Copy(data, offset, payload, 0, (int)size);
            offset += (int)size;

            callback(tagId, level, payload);
        }
    }

    // ── 유틸리티 ────────────────────────────────────────────────────────────

    private static byte[] ReadAllBytes(CfbStream stream)
    {
        using var ms = new MemoryStream();
        stream.CopyTo(ms);
        return ms.ToArray();
    }

    private static byte[] Decompress(byte[] data)
    {
        using var input  = new MemoryStream(data);
        using var output = new MemoryStream();
        using (var deflate = new DeflateStream(input, CompressionMode.Decompress, leaveOpen: false))
            deflate.CopyTo(output);
        return output.ToArray();
    }

    // ── 내부 전용 모델 ─────────────────────────────────────────────────────

    private sealed class HwpFileHeader
    {
        public bool IsCompressed { get; set; }
    }

    private sealed class HwpDocInfo
    {
        public int SectionCount { get; set; }
        public List<string>          FontNames    { get; } = new();
        public List<HwpBinInfo>      BinInfos     { get; } = new();
        public List<HwpCharShape>    CharShapes   { get; } = new();
        public List<HwpParaShape>    ParaShapes   { get; } = new();
        public List<HwpBorderFill>   BorderFills  { get; } = new();
    }

    private sealed class HwpBorderFill
    {
        // 테두리 선 (top/left/bottom/right): kind=0 이면 없음
        public byte TopKind    { get; set; }
        public byte TopWidth   { get; set; }
        public string? TopColor    { get; set; }
        public byte LeftKind   { get; set; }
        public byte LeftWidth  { get; set; }
        public string? LeftColor   { get; set; }
        public byte BottomKind { get; set; }
        public byte BottomWidth{ get; set; }
        public string? BottomColor { get; set; }
        public byte RightKind  { get; set; }
        public byte RightWidth { get; set; }
        public string? RightColor  { get; set; }
        // 배경 채움
        public string? BackgroundColor { get; set; }   // #RRGGBB or null
    }

    private sealed class HwpCharShape
    {
        public string? FontFamily      { get; set; }
        public double  FontSizePt      { get; set; } = 11;
        public bool    Bold            { get; set; }
        public bool    Italic          { get; set; }
        public bool    Underline       { get; set; }
        public bool    Strikethrough   { get; set; }
        public bool    Superscript     { get; set; }
        public bool    Subscript       { get; set; }
        public string? Color           { get; set; }       // #RRGGBB
        public string? Background      { get; set; }       // #RRGGBB
        public double  WidthPercent    { get; set; } = 100;  // 장평
        public double  LetterSpacingPx { get; set; }          // 자간
    }

    private sealed class HwpParaShape
    {
        public Alignment Alignment        { get; set; } = Alignment.Left;
        public double    IndentFirstLineMm { get; set; }
        public double    IndentLeftMm     { get; set; }
        public double    IndentRightMm    { get; set; }
        public double    LineHeightFactor { get; set; } = 1.2;
        public double    SpaceBeforePt    { get; set; }
        public double    SpaceAfterPt     { get; set; }
    }

    private sealed class HwpBinInfo
    {
        public int    Id         { get; set; }
        public bool   IsEmbedded { get; set; }
        public string Format     { get; set; } = "";
        public string LinkPath   { get; set; } = "";
    }

    private sealed class HwpBodyText
    {
        public HwpPageDef?            PageDef    { get; set; }
        public List<HwpBlock>         Blocks     { get; } = new();
        public List<HwpHeaderFooter>  Headers    { get; } = new();
        public List<HwpHeaderFooter>  Footers    { get; } = new();
        public List<HwpTextBox>       TextBoxes  { get; } = new();
        public List<HwpImage>         Images     { get; } = new();
        public List<HwpShape>         Shapes     { get; } = new();

        // 호환용: 이전 Paragraphs API.
        public IEnumerable<HwpParagraph> Paragraphs =>
            Blocks.OfType<HwpParagraphBlock>().Select(b => b.Paragraph);
    }

    // 본문 블록 — 단락, 표, 또는 수평선
    private abstract class HwpBlock { }

    private sealed class HwpParagraphBlock : HwpBlock
    {
        public HwpParagraph Paragraph { get; set; } = new();
    }

    // 인라인/단락 앵커된 LINE GSO 를 수평선 블록으로 변환할 때 사용
    private sealed class HwpThematicBreakBlock : HwpBlock { }

    private sealed class HwpTableBlock : HwpBlock
    {
        public int RowCount { get; set; }
        public int ColCount { get; set; }
        public List<HwpTableCell> Cells { get; } = new();
        public double WidthMm  { get; set; }
        public double HeightMm { get; set; }
    }

    private sealed class HwpTableCell
    {
        public int Row { get; set; }
        public int Col { get; set; }
        public int RowSpan { get; set; } = 1;
        public int ColSpan { get; set; } = 1;
        public double WidthMm { get; set; }          // 셀 너비 (mm)
        public double HeightMm { get; set; }         // 셀 높이 (mm)
        public double PaddingTopMm { get; set; }     // 셀 위 여백
        public double PaddingBottomMm { get; set; }  // 셀 아래 여백
        public double PaddingLeftMm { get; set; }    // 셀 왼쪽 여백
        public double PaddingRightMm { get; set; }   // 셀 오른쪽 여백
        public int BorderFillId { get; set; } = -1;  // DocInfo BorderFills 참조 (1-based)
        public List<HwpParagraph> Paragraphs { get; set; } = new();
        // 중첩 표 지원: 셀이 단락 대신 표를 포함할 수 있음.
        public List<object> Blocks { get; set; } = new();  // HwpParagraph 또는 HwpTableBlock
    }

    private sealed class HwpHeaderFooter
    {
        public List<HwpParagraph> Paragraphs { get; set; } = new();
    }

    private sealed class HwpPageDef
    {
        public double PaperWidthMm   { get; set; }
        public double PaperHeightMm  { get; set; }
        public double MarginLeftMm   { get; set; }
        public double MarginRightMm  { get; set; }
        public double MarginTopMm    { get; set; }
        public double MarginBottomMm { get; set; }
        public double MarginHeaderMm { get; set; }
        public double MarginFooterMm { get; set; }
    }

    private sealed class HwpParagraph
    {
        public string Text { get; set; } = "";
        public int    ParaShapeId { get; set; } = -1;
        public int    CharShapeId { get; set; } = -1;
        public bool   PageBreakBefore { get; set; } = false;
        public bool   IsIntermediateLine { get; set; } = false;  // soft line break 분할 시 중간 줄 표시
    }

    private sealed class HwpTextBox
    {
        public double XMm { get; set; }
        public double YMm { get; set; }
        public double WidthMm  { get; set; }
        public double HeightMm { get; set; }
        public List<HwpParagraph> Paragraphs { get; set; } = new();
        public int    AnchorPageIndex { get; set; }
    }

    private sealed class HwpImage
    {
        public double XMm       { get; set; }
        public double YMm       { get; set; }
        public double WidthMm   { get; set; }
        public double HeightMm  { get; set; }
        public int    BinDataId { get; set; }
        public int    AnchorPageIndex { get; set; }
    }

    private sealed class HwpShape
    {
        public double       XMm       { get; set; }
        public double       YMm       { get; set; }
        public double       WidthMm   { get; set; }
        public double       HeightMm  { get; set; }
        public HwpShapeKind Kind      { get; set; }
        public int          BinDataId { get; set; }  // OLE 도형용
        public int          AnchorPageIndex { get; set; }
    }

    private record struct HwpRecord(uint TagId, uint Level, byte[] Payload);

    private enum HwpShapeKind
    {
        Rectangle, Line, Ellipse, Arc, Polygon, Curve, Ole, Picture, Container, TextBox
    }
}
