using System.IO.Compression;
using System.Text;
using OpenMcdf;
using PolyDonky.Core;

namespace PolyDonky.Codecs.Hwp;

/// <summary>
/// HwpReader 전용 파일 로거. Debug 빌드에서만 동작한다.
/// d:\Temp\PolyDonky-HwpReader.log 에 기록한다.
/// Debug.WriteLine 은 외부 CLI 프로세스에서 VS 출력 창에 표시되지 않으므로
/// 파일 로그로 대체해 진단한다.
/// </summary>
internal static class HwpLog
{
#if DEBUG
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
#endif

    [System.Diagnostics.Conditional("DEBUG")]
    public static void Write(string message)
    {
#if DEBUG
        var line = $"[{DateTime.Now:HH:mm:ss.fff}] {message}";
        System.Diagnostics.Debug.WriteLine(line);
        try { File.AppendAllText(LogPath, line + "\n"); } catch { }
#endif
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
    private const double MmToPt      = 72.0 / 25.4;

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
    private const uint TAG_PAGE_BORDER_FILL     = 0x04B;  // HWPTAG_PAGE_BORDER_FILL
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
                    {
                        var bfIdx = info.BorderFills.Count + 1;
                        if (payload.Length >= 2)
                        {
                            var hexDump = BitConverter.ToString(payload, 0, Math.Min(50, payload.Length));
                            HwpLog.Write($"[BORDER_FILL#{bfIdx}] len={payload.Length} bytes={hexDump}");
                        }
                        info.BorderFills.Add(ParseBorderFill(payload));
                    }
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

            // HWP 섹션은 항상 새 페이지에서 시작 — 두 번째 섹션부터 페이지 나누기 단락 삽입.
            if (i > 0 && body.Blocks.Count > 0)
            {
                body.Blocks.Add(new HwpParagraphBlock
                {
                    Paragraph = new HwpParagraph { IsSectionBreak = true }
                });
                HwpLog.Write($"[ParseBodyText] Section{i} start → 섹션 경계 페이지 나누기 삽입");
            }

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

                case TAG_PARA_CHAR_SHAPE when rec.Level == 1 && current != null:
                    // PARA_CHAR_SHAPE payload: 반복 (position uint32, charShapeId uint32).
                    // 모든 엔트리를 CharRuns 에 수집해 캐릭터별 서식 변화 보존.
                    for (int kp = 0; kp + 8 <= rec.Payload.Length; kp += 8)
                    {
                        int cpos = (int)BitConverter.ToUInt32(rec.Payload, kp);
                        int csid = (int)BitConverter.ToUInt32(rec.Payload, kp + 4);
                        current.CharRuns.Add((cpos, csid));
                        if (current.CharShapeId < 0) current.CharShapeId = csid;
                    }
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
                                // KS X 5700 §4.7.2 GShapeObject: yPos(8-11), xPos(12-15), w(16-19), h(20-23)
                                gsoYMm = BitConverter.ToInt32(p, 8)  * HwpUnitToMm;
                                gsoXMm = BitConverter.ToInt32(p, 12) * HwpUnitToMm;
                                gsoWMm = BitConverter.ToUInt32(p, 16) * HwpUnitToMm;
                                gsoHMm = BitConverter.ToUInt32(p, 20) * HwpUnitToMm;
                            }
                            HwpLog.Write($"[ParseSectionRecords] GSO flags=0x{gsoFlags:X8}, holdSpace={(gsoFlags & 0x10000000u) != 0}, isBehindText={(gsoFlags & 0x4000u) != 0}");
                            // 앵커 단락 마킹: 이 단락은 GSO 앵커 전용이므로 빈 줄 생성 억제.
                            if (current != null) current.HasGsoAnchor = true;
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
        int shapeBorderFillId = -1;  // SHAPE_COMPONENT 의 borderFillId (채우기 색상 참조)
        // CTRL_HEADER flags 비트 레이아웃 (실측 기반):
        //   bit  0: 글자처럼 취급 (TreatAsChar) — 텍스트 흐름 내 인라인 배치
        //   bit  3-4: 세로 위치 기준 (VertRel): 0=종이(paper), 1=쪽(page), 2=단락(paragraph)
        //   bit  8-9: 가로 위치 기준 (HorzRel): 0=종이(paper), 1=쪽(page), 2=단(column), 3=단락
        //   bit 14: 글 뒤에 배치 (isBehindText) — overlay behind text
        //   bit 28: 자리차지 (HoldSpace) — 본문 흐름에 공간을 점유하되 좌표로 배치 (OLE 차트 등)
        //   bits 14/28 모두 0, bit 0도 0 → 글자 앞에 배치 (InFrontOfText overlay)
        bool isBehindText = (ctrlFlags & 0x4000u)    != 0;  // bit 14
        bool holdSpace    = (ctrlFlags & 0x10000000u) != 0; // bit 28: 자리차지
        bool treatAsChar  = (ctrlFlags & 0x1u)        != 0; // bit  0: 글자처럼 취급
        // '자리차지'(holdSpace)는 '글자처럼 취급'과 다르다 — 인라인 흐름에 글자로 끼우면
        // 개체 높이만큼 본문이 밀려 뒤 콘텐츠가 다음 페이지로 넘어간다. holdSpace 개체는
        // overlay(좌표 배치)로 두고, 인라인은 오직 treatAsChar 일 때만.
        bool isInlineFlow = treatAsChar;
        // 위치 기준: '종이'(0) = 페이지 절대 좌표, 그 외 = 본문 영역(여백 안쪽) 기준
        uint horzRel = (ctrlFlags >> 8) & 0x3;  // bits 8-9
        uint vertRel = (ctrlFlags >> 3) & 0x3;  // bits 3-4
        HwpLog.Write($"[ParseGsoControl@{startIdx}] ctrlFlags=0x{ctrlFlags:X8}, isBehindText={isBehindText}, holdSpace={holdSpace}, treatAsChar={treatAsChar}, horzRel={horzRel}, vertRel={vertRel}");
        List<HwpParagraph>? tbContent = null;

        // 도형별 추가 속성 (도형 서브컴포넌트에서 채워짐)
        double shapeRotationDeg = 0;
        ShapeArrow shapeStartArrow = ShapeArrow.None;
        ShapeArrow shapeEndArrow   = ShapeArrow.None;
        double shapeCornerRadiusPct = 0;
        List<(double X, double Y)>? shapePoints = null;
        // SHAPE_COMPONENT 인라인 외곽선/채우기 (그리기 개체 고유 속성)
        string? shapeLineColor   = null;
        double  shapeLineWidthMm = -1;
        string? shapeFillColor   = null;

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
                case TAG_SHAPE_COMPONENT when rec.Payload.Length >= 28:
                    {
                        var p = rec.Payload;
                        // 진단: SHAPE_COMPONENT 전체 바이트 dump (테두리/채우기 인라인 필드 위치 파악용)
                        HwpLog.Write($"[SHAPE_COMPONENT] len={p.Length} dump={BitConverter.ToString(p)}");
                        if (wMm <= 0 && p.Length >= 24)
                            wMm = BitConverter.ToUInt32(p, 20) * HwpUnitToMm;
                        if (hMm <= 0 && p.Length >= 28)
                            hMm = BitConverter.ToUInt32(p, 24) * HwpUnitToMm;

                        // 회전각: 변환 행렬(offset 52) — double[0]=cos(θ), double[1]=sin(θ)
                        if (p.Length >= 68)
                        {
                            double scaleX = BitConverter.ToDouble(p, 52);
                            double shearY = BitConverter.ToDouble(p, 60);
                            if (Math.Abs(scaleX) <= 1.0 && Math.Abs(shearY) <= 1.0 && (Math.Abs(scaleX) > 0.001 || Math.Abs(shearY) > 0.001))
                            {
                                double rad = Math.Atan2(shearY, scaleX);
                                shapeRotationDeg = rad * 180.0 / Math.PI;
                                if (shapeRotationDeg < 0) shapeRotationDeg += 360.0;
                            }
                        }

                        // 인라인 외곽선/채우기 블록: 변환 행렬 뒤에 위치.
                        // 헤더 52바이트 + 변환 행렬(48) + (스케일+회전 행렬쌍) 96 × matrixCount.
                        // matrixCount 는 offset 50 의 UINT16.
                        //   block+0  : lineColor  (COLORREF, 4바이트)
                        //   block+4  : lineWidth  (UINT32, HWPUNIT)
                        //   block+8  : lineAttr   (UINT32)
                        //   block+13 : fillType   (UINT32, 1=단색 채우기, 0=없음)
                        //   block+17 : faceColor  (COLORREF, 4바이트)
                        if (p.Length >= 52)
                        {
                            int matrixCount = BitConverter.ToUInt16(p, 50);
                            int block = 52 + 48 + matrixCount * 96;
                            if (matrixCount is >= 1 and <= 16 && block + 21 <= p.Length)
                            {
                                shapeLineColor   = FormatRgb(BitConverter.ToUInt32(p, block));
                                shapeLineWidthMm = BitConverter.ToUInt32(p, block + 4) * HwpUnitToMm;
                                uint fillType    = BitConverter.ToUInt32(p, block + 13);
                                if (fillType == 1)
                                    shapeFillColor = FormatRgb(BitConverter.ToUInt32(p, block + 17));
                                HwpLog.Write($"[SHAPE_COMPONENT] line/fill block@{block}: lineColor={shapeLineColor}, lineWidth={shapeLineWidthMm:F3}mm, fillType={fillType}, fillColor={shapeFillColor ?? "none"}");
                            }
                        }
                        hasShape = true;
                    }
                    break;

                case TAG_LINE_COMPONENT:
                    kind = HwpShapeKind.Line;
                    // KS X 5700 HWPTAG_SHAPE_COMPONENT_LINE 페이로드:
                    //   0-3:  startPtX (INT32, HWPUNIT) — bounding box 기준
                    //   4-7:  startPtY
                    //   8-11: endPtX
                    //  12-15: endPtY
                    //  16-17: attr (UINT16): arrow types
                    if (rec.Payload.Length >= 16)
                    {
                        int sx = BitConverter.ToInt32(rec.Payload, 0);
                        int sy = BitConverter.ToInt32(rec.Payload, 4);
                        int ex = BitConverter.ToInt32(rec.Payload, 8);
                        int ey = BitConverter.ToInt32(rec.Payload, 12);
                        double sxMm = sx * HwpUnitToMm, syMm = sy * HwpUnitToMm;
                        double exMm = ex * HwpUnitToMm, eyMm = ey * HwpUnitToMm;
                        HwpLog.Write($"[LINE_COMPONENT] raw=({sx},{sy})→({ex},{ey}) mm=({sxMm:F1},{syMm:F1})→({exMm:F1},{eyMm:F1}) hMm={hMm:F1}");
                        // bounding box 좌표는 절대 위치일 수 있어 상대 변환.
                        // HWP LINE_COMPONENT 좌표계는 화면 좌표계(Y=상단 0) 와 동일하므로 Y-flip 불필요.
                        double bbX = Math.Min(sxMm, exMm), bbY = Math.Min(syMm, eyMm);
                        shapePoints = new List<(double X, double Y)>
                        {
                            (sxMm - bbX, syMm - bbY),
                            (exMm - bbX, eyMm - bbY),
                        };
                        HwpLog.Write($"[LINE_COMPONENT] rel=({shapePoints[0].X:F1},{shapePoints[0].Y:F1})→({shapePoints[1].X:F1},{shapePoints[1].Y:F1})");
                    }
                    if (rec.Payload.Length >= 18)
                    {
                        ushort la = BitConverter.ToUInt16(rec.Payload, 16);
                        shapeStartArrow = MapArrow((uint)((la >> 0) & 0x7));
                        shapeEndArrow   = MapArrow((uint)((la >> 4) & 0x7));
                    }
                    else if (rec.Payload.Length >= 2)
                    {
                        // 짧은 페이로드 fallback (구버전): attr 가 처음 2 바이트
                        ushort la = BitConverter.ToUInt16(rec.Payload, 0);
                        shapeStartArrow = MapArrow((uint)((la >> 0) & 0x7));
                        shapeEndArrow   = MapArrow((uint)((la >> 4) & 0x7));
                    }
                    break;
                case TAG_RECT_COMPONENT:
                    kind = HwpShapeKind.Rectangle;
                    if (rec.Payload.Length >= 2)
                    {
                        ushort roundPct = BitConverter.ToUInt16(rec.Payload, 0);
                        if (roundPct > 0 && roundPct <= 100)
                        {
                            shapeCornerRadiusPct = roundPct;
                            kind = HwpShapeKind.RoundedRect;
                        }
                    }
                    break;
                case TAG_ELLIPSE_COMPONENT:
                    kind = HwpShapeKind.Ellipse;
                    break;
                case TAG_ARC_COMPONENT:
                    kind = HwpShapeKind.Arc;
                    break;
                case TAG_POLYGON_COMPONENT:
                    kind = HwpShapeKind.Polygon;
                    // POLYGON_COMPONENT payload (KS X 5700 §5.7.9):
                    //   0-1: nPoints (UINT16)
                    //   2..: points (INT32 x, INT32 y) × nPoints
                    // 좌표계: HWPUNIT 이 아니라 도형 bounding-box 내부 독자 스케일 (~26,600 units/mm).
                    // 정규화: raw 값의 min/max → (0..wMm, 0..hMm) 으로 선형 매핑.
                    if (rec.Payload.Length >= 2)
                    {
                        int nPts = BitConverter.ToUInt16(rec.Payload, 0);
                        int expected = 2 + nPts * 8;
                        int dataOff = 2;
                        if (expected > rec.Payload.Length && 4 + nPts * 8 <= rec.Payload.Length)
                            dataOff = 4;
                        var rawPts = new List<(long X, long Y)>(nPts);
                        for (int j = 0; j < nPts; j++)
                        {
                            int off = dataOff + j * 8;
                            if (off + 8 > rec.Payload.Length) break;
                            rawPts.Add((BitConverter.ToInt32(rec.Payload, off), BitConverter.ToInt32(rec.Payload, off + 4)));
                        }
                        if (rawPts.Count >= 2)
                        {
                            long minX = rawPts.Min(p => p.X), maxX = rawPts.Max(p => p.X);
                            long minY = rawPts.Min(p => p.Y), maxY = rawPts.Max(p => p.Y);
                            double rangeX = maxX - minX, rangeY = maxY - minY;
                            double bbW = wMm > 0 ? wMm : 10.0;
                            double bbH = hMm > 0 ? hMm : 10.0;
                            shapePoints = rawPts.Select(p => (
                                rangeX > 0 ? (p.X - minX) / rangeX * bbW : bbW / 2.0,
                                rangeY > 0 ? (p.Y - minY) / rangeY * bbH : bbH / 2.0
                            )).ToList();
                            var sample = string.Join(", ", shapePoints.Take(6).Select(p => $"({p.Item1:F1},{p.Item2:F1})"));
                            HwpLog.Write($"[POLYGON_COMPONENT] nPts={nPts}, wMm={bbW:F1}, hMm={bbH:F1}, pts(mm)={sample}");
                        }
                    }
                    break;
                case TAG_CURVE_COMPONENT:
                    kind = HwpShapeKind.Curve;
                    // CURVE_COMPONENT: 좌표계 동일 (독자 스케일), 동일한 bounding-box 정규화 적용.
                    if (rec.Payload.Length >= 2)
                    {
                        int nCrvPts = BitConverter.ToUInt16(rec.Payload, 0);
                        var rawCrv = new List<(long X, long Y)>(nCrvPts);
                        for (int j = 0; j < nCrvPts; j++)
                        {
                            int off = 2 + j * 8;
                            if (off + 8 > rec.Payload.Length) break;
                            rawCrv.Add((BitConverter.ToInt32(rec.Payload, off), BitConverter.ToInt32(rec.Payload, off + 4)));
                        }
                        if (rawCrv.Count >= 2)
                        {
                            long minX = rawCrv.Min(p => p.X), maxX = rawCrv.Max(p => p.X);
                            long minY = rawCrv.Min(p => p.Y), maxY = rawCrv.Max(p => p.Y);
                            double rangeX = maxX - minX, rangeY = maxY - minY;
                            double bbW = wMm > 0 ? wMm : 10.0;
                            double bbH = hMm > 0 ? hMm : 10.0;
                            shapePoints = rawCrv.Select(p => (
                                rangeX > 0 ? (p.X - minX) / rangeX * bbW : bbW / 2.0,
                                rangeY > 0 ? (p.Y - minY) / rangeY * bbH : bbH / 2.0
                            )).ToList();
                        }
                    }
                    break;
                case TAG_OLE_COMPONENT:
                    kind = HwpShapeKind.Ole;
                    if (rec.Payload.Length >= 4 && binDataId == 0)
                        binDataId = TryReadBinDataId(rec.Payload);
                    HwpLog.Write($"[OLE_COMPONENT] payloadLen={rec.Payload.Length}, binDataId={binDataId}");
                    break;

                case TAG_PICTURE_COMPONENT when rec.Payload.Length >= 4:
                    kind = HwpShapeKind.Picture;
                    binDataId = TryReadBinDataId(rec.Payload);
                    HwpLog.Write($"[PICTURE_COMPONENT] payloadLen={rec.Payload.Length}, binDataId={binDataId}, head16={BitConverter.ToString(rec.Payload, 0, Math.Min(16, rec.Payload.Length))}");
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
                            case TAG_PARA_CHAR_SHAPE when ir.Level == textLevel && tbCur != null:
                                // PARA_CHAR_SHAPE 페이로드: (charPos UINT32, charShapeId UINT32) × N
                                for (int kp = 0; kp + 8 <= ir.Payload.Length; kp += 8)
                                {
                                    int cpos = (int)BitConverter.ToUInt32(ir.Payload, kp);
                                    int csid = (int)BitConverter.ToUInt32(ir.Payload, kp + 4);
                                    tbCur.CharRuns.Add((cpos, csid));
                                    if (tbCur.CharShapeId < 0) tbCur.CharShapeId = csid;  // 첫 번째 엔트리는 fallback
                                }
                                HwpLog.Write($"[TextBox.PARA_CHAR_SHAPE] entries={tbCur.CharRuns.Count} text='{tbCur.Text}'");
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

        // Line 도형이고 endpoint 가 추출되지 않았으면 기본 하강 대각선 (0,0)→(w,h)
        if (kind == HwpShapeKind.Line && (shapePoints == null || shapePoints.Count < 2) && wMm > 0 && hMm > 0)
        {
            shapePoints = new List<(double X, double Y)> { (0, 0), (wMm, hMm) };
            HwpLog.Write($"[ParseGsoControl] Line endpoints not found → default diagonal (0,0)→({wMm:F1},{hMm:F1})");
        }

        HwpLog.Write($"[ParseGsoControl] GSO at idx={gsoStartIdx}: kind={kind}, hasShape={hasShape}, " +
            $"size={wMm:F1}x{hMm:F1}mm, pos=({xMm:F1},{yMm:F1}), " +
            $"tbContent={(tbContent != null ? tbContent.Count.ToString() : "null")}, " +
            $"children=[{string.Join(",", gsoChildTags.Take(20))}{(gsoChildTags.Count > 20 ? "..." : "")}]");

        if (!hasShape) return i;

        // ── 자리차지 LINE → ThematicBreakBlock ───────────────────────────────
        // 자리차지(holdSpace) 수평선은 본문 흐름 안에 구분선으로 삽입한다.
        // overlay ShapeObject 로 렌더하면 위치가 틀어지므로 ThematicBreakBlock 으로 변환.
        if (kind == HwpShapeKind.Line)
        {
            bool isNonPageAnchor = holdSpace;
            // 폴백: holdSpace 비트가 0이어도 넓은 수평선(너비>100mm, 높이≈0)이
            // 페이지 절대좌표 밖이면 인라인으로 판단.
            if (!isNonPageAnchor && wMm > 100 && Math.Abs(hMm) < 5 && xMm + wMm > 220)
                isNonPageAnchor = true;

            HwpLog.Write($"[ParseGsoControl] LINE holdSpace={holdSpace} ({(isNonPageAnchor ? "inline → ThematicBreak" : "page-absolute → overlay")})");
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

        // 위치 기준 보정: HorzRel/VertRel 이 '종이'(0) 가 아니면 좌표가 본문 영역
        // (여백 안쪽) 기준이므로 페이지 여백만큼 더해 페이지 절대 좌표로 변환.
        // overlay 객체(비-인라인)에만 적용 — 인라인 객체는 흐름에 배치되어 좌표 미사용.
        if (!isInlineFlow && body.PageDef != null)
        {
            double adjX = horzRel != 0 ? body.PageDef.MarginLeftMm : 0;
            // VertRel=2(단락 기준): Y 는 앵커 단락 상대값이므로 여백 보정 불필요.
            // 실제 페이지 좌표 변환은 FlowDocumentPaginationAdapter 가 레이아웃 후 수행.
            double adjY = vertRel == 1 ? body.PageDef.MarginTopMm  : 0;
            if (adjX != 0 || adjY != 0)
            {
                HwpLog.Write($"[ParseGsoControl] 위치 기준 보정: ({xMm:F1},{yMm:F1}) + 여백({adjX:F1},{adjY:F1}) → ({xMm + adjX:F1},{yMm + adjY:F1})");
                xMm += adjX;
                yMm += adjY;
            }
        }

        switch (kind)
        {
            case HwpShapeKind.Picture:
            {
                var img = new HwpImage
                {
                    XMm = xMm, YMm = yMm, WidthMm = wMm, HeightMm = hMm,
                    BinDataId = binDataId,
                    AnchorPageIndex = anchorPageIndex,
                    IsInline = isInlineFlow,
                    VertRel = vertRel,
                };
                body.Images.Add(img);
                body.Blocks.Add(new HwpImageBlock { Image = img });
                break;
            }

            case HwpShapeKind.TextBox:
            {
                var tb = new HwpTextBox
                {
                    XMm = xMm, YMm = yMm, WidthMm = wMm, HeightMm = hMm,
                    Paragraphs = tbContent ?? new List<HwpParagraph>(),
                    AnchorPageIndex = anchorPageIndex,
                    BorderFillId = shapeBorderFillId,
                    LineColor   = shapeLineColor,
                    LineWidthMm = shapeLineWidthMm,
                    FillColor   = shapeFillColor,
                };
                body.TextBoxes.Add(tb);
                body.Blocks.Add(new HwpTextBoxBlock { TextBox = tb });
                break;
            }

            default:
            {
                var sh = new HwpShape
                {
                    XMm = xMm, YMm = yMm, WidthMm = wMm, HeightMm = hMm, Kind = kind,
                    BinDataId = binDataId,
                    AnchorPageIndex = anchorPageIndex,
                    BorderFillId = shapeBorderFillId,
                    IsBehindText = isBehindText,
                    RotationAngleDeg = shapeRotationDeg,
                    StartArrow = shapeStartArrow,
                    EndArrow = shapeEndArrow,
                    CornerRadiusPct = shapeCornerRadiusPct,
                    Points = shapePoints,
                    IsInline = isInlineFlow,
                    VertRel = vertRel,
                    LineColor   = shapeLineColor,
                    LineWidthMm = shapeLineWidthMm,
                    FillColor   = shapeFillColor,
                };
                body.Shapes.Add(sh);
                body.Blocks.Add(new HwpShapeBlock { Shape = sh });
                break;
            }
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

    // HWP5 채우기 색상은 ABGR 형식으로 저장됨 (low byte = alpha/투명도, high byte = R).
    // 파일의 UINT32 little-endian 값 = 0xRRGGBBAA.
    // 투명 마커: alpha(low byte) == 0xFF 또는 전체가 0xFFFFFFFF.
    private static string? FormatAbgr(uint abgr)
    {
        byte a = (byte)(abgr & 0xFF);
        if (a == 0xFF || abgr == 0xFFFFFFFF) return null;  // transparent
        byte r = (byte)((abgr >> 24) & 0xFF);
        byte g = (byte)((abgr >> 16) & 0xFF);
        byte b = (byte)((abgr >> 8) & 0xFF);
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
    ///   2-7  : top border    (HWPLINE = type 1B, width 1B, color 4B)
    ///   8-13 : left border
    ///   14-19: right border
    ///   20-25: bottom border
    ///   26-31: diagonal border
    ///   32-33: fillType (WORD: 0=none, 1=solid color, 2=image, 4=gradient)
    ///   if fillType == 1 (solid):
    ///     34-37: background color (BBGGRR uint32)
    ///     38-41: pattern color    (BBGGRR uint32)
    ///     42-43: pattern kind     (WORD)
    ///   if fillType == 4 (gradient):
    ///     34-37: gradient type (uint32)
    ///     38-41: angle         (uint32, 1/100 degree)
    ///     42-45: startColor    (uint32)
    ///     46-49: endColor      (uint32)
    ///     ... (more gradient params)
    /// </summary>
    private static HwpBorderFill ParseBorderFill(byte[] p)
    {
        var bf = new HwpBorderFill();

        // HWPLINE 파서: type(1B), width(1B), color 0xAABBGGRR(4B) = 6B
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
        if (p.Length >= 20) { var (k, w, c) = ReadLine(p, 14); bf.RightKind = k; bf.RightWidth = w; bf.RightColor = c; }
        if (p.Length >= 26) { var (k, w, c) = ReadLine(p, 20); bf.BottomKind = k; bf.BottomWidth = w; bf.BottomColor = c; }

        // fill: fillType WORD at offset 32
        if (p.Length >= 34)
        {
            ushort fillType = BitConverter.ToUInt16(p, 32);

            if (fillType == 1 && p.Length >= 38)  // solid color
            {
                // HWP5 채우기 색상은 ABGR 형식 (low byte = alpha).
                // FormatAbgr: alpha==0xFF → null(투명), 나머지 → #RRGGBB.
                uint bkColor  = BitConverter.ToUInt32(p, 34);
                uint patColor = p.Length >= 42 ? BitConverter.ToUInt32(p, 38) : 0xFFFFFFFF;
                bf.BackgroundColor = FormatAbgr(bkColor) ?? FormatAbgr(patColor);
            }
            else if (fillType == 4 && p.Length >= 50)  // gradient
            {
                // gradient: type(4B) angle(4B) startColor(4B) endColor(4B) ...
                uint startColor = BitConverter.ToUInt32(p, 42);
                bf.BackgroundColor = FormatAbgr(startColor);
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

                case TAG_PARA_CHAR_SHAPE when rec.Level == minLevel + 1 && cur != null:
                    for (int kp = 0; kp + 8 <= rec.Payload.Length; kp += 8)
                    {
                        int cpos = (int)BitConverter.ToUInt32(rec.Payload, kp);
                        int csid = (int)BitConverter.ToUInt32(rec.Payload, kp + 4);
                        cur.CharRuns.Add((cpos, csid));
                        if (cur.CharShapeId < 0) cur.CharShapeId = csid;
                    }
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
                        // TAG_TABLE layout (KS X 5700):
                        //   0-3: properties (flags)
                        //   4-5: rowCnt, 6-7: colCnt
                        //   8-9: cellSpacing
                        //   10-11: margin left, 12-13: margin right
                        //   14-15: margin top, 16-17: margin bottom
                        //   18-19: width (HWPUNIT), 20-21: height (HWPUNIT)
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

                case TAG_PARA_CHAR_SHAPE when rec.Level == minLevel + 1 && curPara != null:
                    for (int kp = 0; kp + 8 <= rec.Payload.Length; kp += 8)
                    {
                        int cpos = (int)BitConverter.ToUInt32(rec.Payload, kp);
                        int csid = (int)BitConverter.ToUInt32(rec.Payload, kp + 4);
                        curPara.CharRuns.Add((cpos, csid));
                        if (curPara.CharShapeId < 0) curPara.CharShapeId = csid;
                    }
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
            else if (rec.TagId == TAG_PAGE_BORDER_FILL && body.PageBackgroundBorderFillId < 0
                     && rec.Payload.Length >= 14)
            {
                // HWPTAG_PAGE_BORDER_FILL 레이아웃 (KS X 5700, 14바이트):
                //   0-3  : flags (UINT32)
                //            bit 0: ref area (0=text area, 1=paper 기준)
                //            bit 2: position (0=앞, 1=뒤)
                //   4-5  : top margin    (UINT16, HWPUNIT) — 1417 ≈ 5mm
                //   6-7  : bottom margin (UINT16, HWPUNIT)
                //   8-9  : left margin   (UINT16, HWPUNIT)
                //   10-11: right margin  (UINT16, HWPUNIT)
                //   12-13: borderFillId  (UINT16, 1-based index into DocInfo BorderFills)
                var pb = rec.Payload;
                ushort fillId = BitConverter.ToUInt16(pb, 12);
                HwpLog.Write($"[SkipControl] PAGE_BORDER_FILL: borderFillId@12={fillId}");
                if (fillId >= 1)
                    body.PageBackgroundBorderFillId = fillId;
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

    // PICTURE_COMPONENT: binDataId는 offset 50 (uint16, 1-based)
    // OLE_COMPONENT: binDataId는 offset 12 (uint16, 1-based)
    private static int TryReadBinDataId(byte[] p)
    {
        // KS X 5700 실측: PICTURE_COMPONENT @ 50, OLE_COMPONENT @ 12
        foreach (int off in new[] { 50, 12, 4, 2, 52, 6 })
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

        // BinData 스트림 이름에는 확장자가 붙을 수 있음 (예: BIN0001.OLE, BIN0002.bmp).
        // 접두사 "BIN{id:X4}" 로 시작하는 첫 번째 항목을 찾는다.
        var prefix = $"BIN{binId:X4}";
        string? foundName = null;
        foreach (var entry in binDir.EnumerateEntries())
        {
            if (entry.Name.StartsWith(prefix, StringComparison.OrdinalIgnoreCase))
            {
                foundName = entry.Name;
                break;
            }
        }
        if (foundName == null) return null;

        if (!binDir.TryOpenStream(foundName, out var binStream)) return null;
        byte[] raw;
        using (binStream)
            raw = ReadAllBytes(binStream);

        // HWP BinData 는 보통 raw deflate 압축이지만, 일부는 압축 없이 저장됨.
        // 압축 해제와 원본 모두 시도해서 알려진 이미지/OLE 매직이 보이는 쪽을 선택.
        byte[]? decompressed = null;
        try
        {
            using var ms   = new MemoryStream(raw);
            using var ds   = new DeflateStream(ms, CompressionMode.Decompress);
            using var @out = new MemoryStream();
            ds.CopyTo(@out);
            decompressed = @out.ToArray();
        }
        catch { /* 압축되지 않았을 수 있음 */ }

        // OLE 스트림은 종종 4-byte size 헤더가 앞에 붙음 → magic 검색 시 offset 4 도 확인
        static bool HasValidMagic(byte[] d, out int skipBytes)
        {
            skipBytes = 0;
            if (d == null || d.Length < 4) return false;
            // PNG / JPEG / GIF / BMP / OLE2 / EMF / WMF magic at offset 0
            if (IsKnownMagic(d, 0)) return true;
            // 4-byte size prefix?
            if (d.Length >= 8 && IsKnownMagic(d, 4)) { skipBytes = 4; return true; }
            return false;
        }

        if (decompressed != null && HasValidMagic(decompressed, out int skipD))
            return skipD == 0 ? decompressed : decompressed[skipD..];

        if (HasValidMagic(raw, out int skipR))
            return skipR == 0 ? raw : raw[skipR..];

        // 매직을 못 찾으면, 압축 해제가 됐다면 그것을, 아니면 원본을 반환
        HwpLog.Write($"[ReadBinData] {foundName}: no known image magic found (raw={raw.Length}B, decompressed={decompressed?.Length ?? -1}B)");
        return decompressed ?? raw;
    }

    private static bool IsKnownMagic(byte[] d, int off)
    {
        if (d.Length < off + 4) return false;
        // PNG: 89 50 4E 47
        if (d[off] == 0x89 && d[off + 1] == 0x50 && d[off + 2] == 0x4E && d[off + 3] == 0x47) return true;
        // JPEG: FF D8 FF
        if (d[off] == 0xFF && d[off + 1] == 0xD8 && d[off + 2] == 0xFF) return true;
        // GIF: 47 49 46 38 (GIF8)
        if (d[off] == 0x47 && d[off + 1] == 0x49 && d[off + 2] == 0x46 && d[off + 3] == 0x38) return true;
        // BMP: 42 4D (BM) — sanity check via DIB header size at offset 14
        if (d[off] == 0x42 && d[off + 1] == 0x4D && d.Length - off >= 18)
        {
            uint dibSize = BitConverter.ToUInt32(d, off + 14);
            // BITMAPINFOHEADER=40, BITMAPV4HEADER=108, BITMAPV5HEADER=124, etc.
            if (dibSize is >= 12 and <= 256) return true;
        }
        // OLE2 compound document: D0 CF 11 E0
        if (d[off] == 0xD0 && d[off + 1] == 0xCF && d[off + 2] == 0x11 && d[off + 3] == 0xE0) return true;
        // EMF: signature " EMF" at offset 40 of metafile header, but record type "01 00 00 00" at offset 0
        if (d[off] == 0x01 && d[off + 1] == 0x00 && d[off + 2] == 0x00 && d[off + 3] == 0x00 && d.Length - off >= 44)
        {
            if (d[off + 40] == 0x20 && d[off + 41] == 0x45 && d[off + 42] == 0x4D && d[off + 43] == 0x46) return true;
        }
        // WMF placeable header: D7 CD C6 9A
        if (d[off] == 0xD7 && d[off + 1] == 0xCD && d[off + 2] == 0xC6 && d[off + 3] == 0x9A) return true;
        return false;
    }

    private static string DetectMediaType(byte[] data)
    {
        if (data.Length < 4) return "application/octet-stream";
        if (data[0] == 0x89 && data[1] == 0x50) return "image/png";
        if (data[0] == 0xFF && data[1] == 0xD8) return "image/jpeg";
        if (data[0] == 0x47 && data[1] == 0x49) return "image/gif";
        if (data[0] == 0x42 && data[1] == 0x4D) return "image/bmp";
        // OLE2 compound document: D0 CF 11 E0
        if (data[0] == 0xD0 && data[1] == 0xCF && data[2] == 0x11 && data[3] == 0xE0)
            return "application/x-ole-storage";
        // EMF: record type "01 00 00 00" + " EMF" signature at offset 40
        if (data.Length >= 44
            && data[0] == 0x01 && data[1] == 0x00 && data[2] == 0x00 && data[3] == 0x00
            && data[40] == 0x20 && data[41] == 0x45 && data[42] == 0x4D && data[43] == 0x46)
            return "image/x-emf";
        // WMF placeable: D7 CD C6 9A
        if (data[0] == 0xD7 && data[1] == 0xCD && data[2] == 0xC6 && data[3] == 0x9A)
            return "image/x-wmf";
        return "application/octet-stream";
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

        // ── Page background color (SECD PAGE_BORDER_FILL에서 추출) ──────────
        if (body.PageBackgroundBorderFillId >= 1 &&
            body.PageBackgroundBorderFillId - 1 < docInfo.BorderFills.Count)
        {
            var bgFill = docInfo.BorderFills[body.PageBackgroundBorderFillId - 1];
            if (!string.IsNullOrEmpty(bgFill.BackgroundColor))
            {
                section.Page.PaperColor = bgFill.BackgroundColor;
                HwpLog.Write($"[HwpReader] 페이지 배경색: {bgFill.BackgroundColor}");
            }
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

        // ── Body blocks (모든 요소 원본 순서대로) ────────────────────────────
        // HWP 파일에서 단락/표/도형/이미지는 하나의 스트림에 섞여 있으므로
        // body.Blocks 에 파싱 순서대로 모든 타입이 들어 있다.
        var usedBinIds = new HashSet<int>();
        int nextSeqBinId = 1;

        // VertRel=2(단락 기준) 오버레이: 앵커 단락 발견 전까지 AnchorParagraphBlockIndex=-2 상태.
        // 앵커 단락 발견 시 실제 인덱스로 교체한다.
        var pendingParaAnchored = new List<IParaAnchoredOverlay>();

        foreach (var block in body.Blocks)
        {
            switch (block)
            {
                case HwpParagraphBlock pb:
                {
                    int anchorStartIdx = section.Blocks.Count;
                    foreach (var para in ConvertHwpParagraphMulti(pb.Paragraph, docInfo))
                        section.Blocks.Add(para);
                    // 앵커 단락 발견 → 대기 중인 오버레이에 인덱스 기록.
                    // 앵커 단락은 보통 텍스트 없는 GSO 전용 단락이라 ConvertHwpParagraphMulti
                    // 가 스킵한다 → anchorStartIdx 는 그 다음 실제 블록을 가리키게 된다.
                    // 스킵된 단락은 레이아웃 높이 0 이므로 다음 블록이 앵커 단락 위치에서
                    // 시작하며, 그 블록의 Y 가 곧 앵커 단락의 Y 와 같다.
                    if (pb.Paragraph.HasGsoAnchor && pendingParaAnchored.Count > 0)
                    {
                        foreach (var pi in pendingParaAnchored)
                            pi.AnchorParagraphBlockIndex = anchorStartIdx;
                        pendingParaAnchored.Clear();
                    }
                    break;
                }

                case HwpTableBlock tb:
                    var table = ConvertHwpTable(tb, docInfo);
                    if (table != null) section.Blocks.Add(table);
                    break;

                case HwpThematicBreakBlock:
                    section.Blocks.Add(new ThematicBreakBlock { LineColor = "#000000" });
                    break;

                case HwpTextBoxBlock tbb:
                {
                    var tb2 = tbb.TextBox;
                    var tbo = new TextBoxObject
                    {
                        WrapMode        = ImageWrapMode.InFrontOfText,
                        OverlayXMm      = tb2.XMm,
                        OverlayYMm      = tb2.YMm,
                        WidthMm         = tb2.WidthMm  > 1 ? tb2.WidthMm  : 60,
                        HeightMm        = tb2.HeightMm > 1 ? tb2.HeightMm : 30,
                        AnchorPageIndex = tb2.AnchorPageIndex,
                    };
                    // 채우기·외곽선은 SHAPE_COMPONENT 인라인 블록(그리기 개체 고유 속성)을 사용.
                    if (!string.IsNullOrEmpty(tb2.FillColor))
                        tbo.BackgroundColor = tb2.FillColor;
                    if (!string.IsNullOrEmpty(tb2.LineColor) && tb2.LineWidthMm > 0)
                    {
                        tbo.BorderColor = tb2.LineColor;
                        tbo.BorderThicknessPt = tb2.LineWidthMm * MmToPt;
                    }
                    HwpLog.Write($"[BuildDocument] TextBox → bg={tbo.BackgroundColor ?? "null"}, border={tbo.BorderColor ?? "null"}/{tbo.BorderThicknessPt:F2}pt (lineWidth={tb2.LineWidthMm:F3}mm)");
                    foreach (var tp in tb2.Paragraphs)
                        foreach (var para in ConvertHwpParagraphMulti(tp, docInfo))
                            tbo.Content.Add(para);
                    section.Blocks.Add(tbo);
                    break;
                }

                case HwpImageBlock imgBlk:
                {
                    var img = imgBlk.Image;
                    byte[]? imgData = null;
                    int effectiveId = 0;

                    HwpLog.Write($"[BuildDocument] Image BinDataId={img.BinDataId}, IsInline={img.IsInline}, size={img.WidthMm:F1}x{img.HeightMm:F1}mm");

                    if (img.BinDataId > 0)
                    {
                        imgData = ReadBinData(root, img.BinDataId);
                        if (imgData != null) effectiveId = img.BinDataId;
                        HwpLog.Write($"[BuildDocument] Image ReadBinData(id={img.BinDataId}) → {(imgData == null ? "null" : $"{imgData.Length}B, magic={(imgData.Length >= 4 ? imgData[0].ToString("X2") + imgData[1].ToString("X2") + imgData[2].ToString("X2") + imgData[3].ToString("X2") : "?")}")}");
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
                                imgData = tryData; effectiveId = candidate;
                                break;
                            }
                        }
                    }
                    if (imgData == null || imgData.Length == 0)
                    {
                        HwpLog.Write($"[BuildDocument] Image SKIPPED (no data found for any BinDataId)");
                        break;
                    }
                    usedBinIds.Add(effectiveId);

                    var detectedType = DetectMediaType(imgData);
                    HwpLog.Write($"[BuildDocument] Image MediaType={detectedType}, dataLen={imgData.Length}B");

                    // OLE2 / unknown binary은 이미지로 표시 불가 → 건너뜀
                    if (detectedType is "application/x-ole-storage" or "application/octet-stream")
                    {
                        HwpLog.Write($"[BuildDocument] Image SKIPPED (non-image mediaType: {detectedType})");
                        break;
                    }

                    var ib = new ImageBlock
                    {
                        Data      = imgData,
                        MediaType = detectedType,
                        WidthMm   = img.WidthMm  > 1 ? img.WidthMm  : 80,
                        HeightMm  = img.HeightMm > 1 ? img.HeightMm : 60,
                    };
                    // anchorType=1(단락앵커): HWP의 X/Y는 페이지 절대 좌표가 아니라
                    // 단락 기준 상대 오프셋이므로 overlay 좌표로 쓰면 안 됨.
                    // 인라인 흐름에 배치하고 X 오프셋으로 가로 정렬만 근사한다.
                    if (img.IsInline)
                    {
                        ib.WrapMode = ImageWrapMode.Inline;
                        // X 오프셋으로 좌/중앙/우 근사 (본문 너비 ~180mm 가정)
                        const double bodyW = 180.0;
                        var cx = img.XMm + img.WidthMm / 2.0;
                        ib.HAlign = cx < bodyW * 0.33 ? ImageHAlign.Left
                                  : cx > bodyW * 0.67 ? ImageHAlign.Right
                                  : ImageHAlign.Center;
                    }
                    else if (img.VertRel == 2)
                    {
                        // 단락 기준 Y: 앵커 단락 위치를 알아야 최종 페이지 좌표를 확정할 수 있다.
                        // FlowDocumentPaginationAdapter.Paginate() 가 레이아웃 후 OverlayYMm 을 채워준다.
                        ib.WrapMode               = ImageWrapMode.InFrontOfText;
                        ib.OverlayXMm             = img.XMm;
                        ib.OverlayYMm             = 0;   // 레이아웃 후 해소됨
                        ib.AnchorPageIndex        = img.AnchorPageIndex;
                        ib.AnchorRelativeYMm      = img.YMm; // 앵커 단락 상단 기준 오프셋
                        ib.AnchorParagraphBlockIndex = -2;   // 보류 sentinel (다음 앵커 단락이 확정)
                        pendingParaAnchored.Add(ib);
                        HwpLog.Write($"[BuildDocument] Image (VertRel=para) deferred: relY={img.YMm:F1}mm, xMm={img.XMm:F1}mm");
                    }
                    else
                    {
                        ib.WrapMode         = ImageWrapMode.InFrontOfText;
                        ib.OverlayXMm       = img.XMm;
                        ib.OverlayYMm       = img.YMm;
                        ib.AnchorPageIndex  = img.AnchorPageIndex;
                    }
                    section.Blocks.Add(ib);
                    break;
                }

                case HwpShapeBlock shBlk:
                {
                    var sh = shBlk.Shape;
                    byte[]? oleData = null;
                    if (sh.Kind == HwpShapeKind.Ole && sh.BinDataId > 0)
                    {
                        oleData = ReadBinData(root, sh.BinDataId);
                        // OLE binDataId를 사용 완료로 표시 → 이후 이미지 sequential scan이 건너뜀
                        usedBinIds.Add(sh.BinDataId);
                    }

                    var soWidthMm  = sh.WidthMm  > 1 ? sh.WidthMm  : 40;
                    var soHeightMm = sh.HeightMm > 1 ? sh.HeightMm : 20;

                    ImageWrapMode wrapMode;
                    if (sh.IsInline)
                        wrapMode = ImageWrapMode.Inline;
                    else if (sh.IsBehindText)
                        wrapMode = ImageWrapMode.BehindText;
                    else
                        wrapMode = ImageWrapMode.InFrontOfText;

                    var so = new ShapeObject
                    {
                        Kind              = MapShapeKind(sh.Kind),
                        WrapMode          = wrapMode,
                        OverlayXMm        = sh.IsInline ? 0 : sh.XMm,
                        OverlayYMm        = sh.IsInline ? 0 : sh.YMm,
                        WidthMm           = soWidthMm,
                        HeightMm          = soHeightMm,
                        AnchorPageIndex   = sh.AnchorPageIndex,
                        RotationAngleDeg  = sh.RotationAngleDeg,
                        StartArrow        = sh.StartArrow,
                        EndArrow          = sh.EndArrow,
                        StrokeColor       = "#000000",
                        StrokeThicknessPt = 1.0,
                        OleData           = oleData,
                    };
                    // 인라인 도형/차트: HWP X 오프셋(단락 상대)으로 가로 정렬 근사
                    if (sh.IsInline)
                    {
                        const double bodyW = 180.0;
                        var cx = sh.XMm + soWidthMm / 2.0;
                        so.HAlign = cx < bodyW * 0.33 ? ImageHAlign.Left
                                  : cx > bodyW * 0.67 ? ImageHAlign.Right
                                  : ImageHAlign.Center;
                    }
                    // VertRel=2(단락 기준) 오버레이 도형(자리차지 OLE 차트 등): 앵커 단락
                    // 위치를 알아야 최종 Y 를 확정 → 레이아웃 후 ResolveParaAnchoredImages 가 채움.
                    else if (sh.VertRel == 2)
                    {
                        so.OverlayYMm                = 0;
                        so.AnchorRelativeYMm         = sh.YMm;
                        so.AnchorParagraphBlockIndex = -2;
                        pendingParaAnchored.Add(so);
                        HwpLog.Write($"[BuildDocument] Shape (VertRel=para) deferred: kind={sh.Kind}, relY={sh.YMm:F1}mm, xMm={sh.XMm:F1}mm");
                    }

                    // OLE → Chart 미리보기: Hancom Chart(HCH) 포맷에는 embedded preview image가 없으므로,
                    // 데이터 추출이 가능하면 막대 그래프로, 아니면 일반적인 차트 외형 placeholder로 렌더링.
                    if (sh.Kind == HwpShapeKind.Ole && oleData != null && oleData.Length > 16)
                    {
                        HwpChartParser.PopulateChart(so, oleData);
                        HwpLog.Write($"[BuildDocument] OLE chart populated: categories={so.ChartCategories?.Count ?? 0}, series={so.ChartSeries?.Count ?? 0}");
                    }

                    if (sh.Kind == HwpShapeKind.RoundedRect && sh.CornerRadiusPct > 0)
                        so.CornerRadiusMm = Math.Min(soWidthMm, soHeightMm) * sh.CornerRadiusPct / 100.0;

                    if (sh.Points != null && sh.Points.Count >= 2)
                        foreach (var (px, py) in sh.Points)
                            so.Points.Add(new ShapePoint { X = px, Y = py });

                    // 채우기·외곽선: SHAPE_COMPONENT 인라인 블록 (그리기 개체 고유 속성)
                    if (!string.IsNullOrEmpty(sh.FillColor))
                        so.FillColor = sh.FillColor;
                    if (!string.IsNullOrEmpty(sh.LineColor))
                        so.StrokeColor = sh.LineColor;
                    if (sh.LineWidthMm > 0)
                        so.StrokeThicknessPt = sh.LineWidthMm * MmToPt;
                    // lineWidth==0 은 '테두리 없음'이 아니라 헤어라인 → ShapeObject 기본 두께 유지.
                    HwpLog.Write($"[BuildDocument] Shape line/fill: lineColor={sh.LineColor ?? "null"} " +
                        $"lineWidth={sh.LineWidthMm:F3}mm strokePt={so.StrokeThicknessPt:F2} fill={so.FillColor ?? "null"}");
                    HwpLog.Write($"[BuildDocument] Shape → kind={so.Kind} size={so.WidthMm:F1}x{so.HeightMm:F1}mm " +
                        $"pos=({so.OverlayXMm:F1},{so.OverlayYMm:F1}) rot={so.RotationAngleDeg:F1}deg " +
                        $"wrapMode={so.WrapMode} isInline={sh.IsInline} fillColor={so.FillColor ?? "null"} " +
                        $"strokeColor={so.StrokeColor} anchorPage={so.AnchorPageIndex} points={so.Points.Count}");

                    section.Blocks.Add(so);
                    break;
                }
            }
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
        // 섹션 경계 단락: 텍스트 없이 페이지 나누기만 발생시킴.
        if (hp.IsSectionBreak)
        {
            yield return new Core.Paragraph { Style = { ForcePageBreakBefore = true } };
            yield break;
        }
        if (string.IsNullOrWhiteSpace(hp.Text))
        {
            // GSO 앵커 전용 빈 단락은 block 흐름에서 제거 (불필요한 빈 줄 누적 방지).
            if (hp.HasGsoAnchor) yield break;
            // 그 외 빈 단락은 페이지 구조(anchorPage 기준 페이지 생성)를 위해 보존.
            yield return new Core.Paragraph { Style = { ForcePageBreakBefore = hp.PageBreakBefore } };
            yield break;
        }

        var fullText = hp.Text.Replace("\r", "");
        if (string.IsNullOrWhiteSpace(fullText))
        {
            if (!hp.HasGsoAnchor)
                yield return new Core.Paragraph { Style = { ForcePageBreakBefore = hp.PageBreakBefore } };
            yield break;
        }

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
                CharRuns        = hp.CharRuns,  // 캐릭터별 서식 보존
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
            // 다중 CharShape 구간이 있으면 텍스트를 구간별로 잘라 별도 Run 생성.
            var runs = hp.CharRuns;
            if (runs != null && runs.Count > 1)
            {
                // charPos 기준 오름차순 정렬 (이미 정렬되어 있겠지만 안전을 위해)
                var sorted = runs.OrderBy(r => r.CharPos).ToList();
                for (int ri = 0; ri < sorted.Count; ri++)
                {
                    int startPos = sorted[ri].CharPos;
                    int endPos = (ri + 1 < sorted.Count) ? sorted[ri + 1].CharPos : text.Length;
                    if (startPos >= text.Length) break;
                    if (endPos > text.Length) endPos = text.Length;
                    if (endPos <= startPos) continue;
                    var segText = text.Substring(startPos, endPos - startPos);
                    var segRun  = new Run { Text = segText };
                    ApplyCharShape(segRun, docInfo, sorted[ri].CharShapeId);
                    paragraph.Runs.Add(segRun);
                }
                if (paragraph.Runs.Count > 0) return paragraph;
            }

            // 단일 CharShape — 기존 동작.
            ApplyCharShape(run, docInfo, hp.CharShapeId);
        }

        paragraph.Runs.Add(run);
        return paragraph;
    }

    private static void ApplyCharShape(Run run, HwpDocInfo docInfo, int charShapeId)
    {
        if (charShapeId < 0 || charShapeId >= docInfo.CharShapes.Count) return;
        var cs = docInfo.CharShapes[charShapeId];
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

        // 표 너비 설정: HWP 파일에서 저장된 너비가 없으므로,
        // 모든 열 너비의 합으로 표 전체 너비 계산
        double tableWidthMm = colWidths.Sum();
        if (tableWidthMm > 0)
            table.WidthMm = tableWidthMm;

        if (ht.HeightMm > 0)
            table.HeightMm = ht.HeightMm;

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
        HwpShapeKind.Line        => ShapeKind.Line,
        HwpShapeKind.Ellipse     => ShapeKind.Ellipse,
        HwpShapeKind.Polygon     => ShapeKind.Polygon,
        HwpShapeKind.Curve       => ShapeKind.Spline,
        HwpShapeKind.Arc         => ShapeKind.HalfCircle,
        HwpShapeKind.Ole         => ShapeKind.Ole,
        HwpShapeKind.RoundedRect => ShapeKind.RoundedRect,
        _                        => ShapeKind.Rectangle,
    };

    private static StrokeDash HwpLineKindToDash(byte kind) => kind switch
    {
        2 => StrokeDash.Dashed,
        3 => StrokeDash.Dotted,
        4 => StrokeDash.DashDot,
        _ => StrokeDash.Solid,
    };

    private static ShapeArrow MapArrow(uint kind) => kind switch
    {
        1 => ShapeArrow.Open,
        2 => ShapeArrow.Filled,
        3 => ShapeArrow.Diamond,
        4 => ShapeArrow.Circle,
        _ => ShapeArrow.None,
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
        public int                    PageBackgroundBorderFillId { get; set; } = -1; // SECD PAGE_BORDER_FILL 에서 추출
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

    private sealed class HwpShapeBlock : HwpBlock
    {
        public HwpShape Shape { get; set; } = new();
    }

    private sealed class HwpImageBlock : HwpBlock
    {
        public HwpImage Image { get; set; } = new();
    }

    private sealed class HwpTextBoxBlock : HwpBlock
    {
        public HwpTextBox TextBox { get; set; } = new();
    }

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
        /// <summary>
        /// PARA_CHAR_SHAPE 다중 엔트리: (charPos, charShapeId) 쌍의 시퀀스.
        /// 각 엔트리는 charPos 부터 적용되는 캐릭터 서식 ID를 정의 (다음 엔트리의 charPos 까지 유효).
        /// 비어 있으면 단일 CharShapeId 사용.
        /// </summary>
        public List<(int CharPos, int CharShapeId)> CharRuns { get; set; } = new();
        public bool   PageBreakBefore { get; set; } = false;
        public bool   IsIntermediateLine { get; set; } = false;  // soft line break 분할 시 중간 줄 표시
        public bool   IsSectionBreak { get; set; } = false;      // HWP 섹션 경계 → ForcePageBreakBefore 단락 생성
        /// <summary>
        /// GSO(도형·글상자·OLE·이미지) 앵커 단락 여부.
        /// 텍스트가 없고 GSO 앵커만 있는 단락은 block 흐름에서 제거해 빈 줄 누적을 방지한다.
        /// </summary>
        public bool   HasGsoAnchor { get; set; } = false;
    }

    private sealed class HwpTextBox
    {
        public double XMm { get; set; }
        public double YMm { get; set; }
        public double WidthMm  { get; set; }
        public double HeightMm { get; set; }
        public List<HwpParagraph> Paragraphs { get; set; } = new();
        public int    AnchorPageIndex { get; set; }
        public int    BorderFillId { get; set; } = -1;
        public string? LineColor    { get; set; }   // SHAPE_COMPONENT 인라인 외곽선 색
        public double  LineWidthMm  { get; set; } = -1;  // 외곽선 두께 (mm), -1=미설정
        public string? FillColor    { get; set; }   // SHAPE_COMPONENT 인라인 채우기 색
    }

    private sealed class HwpImage
    {
        public double XMm       { get; set; }
        public double YMm       { get; set; }
        public double WidthMm   { get; set; }
        public double HeightMm  { get; set; }
        public int    BinDataId { get; set; }
        public int    AnchorPageIndex { get; set; }
        public bool   IsInline  { get; set; }  // anchorType=1/2 → 인라인 흐름
        public uint   VertRel   { get; set; }   // bits 3-4 of ctrlFlags: 0=paper,1=page,2=paragraph
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
        public int          BorderFillId { get; set; } = -1;  // SHAPE_COMPONENT 의 borderFillId (1-based)
        public bool         IsBehindText { get; set; }
        public bool         IsInline  { get; set; }  // anchorType != 0
        public uint         VertRel   { get; set; }  // bits 3-4 of ctrlFlags: 0=paper,1=page,2=paragraph
        public double       RotationAngleDeg { get; set; }
        public ShapeArrow   StartArrow { get; set; }
        public ShapeArrow   EndArrow   { get; set; }
        public double       CornerRadiusPct { get; set; }     // RECT_COMPONENT roundedCornerPercent (0-100)
        public List<(double X, double Y)>? Points { get; set; }  // POLYGON/CURVE 꼭짓점 (도형 내부 좌표 mm)
        public string?      LineColor   { get; set; }   // SHAPE_COMPONENT 인라인 외곽선 색
        public double       LineWidthMm { get; set; } = -1;  // 외곽선 두께 (mm), -1=미설정
        public string?      FillColor   { get; set; }   // SHAPE_COMPONENT 인라인 채우기 색
    }

    private record struct HwpRecord(uint TagId, uint Level, byte[] Payload);

    private enum HwpShapeKind
    {
        Rectangle, RoundedRect, Line, Ellipse, Arc, Polygon, Curve, Ole, Picture, Container, TextBox
    }
}
