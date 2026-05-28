using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using PolyDonky.Core;

namespace PolyDonky.Convert.Doc;

/// <summary>
/// IWPF → RTF (Rich Text Format) 변환기.
/// 지원: 글자 서식·단락 서식·위첨자/아래첨자·들여쓰기·리스트·이미지·표·메타데이터·
///       도형(\shp, 위치·크기·종류·색상 아웃라인)·OLE 개체(OpaqueBlock 재출력 또는 플레이스holder)·
///       하이퍼링크(\field HYPERLINK)·자동 필드(PAGE/NUMPAGES/DATE/TIME/AUTHOR/TITLE 등)·
///       책갈피(\*\bkmkstart / \*\bkmkend)·변경추적(\revised / \deleted + \revtbl)·
///       각주·미주(\chftn / \footnote [\ftnalt])·주석(\chatn / \annotation + \atnauthor / \atndate)·
///       페이지 설정(\paperw / \paperh / \margl-r-t-b / \landscape / \titlepg / \facingp)·
///       머리말·꼬리말(\header / \footer, Left·Center·Right 슬롯 stacked 출력).
/// v1.0.0 이후 계획: \shp 전체 속성(그림자·3D·꼭짓점 경로 등) + OLE 데이터 완전 직렬화.
/// </summary>
public class DocWriter
{
    // ── 테이블 ──────────────────────────────────────────────────────────────────
    private readonly List<string>   _fonts      = new();
    private readonly List<RtfColor> _colors     = new();
    /// <summary>변경추적 작성자 테이블 — index 0 = "Unknown".</summary>
    private readonly List<string>   _revAuthors = new();

    /// <summary>현재 Write 중인 PolyDonkyument — 각주/미주/주석 본문 lookup 에 사용.</summary>
    private PolyDonkyument? _doc;

    /// <summary>각주/주석 본문을 emit 중인지 — 중첩 footnote ref 무한 재귀 방지.</summary>
    private bool _inFootnoteOrComment;

    private const string DefaultFont = "Arial";
    private const double MmToTwips   = 56.692;  // 1mm ≈ 56.692 twips (1440/25.4)
    private const double PtToTwips   = 20.0;

    // ── public entry ────────────────────────────────────────────────────────────

    public void Write(PolyDonkyument doc, Stream output)
    {
        _doc = doc;
        _fonts.Clear();
        _colors.Clear();
        _revAuthors.Clear();
        _fonts.Add(DefaultFont);
        _colors.Add(new RtfColor(0, 0, 0));   // 색상 0: 기본 검정
        _revAuthors.Add("Unknown");           // 작성자 0: Unknown

        // 1패스: 폰트/색상/작성자 수집
        foreach (var sec in doc.Sections)
            foreach (var blk in sec.Blocks)
                ScanBlock(blk);
        // 각주/미주/주석 본문 안의 폰트/색상/작성자도 수집
        foreach (var fn in doc.Footnotes.Concat(doc.Endnotes))
            foreach (var b in fn.Blocks) ScanBlock(b);
        foreach (var cm in doc.Comments)
            foreach (var b in cm.Blocks) ScanBlock(b);

        // 2패스: RTF 생성
        var sb = new StringBuilder(4096);
        sb.AppendLine(@"{\rtf1\ansi\ansicpg1252\deff0");
        WriteFontTable(sb);
        WriteColorTable(sb);
        WriteRevAuthorTable(sb);
        sb.AppendLine(@"\viewkind4\uc1");

        WriteInfo(doc.Metadata, sb);

        // 페이지 설정 — 첫 번째 섹션의 PageSettings 를 문서 레벨 \paperw/\paperh/\margXX 로 출력.
        if (doc.Sections.Count > 0)
            WriteDocumentPageSetup(doc.Sections[0].Page, sb);

        for (int i = 0; i < doc.Sections.Count; i++)
        {
            var sec = doc.Sections[i];

            // 섹션 속성 — 머리말/꼬리말은 \sectd 다음, 본문 단락보다 앞.
            if (i > 0) sb.Append(@"\sect\sectd").AppendLine();
            WriteSectionProperties(sec.Page, sb);
            WriteHeaderFooter(sec.Page, sb);

            foreach (var blk in sec.Blocks)
                WriteBlock(blk, sb, inTable: false);
        }

        sb.Append('}');

        using var sw = new StreamWriter(output, Encoding.Default, leaveOpen: true);
        sw.Write(sb.ToString());
    }

    // ── 폰트·색상 테이블 스캔 ──────────────────────────────────────────────────

    private void ScanBlock(Block block)
    {
        switch (block)
        {
            case Paragraph p:
                foreach (var r in p.Runs)
                {
                    ScanRunStyle(r.Style, p.Style);
                    if (!string.IsNullOrEmpty(r.RevisionAuthor))
                        RegisterRevAuthor(r.RevisionAuthor!);
                }
                break;
            case Table t:
                foreach (var row in t.Rows)
                    foreach (var cell in row.Cells)
                        foreach (var b in cell.Blocks) ScanBlock(b);
                break;
            case ShapeObject s:
                if (!string.IsNullOrEmpty(s.FillColor))
                    try { RegisterColor(Color.FromHex(s.FillColor)); } catch { }
                if (!string.IsNullOrEmpty(s.StrokeColor))
                    try { RegisterColor(Color.FromHex(s.StrokeColor)); } catch { }
                break;
            case TextBoxObject tb:
                if (!string.IsNullOrEmpty(tb.BorderColor))
                    try { RegisterColor(Color.FromHex(tb.BorderColor)); } catch { }
                if (!string.IsNullOrEmpty(tb.BackgroundColor))
                    try { RegisterColor(Color.FromHex(tb.BackgroundColor)); } catch { }
                foreach (var b in tb.Content) ScanBlock(b);
                break;
            case ContainerBlock c:
                foreach (var b in c.Children) ScanBlock(b);
                break;
        }
    }

    private int RegisterRevAuthor(string name)
    {
        int i = _revAuthors.IndexOf(name);
        if (i >= 0) return i;
        _revAuthors.Add(name);
        return _revAuthors.Count - 1;
    }

    private void ScanRunStyle(RunStyle? rs, ParagraphStyle? ps)
    {
        if (rs is null) return;
        RegisterFont(!string.IsNullOrEmpty(rs.FontFamily) ? rs.FontFamily! : DefaultFont);
        if (rs.Foreground.HasValue) RegisterColor(rs.Foreground.Value);
        if (rs.Background.HasValue) { RegisterColor(rs.Background.Value); return; }
        if (ps?.BackgroundColor is { Length: > 0 } hex)
        {
            try { RegisterColor(Color.FromHex(hex)); } catch { }
        }
    }

    private int RegisterFont(string name)
    {
        int i = _fonts.IndexOf(name);
        if (i >= 0) return i;
        _fonts.Add(name);
        return _fonts.Count - 1;
    }

    private int RegisterColor(Color c)
    {
        for (int i = 0; i < _colors.Count; i++)
            if (_colors[i].R == c.R && _colors[i].G == c.G && _colors[i].B == c.B) return i;
        _colors.Add(new RtfColor(c.R, c.G, c.B));
        return _colors.Count - 1;
    }

    // ── RTF 헤더 ────────────────────────────────────────────────────────────────

    private void WriteFontTable(StringBuilder sb)
    {
        sb.Append(@"{\fonttbl");
        for (int i = 0; i < _fonts.Count; i++)
            sb.Append($@"{{\f{i}\fnil\fcharset0 {_fonts[i]};}}");
        sb.AppendLine("}");
    }

    private void WriteColorTable(StringBuilder sb)
    {
        sb.Append(@"{\colortbl");
        foreach (var c in _colors)
            sb.Append($@"\red{c.R}\green{c.G}\blue{c.B};");
        sb.AppendLine("}");
    }

    /// <summary>변경추적 작성자 테이블. <c>\revauthN</c> / <c>\revauthdelN</c> 가 인덱스로 참조.</summary>
    private void WriteRevAuthorTable(StringBuilder sb)
    {
        if (_revAuthors.Count <= 1) return;  // Unknown 한 명뿐이면 생략
        sb.Append(@"{\*\revtbl");
        foreach (var name in _revAuthors)
            sb.Append('{').Append(EscapeRtf(name)).Append(";}");
        sb.AppendLine("}");
    }

    /// <summary>DateTimeOffset → Word DTTM 4-byte packed (\revdttm / \atndate 가 사용).</summary>
    private static int PackDttm(DateTimeOffset d)
    {
        int min  = d.Minute & 0x3F;
        int hour = d.Hour   & 0x1F;
        int day  = d.Day    & 0x1F;
        int mon  = d.Month  & 0x0F;
        int year = (d.Year - 1900) & 0x1FF;
        int wd   = (int)d.DayOfWeek & 0x07;
        return min | (hour << 6) | (day << 11) | (mon << 16) | (year << 20) | (wd << 29);
    }

    private static void WriteInfo(DocumentMetadata meta, StringBuilder sb)
    {
        sb.Append(@"{\info");
        if (!string.IsNullOrEmpty(meta.Title))
            sb.Append($@"{{\title {EscapeRtf(meta.Title)}}}");
        if (!string.IsNullOrEmpty(meta.Author))
            sb.Append($@"{{\author {EscapeRtf(meta.Author)}}}");
        if (!string.IsNullOrEmpty(meta.Application))
            sb.Append($@"{{\operator {EscapeRtf(meta.Application)}}}");

        var cr = meta.Created;
        sb.Append($@"{{\creatim\yr{cr.Year}\mo{cr.Month}\dy{cr.Day}\hr{cr.Hour}\min{cr.Minute}\sec{cr.Second}}}");
        var mo = meta.Modified;
        sb.Append($@"{{\revtim\yr{mo.Year}\mo{mo.Month}\dy{mo.Day}\hr{mo.Hour}\min{mo.Minute}\sec{mo.Second}}}");
        sb.AppendLine("}");
    }

    // ── 페이지 설정 / 섹션 속성 ────────────────────────────────────────────────

    /// <summary>문서 전체 페이지 크기·여백을 emit (\paperw, \paperh, \margl/r/t/b, \landscape).</summary>
    private static void WriteDocumentPageSetup(PageSettings page, StringBuilder sb)
    {
        // PageOrientation 에 따라 자동 swap — EffectiveWidth/Height 사용
        int paperW = T(page.EffectiveWidthMm);
        int paperH = T(page.EffectiveHeightMm);
        sb.Append($@"\paperw{paperW}\paperh{paperH}");
        sb.Append($@"\margl{T(page.MarginLeftMm)}\margr{T(page.MarginRightMm)}");
        sb.Append($@"\margt{T(page.MarginTopMm)}\margb{T(page.MarginBottomMm)}");
        if (page.MarginHeaderMm > 0) sb.Append($@"\headery{T(page.MarginHeaderMm)}");
        if (page.MarginFooterMm > 0) sb.Append($@"\footery{T(page.MarginFooterMm)}");
        if (page.Orientation == PageOrientation.Landscape) sb.Append(@"\landscape");
        if (page.DifferentOddEven) sb.Append(@"\facingp");
        sb.AppendLine();
    }

    /// <summary>섹션 속성 — \sectd 다음에 페이지 설정·다단·첫페이지/홀짝 다름 플래그·헤더/푸터 거리 emit.</summary>
    private static void WriteSectionProperties(PageSettings page, StringBuilder sb)
    {
        sb.Append(@"\sectd");
        // 섹션마다 페이지 크기·여백 반복 — Word 의 multi-section 호환을 위해
        sb.Append($@"\pgwsxn{T(page.EffectiveWidthMm)}\pghsxn{T(page.EffectiveHeightMm)}");
        sb.Append($@"\marglsxn{T(page.MarginLeftMm)}\margrsxn{T(page.MarginRightMm)}");
        sb.Append($@"\margtsxn{T(page.MarginTopMm)}\margbsxn{T(page.MarginBottomMm)}");
        if (page.MarginHeaderMm > 0) sb.Append($@"\headery{T(page.MarginHeaderMm)}");
        if (page.MarginFooterMm > 0) sb.Append($@"\footery{T(page.MarginFooterMm)}");
        if (page.Orientation == PageOrientation.Landscape) sb.Append(@"\lndscpsxn");
        if (page.DifferentFirstPage) sb.Append(@"\titlepg");
        if (page.ColumnCount > 1)
        {
            sb.Append($@"\cols{page.ColumnCount}");
            if (page.ColumnGapMm > 0) sb.Append($@"\colsx{T(page.ColumnGapMm)}");
        }
        sb.AppendLine();
    }

    // ── 머리말 / 꼬리말 ────────────────────────────────────────────────────────

    /// <summary>섹션의 머리말·꼬리말을 RTF 그룹으로 emit. Left/Center/Right 슬롯은 정렬 단락으로 stacked 출력.
    /// 향후 개선: 단일 단락만 들어 있을 때 tab-stop 기반 single-line 레이아웃으로 합칠 수 있음.</summary>
    private void WriteHeaderFooter(PageSettings page, StringBuilder sb)
    {
        if (!page.Header.IsEmpty) WriteHeaderFooterGroup("header", page.Header, sb);
        if (!page.Footer.IsEmpty) WriteHeaderFooterGroup("footer", page.Footer, sb);
    }

    private void WriteHeaderFooterGroup(string kind, HeaderFooterContent content, StringBuilder sb)
    {
        sb.Append('{').Append('\\').Append(kind).Append(' ');

        WriteSlotParagraphs(content.Left,   Alignment.Left,   sb);
        WriteSlotParagraphs(content.Center, Alignment.Center, sb);
        WriteSlotParagraphs(content.Right,  Alignment.Right,  sb);

        sb.Append('}').AppendLine();
    }

    private void WriteSlotParagraphs(HeaderFooterSlot slot, Alignment align, StringBuilder sb)
    {
        if (slot.IsEmpty) return;
        foreach (var src in slot.Paragraphs)
        {
            var p = src.Clone();
            p.Style.Alignment = align;  // 슬롯 정렬을 단락 정렬보다 우선
            // 머리말/꼬리말 텍스트 안의 {PAGE} 등 토큰을 RTF 필드 Run 으로 분해.
            ExpandHeaderFooterTokens(p);
            WriteParagraph(p, sb, inTable: false);
        }
    }

    /// <summary>머리말/꼬리말 Run.Text 안의 <c>{PAGE}</c>/<c>{페이지}</c> 등 토큰을 RTF 필드로 변환한다.
    /// HeaderFooterTokens 와 같은 토큰 집합을 인식 — 토큰 Run 은 <see cref="Run.Field"/> 로,
    /// 일반 텍스트는 그대로 둔다. <c>{PAGE}</c> 가 리터럴로 찍히던 회귀를 방지.</summary>
    private static void ExpandHeaderFooterTokens(Paragraph p)
    {
        var newRuns = new List<Run>();
        foreach (var run in p.Runs)
        {
            if (string.IsNullOrEmpty(run.Text) || run.Text.IndexOf('{') < 0)
            {
                newRuns.Add(run);
                continue;
            }
            SplitTokenRuns(run, newRuns);
        }
        p.Runs = newRuns;
    }

    private static void SplitTokenRuns(Run run, List<Run> output)
    {
        string text = run.Text;
        int i = 0;
        var literal = new StringBuilder();

        void FlushLiteral()
        {
            if (literal.Length == 0) return;
            var r = run.Clone();
            r.Text = literal.ToString();
            r.Field = null; r.FieldArg = null;
            output.Add(r);
            literal.Clear();
        }

        while (i < text.Length)
        {
            if (text[i] == '{')
            {
                int end = text.IndexOf('}', i + 1);
                if (end > i + 1)
                {
                    var name = text.Substring(i + 1, end - i - 1).Trim();
                    if (TryMapHeaderFooterToken(name, out var fieldType))
                    {
                        FlushLiteral();
                        var fr = run.Clone();
                        fr.Text = string.Empty;   // 결과는 Word 가 채움
                        fr.Field = fieldType;
                        fr.FieldArg = null;
                        output.Add(fr);
                        i = end + 1;
                        continue;
                    }
                }
            }
            literal.Append(text[i]);
            i++;
        }
        FlushLiteral();
    }

    /// <summary>HeaderFooterTokens 와 동일한 토큰명 → FieldType 매핑 (영문 + 한국어 별칭).</summary>
    private static bool TryMapHeaderFooterToken(string name, out FieldType fieldType)
    {
        switch (name.ToUpperInvariant())
        {
            case "PAGE":     case "페이지":     fieldType = FieldType.Page;     return true;
            case "NUMPAGES": case "전체페이지": fieldType = FieldType.NumPages; return true;
            case "DATE":     case "날짜":       fieldType = FieldType.Date;     return true;
            case "TIME":     case "시간":       fieldType = FieldType.Time;     return true;
            case "TITLE":    case "제목":       fieldType = FieldType.Title;    return true;
            case "AUTHOR":   case "저자":       fieldType = FieldType.Author;   return true;
            case "FILENAME": case "파일명":     fieldType = FieldType.FileName; return true;
            default:         fieldType = default; return false;
        }
    }

    // ── 블록 디스패치 ───────────────────────────────────────────────────────────

    private void WriteBlock(Block block, StringBuilder sb, bool inTable)
    {
        switch (block)
        {
            case Paragraph p:      WriteParagraph(p, sb, inTable); break;
            case Table t:          WriteTable(t, sb); break;
            case ImageBlock img:   WriteImage(img, sb); break;
            case ShapeObject s:    WriteShape(s, sb); break;
            case TextBoxObject tb: WriteTextBox(tb, sb); break;
            case OpaqueBlock o:    WriteOpaque(o, sb); break;
            case ContainerBlock c:
                foreach (var b in c.Children) WriteBlock(b, sb, inTable);
                break;
        }
    }

    // ── 글상자 (TextBox) ─────────────────────────────────────────────────────────

    /// <summary>글상자 → 위치 지정 프레임 + 4면 테두리 단락.
    /// RTF Office Drawing shape(\shptxt)는 Word 가 까다롭게 처리해 내부 텍스트가 안 보이는 경우가 많다.
    /// 대신 Word 가 오래전부터 안정적으로 지원하는 "positioned frame"(\posx/\posy/\absw/\absh) +
    /// 단락 테두리(\brdrt/l/b/r)로 글상자를 표현한다 — 텍스트가 단락 안에 직접 들어가 항상 렌더링된다.</summary>
    private void WriteTextBox(TextBoxObject tb, StringBuilder sb)
    {
        int x = T(tb.OverlayXMm);
        int y = T(tb.OverlayYMm);
        int w = T(Math.Max(1, tb.WidthMm));
        int h = T(Math.Max(1, tb.HeightMm));

        // 4면 테두리 (border) — \brdrwN 은 twips 단위.
        string border = string.Empty;
        if (tb.BorderThicknessPt > 0)
        {
            int bw = Math.Max(10, (int)(tb.BorderThicknessPt * 20));
            int ci = 0;
            if (!string.IsNullOrEmpty(tb.BorderColor))
                try { ci = RegisterColor(Color.FromHex(tb.BorderColor)); } catch { }
            string side = $@"\brdrs\brdrw{bw}\brdrcf{ci}";
            border = $@"\brdrt{side}\brdrl{side}\brdrb{side}\brdrr{side} ";
        }

        // 배경색 (\cbpat = paragraph shading background color index)
        string shading = string.Empty;
        if (!string.IsNullOrEmpty(tb.BackgroundColor))
            try { shading = $@"\cbpat{RegisterColor(Color.FromHex(tb.BackgroundColor))} "; } catch { }

        string align = tb.HAlign switch
        {
            TextBoxHAlign.Center  => @"\qc",
            TextBoxHAlign.Right   => @"\qr",
            TextBoxHAlign.Justify => @"\qj",
            _                     => @"\ql",
        };

        var paras = tb.Content.OfType<Paragraph>().ToList();
        if (paras.Count == 0) paras.Add(new Paragraph());

        // 동일한 frame 좌표를 가진 연속 단락은 Word 가 한 프레임으로 병합한다.
        foreach (var src in paras)
        {
            var p = src.Clone();
            var ps = p.Style ?? new ParagraphStyle();
            sb.Append($@"{{\pard\phpg\pvpg\posx{x}\posy{y}\absw{w}\absh{h}\dxfrtext0\dfrmtxtx0\dfrmtxty0 ");
            sb.Append(border);
            sb.Append(shading);
            sb.Append(align);
            foreach (var run in p.Runs) WriteRun(run, ps, sb);
            sb.Append(@"\par}");
            sb.AppendLine();
        }
    }

    // ── 단락 ────────────────────────────────────────────────────────────────────

    private void WriteParagraph(Paragraph para, StringBuilder sb, bool inTable)
    {
        var ps = para.Style ?? new ParagraphStyle();

        sb.Append(@"\pard");
        if (inTable) sb.Append(@"\intbl");

        // 정렬
        sb.Append(ps.Alignment switch
        {
            Alignment.Center  => @"\qc",
            Alignment.Right   => @"\qr",
            Alignment.Justify => @"\qj",
            _                 => @"\ql",
        });

        // 들여쓰기 (twips)
        if (ps.IndentLeftMm  > 0) sb.Append($@"\li{T(ps.IndentLeftMm)}");
        if (ps.IndentRightMm > 0) sb.Append($@"\ri{T(ps.IndentRightMm)}");
        if (ps.IndentFirstLineMm != 0) sb.Append($@"\fi{T(ps.IndentFirstLineMm)}");

        // 문단 간격 (twips)
        if (ps.SpaceBeforePt > 0) sb.Append($@"\sb{(int)(ps.SpaceBeforePt * PtToTwips)}");
        if (ps.SpaceAfterPt  > 0) sb.Append($@"\sa{(int)(ps.SpaceAfterPt  * PtToTwips)}");

        // 줄 간격
        if (ps.LineHeightFactor > 0)
            sb.Append($@"\sl{(int)(ps.LineHeightFactor * 240)}\slmult1");

        // 글머리 기호·번호 (간단 구현 — 들여쓰기 + 마커 Run 선행)
        if (ps.ListMarker is { } lm)
            WriteListPreamble(lm, ps, sb);

        // Run들
        foreach (var run in para.Runs)
            WriteRun(run, ps, sb);

        // 단락 마크(\r) 자체의 변경추적 — Phase 3h-3 와 짝.
        if (para.IsInsertedRevision)
            sb.Append(@"\revised\revauth0");
        else if (para.IsDeletedRevision)
            sb.Append(@"\deleted\revauthdel0");

        sb.Append(inTable ? @"\cell" : @"\par");
        sb.AppendLine();
    }

    private void WriteListPreamble(ListMarker lm, ParagraphStyle ps, StringBuilder sb)
    {
        // 레벨별 들여쓰기 (기본값이 없으면 보강)
        int level = Math.Max(0, lm.Level);
        int liTwips = T(ps.IndentLeftMm > 0 ? ps.IndentLeftMm : (level + 1) * 6.35);
        int fiTwips = -T(3.0);  // 내어쓰기

        sb.Append($@"\li{liTwips}\fi{fiTwips}");
    }

    private void WriteRun(Run run, ParagraphStyle ps, StringBuilder sb)
    {
        bool hasText       = !string.IsNullOrEmpty(run.Text);
        bool hasBkmkStart  = !string.IsNullOrEmpty(run.BookmarkStart);
        bool hasBkmkEnd    = !string.IsNullOrEmpty(run.BookmarkEnd);
        bool isHyperlink   = !string.IsNullOrEmpty(run.Url);
        var  fieldType     = run.Field;
        bool hasFootnote   = !string.IsNullOrEmpty(run.FootnoteId);
        bool hasEndnote    = !string.IsNullOrEmpty(run.EndnoteId);
        bool hasComment    = !string.IsNullOrEmpty(run.CommentId);
        bool isRevised     = run.IsInsertedRevision || run.IsDeletedRevision;

        if (!hasText && !hasBkmkStart && !hasBkmkEnd && !isHyperlink && fieldType is null
            && !hasFootnote && !hasEndnote && !hasComment)
            return;

        if (hasBkmkStart)
            sb.Append($@"{{\*\bkmkstart {SanitizeBookmarkName(run.BookmarkStart!)}}}");

        // 변경추적 wrap — { \revised\revauthN [\revdttmDTTM] <body> } 또는 { \deleted... }
        bool openedRevGroup = false;
        if (isRevised)
        {
            int authIdx = string.IsNullOrEmpty(run.RevisionAuthor)
                ? 0 : RegisterRevAuthor(run.RevisionAuthor!);
            string revToken  = run.IsInsertedRevision ? "revised"   : "deleted";
            string authToken = run.IsInsertedRevision ? "revauth"   : "revauthdel";
            sb.Append($@"{{\{revToken}\{authToken}{authIdx}");
            if (run.RevisionDate is { } dt)
            {
                string dttmToken = run.IsInsertedRevision ? "revdttm" : "revdttmdel";
                sb.Append($@"\{dttmToken}{PackDttm(dt)}");
            }
            sb.Append(' ');
            openedRevGroup = true;
        }

        if (hasFootnote && !_inFootnoteOrComment)
            WriteFootnoteRef(run.FootnoteId!, isEndnote: false, sb);
        else if (hasEndnote && !_inFootnoteOrComment)
            WriteFootnoteRef(run.EndnoteId!, isEndnote: true, sb);
        else if (hasComment && !_inFootnoteOrComment)
            WriteCommentRef(run.CommentId!, sb);
        else if (isHyperlink || fieldType is not null)
            WriteFieldRun(run, ps, sb, isHyperlink, fieldType);
        else if (hasText)
            WriteStyledRunBody(run, ps, sb);

        if (openedRevGroup) sb.Append('}');

        if (hasBkmkEnd)
            sb.Append($@"{{\*\bkmkend {SanitizeBookmarkName(run.BookmarkEnd!)}}}");
    }

    /// <summary>각주/미주 참조 marker — `{\super\chftn {\*\footnote [\ftnalt] \chftn ...body... }}`.
    /// 본문 lookup 은 <see cref="_doc"/> 의 Footnotes/Endnotes 에서 Id 매칭. dangling ref 면 marker 만 emit.</summary>
    private void WriteFootnoteRef(string id, bool isEndnote, StringBuilder sb)
    {
        var list = isEndnote ? _doc?.Endnotes : _doc?.Footnotes;
        var entry = list?.FirstOrDefault(e => string.Equals(e.Id, id, StringComparison.Ordinal));

        sb.Append(@"{\super\chftn ");
        sb.Append(@"{\*\footnote ");
        if (isEndnote) sb.Append(@"\ftnalt ");
        sb.Append(@"\chftn ");

        if (entry is not null)
        {
            _inFootnoteOrComment = true;
            try
            {
                foreach (var b in entry.Blocks) WriteBlock(b, sb, inTable: false);
            }
            finally { _inFootnoteOrComment = false; }
        }

        sb.Append("}}");
        sb.Append(' ');
    }

    /// <summary>주석(annotation) 참조 marker — `{\*\atnid …}{\*\atnauthor …}\chatn{\*\annotation …}`.</summary>
    private void WriteCommentRef(string id, StringBuilder sb)
    {
        var entry = _doc?.Comments.FirstOrDefault(c => string.Equals(c.Id, id, StringComparison.Ordinal));

        if (entry?.Author is { Length: > 0 } author)
        {
            // \*\atnid 는 짧은 이니셜용이지만 풀네임으로 채워도 Word 가 허용
            sb.Append($@"{{\*\atnid {EscapeRtf(author)}}}");
            sb.Append($@"{{\*\atnauthor {EscapeRtf(author)}}}");
        }
        if (entry?.Date is { } dt)
            sb.Append($@"{{\*\atndate {PackDttm(dt)}}}");

        sb.Append(@"\chatn ");
        sb.Append(@"{\*\annotation \chatn ");

        if (entry is not null)
        {
            _inFootnoteOrComment = true;
            try
            {
                foreach (var b in entry.Blocks) WriteBlock(b, sb, inTable: false);
            }
            finally { _inFootnoteOrComment = false; }
        }

        sb.Append('}');
        sb.Append(' ');
    }

    private void WriteFieldRun(Run run, ParagraphStyle ps, StringBuilder sb, bool isHyperlink, FieldType? fieldType)
    {
        string fldinst = isHyperlink
            ? $"HYPERLINK \"{EscapeFieldArg(run.Url!)}\""
            : BuildFieldInstruction(fieldType!.Value, run.FieldArg);

        sb.Append(@"{\field{\*\fldinst ");
        sb.Append(fldinst);
        sb.Append(@"}{\fldrslt ");
        if (!string.IsNullOrEmpty(run.Text))
        {
            sb.Append('{');
            WriteStyledRunBody(run, ps, sb);
            sb.Append('}');
        }
        sb.Append("}}");
        sb.Append(' ');
    }

    private void WriteStyledRunBody(Run run, ParagraphStyle ps, StringBuilder sb)
    {
        var rs = run.Style ?? new RunStyle();

        // 글꼴
        int fi = RegisterFont(!string.IsNullOrEmpty(rs.FontFamily) ? rs.FontFamily! : DefaultFont);
        sb.Append($@"\f{fi}");

        // 전경색
        int ci = rs.Foreground.HasValue ? RegisterColor(rs.Foreground.Value) : 0;
        sb.Append($@"\cf{ci}");

        // 배경색
        Color? bg = rs.Background;
        if (!bg.HasValue && ps.BackgroundColor is { Length: > 0 } hex)
            try { bg = Color.FromHex(hex); } catch { }
        bool hasBg = bg.HasValue;
        if (hasBg) sb.Append($@"\cb{RegisterColor(bg!.Value)}");

        // 글자 크기 (half-point)
        double fsz = rs.FontSizePt > 0 ? rs.FontSizePt : 11;
        sb.Append($@"\fs{(int)(fsz * 2)}");

        // 서식 on
        if (rs.Bold)          sb.Append(@"\b");
        if (rs.Italic)        sb.Append(@"\i");
        if (rs.Underline)     sb.Append(@"\ul");
        if (rs.Strikethrough) sb.Append(@"\strike");
        if (rs.Superscript)   sb.Append(@"\super");
        if (rs.Subscript)     sb.Append(@"\sub");

        sb.Append(' ');
        sb.Append(EscapeRtf(run.Text));

        // 서식 off
        if (rs.Bold)          sb.Append(@"\b0");
        if (rs.Italic)        sb.Append(@"\i0");
        if (rs.Underline)     sb.Append(@"\ulnone");
        if (rs.Strikethrough) sb.Append(@"\strike0");
        if (rs.Superscript || rs.Subscript) sb.Append(@"\nosupersub");
        if (hasBg)            sb.Append(@"\cb0");

        sb.Append(' ');
    }

    /// <summary>FieldType → RTF \fldinst 명령문 헤더. Reader 의 ParseFieldInstr 와 역매핑.</summary>
    private static string BuildFieldInstruction(FieldType type, string? arg)
    {
        string head = type switch
        {
            FieldType.Page        => "PAGE",
            FieldType.NumPages    => "NUMPAGES",
            FieldType.Date        => "DATE",
            FieldType.Time        => "TIME",
            FieldType.Author      => "AUTHOR",
            FieldType.Title       => "TITLE",
            FieldType.NumChars    => "NUMCHARS",
            FieldType.FileName    => "FILENAME",
            FieldType.Subject     => "SUBJECT",
            FieldType.Keywords    => "KEYWORDS",
            FieldType.Comments    => "COMMENTS",
            FieldType.Seq         => "SEQ",
            FieldType.Ref         => "REF",
            FieldType.StyleRef    => "STYLEREF",
            FieldType.IncludeText => "INCLUDETEXT",
            FieldType.If          => "IF",
            _                     => "MERGEFIELD",
        };
        if (string.IsNullOrEmpty(arg)) return head;

        // STYLEREF / INCLUDETEXT 는 인자를 따옴표로 감싸는 게 표준
        bool quote = type is FieldType.StyleRef or FieldType.IncludeText;
        return quote ? $"{head} \"{EscapeFieldArg(arg)}\"" : $"{head} {EscapeFieldArg(arg)}";
    }

    private static string EscapeFieldArg(string s)
    {
        if (string.IsNullOrEmpty(s)) return s;
        var sb = new StringBuilder(s.Length + 8);
        foreach (char c in s)
        {
            switch (c)
            {
                case '\\': sb.Append(@"\\"); break;
                case '{':  sb.Append(@"\{"); break;
                case '}':  sb.Append(@"\}"); break;
                case '"':  sb.Append("'");   break;  // RTF \fldinst 인자에서 내부 따옴표 회피
                default:
                    if (c < 128) sb.Append(c);
                    else         sb.Append($@"\u{(short)c}?");
                    break;
            }
        }
        return sb.ToString();
    }

    /// <summary>책갈피 이름은 RTF 에서 공백/특수문자를 허용하지 않으므로 ASCII identifier 로 정리.</summary>
    private static string SanitizeBookmarkName(string name)
    {
        var sb = new StringBuilder(name.Length);
        foreach (char c in name)
        {
            if ((c >= 'A' && c <= 'Z') || (c >= 'a' && c <= 'z') ||
                (c >= '0' && c <= '9') || c == '_')
                sb.Append(c);
        }
        if (sb.Length == 0) sb.Append("bm");
        return sb.ToString();
    }

    // ── 표 ──────────────────────────────────────────────────────────────────────

    private void WriteTable(Table table, StringBuilder sb)
    {
        // 열 너비 계산 (twips). Columns 목록이 없으면 균등 분배.
        var colWidths = BuildColWidths(table);
        int colCount  = colWidths.Count;

        for (int ri = 0; ri < table.Rows.Count; ri++)
        {
            var row = table.Rows[ri];

            // trowd — 행 정의
            sb.Append(@"\trowd");

            // 표 정렬
            sb.Append(table.HAlign switch
            {
                TableHAlign.Center => @"\trqc",
                TableHAlign.Right  => @"\trqr",
                _                  => @"\trql",
            });

            if (row.HeightMm > 0)
                sb.Append($@"\trrh{T(row.HeightMm)}");

            // 셀 경계 정의
            int cumWidth = 0;
            for (int ci = 0; ci < row.Cells.Count && ci < colCount; ci++)
            {
                var cell = row.Cells[ci];
                int span = Math.Max(1, cell.ColumnSpan);
                int cw   = 0;
                for (int k = ci; k < ci + span && k < colWidths.Count; k++) cw += colWidths[k];
                cumWidth += cw;

                WriteCellDef(cell, table, cumWidth, sb);
            }
            sb.AppendLine();

            // 셀 콘텐츠
            for (int ci = 0; ci < row.Cells.Count; ci++)
            {
                var cell = row.Cells[ci];
                if (cell.Blocks.Count == 0)
                {
                    // 빈 셀
                    sb.Append(@"\pard\intbl\ql ");
                    sb.AppendLine(@"\cell");
                }
                else
                {
                    foreach (var blk in cell.Blocks)
                        WriteBlock(blk, sb, inTable: true);
                }
            }

            sb.AppendLine(@"\row\pard");
        }
    }

    private void WriteCellDef(TableCell cell, Table table, int rightEdgeTwips, StringBuilder sb)
    {
        // 수직 정렬
        sb.Append(cell.VerticalAlign switch
        {
            CellVerticalAlign.Middle => @"\clvertalc",
            CellVerticalAlign.Bottom => @"\clvertalb",
            _                        => @"\clvertalt",
        });

        // 배경색
        if (!string.IsNullOrEmpty(cell.BackgroundColor))
        {
            try
            {
                var bg = Color.FromHex(cell.BackgroundColor);
                sb.Append($@"\clcbpat{RegisterColor(bg)}");
            }
            catch { }
        }

        // 셀 테두리
        WriteCellBorder(@"\clbrdrt", cell.BorderTop   ?? table.BorderTop,   sb);
        WriteCellBorder(@"\clbrdrb", cell.BorderBottom?? table.BorderBottom, sb);
        WriteCellBorder(@"\clbrdrl", cell.BorderLeft  ?? table.BorderLeft,   sb);
        WriteCellBorder(@"\clbrdrr", cell.BorderRight ?? table.BorderRight,  sb);

        // 패딩 (twips)
        double pt = cell.PaddingTopMm    > 0 ? cell.PaddingTopMm    : table.DefaultCellPaddingTopMm;
        double pb = cell.PaddingBottomMm > 0 ? cell.PaddingBottomMm : table.DefaultCellPaddingBottomMm;
        double pl = cell.PaddingLeftMm   > 0 ? cell.PaddingLeftMm   : table.DefaultCellPaddingLeftMm;
        double pr = cell.PaddingRightMm  > 0 ? cell.PaddingRightMm  : table.DefaultCellPaddingRightMm;
        if (pt > 0) sb.Append($@"\clpadft3\clpadt{T(pt)}");
        if (pb > 0) sb.Append($@"\clpadfb3\clpadb{T(pb)}");
        if (pl > 0) sb.Append($@"\clpadfl3\clpadl{T(pl)}");
        if (pr > 0) sb.Append($@"\clpadfr3\clpadr{T(pr)}");

        sb.Append($@"\cellx{rightEdgeTwips}");
    }

    private static void WriteCellBorder(string rtfKey, CellBorderSide? side, StringBuilder sb)
    {
        if (side is null) return;
        sb.Append(rtfKey);
        sb.Append(side.Value.LineStyle switch
        {
            BorderLineStyle.Dashed  => @"\brdrdash",
            BorderLineStyle.Dotted  => @"\brdrdot",
            BorderLineStyle.Double  => @"\brdrdb",
            _                       => @"\brdrs",
        });
        int w = (int)(side.Value.ThicknessPt * PtToTwips / 10); // brdrw는 twips/10
        if (w > 0) sb.Append($@"\brdrw{w}");
        if (!string.IsNullOrEmpty(side.Value.Color))
        {
            try
            { /* color index 등록 불가 (이미 scan 완료 전) — 색은 생략 */ }
            catch { }
        }
    }

    private static List<int> BuildColWidths(Table table)
    {
        if (table.Columns.Count > 0)
            return table.Columns.Select(c => T(c.WidthMm > 0 ? c.WidthMm : 30.0)).ToList();

        // 행에서 열 수 추론
        int cols = table.Rows.Count > 0
            ? table.Rows.Max(r => r.Cells.Sum(c => Math.Max(1, c.ColumnSpan)))
            : 1;
        double eachMm = table.WidthMm > 0 ? table.WidthMm / cols : 160.0 / cols;
        return Enumerable.Repeat(T(eachMm), cols).ToList();
    }

    // ── 이미지 ──────────────────────────────────────────────────────────────────

    private static void WriteImage(ImageBlock img, StringBuilder sb)
    {
        if (img.Data is not { Length: > 0 }) return;

        bool floating = img.WrapMode is ImageWrapMode.InFrontOfText or ImageWrapMode.BehindText
                        && (img.OverlayXMm > 0 || img.OverlayYMm > 0);

        if (floating)
        {
            // 부유 이미지 — 페이지 절대 위치 frame 단락 안에 \pict 배치.
            //   \phpg/\pvpg = 페이지 기준, \posx/\posy = 절대 위치(twips), \absw/\absh = frame 크기.
            int x = T(img.OverlayXMm), y = T(img.OverlayYMm);
            int w = img.WidthMm  > 0 ? T(img.WidthMm)  : 5040;
            int h = img.HeightMm > 0 ? T(img.HeightMm) : 3780;
            sb.Append($@"{{\pard\phpg\pvpg\posx{x}\posy{y}\absw{w}\absh{h}\dxfrtext0 ");
            WritePictBody(img, sb);
            sb.Append(@"\par}");
            sb.AppendLine();
            return;
        }

        WritePictBody(img, sb);
    }

    private static void WritePictBody(ImageBlock img, StringBuilder sb)
    {
        bool isPng = img.MediaType?.Contains("png",  StringComparison.OrdinalIgnoreCase) == true;
        bool isJpg = img.MediaType?.Contains("jpeg", StringComparison.OrdinalIgnoreCase) == true
                  || img.MediaType?.Contains("jpg",  StringComparison.OrdinalIgnoreCase) == true;
        bool isBmp = img.MediaType?.Contains("bmp",  StringComparison.OrdinalIgnoreCase) == true;

        string blipTag = isPng ? @"\pngblip"
                       : isJpg ? @"\jpegblip"
                       : isBmp ? @"\dibitmap0"
                       :         @"\pngblip";  // 폴백

        int picwTwips = img.WidthMm  > 0 ? T(img.WidthMm)  : 5040;
        int pichTwips = img.HeightMm > 0 ? T(img.HeightMm) : 3780;

        sb.Append($@"{{\pict{blipTag}\picwgoal{picwTwips}\pichgoal{pichTwips}");
        sb.AppendLine();
        // hex 인코딩 (한 줄에 64바이트씩)
        var hex = System.Convert.ToHexString(img.Data);
        for (int i = 0; i < hex.Length; i += 128)
        {
            sb.AppendLine(hex.Substring(i, Math.Min(128, hex.Length - i)));
        }
        sb.AppendLine("}");
    }

    // ── 도형 ────────────────────────────────────────────────────────────────────

    private void WriteShape(ShapeObject shape, StringBuilder sb)
    {
        int left   = T(shape.OverlayXMm);
        int top    = T(shape.OverlayYMm);
        int right  = T(shape.OverlayXMm + Math.Max(1, shape.WidthMm));
        int bottom = T(shape.OverlayYMm + Math.Max(1, shape.HeightMm));

        int shapeType = shape.Kind switch
        {
            ShapeKind.Rectangle   => 1,
            ShapeKind.RoundedRect => 2,
            ShapeKind.Ellipse     => 3,
            ShapeKind.Triangle    => 5,
            ShapeKind.Star        => 75,
            ShapeKind.Line        => 20,
            ShapeKind.Polyline    => 20,
            _                     => 1,
        };

        sb.Append(@"{\shp");
        sb.Append($@"\shpleft{left}\shptop{top}\shpright{right}\shpbottom{bottom}");
        // OverlayXMm/OverlayYMm 는 페이지 기준 좌표이므로 bx/by 모두 page-relative.
        sb.Append(@"\shpfhdr0\shpbxpage\shpbypage\shpwr3");
        sb.Append(@"{\*\shpinst");

        // 도형 종류
        sb.Append($@"{{\sp{{\sn shapeType}}{{\sv {shapeType}}}}}");

        // 채우기 색상 (BGR int)
        if (!string.IsNullOrEmpty(shape.FillColor))
        {
            try
            {
                var c = Color.FromHex(shape.FillColor);
                int abgr = c.R | (c.G << 8) | (c.B << 16);
                sb.Append($@"{{\sp{{\sn fillColor}}{{\sv {abgr}}}}}");
            }
            catch { }
        }

        // 선 색상
        if (!string.IsNullOrEmpty(shape.StrokeColor))
        {
            try
            {
                var c = Color.FromHex(shape.StrokeColor);
                int abgr = c.R | (c.G << 8) | (c.B << 16);
                sb.Append($@"{{\sp{{\sn lineColor}}{{\sv {abgr}}}}}");
            }
            catch { }
        }

        // 선 두께 (pt → EMU: 1pt = 12700 EMU)
        if (shape.StrokeThicknessPt > 0)
        {
            int emu = (int)(shape.StrokeThicknessPt * 12700);
            sb.Append($@"{{\sp{{\sn lineWidth}}{{\sv {emu}}}}}");
        }

        sb.Append("}}");
        sb.AppendLine();
    }

    // ── OpaqueBlock ─────────────────────────────────────────────────────────────

    private static void WriteOpaque(OpaqueBlock opaque, StringBuilder sb)
    {
        // RTF 포맷 OpaqueBlock — 원본 그대로 재출력
        if (opaque.Format == "rtf" && !string.IsNullOrEmpty(opaque.Xml))
        {
            sb.AppendLine(opaque.Xml);
            return;
        }
        // 다른 포맷 — 플레이스홀더 단락
        sb.Append(@"\pard\ql ");
        sb.Append($@"{{\b {EscapeRtf(opaque.DisplayLabel)}\b0}}");
        sb.AppendLine(@"\par");
    }

    // ── 유틸 ────────────────────────────────────────────────────────────────────

    private static int T(double mm) => (int)Math.Round(mm * MmToTwips);

    private static string EscapeRtf(string text)
    {
        if (string.IsNullOrEmpty(text)) return text;
        var sb = new StringBuilder(text.Length + 16);
        foreach (char c in text)
        {
            switch (c)
            {
                case '\\': sb.Append(@"\\");      break;
                case '{':  sb.Append(@"\{");      break;
                case '}':  sb.Append(@"\}");      break;
                case '\n': sb.Append(@"\line ");  break;
                case '\r':                        break;
                case '\t': sb.Append(@"\tab ");   break;
                default:
                    if (c < 128) sb.Append(c);
                    else         sb.Append($@"\u{(short)c}?");
                    break;
            }
        }
        return sb.ToString();
    }

    // ── 내부 타입 ────────────────────────────────────────────────────────────────

    private readonly record struct RtfColor(byte R, byte G, byte B);
}
