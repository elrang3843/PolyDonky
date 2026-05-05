using PolyDonky.Codecs.Html;
using PolyDonky.Core;

namespace PolyDonky.Codecs.Html.Tests;

public class HtmlTests
{
    // ── Reader ─────────────────────────────────────────────────────

    [Fact]
    public void Reader_HeadingsHaveOutlineLevel()
    {
        var doc = HtmlReader.FromHtml("<h1>A</h1><h2>B</h2><h6>F</h6>");
        var ps  = doc.EnumerateParagraphs().ToList();
        Assert.Equal(OutlineLevel.H1, ps[0].Style.Outline);
        Assert.Equal(OutlineLevel.H2, ps[1].Style.Outline);
        Assert.Equal(OutlineLevel.H6, ps[2].Style.Outline);
    }

    [Fact]
    public void Reader_BoldItalicStrikeAndSubSup()
    {
        var doc = HtmlReader.FromHtml(
            "<p><strong>b</strong> <em>i</em> <s>x</s> H<sub>2</sub>O <sup>2</sup></p>");
        var p   = doc.EnumerateParagraphs().Single();
        Assert.Contains(p.Runs, r => r.Style.Bold          && r.Text == "b");
        Assert.Contains(p.Runs, r => r.Style.Italic        && r.Text == "i");
        Assert.Contains(p.Runs, r => r.Style.Strikethrough && r.Text == "x");
        Assert.Contains(p.Runs, r => r.Style.Subscript     && r.Text == "2");
        Assert.Contains(p.Runs, r => r.Style.Superscript   && r.Text == "2");
    }

    [Fact]
    public void Reader_ListUlAndOl()
    {
        var doc = HtmlReader.FromHtml("<ul><li>a</li><li>b</li></ul><ol><li>1</li><li>2</li></ol>");
        var ps  = doc.EnumerateParagraphs().ToList();
        Assert.Equal(4, ps.Count);
        Assert.Equal(ListKind.Bullet,         ps[0].Style.ListMarker!.Kind);
        Assert.Equal(ListKind.Bullet,         ps[1].Style.ListMarker!.Kind);
        Assert.Equal(ListKind.OrderedDecimal, ps[2].Style.ListMarker!.Kind);
        Assert.Equal(2,                       ps[3].Style.ListMarker!.OrderedNumber);
    }

    [Fact]
    public void Reader_TaskListCheckbox()
    {
        var doc = HtmlReader.FromHtml(
            "<ul><li><input type=\"checkbox\" checked> done</li>" +
            "<li><input type=\"checkbox\"> todo</li></ul>");
        var ps  = doc.EnumerateParagraphs().ToList();
        Assert.True (ps[0].Style.ListMarker!.Checked);
        Assert.False(ps[1].Style.ListMarker!.Checked);
    }

    [Fact]
    public void Reader_BlockquoteNested()
    {
        var doc = HtmlReader.FromHtml("<blockquote><p>1</p><blockquote><p>2</p></blockquote></blockquote>");
        var ps  = doc.EnumerateParagraphs().ToList();
        Assert.Contains(ps, p => p.Style.QuoteLevel == 1 && p.GetPlainText() == "1");
        Assert.Contains(ps, p => p.Style.QuoteLevel == 2 && p.GetPlainText() == "2");
    }

    [Fact]
    public void Reader_HrIsThematicBreak()
    {
        var doc = HtmlReader.FromHtml("<p>before</p><hr><p>after</p>");
        var ps  = doc.EnumerateParagraphs().ToList();
        Assert.Contains(ps, p => p.Style.IsThematicBreak);
    }

    [Fact]
    public void Reader_PreCodeWithLanguageClass()
    {
        var doc = HtmlReader.FromHtml(
            "<pre><code class=\"language-python\">print('hi')</code></pre>");
        var p = doc.EnumerateParagraphs().Single();
        Assert.Equal("python", p.Style.CodeLanguage);
        Assert.Contains("print", p.GetPlainText());
    }

    [Fact]
    public void Reader_LinkStoresUrlAndUnderline()
    {
        var doc = HtmlReader.FromHtml("<p>방문 <a href=\"https://example.com\">사이트</a> 하기</p>");
        var p   = doc.EnumerateParagraphs().Single();
        Assert.Contains(p.Runs, r => r.Url == "https://example.com" && r.Style.Underline);
    }

    [Fact]
    public void Reader_TableHeaderAndCellAlignment()
    {
        var doc = HtmlReader.FromHtml(
            "<table>" +
            "<thead><tr><th>이름</th><th>나이</th></tr></thead>" +
            "<tbody><tr><td>홍길동</td><td style=\"text-align:right\">30</td></tr></tbody>" +
            "</table>");
        var t = doc.Sections[0].Blocks.OfType<PolyDonky.Core.Table>().Single();
        Assert.Equal(2, t.Rows.Count);
        Assert.True (t.Rows[0].IsHeader);
        Assert.Equal(CellTextAlign.Right, t.Rows[1].Cells[1].TextAlign);
    }

    [Fact]
    public void Reader_ImgBecomesImageBlock()
    {
        var doc = HtmlReader.FromHtml("<img src=\"img/a.png\" alt=\"대체\" width=\"200\" height=\"100\">");
        var img = doc.Sections[0].Blocks.OfType<ImageBlock>().Single();
        Assert.Equal("img/a.png", img.ResourcePath);
        Assert.Equal("대체", img.Description);
        Assert.Equal("image/png", img.MediaType);
        Assert.True(img.WidthMm > 0);
    }

    [Fact]
    public void Reader_FigureWithCaptionUsesImageTitle()
    {
        var doc = HtmlReader.FromHtml(
            "<figure><img src=\"x.png\" alt=\"a\"><figcaption>제목</figcaption></figure>");
        var img = doc.Sections[0].Blocks.OfType<ImageBlock>().Single();
        Assert.True(img.ShowTitle);
        Assert.Equal("제목", img.Title);
    }

    [Fact]
    public void Reader_DataUriImageDecodesBytes()
    {
        // 1x1 transparent PNG, base64.
        var pngB64 = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNkYAAAAAYAAjCB0C8AAAAASUVORK5CYII=";
        var doc = HtmlReader.FromHtml($"<img src=\"data:image/png;base64,{pngB64}\" alt=\"x\">");
        var img = doc.Sections[0].Blocks.OfType<ImageBlock>().Single();
        Assert.Equal("image/png", img.MediaType);
        Assert.NotEmpty(img.Data);
    }

    [Fact]
    public void Reader_SpanInlineStyleParsesColors()
    {
        var doc = HtmlReader.FromHtml(
            "<p><span style=\"color:#FF0000;background-color:rgb(0,255,0);font-weight:bold\">x</span></p>");
        var run = doc.EnumerateParagraphs().Single().Runs.Single(r => r.Text == "x");
        Assert.True(run.Style.Bold);
        Assert.Equal(new Color(0xFF, 0, 0),   run.Style.Foreground);
        Assert.Equal(new Color(0,    0xFF, 0), run.Style.Background);
    }

    [Fact]
    public void Reader_BrInsideParagraphYieldsNewline()
    {
        var doc = HtmlReader.FromHtml("<p>a<br>b</p>");
        var p   = doc.EnumerateParagraphs().Single();
        Assert.Contains(p.Runs, r => r.Text == "\n");
    }

    [Fact]
    public void Reader_CodeInlineUsesMonospace()
    {
        var doc = HtmlReader.FromHtml("<p>call <code>foo()</code></p>");
        var run = doc.EnumerateParagraphs().Single().Runs.Single(r => r.Text == "foo()");
        Assert.Contains("monospace", run.Style.FontFamily, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Reader_IgnoresScriptStyleHead()
    {
        var doc = HtmlReader.FromHtml(
            "<head><title>T</title><style>p{color:red}</style></head>" +
            "<body><script>alert('x')</script><p>본문</p></body>");
        var p = doc.EnumerateParagraphs().Single();
        Assert.Equal("본문", p.GetPlainText());
    }

    // ── Writer ─────────────────────────────────────────────────────

    [Fact]
    public void Writer_FullDocumentEmitsDoctypeAndHead()
    {
        var pd  = new PolyDonkyument();
        var sec = new Section(); pd.Sections.Add(sec);
        sec.Blocks.Add(Paragraph.Of("hi"));
        var html = HtmlWriter.ToHtml(pd);
        Assert.StartsWith("<!DOCTYPE html>", html);
        Assert.Contains("<meta charset=\"utf-8\">", html);
        Assert.Contains("<title>", html);
    }

    [Fact]
    public void Writer_HeadingTags()
    {
        var pd  = new PolyDonkyument();
        var sec = new Section(); pd.Sections.Add(sec);
        var h2 = new Paragraph { Style = { Outline = OutlineLevel.H2 } };
        h2.AddText("section");
        sec.Blocks.Add(h2);

        var html = HtmlWriter.ToHtml(pd, fullDocument: false);
        Assert.Contains("<h2>section</h2>", html);
    }

    [Fact]
    public void Writer_BoldItalicStrikeSubSup()
    {
        var pd  = new PolyDonkyument();
        var sec = new Section(); pd.Sections.Add(sec);
        var p = new Paragraph();
        p.AddText("b", new RunStyle { Bold          = true });
        p.AddText("i", new RunStyle { Italic        = true });
        p.AddText("s", new RunStyle { Strikethrough = true });
        p.AddText("u", new RunStyle { Underline     = true });
        p.AddText("x", new RunStyle { Subscript     = true });
        p.AddText("y", new RunStyle { Superscript   = true });
        sec.Blocks.Add(p);

        var html = HtmlWriter.ToHtml(pd, fullDocument: false);
        Assert.Contains("<strong>b</strong>", html);
        Assert.Contains("<em>i</em>", html);
        Assert.Contains("<s>s</s>", html);
        Assert.Contains("<u>u</u>", html);
        Assert.Contains("<sub>x</sub>", html);
        Assert.Contains("<sup>y</sup>", html);
    }

    [Fact]
    public void Writer_LinksAndCode()
    {
        var pd  = new PolyDonkyument();
        var sec = new Section(); pd.Sections.Add(sec);
        var p = new Paragraph();
        p.Runs.Add(new Run { Text = "site", Style = new RunStyle(), Url = "https://x" });
        p.AddText("foo()", new RunStyle { FontFamily = "Consolas, monospace" });
        sec.Blocks.Add(p);

        var html = HtmlWriter.ToHtml(pd, fullDocument: false);
        Assert.Contains("<a href=\"https://x\">site</a>", html);
        Assert.Contains("<code>foo()</code>", html);
    }

    [Fact]
    public void Writer_BulletAndOrderedList()
    {
        var pd  = new PolyDonkyument();
        var sec = new Section(); pd.Sections.Add(sec);
        var p1 = new Paragraph();
        p1.Style.ListMarker = new ListMarker { Kind = ListKind.Bullet };
        p1.AddText("a");
        var p2 = new Paragraph();
        p2.Style.ListMarker = new ListMarker { Kind = ListKind.OrderedDecimal, OrderedNumber = 1 };
        p2.AddText("1st");
        sec.Blocks.Add(p1); sec.Blocks.Add(p2);

        var html = HtmlWriter.ToHtml(pd, fullDocument: false);
        Assert.Contains("<ul>",  html);
        Assert.Contains("<li>a</li>", html);
        Assert.Contains("<ol>",  html);
        Assert.Contains("<li>1st</li>", html);
    }

    [Fact]
    public void Writer_TaskListEmitsCheckbox()
    {
        var pd  = new PolyDonkyument();
        var sec = new Section(); pd.Sections.Add(sec);
        var p = new Paragraph();
        p.Style.ListMarker = new ListMarker { Kind = ListKind.Bullet, Checked = true };
        p.AddText("done");
        sec.Blocks.Add(p);

        var html = HtmlWriter.ToHtml(pd, fullDocument: false);
        Assert.Contains("<input type=\"checkbox\" disabled checked>", html);
    }

    [Fact]
    public void Writer_BlockquoteNested()
    {
        var pd  = new PolyDonkyument();
        var sec = new Section(); pd.Sections.Add(sec);
        var p1 = new Paragraph { Style = { QuoteLevel = 1 } }; p1.AddText("외");
        var p2 = new Paragraph { Style = { QuoteLevel = 2 } }; p2.AddText("내");
        sec.Blocks.Add(p1); sec.Blocks.Add(p2);

        var html = HtmlWriter.ToHtml(pd, fullDocument: false);
        Assert.Contains("<blockquote>", html);
        // 2단 인용은 두 개의 <blockquote> 가 중첩되어야 함.
        Assert.Equal(2, System.Text.RegularExpressions.Regex.Matches(html, "<blockquote>").Count);
    }

    [Fact]
    public void Writer_PreCodeWithLanguageClass()
    {
        var pd  = new PolyDonkyument();
        var sec = new Section(); pd.Sections.Add(sec);
        var p = new Paragraph { Style = { CodeLanguage = "cs" } };
        p.AddText("var x = 1;");
        sec.Blocks.Add(p);

        var html = HtmlWriter.ToHtml(pd, fullDocument: false);
        Assert.Contains("<pre><code class=\"language-cs\">var x = 1;</code></pre>", html);
    }

    [Fact]
    public void Writer_TableHeaderBodyAlignment()
    {
        var pd  = new PolyDonkyument();
        var sec = new Section(); pd.Sections.Add(sec);
        var t = new PolyDonky.Core.Table();
        t.Columns.Add(new TableColumn());
        t.Columns.Add(new TableColumn());
        var hdr = new TableRow { IsHeader = true };
        var c1 = new TableCell(); c1.Blocks.Add(Paragraph.Of("H1"));
        var c2 = new TableCell { TextAlign = CellTextAlign.Right }; c2.Blocks.Add(Paragraph.Of("H2"));
        hdr.Cells.Add(c1); hdr.Cells.Add(c2);
        var body = new TableRow();
        var b1 = new TableCell(); b1.Blocks.Add(Paragraph.Of("a"));
        var b2 = new TableCell { TextAlign = CellTextAlign.Right }; b2.Blocks.Add(Paragraph.Of("b"));
        body.Cells.Add(b1); body.Cells.Add(b2);
        t.Rows.Add(hdr); t.Rows.Add(body);
        sec.Blocks.Add(t);

        var html = HtmlWriter.ToHtml(pd, fullDocument: false);
        Assert.Contains("<thead>", html);
        Assert.Contains("<tbody>", html);
        Assert.Contains("<th style=\"text-align:right\">H2</th>", html);
        Assert.Contains("<td style=\"text-align:right\">b</td>",  html);
    }

    [Fact]
    public void Writer_ImageEmitsImg()
    {
        var pd  = new PolyDonkyument();
        var sec = new Section(); pd.Sections.Add(sec);
        sec.Blocks.Add(new ImageBlock
        {
            Description  = "alt",
            ResourcePath = "img/foo.png",
            MediaType    = "image/png",
        });

        var html = HtmlWriter.ToHtml(pd, fullDocument: false);
        Assert.Contains("<img src=\"img/foo.png\" alt=\"alt\"", html);
    }

    [Fact]
    public void Writer_FigureWithCaption()
    {
        var pd  = new PolyDonkyument();
        var sec = new Section(); pd.Sections.Add(sec);
        sec.Blocks.Add(new ImageBlock
        {
            ResourcePath = "x.png",
            ShowTitle    = true,
            Title        = "캡션",
        });

        var html = HtmlWriter.ToHtml(pd, fullDocument: false);
        Assert.Contains("<figure>", html);
        Assert.Contains("<figcaption>캡션</figcaption>", html);
    }

    // ── 안전 한도 (대용량 HTML) ─────────────────────────────────────

    [Fact]
    public void Reader_BlockLimit_TruncatesAndAppendsWarning()
    {
        // 50,000 개의 단락이 있는 거대한 HTML — 기본 한도(10,000) 초과해야 함.
        var sb = new System.Text.StringBuilder("<html><body>");
        for (int i = 0; i < 50_000; i++) sb.Append("<p>x</p>");
        sb.Append("</body></html>");

        var doc = HtmlReader.FromHtml(sb.ToString());
        var ps  = doc.EnumerateParagraphs().ToList();

        // 한도 + 마지막 경고 단락 = 약 10,001.
        Assert.True(ps.Count <= 10_001, $"잘림 후 단락 수 {ps.Count} 가 한도(10,001)를 초과");
        Assert.Contains(ps, p => p.GetPlainText().Contains("잘림") || p.GetPlainText().Contains("초과"));
    }

    [Fact]
    public void Reader_CustomMaxBlocks_RespectsCap()
    {
        var sb = new System.Text.StringBuilder("<html><body>");
        for (int i = 0; i < 200; i++) sb.Append("<p>x</p>");
        sb.Append("</body></html>");

        var doc = HtmlReader.FromHtml(sb.ToString(), maxBlocks: 50);
        var ps  = doc.EnumerateParagraphs().ToList();
        Assert.True(ps.Count <= 51, $"단락 수 {ps.Count} 가 50+1 을 초과");
    }

    [Fact]
    public void Writer_EscapesAngleBracketsAndAmpersand()
    {
        var pd  = new PolyDonkyument();
        var sec = new Section(); pd.Sections.Add(sec);
        var p = new Paragraph();
        p.AddText("a < b & c > d");
        sec.Blocks.Add(p);

        var html = HtmlWriter.ToHtml(pd, fullDocument: false);
        Assert.Contains("a &lt; b &amp; c &gt; d", html);
    }

    // ── Round-trip ──────────────────────────────────────────────────

    [Fact]
    public void RoundTrip_PreservesStructure()
    {
        const string source =
            "<!DOCTYPE html><html><body>" +
            "<h1>제목</h1>" +
            "<p>본문 <strong>굵게</strong> <em>기울임</em> <a href=\"https://x\">링크</a></p>" +
            "<ul><li>항목 1</li><li>항목 2</li></ul>" +
            "<blockquote><p>인용</p></blockquote>" +
            "<pre><code class=\"language-py\">print(1)</code></pre>" +
            "<hr>" +
            "<table><thead><tr><th>A</th><th>B</th></tr></thead>" +
            "<tbody><tr><td>1</td><td>2</td></tr></tbody></table>" +
            "</body></html>";

        var doc      = HtmlReader.FromHtml(source);
        var rendered = HtmlWriter.ToHtml(doc, fullDocument: false);
        var reread   = HtmlReader.FromHtml(rendered);

        Assert.Equal(OutlineLevel.H1, reread.EnumerateParagraphs().First().Style.Outline);
        Assert.Contains(reread.EnumerateParagraphs(), p => p.Style.QuoteLevel >= 1);
        Assert.Contains(reread.EnumerateParagraphs(), p => p.Style.IsThematicBreak);
        Assert.Contains(reread.EnumerateParagraphs(), p => p.Style.CodeLanguage == "py");
        Assert.Contains(reread.EnumerateParagraphs(), p => p.Style.ListMarker?.Kind == ListKind.Bullet);
        Assert.Single(reread.Sections[0].Blocks.OfType<PolyDonky.Core.Table>());
        Assert.Contains(reread.EnumerateParagraphs().SelectMany(p => p.Runs),
            r => r.Url == "https://x");
    }
}
