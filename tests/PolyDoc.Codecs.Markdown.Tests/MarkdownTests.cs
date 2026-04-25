using PolyDoc.Codecs.Markdown;
using PolyDoc.Core;

namespace PolyDoc.Codecs.Markdown.Tests;

public class MarkdownTests
{
    [Fact]
    public void Reader_ExtractsAtxHeaderLevels()
    {
        const string source = "# H1\n\n## H2\n\n### H3\n";
        var doc = MarkdownReader.FromMarkdown(source);

        var paragraphs = doc.EnumerateParagraphs().ToList();
        Assert.Equal(3, paragraphs.Count);
        Assert.Equal(OutlineLevel.H1, paragraphs[0].Style.Outline);
        Assert.Equal(OutlineLevel.H2, paragraphs[1].Style.Outline);
        Assert.Equal(OutlineLevel.H3, paragraphs[2].Style.Outline);
    }

    [Fact]
    public void Reader_BuildsBoldAndItalicRuns()
    {
        var doc = MarkdownReader.FromMarkdown("plain **bold** and *italic* end");
        var p = doc.EnumerateParagraphs().Single();

        Assert.Contains(p.Runs, r => r.Style.Bold && r.Text == "bold");
        Assert.Contains(p.Runs, r => r.Style.Italic && r.Text == "italic");
    }

    [Fact]
    public void Reader_DetectsBulletAndOrderedLists()
    {
        const string source = "- item one\n- item two\n\n1. first\n2. second\n";
        var doc = MarkdownReader.FromMarkdown(source);
        var paragraphs = doc.EnumerateParagraphs().ToList();

        Assert.Equal(4, paragraphs.Count);
        Assert.Equal(ListKind.Bullet, paragraphs[0].Style.ListMarker!.Kind);
        Assert.Equal(ListKind.Bullet, paragraphs[1].Style.ListMarker!.Kind);
        Assert.Equal(ListKind.OrderedDecimal, paragraphs[2].Style.ListMarker!.Kind);
        Assert.Equal(2, paragraphs[3].Style.ListMarker!.OrderedNumber);
    }

    [Fact]
    public void Writer_RendersHeaderHashesByOutlineLevel()
    {
        var doc = new PolyDocument();
        var section = new Section();
        doc.Sections.Add(section);
        var h2 = new Paragraph { Style = { Outline = OutlineLevel.H2 } };
        h2.AddText("section");
        section.Blocks.Add(h2);

        var rendered = MarkdownWriter.ToMarkdown(doc);

        Assert.StartsWith("## section", rendered);
    }

    [Fact]
    public void Writer_EmitsBulletAndOrderedMarkers()
    {
        var doc = new PolyDocument();
        var section = new Section();
        doc.Sections.Add(section);
        var bullet = new Paragraph();
        bullet.Style.ListMarker = new ListMarker { Kind = ListKind.Bullet };
        bullet.AddText("a");
        section.Blocks.Add(bullet);
        var ordered = new Paragraph();
        ordered.Style.ListMarker = new ListMarker { Kind = ListKind.OrderedDecimal, OrderedNumber = 3 };
        ordered.AddText("b");
        section.Blocks.Add(ordered);

        var rendered = MarkdownWriter.ToMarkdown(doc);

        Assert.Contains("- a", rendered);
        Assert.Contains("3. b", rendered);
    }

    [Fact]
    public void RoundTrip_HeadersAndListsArePreserved()
    {
        const string source =
            "# 제목\n\n본문 단락\n\n- 항목 A\n- 항목 B\n\n1. 첫째\n2. 둘째\n";
        var doc = MarkdownReader.FromMarkdown(source);
        var rendered = MarkdownWriter.ToMarkdown(doc);
        var reparsed = MarkdownReader.FromMarkdown(rendered);

        var original = doc.EnumerateParagraphs().Select(p => (p.Style.Outline, p.Style.ListMarker?.Kind)).ToList();
        var roundTripped = reparsed.EnumerateParagraphs().Select(p => (p.Style.Outline, p.Style.ListMarker?.Kind)).ToList();
        Assert.Equal(original, roundTripped);
    }
}
