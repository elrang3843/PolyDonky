using System.IO.Compression;
using PolyDoc.Codecs.Hwpx;
using PolyDoc.Core;

namespace PolyDoc.Codecs.Hwpx.Tests;

public class HwpxRoundTripTests
{
    [Fact]
    public void RoundTrip_PreservesParagraphsAndHeadings()
    {
        var doc = new PolyDocument();
        doc.Metadata.Title = "HWPX 라운드트립";
        doc.Metadata.Author = "Noh JinMoon";

        var section = new Section();
        doc.Sections.Add(section);

        var heading = new Paragraph { Style = { Outline = OutlineLevel.H1 } };
        heading.AddText("제목");
        section.Blocks.Add(heading);

        var body = new Paragraph();
        body.AddText("본문 ");
        body.AddText("강조", new RunStyle { Bold = true });
        body.AddText(" 와 ");
        body.AddText("기울임", new RunStyle { Italic = true });
        body.AddText(" 끝.");
        section.Blocks.Add(body);

        var roundTripped = WriteThenRead(doc);

        Assert.Equal("HWPX 라운드트립", roundTripped.Metadata.Title);
        Assert.Equal("Noh JinMoon", roundTripped.Metadata.Author);

        var paragraphs = roundTripped.EnumerateParagraphs().ToList();
        Assert.Equal(2, paragraphs.Count);
        Assert.Equal(OutlineLevel.H1, paragraphs[0].Style.Outline);
        Assert.Equal("제목", paragraphs[0].GetPlainText());
        Assert.Equal(OutlineLevel.Body, paragraphs[1].Style.Outline);
        Assert.Contains(paragraphs[1].Runs, r => r.Style.Bold && r.Text == "강조");
        Assert.Contains(paragraphs[1].Runs, r => r.Style.Italic && r.Text == "기울임");
    }

    [Fact]
    public void RoundTrip_PreservesAlignment()
    {
        var doc = new PolyDocument();
        var section = new Section();
        doc.Sections.Add(section);

        foreach (var alignment in new[] { Alignment.Left, Alignment.Center, Alignment.Right, Alignment.Justify })
        {
            var p = new Paragraph { Style = { Alignment = alignment } };
            p.AddText($"align={alignment}");
            section.Blocks.Add(p);
        }

        var roundTripped = WriteThenRead(doc);
        var paragraphs = roundTripped.EnumerateParagraphs().ToList();

        Assert.Equal(Alignment.Left, paragraphs[0].Style.Alignment);
        Assert.Equal(Alignment.Center, paragraphs[1].Style.Alignment);
        Assert.Equal(Alignment.Right, paragraphs[2].Style.Alignment);
        Assert.Equal(Alignment.Justify, paragraphs[3].Style.Alignment);
    }

    [Fact]
    public void RoundTrip_PreservesUnderlineAndStrikethrough()
    {
        var doc = new PolyDocument();
        var section = new Section();
        doc.Sections.Add(section);

        var p = new Paragraph();
        p.AddText("u", new RunStyle { Underline = true });
        p.AddText("s", new RunStyle { Strikethrough = true });
        section.Blocks.Add(p);

        var run = WriteThenRead(doc).EnumerateParagraphs().Single().Runs;
        Assert.True(run.Single(r => r.Text == "u").Style.Underline);
        Assert.True(run.Single(r => r.Text == "s").Style.Strikethrough);
    }

    [Fact]
    public void Package_HasMimetypeAsFirstStoredEntry()
    {
        var doc = PolyDocument.Empty();
        doc.Sections[0].Blocks.Add(Paragraph.Of("content"));

        var bytes = WriteToBytes(doc);
        using var ms = new MemoryStream(bytes);
        using var archive = new ZipArchive(ms, ZipArchiveMode.Read);

        // HWPX 사양: 첫 entry 는 mimetype 이며 STORED (압축 없음).
        var first = archive.Entries[0];
        Assert.Equal("mimetype", first.FullName);
        Assert.Equal(first.CompressedLength, first.Length);
    }

    [Fact]
    public void Package_ContainsRequiredParts()
    {
        var doc = PolyDocument.Empty();
        var bytes = WriteToBytes(doc);
        using var ms = new MemoryStream(bytes);
        using var archive = new ZipArchive(ms, ZipArchiveMode.Read);

        Assert.NotNull(archive.GetEntry("mimetype"));
        Assert.NotNull(archive.GetEntry("META-INF/container.xml"));
        Assert.NotNull(archive.GetEntry("Contents/content.hpf"));
        Assert.NotNull(archive.GetEntry("Contents/header.xml"));
        Assert.NotNull(archive.GetEntry("Contents/section0.xml"));
        Assert.NotNull(archive.GetEntry("version.xml"));
    }

    [Fact]
    public void Read_ThrowsOnWrongMimetype()
    {
        // 임의 ZIP 에 잘못된 mimetype 을 넣어 우리 reader 가 거부하는지 확인.
        using var ms = new MemoryStream();
        using (var archive = new ZipArchive(ms, ZipArchiveMode.Create, leaveOpen: true))
        {
            var entry = archive.CreateEntry("mimetype", CompressionLevel.NoCompression);
            using var stream = entry.Open();
            var bytes = System.Text.Encoding.UTF8.GetBytes("application/epub+zip");
            stream.Write(bytes, 0, bytes.Length);
        }
        ms.Position = 0;
        Assert.Throws<InvalidDataException>(() => new HwpxReader().Read(ms));
    }

    private static byte[] WriteToBytes(PolyDocument doc)
    {
        using var ms = new MemoryStream();
        new HwpxWriter().Write(doc, ms);
        return ms.ToArray();
    }

    private static PolyDocument WriteThenRead(PolyDocument doc)
    {
        var bytes = WriteToBytes(doc);
        using var ms = new MemoryStream(bytes);
        return new HwpxReader().Read(ms);
    }
}
