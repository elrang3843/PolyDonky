using System.IO.Compression;
using PolyDoc.Core;
using PolyDoc.Iwpf;

namespace PolyDoc.Iwpf.Tests;

public class IwpfRoundTripTests
{
    [Fact]
    public void RoundTrip_PreservesMetadataAndBody()
    {
        var doc = new PolyDocument();
        doc.Metadata.Title = "라운드트립";
        doc.Metadata.Author = "Noh JinMoon";
        var section = new Section();
        doc.Sections.Add(section);
        var heading = new Paragraph { Style = { Outline = OutlineLevel.H1 } };
        heading.AddText("제목");
        section.Blocks.Add(heading);
        var body = new Paragraph();
        body.AddText("본문 ");
        body.AddText("강조", new RunStyle { Bold = true });
        body.AddText(" 끝.");
        section.Blocks.Add(body);

        var bytes = WriteToBytes(doc);
        var read = ReadFromBytes(bytes);

        Assert.Equal("라운드트립", read.Metadata.Title);
        Assert.Equal("Noh JinMoon", read.Metadata.Author);
        var paragraphs = read.EnumerateParagraphs().ToList();
        Assert.Equal(2, paragraphs.Count);
        Assert.Equal(OutlineLevel.H1, paragraphs[0].Style.Outline);
        Assert.Equal("제목", paragraphs[0].GetPlainText());
        Assert.Contains(paragraphs[1].Runs, r => r.Style.Bold && r.Text == "강조");
    }

    [Fact]
    public void Package_ContainsRequiredParts()
    {
        var doc = PolyDocument.Empty();
        var bytes = WriteToBytes(doc);

        using var ms = new MemoryStream(bytes);
        using var archive = new ZipArchive(ms, ZipArchiveMode.Read);

        Assert.NotNull(archive.GetEntry("manifest.json"));
        Assert.NotNull(archive.GetEntry("content/document.json"));
        Assert.NotNull(archive.GetEntry("content/styles.json"));
    }

    [Fact]
    public void Read_ThrowsOnTamperedDocumentJson()
    {
        var doc = PolyDocument.Empty();
        doc.Sections[0].Blocks.Add(Paragraph.Of("원본 데이터"));
        var bytes = WriteToBytes(doc);

        var tampered = TamperPart(bytes, "content/document.json");

        using var ms = new MemoryStream(tampered);
        Assert.Throws<InvalidDataException>(() => new IwpfReader().Read(ms));
    }

    [Fact]
    public void Read_RejectsUnknownPackageType()
    {
        // 매니페스트의 packageType 을 다른 값으로 바꿔 위장한다.
        var doc = PolyDocument.Empty();
        var bytes = WriteToBytes(doc);
        var faked = ReplaceInManifest(bytes, "polydoc.iwpf", "vendor.unknown");

        using var ms = new MemoryStream(faked);
        Assert.Throws<InvalidDataException>(() => new IwpfReader().Read(ms));
    }

    [Fact]
    public void Manifest_RecordsSha256ForEveryPart()
    {
        var doc = PolyDocument.Empty();
        var bytes = WriteToBytes(doc);
        using var ms = new MemoryStream(bytes);
        using var archive = new ZipArchive(ms, ZipArchiveMode.Read);

        var manifestEntry = archive.GetEntry("manifest.json");
        Assert.NotNull(manifestEntry);
        using var stream = manifestEntry!.Open();
        using var memory = new MemoryStream();
        stream.CopyTo(memory);
        var manifest = System.Text.Json.JsonSerializer.Deserialize<IwpfManifest>(memory.ToArray(), JsonDefaults.Options);
        Assert.NotNull(manifest);
        Assert.NotEmpty(manifest!.Parts);
        Assert.All(manifest.Parts.Values, e => Assert.Matches("^[0-9a-f]{64}$", e.Sha256));
    }

    private static byte[] WriteToBytes(PolyDocument doc)
    {
        using var ms = new MemoryStream();
        new IwpfWriter().Write(doc, ms);
        return ms.ToArray();
    }

    private static PolyDocument ReadFromBytes(byte[] bytes)
    {
        using var ms = new MemoryStream(bytes);
        return new IwpfReader().Read(ms);
    }

    private static byte[] TamperPart(byte[] original, string path)
    {
        using var input = new MemoryStream(original);
        using var output = new MemoryStream();
        using (var inputZip = new ZipArchive(input, ZipArchiveMode.Read, leaveOpen: true))
        using (var outputZip = new ZipArchive(output, ZipArchiveMode.Create, leaveOpen: true))
        {
            foreach (var entry in inputZip.Entries)
            {
                byte[] payload;
                using (var es = entry.Open())
                using (var ms = new MemoryStream())
                {
                    es.CopyTo(ms);
                    payload = ms.ToArray();
                }
                if (entry.FullName == path && payload.Length > 1)
                {
                    payload[^2] ^= 0x55;
                }
                var newEntry = outputZip.CreateEntry(entry.FullName, CompressionLevel.Optimal);
                using var ws = newEntry.Open();
                ws.Write(payload, 0, payload.Length);
            }
        }
        return output.ToArray();
    }

    private static byte[] ReplaceInManifest(byte[] original, string find, string replace)
    {
        using var input = new MemoryStream(original);
        using var output = new MemoryStream();
        using (var inputZip = new ZipArchive(input, ZipArchiveMode.Read, leaveOpen: true))
        using (var outputZip = new ZipArchive(output, ZipArchiveMode.Create, leaveOpen: true))
        {
            foreach (var entry in inputZip.Entries)
            {
                byte[] payload;
                using (var es = entry.Open())
                using (var ms = new MemoryStream())
                {
                    es.CopyTo(ms);
                    payload = ms.ToArray();
                }
                if (entry.FullName == "manifest.json")
                {
                    var text = System.Text.Encoding.UTF8.GetString(payload);
                    text = text.Replace(find, replace, StringComparison.Ordinal);
                    payload = System.Text.Encoding.UTF8.GetBytes(text);
                }
                var newEntry = outputZip.CreateEntry(entry.FullName, CompressionLevel.Optimal);
                using var ws = newEntry.Open();
                ws.Write(payload, 0, payload.Length);
            }
        }
        return output.ToArray();
    }
}
