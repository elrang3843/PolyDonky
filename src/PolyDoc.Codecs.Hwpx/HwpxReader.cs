using System.IO.Compression;
using System.Text;
using System.Xml.Linq;
using PolyDoc.Core;

namespace PolyDoc.Codecs.Hwpx;

/// <summary>
/// HWPX (KS X 6101) → PolyDocument 리더.
///
/// HwpxWriter 와 같은 charPr/paraPr/style ID 약속을 따른다 (writer 의 doc-comment 참조).
/// 다른 도구(한컴 오피스 등)가 만든 HWPX 의 임의 ID 매핑은 다음 사이클에서 header.xml 의
/// charPr/paraPr 정의를 풀어서 정확하게 회수한다.
/// </summary>
public sealed class HwpxReader : IDocumentReader
{
    public string FormatId => "hwpx";

    private static readonly XNamespace OpfContainer = HwpxNamespaces.OpfContainer;
    private static readonly XNamespace Opf = HwpxNamespaces.OpfPackage;
    private static readonly XNamespace Dc = HwpxNamespaces.DcMetadata;
    private static readonly XNamespace Hh = HwpxNamespaces.Head;
    private static readonly XNamespace Hp = HwpxNamespaces.Paragraph;
    private static readonly XNamespace Hs = HwpxNamespaces.Section;

    public PolyDocument Read(Stream input)
    {
        ArgumentNullException.ThrowIfNull(input);

        Stream zipStream = input;
        MemoryStream? buffered = null;
        if (!input.CanSeek)
        {
            buffered = new MemoryStream();
            input.CopyTo(buffered);
            buffered.Position = 0;
            zipStream = buffered;
        }

        try
        {
            using var archive = new ZipArchive(zipStream, ZipArchiveMode.Read, leaveOpen: true);

            ValidateMimetype(archive);

            var rootHpfPath = ResolveContentHpf(archive);
            var (sectionPaths, metadata) = ReadOpfManifest(archive, rootHpfPath);

            // 의도적으로 header.xml 은 1차 사이클에서 깊이 파싱하지 않는다.
            // (writer 가 만든 charPr/paraPr ID 매핑을 약속으로 사용)

            var document = new PolyDocument { Metadata = metadata };
            foreach (var path in sectionPaths)
            {
                document.Sections.Add(ReadSection(archive, path));
            }
            if (document.Sections.Count == 0)
            {
                document.Sections.Add(new Section());
            }
            return document;
        }
        finally
        {
            buffered?.Dispose();
        }
    }

    private static void ValidateMimetype(ZipArchive archive)
    {
        var entry = archive.GetEntry(HwpxPaths.Mimetype)
            ?? throw new InvalidDataException("HWPX package is missing 'mimetype' entry.");
        using var stream = entry.Open();
        using var reader = new StreamReader(stream, Encoding.UTF8);
        var content = reader.ReadToEnd().Trim();
        if (content != HwpxPaths.MimetypeContent)
        {
            throw new InvalidDataException(
                $"Unexpected HWPX mimetype: '{content}'. Expected '{HwpxPaths.MimetypeContent}'.");
        }
    }

    private static string ResolveContentHpf(ZipArchive archive)
    {
        var container = LoadXml(archive, HwpxPaths.ContainerXml)
            ?? throw new InvalidDataException("HWPX package is missing META-INF/container.xml.");

        var rootfile = container.Root
            ?.Element(OpfContainer + "rootfiles")
            ?.Element(OpfContainer + "rootfile");

        var fullPath = rootfile?.Attribute("full-path")?.Value;
        return string.IsNullOrEmpty(fullPath) ? HwpxPaths.ContentHpf : fullPath;
    }

    private static (List<string> sectionPaths, DocumentMetadata metadata) ReadOpfManifest(ZipArchive archive, string rootHpfPath)
    {
        var doc = LoadXml(archive, rootHpfPath);
        var metadata = new DocumentMetadata();
        var sectionPaths = new List<string>();
        if (doc?.Root is null)
        {
            return (sectionPaths, metadata);
        }

        var packageElem = doc.Root;
        var meta = packageElem.Element(Opf + "metadata");
        if (meta is not null)
        {
            metadata.Title = meta.Element(Dc + "title")?.Value;
            metadata.Author = meta.Element(Dc + "creator")?.Value;
            var lang = meta.Element(Dc + "language")?.Value;
            if (!string.IsNullOrEmpty(lang)) metadata.Language = lang;
        }

        // manifest 의 item 중 section{N}.xml 패턴을 spine 순서대로 모은다.
        var manifestItems = packageElem.Element(Opf + "manifest")?.Elements(Opf + "item")
            .ToDictionary(
                e => e.Attribute("id")?.Value ?? string.Empty,
                e => e.Attribute("href")?.Value ?? string.Empty,
                StringComparer.Ordinal)
            ?? new Dictionary<string, string>(StringComparer.Ordinal);

        var spine = packageElem.Element(Opf + "spine")?.Elements(Opf + "itemref");
        if (spine is not null)
        {
            foreach (var itemref in spine)
            {
                var idref = itemref.Attribute("idref")?.Value;
                if (!string.IsNullOrEmpty(idref) && manifestItems.TryGetValue(idref!, out var href))
                {
                    sectionPaths.Add(CombineRoot(rootHpfPath, href));
                }
            }
        }

        // spine 이 비어 있으면 manifest 의 section* 추정.
        if (sectionPaths.Count == 0)
        {
            foreach (var (id, href) in manifestItems)
            {
                if (id.StartsWith("section", StringComparison.OrdinalIgnoreCase))
                {
                    sectionPaths.Add(CombineRoot(rootHpfPath, href));
                }
            }
        }

        return (sectionPaths, metadata);
    }

    private static string CombineRoot(string rootHpfPath, string relative)
    {
        var slash = rootHpfPath.LastIndexOf('/');
        var dir = slash >= 0 ? rootHpfPath[..(slash + 1)] : string.Empty;
        return dir + relative;
    }

    private static Section ReadSection(ZipArchive archive, string path)
    {
        var section = new Section();
        var doc = LoadXml(archive, path);
        if (doc?.Root is null)
        {
            return section;
        }

        // 어떤 namespace 든 paragraph 요소만 정확히 추리려고 LocalName 비교.
        foreach (var elem in doc.Root.Elements())
        {
            if (elem.Name.LocalName == "p")
            {
                section.Blocks.Add(ReadParagraph(elem));
            }
        }
        return section;
    }

    private static Paragraph ReadParagraph(XElement wp)
    {
        var paragraph = new Paragraph();
        if (int.TryParse(wp.Attribute("paraPrIDRef")?.Value, out var paraPrId))
        {
            paragraph.Style.Alignment = ParaPrIdToAlignment(paraPrId);
        }
        if (int.TryParse(wp.Attribute("styleIDRef")?.Value, out var styleId)
            && styleId is >= 1 and <= 6)
        {
            paragraph.Style.Outline = (OutlineLevel)styleId;
        }

        foreach (var elem in wp.Elements())
        {
            if (elem.Name.LocalName == "run")
            {
                ReadRun(paragraph, elem);
            }
        }

        if (paragraph.Runs.Count == 0)
        {
            paragraph.AddText(string.Empty);
        }
        return paragraph;
    }

    private static void ReadRun(Paragraph paragraph, XElement run)
    {
        var style = new RunStyle();
        if (int.TryParse(run.Attribute("charPrIDRef")?.Value, out var charPrId))
        {
            ApplyCharPrIdToStyle(charPrId, style);
        }

        var sb = new StringBuilder();
        foreach (var elem in run.Elements())
        {
            if (elem.Name.LocalName == "t")
            {
                sb.Append(elem.Value);
            }
        }
        if (sb.Length > 0)
        {
            paragraph.AddText(sb.ToString(), style);
        }
    }

    private static void ApplyCharPrIdToStyle(int id, RunStyle style)
    {
        // HwpxWriter 의 charPr id 매핑과 정확히 대응:
        // 0=plain, 1=bold, 2=italic, 3=bold+italic, 4=underline, 5=strikeout
        switch (id)
        {
            case 1: style.Bold = true; break;
            case 2: style.Italic = true; break;
            case 3: style.Bold = true; style.Italic = true; break;
            case 4: style.Underline = true; break;
            case 5: style.Strikethrough = true; break;
        }
    }

    private static Alignment ParaPrIdToAlignment(int id) => id switch
    {
        1 => Alignment.Center,
        2 => Alignment.Right,
        3 => Alignment.Justify,
        _ => Alignment.Left,
    };

    private static XDocument? LoadXml(ZipArchive archive, string path)
    {
        var entry = archive.GetEntry(path);
        if (entry is null) return null;
        using var stream = entry.Open();
        return XDocument.Load(stream);
    }
}
