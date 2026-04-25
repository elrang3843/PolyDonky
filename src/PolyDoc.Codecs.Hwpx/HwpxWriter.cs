using System.IO.Compression;
using System.Text;
using System.Xml.Linq;
using PolyDoc.Core;

namespace PolyDoc.Codecs.Hwpx;

/// <summary>
/// PolyDocument → HWPX (KS X 6101) 라이터.
///
/// 1차 사이클 목표: PolyDoc 자체 라운드트립 (PolyDoc → HWPX → PolyDoc) 으로 본문·기본 서식 보존.
/// 한컴 오피스 호환은 G3 검증 후 다음 사이클에서 정밀화한다.
///
/// 매핑:
///   - Section → section{N}.xml
///   - Paragraph → hp:p (paraPrIDRef + styleIDRef)
///   - Run → hp:run (charPrIDRef) + hp:t
///   - 굵게/기울임/밑줄/취소선 → charPr 의 bold/italic/underline/strikeout 속성
///   - 정렬 → paraPr 의 align horizontal
///   - 헤더(H1~H6) → 별도 styleID 를 부여 (1차 사이클은 charPr 의 height 와 bold 로 시각화)
/// </summary>
public sealed class HwpxWriter : IDocumentWriter
{
    public string FormatId => "hwpx";

    private static readonly XNamespace OpfContainer = HwpxNamespaces.OpfContainer;
    private static readonly XNamespace Opf = HwpxNamespaces.OpfPackage;
    private static readonly XNamespace Dc = HwpxNamespaces.DcMetadata;
    private static readonly XNamespace Hh = HwpxNamespaces.Head;
    private static readonly XNamespace Hp = HwpxNamespaces.Paragraph;
    private static readonly XNamespace Hs = HwpxNamespaces.Section;
    private static readonly XNamespace Hc = HwpxNamespaces.Common;

    public void Write(PolyDocument document, Stream output)
    {
        ArgumentNullException.ThrowIfNull(document);
        ArgumentNullException.ThrowIfNull(output);

        // HWPX 사양: 첫 ZIP entry 는 'mimetype' 이며 압축 없이(STORED) 저장되어야 한다.
        // System.IO.Compression.ZipArchive 는 entry 별 CompressionLevel 을 받으므로 가능.
        using var archive = new ZipArchive(output, ZipArchiveMode.Create, leaveOpen: true);

        WriteRawText(archive, HwpxPaths.Mimetype, HwpxPaths.MimetypeContent, CompressionLevel.NoCompression);
        WriteContainerXml(archive);
        WriteContentHpf(archive, document);
        WriteVersionXml(archive);
        WriteHeaderXml(archive);

        for (int i = 0; i < document.Sections.Count; i++)
        {
            WriteSectionXml(archive, i, document.Sections[i]);
        }

        // 빈 섹션이면 적어도 하나는 넣어 매니페스트 일관성 유지.
        if (document.Sections.Count == 0)
        {
            WriteSectionXml(archive, 0, new Section());
        }
    }

    private static void WriteRawText(ZipArchive archive, string path, string content, CompressionLevel level)
    {
        var entry = archive.CreateEntry(path, level);
        using var stream = entry.Open();
        var bytes = Encoding.UTF8.GetBytes(content);
        stream.Write(bytes, 0, bytes.Length);
    }

    private void WriteContainerXml(ZipArchive archive)
    {
        var doc = new XDocument(new XDeclaration("1.0", "utf-8", null),
            new XElement(OpfContainer + "container",
                new XAttribute("version", "1.0"),
                new XElement(OpfContainer + "rootfiles",
                    new XElement(OpfContainer + "rootfile",
                        new XAttribute("full-path", HwpxPaths.ContentHpf),
                        new XAttribute("media-type", "application/hwpml-package+xml")))));
        WriteXml(archive, HwpxPaths.ContainerXml, doc);
    }

    private void WriteContentHpf(ZipArchive archive, PolyDocument document)
    {
        var manifest = new XElement(Opf + "manifest",
            new XElement(Opf + "item",
                new XAttribute("id", "header"),
                new XAttribute("href", "header.xml"),
                new XAttribute("media-type", "application/xml")));

        var spine = new XElement(Opf + "spine");

        int sectionCount = Math.Max(document.Sections.Count, 1);
        for (int i = 0; i < sectionCount; i++)
        {
            var id = $"section{i}";
            manifest.Add(new XElement(Opf + "item",
                new XAttribute("id", id),
                new XAttribute("href", $"section{i}.xml"),
                new XAttribute("media-type", "application/xml")));
            spine.Add(new XElement(Opf + "itemref", new XAttribute("idref", id)));
        }

        var metadata = new XElement(Opf + "metadata",
            new XAttribute(XNamespace.Xmlns + "dc", Dc.NamespaceName),
            new XAttribute(XNamespace.Xmlns + "opf", Opf.NamespaceName));

        if (!string.IsNullOrEmpty(document.Metadata.Title))
        {
            metadata.Add(new XElement(Dc + "title", document.Metadata.Title));
        }
        if (!string.IsNullOrEmpty(document.Metadata.Author))
        {
            metadata.Add(new XElement(Dc + "creator", document.Metadata.Author));
        }
        metadata.Add(new XElement(Dc + "language", document.Metadata.Language ?? "ko"));
        metadata.Add(new XElement(Opf + "meta",
            new XAttribute("name", "producer"),
            new XAttribute("content", $"{HwpxFormat.ProducedBy}/{HwpxFormat.Version}")));

        var package = new XElement(Opf + "package",
            new XAttribute("version", HwpxFormat.Version),
            new XAttribute("unique-identifier", "polydoc-id"),
            metadata,
            manifest,
            spine);

        var doc = new XDocument(new XDeclaration("1.0", "utf-8", null), package);
        WriteXml(archive, HwpxPaths.ContentHpf, doc);
    }

    private void WriteVersionXml(ZipArchive archive)
    {
        var doc = new XDocument(new XDeclaration("1.0", "utf-8", null),
            new XElement(XName.Get("HCFVersion", HwpxNamespaces.Hancom + "version"),
                new XAttribute("targetApplication", "WORDPROCESSOR"),
                new XAttribute("major", "5"),
                new XAttribute("minor", "1"),
                new XAttribute("micro", "1"),
                new XAttribute("buildNumber", "0"),
                new XAttribute("os", "0"),
                new XAttribute("xmlVersion", HwpxFormat.Version)));
        WriteXml(archive, HwpxPaths.VersionXml, doc);
    }

    /// <summary>
    /// header.xml — charPr / paraPr / style 정의. 본 codec 의 특수 ID 매핑:
    ///   styleID 0 = bodyText (기본 본문)
    ///   styleID 1~6 = Heading 1~6
    ///   charPrID 0 = 기본 글자
    ///   charPrID 1 = 굵게 / 2 = 기울임 / 3 = 굵게+기울임 / 4 = 밑줄 / 5 = 취소선
    ///   paraPrID 0 = Left / 1 = Center / 2 = Right / 3 = Justify
    /// HwpxReader 와 정확히 같은 매핑을 공유한다.
    /// </summary>
    private void WriteHeaderXml(ZipArchive archive)
    {
        var fonts = new XElement(Hh + "fontfaces",
            new XAttribute("itemCnt", "1"),
            new XElement(Hh + "fontface",
                new XAttribute("lang", "HANGUL"),
                new XAttribute("count", "1"),
                new XElement(Hh + "font",
                    new XAttribute("id", "0"),
                    new XAttribute("face", "맑은 고딕"),
                    new XAttribute("type", "ttf"),
                    new XAttribute("isEmbedded", "0"))));

        var charProps = new XElement(Hh + "charProperties", new XAttribute("itemCnt", "6"));
        for (int id = 0; id < 6; id++)
        {
            charProps.Add(BuildCharPr(id));
        }

        var paraProps = new XElement(Hh + "paraProperties", new XAttribute("itemCnt", "4"));
        foreach (var (id, align) in new[] { (0, "LEFT"), (1, "CENTER"), (2, "RIGHT"), (3, "JUSTIFY") })
        {
            paraProps.Add(new XElement(Hh + "paraPr",
                new XAttribute("id", id.ToString()),
                new XAttribute("tabPrIDRef", "0"),
                new XAttribute("condense", "0"),
                new XAttribute("fontLineHeight", "0"),
                new XAttribute("snapToGrid", "0"),
                new XAttribute("suppressLineNumbers", "0"),
                new XAttribute("checked", "0"),
                new XElement(Hh + "align", new XAttribute("horizontal", align), new XAttribute("vertical", "BASELINE"))));
        }

        var styles = new XElement(Hh + "styles", new XAttribute("itemCnt", "7"));
        styles.Add(new XElement(Hh + "style",
            new XAttribute("id", "0"), new XAttribute("type", "PARA"),
            new XAttribute("name", "Normal"), new XAttribute("engName", "Normal"),
            new XAttribute("paraPrIDRef", "0"), new XAttribute("charPrIDRef", "0"),
            new XAttribute("nextStyleIDRef", "0"), new XAttribute("langID", "1042"),
            new XAttribute("lockForm", "0")));
        for (int level = 1; level <= 6; level++)
        {
            styles.Add(new XElement(Hh + "style",
                new XAttribute("id", level.ToString()), new XAttribute("type", "PARA"),
                new XAttribute("name", $"Heading{level}"), new XAttribute("engName", $"Heading{level}"),
                new XAttribute("paraPrIDRef", "0"),
                new XAttribute("charPrIDRef", "1"),     // bold for headings
                new XAttribute("nextStyleIDRef", "0"), new XAttribute("langID", "1042"),
                new XAttribute("lockForm", "0")));
        }

        var refList = new XElement(Hh + "refList", fonts, charProps, paraProps, styles);

        var head = new XElement(Hh + "head",
            new XAttribute(XNamespace.Xmlns + HwpxNamespaces.PrefixHh, Hh.NamespaceName),
            new XAttribute(XNamespace.Xmlns + HwpxNamespaces.PrefixHc, Hc.NamespaceName),
            new XAttribute("version", HwpxFormat.Version),
            new XAttribute("secCnt", "1"),
            new XElement(Hh + "beginNum",
                new XAttribute("page", "1"),
                new XAttribute("footnote", "1"),
                new XAttribute("endnote", "1"),
                new XAttribute("pic", "1"),
                new XAttribute("tbl", "1"),
                new XAttribute("equation", "1")),
            refList);

        var doc = new XDocument(new XDeclaration("1.0", "utf-8", null), head);
        WriteXml(archive, HwpxPaths.HeaderXml, doc);
    }

    private XElement BuildCharPr(int id)
    {
        // id 0=plain, 1=bold, 2=italic, 3=bold+italic, 4=underline, 5=strikeout
        bool bold = id is 1 or 3;
        bool italic = id is 2 or 3;
        bool underline = id == 4;
        bool strike = id == 5;

        var charPr = new XElement(Hh + "charPr",
            new XAttribute("id", id.ToString()),
            new XAttribute("height", "1000"),     // 10pt — HWPX 는 hpsize(1pt = 100)
            new XAttribute("textColor", "#000000"),
            new XAttribute("shadeColor", "none"),
            new XAttribute("useFontSpace", "0"),
            new XAttribute("useKerning", "0"),
            new XAttribute("symMark", "NONE"),
            new XAttribute("borderFillIDRef", "0"));

        if (bold)
        {
            charPr.Add(new XElement(Hh + "bold"));
        }
        if (italic)
        {
            charPr.Add(new XElement(Hh + "italic"));
        }
        if (underline)
        {
            charPr.Add(new XElement(Hh + "underline",
                new XAttribute("type", "BOTTOM"),
                new XAttribute("shape", "SOLID"),
                new XAttribute("color", "#000000")));
        }
        if (strike)
        {
            charPr.Add(new XElement(Hh + "strikeout",
                new XAttribute("shape", "SOLID"),
                new XAttribute("color", "#000000")));
        }
        return charPr;
    }

    private void WriteSectionXml(ZipArchive archive, int sectionIndex, Section section)
    {
        var sec = new XElement(Hs + "sec",
            new XAttribute(XNamespace.Xmlns + HwpxNamespaces.PrefixHs, Hs.NamespaceName),
            new XAttribute(XNamespace.Xmlns + HwpxNamespaces.PrefixHp, Hp.NamespaceName),
            new XAttribute(XNamespace.Xmlns + HwpxNamespaces.PrefixHc, Hc.NamespaceName));

        if (section.Blocks.Count == 0)
        {
            sec.Add(BuildEmptyParagraph());
        }
        else
        {
            foreach (var block in section.Blocks)
            {
                if (block is Paragraph p)
                {
                    sec.Add(BuildParagraph(p));
                }
                // Phase C 의 Table/ImageBlock/OpaqueBlock 은 다음 사이클에서 HWPX 매핑 추가.
            }
        }

        var doc = new XDocument(new XDeclaration("1.0", "utf-8", null), sec);
        WriteXml(archive, HwpxPaths.SectionXml(sectionIndex), doc);
    }

    private XElement BuildEmptyParagraph()
        => new(Hp + "p",
            new XAttribute("paraPrIDRef", "0"),
            new XAttribute("styleIDRef", "0"),
            new XElement(Hp + "run", new XAttribute("charPrIDRef", "0")));

    private XElement BuildParagraph(Paragraph p)
    {
        int paraPrID = AlignmentToParaPrId(p.Style.Alignment);
        int styleID = p.Style.Outline switch
        {
            OutlineLevel.H1 => 1,
            OutlineLevel.H2 => 2,
            OutlineLevel.H3 => 3,
            OutlineLevel.H4 => 4,
            OutlineLevel.H5 => 5,
            OutlineLevel.H6 => 6,
            _ => 0,
        };

        var para = new XElement(Hp + "p",
            new XAttribute("paraPrIDRef", paraPrID.ToString()),
            new XAttribute("styleIDRef", styleID.ToString()));

        foreach (var run in p.Runs)
        {
            para.Add(BuildRun(run));
        }
        if (p.Runs.Count == 0)
        {
            para.Add(new XElement(Hp + "run", new XAttribute("charPrIDRef", "0")));
        }
        return para;
    }

    private XElement BuildRun(Run run)
    {
        int charPrID = SelectCharPrId(run.Style);
        var elem = new XElement(Hp + "run",
            new XAttribute("charPrIDRef", charPrID.ToString()));
        if (run.Text.Length > 0)
        {
            elem.Add(new XElement(Hp + "t", run.Text));
        }
        return elem;
    }

    private static int SelectCharPrId(RunStyle s)
    {
        if (s.Strikethrough) return 5;
        if (s.Underline) return 4;
        if (s.Bold && s.Italic) return 3;
        if (s.Italic) return 2;
        if (s.Bold) return 1;
        return 0;
    }

    private static int AlignmentToParaPrId(Alignment a) => a switch
    {
        Alignment.Center => 1,
        Alignment.Right => 2,
        Alignment.Justify or Alignment.Distributed => 3,
        _ => 0,
    };

    private static void WriteXml(ZipArchive archive, string path, XDocument doc)
    {
        var entry = archive.CreateEntry(path, CompressionLevel.Optimal);
        using var stream = entry.Open();
        using var writer = new StreamWriter(stream, new UTF8Encoding(encoderShouldEmitUTF8Identifier: false));
        doc.Save(writer, SaveOptions.None);
    }
}
