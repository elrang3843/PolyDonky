using System.Text;
using System.Text.RegularExpressions;
using PolyDoc.Core;

namespace PolyDoc.Codecs.Markdown;

/// <summary>
/// CommonMark 의 실용적 서브셋만 처리하는 가벼운 Markdown 리더.
/// 지원 범위: ATX 헤더(#~######), 순서/비순서 리스트, 본문 문단, 강조(**굵게**, *기울임*).
/// 코드블록·인용·표·이미지·링크는 Phase A 범위 밖이며, 추후 Markdig 도입 시 교체한다.
/// </summary>
public sealed partial class MarkdownReader : IDocumentReader
{
    public string FormatId => "md";

    public PolyDocument Read(Stream input)
    {
        ArgumentNullException.ThrowIfNull(input);
        using var reader = new StreamReader(input, Encoding.UTF8, detectEncodingFromByteOrderMarks: true, leaveOpen: true);
        return FromMarkdown(reader.ReadToEnd());
    }

    public static PolyDocument FromMarkdown(string source)
    {
        ArgumentNullException.ThrowIfNull(source);
        var document = new PolyDocument();
        var section = new Section();
        document.Sections.Add(section);

        var lines = source.Replace("\r\n", "\n", StringComparison.Ordinal)
                           .Replace('\r', '\n')
                           .Split('\n');

        var paragraphBuffer = new List<string>();

        foreach (var rawLine in lines)
        {
            var line = rawLine;

            if (string.IsNullOrWhiteSpace(line))
            {
                FlushParagraph(section, paragraphBuffer);
                continue;
            }

            // ATX 헤더 (# ~ ######)
            var headerMatch = HeaderRegex().Match(line);
            if (headerMatch.Success)
            {
                FlushParagraph(section, paragraphBuffer);
                int level = headerMatch.Groups["hashes"].Value.Length;
                string text = headerMatch.Groups["text"].Value.Trim();
                section.Blocks.Add(BuildHeader(level, text));
                continue;
            }

            // 순서 리스트 (1. , 2. ...)
            var orderedMatch = OrderedListRegex().Match(line);
            if (orderedMatch.Success)
            {
                FlushParagraph(section, paragraphBuffer);
                int number = int.Parse(orderedMatch.Groups["num"].Value, System.Globalization.CultureInfo.InvariantCulture);
                section.Blocks.Add(BuildListItem(orderedMatch.Groups["text"].Value, ListKind.OrderedDecimal, number));
                continue;
            }

            // 비순서 리스트 (- , * , + )
            var bulletMatch = BulletListRegex().Match(line);
            if (bulletMatch.Success)
            {
                FlushParagraph(section, paragraphBuffer);
                section.Blocks.Add(BuildListItem(bulletMatch.Groups["text"].Value, ListKind.Bullet, null));
                continue;
            }

            paragraphBuffer.Add(line);
        }

        FlushParagraph(section, paragraphBuffer);
        return document;
    }

    private static void FlushParagraph(Section section, List<string> buffer)
    {
        if (buffer.Count == 0)
        {
            return;
        }
        // 줄바꿈은 단일 공백으로 합친다 (CommonMark soft break 와 호환).
        var text = string.Join(' ', buffer);
        var paragraph = new Paragraph();
        paragraph.Style.Outline = OutlineLevel.Body;
        AppendInline(paragraph, text);
        section.Blocks.Add(paragraph);
        buffer.Clear();
    }

    private static Paragraph BuildHeader(int level, string text)
    {
        var paragraph = new Paragraph();
        paragraph.Style.Outline = (OutlineLevel)Math.Clamp(level, 1, 6);
        AppendInline(paragraph, text);
        return paragraph;
    }

    private static Paragraph BuildListItem(string text, ListKind kind, int? orderedNumber)
    {
        var paragraph = new Paragraph();
        paragraph.Style.ListMarker = new ListMarker
        {
            Kind = kind,
            Level = 0,
            OrderedNumber = orderedNumber,
        };
        AppendInline(paragraph, text);
        return paragraph;
    }

    /// <summary>인라인 강조(**bold**, *italic*) 만 분리해 Run 으로 변환한다.</summary>
    private static void AppendInline(Paragraph paragraph, string text)
    {
        // 매우 단순한 토크나이저: ** 우선, * 차순.
        int i = 0;
        var literal = new StringBuilder();

        void Flush(RunStyle? style = null)
        {
            if (literal.Length == 0)
            {
                return;
            }
            paragraph.AddText(literal.ToString(), style);
            literal.Clear();
        }

        while (i < text.Length)
        {
            // **bold** 매칭
            if (i + 1 < text.Length && text[i] == '*' && text[i + 1] == '*')
            {
                int end = text.IndexOf("**", i + 2, StringComparison.Ordinal);
                if (end > i + 2)
                {
                    Flush();
                    paragraph.AddText(text.Substring(i + 2, end - (i + 2)), new RunStyle { Bold = true });
                    i = end + 2;
                    continue;
                }
            }

            // *italic* 매칭 (단, ** 가 아닐 때)
            if (text[i] == '*')
            {
                int end = text.IndexOf('*', i + 1);
                if (end > i + 1)
                {
                    Flush();
                    paragraph.AddText(text.Substring(i + 1, end - (i + 1)), new RunStyle { Italic = true });
                    i = end + 1;
                    continue;
                }
            }

            literal.Append(text[i]);
            i++;
        }

        Flush();

        // 빈 단락 방지 — 최소 하나의 빈 Run 보장.
        if (paragraph.Runs.Count == 0)
        {
            paragraph.AddText(string.Empty);
        }
    }

    [GeneratedRegex(@"^(?<hashes>\#{1,6})\s+(?<text>.+?)\s*\#*\s*$", RegexOptions.CultureInvariant)]
    private static partial Regex HeaderRegex();

    [GeneratedRegex(@"^(?<num>\d{1,9})\.\s+(?<text>.+)$", RegexOptions.CultureInvariant)]
    private static partial Regex OrderedListRegex();

    [GeneratedRegex(@"^[-*+]\s+(?<text>.+)$", RegexOptions.CultureInvariant)]
    private static partial Regex BulletListRegex();
}
