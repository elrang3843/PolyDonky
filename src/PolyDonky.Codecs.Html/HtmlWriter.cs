using System.Globalization;
using System.Text;
using PolyDonky.Core;

namespace PolyDonky.Codecs.Html;

/// <summary>
/// PolyDonkyument 를 HTML5 로 직렬화한다 — UTF-8, &lt;!DOCTYPE html&gt;, semantic 마크업.
///
/// 매핑:
///   - OutlineLevel.H1~H6   → &lt;h1&gt;~&lt;h6&gt;
///   - 일반 단락             → &lt;p&gt;
///   - QuoteLevel ≥ 1       → 중첩 &lt;blockquote&gt;
///   - IsThematicBreak      → &lt;hr&gt;
///   - CodeLanguage non-null → &lt;pre&gt;&lt;code class="language-xxx"&gt;...&lt;/code&gt;&lt;/pre&gt;
///   - ListMarker (bullet/ordered, nested by Level) → &lt;ul&gt;/&lt;ol&gt; + &lt;li&gt;
///   - ListMarker.Checked    → &lt;input type="checkbox" disabled checked?&gt; 접두
///   - Run.Bold/Italic/Strike/Sub/Super/Underline → &lt;strong&gt;&lt;em&gt;&lt;s&gt;&lt;sub&gt;&lt;sup&gt;&lt;u&gt;
///   - 모노스페이스 FontFamily → &lt;code&gt;
///   - Run.Url               → &lt;a href="..."&gt;
///   - Run.Foreground/Background/FontSizePt/FontFamily → &lt;span style="..."&gt;
///   - Table                 → &lt;table&gt;&lt;thead&gt;/&lt;tbody&gt; + 셀 정렬 style
///   - ImageBlock            → &lt;img src alt width height&gt; (또는 &lt;figure&gt; + &lt;figcaption&gt;)
/// </summary>
public sealed class HtmlWriter : IDocumentWriter
{
    public string FormatId => "html";

    public Encoding Encoding { get; init; } = new UTF8Encoding(encoderShouldEmitUTF8Identifier: false);

    /// <summary>완전한 HTML5 문서로 출력할지 (기본 true) — false 면 fragment.</summary>
    public bool FullDocument { get; init; } = true;

    /// <summary>문서 제목 — &lt;title&gt; 에 사용. null 이면 첫 H1 텍스트 사용.</summary>
    public string? DocumentTitle { get; init; }

    public void Write(PolyDonkyument document, Stream output)
    {
        ArgumentNullException.ThrowIfNull(document);
        ArgumentNullException.ThrowIfNull(output);

        using var writer = new StreamWriter(output, Encoding, leaveOpen: true) { NewLine = "\n" };
        writer.Write(ToHtml(document, FullDocument, DocumentTitle));
    }

    public static string ToHtml(PolyDonkyument document, bool fullDocument = true, string? title = null)
    {
        ArgumentNullException.ThrowIfNull(document);
        var sb = new StringBuilder();

        if (fullDocument)
        {
            var docTitle = title ?? document.EnumerateParagraphs()
                .FirstOrDefault(p => p.Style.Outline == OutlineLevel.H1)?.GetPlainText()
                ?? "PolyDonky 문서";

            sb.Append("<!DOCTYPE html>\n");
            sb.Append("<html lang=\"ko\">\n");
            sb.Append("<head>\n");
            sb.Append("  <meta charset=\"utf-8\">\n");
            sb.Append("  <meta name=\"viewport\" content=\"width=device-width, initial-scale=1\">\n");
            sb.Append("  <meta name=\"generator\" content=\"PolyDonky\">\n");
            sb.Append("  <title>").Append(EscapeHtml(docTitle)).Append("</title>\n");
            sb.Append("</head>\n");
            sb.Append("<body>\n");
        }

        foreach (var section in document.Sections)
            WriteBlocks(sb, section.Blocks, indent: fullDocument ? "  " : "");

        if (fullDocument)
        {
            sb.Append("</body>\n</html>\n");
        }

        return sb.ToString();
    }

    // ── 블록 렌더링 ─────────────────────────────────────────────────────

    private static void WriteBlocks(StringBuilder sb, IList<Block> blocks, string indent)
    {
        // 인접 리스트·인용은 묶어서 처리한다.
        int i = 0;
        while (i < blocks.Count)
        {
            var b = blocks[i];

            // 연속된 같은 인용 깊이의 블록을 <blockquote> 로 감싼다.
            int qLvl = QuoteLevelOf(b);
            if (qLvl > 0)
            {
                int j = i;
                while (j < blocks.Count && QuoteLevelOf(blocks[j]) >= qLvl) j++;
                var inner = blocks.Skip(i).Take(j - i).Select(StripQuoteLevel).ToList();
                sb.Append(indent).Append("<blockquote>\n");
                WriteBlocks(sb, inner, indent + "  ");
                sb.Append(indent).Append("</blockquote>\n");
                i = j;
                continue;
            }

            // 연속된 리스트 단락 묶음.
            if (b is Paragraph p && p.Style.ListMarker is { } lm0)
            {
                int j = i;
                while (j < blocks.Count
                       && blocks[j] is Paragraph pj
                       && pj.Style.ListMarker is { } lmj
                       && lmj.Kind == lm0.Kind
                       && lmj.Level == lm0.Level)
                    j++;
                WriteListGroup(sb, blocks, i, j, indent);
                i = j;
                continue;
            }

            switch (b)
            {
                case Paragraph para: WriteParagraph(sb, para, indent); break;
                case Table table:    WriteTable(sb, table, indent);     break;
                case ImageBlock img: WriteImage(sb, img, indent);       break;
            }
            i++;
        }
    }

    private static int QuoteLevelOf(Block b) => b is Paragraph p ? p.Style.QuoteLevel : 0;

    private static Block StripQuoteLevel(Block b)
    {
        if (b is not Paragraph p) return b;
        var copy = new Paragraph
        {
            StyleId = p.StyleId,
            Style   = new ParagraphStyle
            {
                Alignment         = p.Style.Alignment,
                LineHeightFactor  = p.Style.LineHeightFactor,
                SpaceBeforePt     = p.Style.SpaceBeforePt,
                SpaceAfterPt      = p.Style.SpaceAfterPt,
                IndentFirstLineMm = p.Style.IndentFirstLineMm,
                IndentLeftMm      = p.Style.IndentLeftMm,
                IndentRightMm     = p.Style.IndentRightMm,
                Outline           = p.Style.Outline,
                ListMarker        = p.Style.ListMarker,
                QuoteLevel        = Math.Max(0, p.Style.QuoteLevel - 1),
                CodeLanguage      = p.Style.CodeLanguage,
                IsThematicBreak   = p.Style.IsThematicBreak,
            },
            Runs    = p.Runs,
        };
        return copy;
    }

    private static void WriteParagraph(StringBuilder sb, Paragraph p, string indent)
    {
        if (p.Style.IsThematicBreak)
        {
            sb.Append(indent).Append("<hr>\n");
            return;
        }

        if (p.Style.CodeLanguage is not null)
        {
            var code = EscapeHtml(p.GetPlainText());
            var langAttr = p.Style.CodeLanguage.Length > 0
                ? $" class=\"language-{EscapeAttr(p.Style.CodeLanguage)}\""
                : "";
            sb.Append(indent).Append("<pre><code").Append(langAttr).Append('>')
              .Append(code).Append("</code></pre>\n");
            return;
        }

        if (p.Style.Outline > OutlineLevel.Body)
        {
            int lvl = (int)p.Style.Outline;
            var styleAttr = ParagraphStyleAttr(p.Style);
            sb.Append(indent).Append('<').Append('h').Append(lvl).Append(styleAttr).Append('>');
            sb.Append(RenderRuns(p.Runs));
            sb.Append("</h").Append(lvl).Append(">\n");
            return;
        }

        var pStyleAttr = ParagraphStyleAttr(p.Style);
        sb.Append(indent).Append("<p").Append(pStyleAttr).Append('>');
        sb.Append(RenderRuns(p.Runs));
        sb.Append("</p>\n");
    }

    private static string ParagraphStyleAttr(ParagraphStyle s)
    {
        if (s.Alignment is Alignment.Left or Alignment.Distributed) return "";
        var ta = s.Alignment switch
        {
            Alignment.Center  => "center",
            Alignment.Right   => "right",
            Alignment.Justify => "justify",
            _                 => null,
        };
        return ta is null ? "" : $" style=\"text-align:{ta}\"";
    }

    private static void WriteListGroup(StringBuilder sb, IList<Block> blocks, int from, int to, string indent)
    {
        var p0 = (Paragraph)blocks[from];
        var lm = p0.Style.ListMarker!;
        string tag = lm.Kind == ListKind.Bullet ? "ul" : "ol";
        string startAttr = "";
        if (lm.Kind != ListKind.Bullet && lm.OrderedNumber is { } start && start != 1)
            startAttr = $" start=\"{start}\"";

        sb.Append(indent).Append('<').Append(tag).Append(startAttr).Append(">\n");
        var inner = indent + "  ";

        for (int k = from; k < to; k++)
        {
            var p = (Paragraph)blocks[k];
            var marker = p.Style.ListMarker!;
            sb.Append(inner).Append("<li>");
            if (marker.Checked.HasValue)
            {
                var ck = marker.Checked.Value ? " checked" : "";
                sb.Append("<input type=\"checkbox\" disabled").Append(ck).Append("> ");
            }
            sb.Append(RenderRuns(p.Runs));
            sb.Append("</li>\n");
        }

        sb.Append(indent).Append("</").Append(tag).Append(">\n");
    }

    private static void WriteTable(StringBuilder sb, Table t, string indent)
    {
        if (t.Rows.Count == 0) return;
        sb.Append(indent).Append("<table>\n");

        // <thead> 모음.
        var headerRows = t.Rows.Where(r => r.IsHeader).ToList();
        var bodyRows   = t.Rows.Where(r => !r.IsHeader).ToList();

        if (headerRows.Count > 0)
        {
            sb.Append(indent).Append("  <thead>\n");
            foreach (var r in headerRows) WriteRow(sb, r, indent + "    ", isHeader: true);
            sb.Append(indent).Append("  </thead>\n");
        }
        if (bodyRows.Count > 0)
        {
            sb.Append(indent).Append("  <tbody>\n");
            foreach (var r in bodyRows) WriteRow(sb, r, indent + "    ", isHeader: false);
            sb.Append(indent).Append("  </tbody>\n");
        }

        sb.Append(indent).Append("</table>\n");
    }

    private static void WriteRow(StringBuilder sb, TableRow row, string indent, bool isHeader)
    {
        sb.Append(indent).Append("<tr>\n");
        foreach (var cell in row.Cells)
        {
            var tag = isHeader ? "th" : "td";
            var attrs = new StringBuilder();
            if (cell.ColumnSpan > 1) attrs.Append(" colspan=\"").Append(cell.ColumnSpan).Append('"');
            if (cell.RowSpan    > 1) attrs.Append(" rowspan=\"").Append(cell.RowSpan).Append('"');
            var alignStyle = cell.TextAlign switch
            {
                CellTextAlign.Center  => " style=\"text-align:center\"",
                CellTextAlign.Right   => " style=\"text-align:right\"",
                CellTextAlign.Justify => " style=\"text-align:justify\"",
                _                     => "",
            };
            sb.Append(indent).Append("  <").Append(tag).Append(attrs).Append(alignStyle).Append('>');
            // 셀 안 블록을 인라인적으로 렌더 — 단락은 <br> 로 구분.
            bool first = true;
            foreach (var b in cell.Blocks)
            {
                if (b is Paragraph p)
                {
                    if (!first) sb.Append("<br>");
                    sb.Append(RenderRuns(p.Runs));
                    first = false;
                }
            }
            sb.Append("</").Append(tag).Append(">\n");
        }
        sb.Append(indent).Append("</tr>\n");
    }

    private static void WriteImage(StringBuilder sb, ImageBlock img, string indent)
    {
        var src    = img.ResourcePath ?? BuildDataUri(img);
        var alt    = EscapeAttr(img.Description ?? "");
        var sizeAttr = new StringBuilder();
        if (img.WidthMm  > 0) sizeAttr.Append(" width=\"")  .Append(MmToPx(img.WidthMm) .ToString("0", CultureInfo.InvariantCulture)).Append('"');
        if (img.HeightMm > 0) sizeAttr.Append(" height=\"") .Append(MmToPx(img.HeightMm).ToString("0", CultureInfo.InvariantCulture)).Append('"');

        if (img.ShowTitle && !string.IsNullOrEmpty(img.Title))
        {
            sb.Append(indent).Append("<figure>\n");
            sb.Append(indent).Append("  <img src=\"").Append(EscapeAttr(src)).Append("\" alt=\"")
              .Append(alt).Append('"').Append(sizeAttr).Append(">\n");
            sb.Append(indent).Append("  <figcaption>").Append(EscapeHtml(img.Title!)).Append("</figcaption>\n");
            sb.Append(indent).Append("</figure>\n");
        }
        else
        {
            sb.Append(indent).Append("<img src=\"").Append(EscapeAttr(src)).Append("\" alt=\"")
              .Append(alt).Append('"').Append(sizeAttr).Append(">\n");
        }
    }

    private static string BuildDataUri(ImageBlock img)
    {
        if (img.Data.Length == 0) return "";
        var b64 = Convert.ToBase64String(img.Data);
        return $"data:{img.MediaType};base64,{b64}";
    }

    private static double MmToPx(double mm) => mm * 96.0 / 25.4;

    // ── 인라인 렌더링 ─────────────────────────────────────────────────

    private static string RenderRuns(IList<Run> runs)
    {
        var sb = new StringBuilder();
        foreach (var r in runs) sb.Append(RenderRun(r));
        return sb.ToString();
    }

    private static string RenderRun(Run run)
    {
        var s    = run.Style;
        var text = EscapeHtml(run.Text).Replace("\n", "<br>");

        bool isMono = !string.IsNullOrEmpty(s.FontFamily) &&
                      s.FontFamily.Contains("monospace", StringComparison.OrdinalIgnoreCase);

        // 인라인 스타일 (font-family 제외 — 모노스페이스는 <code> 로 표현).
        var styleAttr = BuildSpanStyle(s, includeMono: false);

        if (!string.IsNullOrEmpty(styleAttr))
            text = $"<span style=\"{styleAttr}\">{text}</span>";

        if (isMono)            text = $"<code>{text}</code>";
        if (s.Subscript)       text = $"<sub>{text}</sub>";
        if (s.Superscript)     text = $"<sup>{text}</sup>";
        if (s.Underline)       text = $"<u>{text}</u>";
        if (s.Strikethrough)   text = $"<s>{text}</s>";
        if (s.Italic)          text = $"<em>{text}</em>";
        if (s.Bold)            text = $"<strong>{text}</strong>";

        if (run.Url is { Length: > 0 } href)
            text = $"<a href=\"{EscapeAttr(href)}\">{text}</a>";

        return text;
    }

    private static string BuildSpanStyle(RunStyle s, bool includeMono)
    {
        var parts = new List<string>(4);
        if (!string.IsNullOrEmpty(s.FontFamily) &&
            (includeMono || !s.FontFamily.Contains("monospace", StringComparison.OrdinalIgnoreCase)))
        {
            parts.Add($"font-family:{s.FontFamily}");
        }
        // 11 = 기본값 — 명시 변경 시에만 출력.
        if (Math.Abs(s.FontSizePt - 11) > 0.01 && s.FontSizePt > 0)
            parts.Add($"font-size:{s.FontSizePt.ToString("0.##", CultureInfo.InvariantCulture)}pt");
        if (s.Foreground is { } fg) parts.Add($"color:{ColorHex(fg)}");
        if (s.Background is { } bg) parts.Add($"background-color:{ColorHex(bg)}");
        return string.Join(';', parts);
    }

    private static string ColorHex(Color c) => $"#{c.R:X2}{c.G:X2}{c.B:X2}";

    // ── 이스케이프 ────────────────────────────────────────────────────

    private static string EscapeHtml(string text)
    {
        if (string.IsNullOrEmpty(text)) return text;
        var sb = new StringBuilder(text.Length);
        foreach (var ch in text)
        {
            switch (ch)
            {
                case '&': sb.Append("&amp;");  break;
                case '<': sb.Append("&lt;");   break;
                case '>': sb.Append("&gt;");   break;
                default:  sb.Append(ch);       break;
            }
        }
        return sb.ToString();
    }

    private static string EscapeAttr(string text)
    {
        if (string.IsNullOrEmpty(text)) return text;
        var sb = new StringBuilder(text.Length);
        foreach (var ch in text)
        {
            switch (ch)
            {
                case '&':  sb.Append("&amp;");  break;
                case '<':  sb.Append("&lt;");   break;
                case '>':  sb.Append("&gt;");   break;
                case '"':  sb.Append("&quot;"); break;
                case '\'': sb.Append("&#39;");  break;
                default:   sb.Append(ch);       break;
            }
        }
        return sb.ToString();
    }
}
