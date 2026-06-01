using System.Linq;
using System.Threading;
using System.Windows;
using PolyDonky.App.Services;
using PolyDonky.Core;
using Wpf = System.Windows.Documents;
using WpfMedia = System.Windows.Media;

namespace PolyDonky.App.Tests;

public class FlowDocumentRoundTripTests
{
    [Fact]
    public void Build_PreservesFontAndSizeOnRun()
    {
        var doc = SingleParagraph(text: "큰 글씨", style: new RunStyle
        {
            FontFamily = "맑은 고딕",
            FontSizePt = 18,
        });

        var fd = FlowDocumentBuilder.Build(doc);
        var run = FirstWpfRun(fd);

        Assert.Equal("맑은 고딕", run.FontFamily.Source);
        Assert.Equal(FlowDocumentBuilder.PtToDip(18), run.FontSize, precision: 1);
    }

    [Fact]
    public void Build_PreservesBoldItalicUnderline()
    {
        var doc = SingleParagraph(text: "ABC", style: new RunStyle
        {
            Bold = true,
            Italic = true,
            Underline = true,
            Strikethrough = true,
        });

        var run = FirstWpfRun(FlowDocumentBuilder.Build(doc));

        Assert.Equal(FontWeights.Bold, run.FontWeight);
        Assert.Equal(FontStyles.Italic, run.FontStyle);
        Assert.NotNull(run.TextDecorations);
        Assert.Contains(run.TextDecorations!, d => d.Location == TextDecorationLocation.Underline);
        Assert.Contains(run.TextDecorations!, d => d.Location == TextDecorationLocation.Strikethrough);
    }

    [Fact]
    public void Build_PreservesForegroundColor()
    {
        var doc = SingleParagraph(text: "red", style: new RunStyle
        {
            Foreground = Color.FromHex("#FF3300"),
        });

        var run = FirstWpfRun(FlowDocumentBuilder.Build(doc));

        Assert.IsType<WpfMedia.SolidColorBrush>(run.Foreground);
        var brush = (WpfMedia.SolidColorBrush)run.Foreground!;
        Assert.Equal(0xFF, brush.Color.R);
        Assert.Equal(0x33, brush.Color.G);
        Assert.Equal(0x00, brush.Color.B);
    }

    [Fact]
    public void Build_HeadingParagraphHasLargerFontAndBold()
    {
        var p = new Paragraph();
        p.Style.Outline = OutlineLevel.H1;
        p.AddText("제목");
        var doc = WrapInDocument(p);

        var fd = FlowDocumentBuilder.Build(doc);
        var wpfPara = (Wpf.Paragraph)fd.Blocks.First();

        Assert.True(wpfPara.FontSize > FlowDocumentBuilder.PtToDip(12));
        // FlowDocumentBuilder:3207 의 매핑 — charStyle.Bold ? Bold : SemiBold.
        // OutlineStyleSet.DefaultForLevel(H1).Char.Bold = true 이므로 Bold 가 정상.
        Assert.Equal(FontWeights.Bold, wpfPara.FontWeight);
    }

    [Fact]
    public void Build_AlignmentMappedCorrectly()
    {
        var doc = new PolyDonkyument();
        var section = new Section();
        doc.Sections.Add(section);
        foreach (var alignment in new[] { Alignment.Left, Alignment.Center, Alignment.Right, Alignment.Justify })
        {
            var p = new Paragraph { Style = { Alignment = alignment } };
            p.AddText(alignment.ToString());
            section.Blocks.Add(p);
        }

        var fd = FlowDocumentBuilder.Build(doc);
        var wpfParas = fd.Blocks.OfType<Wpf.Paragraph>().ToList();

        Assert.Equal(TextAlignment.Left, wpfParas[0].TextAlignment);
        Assert.Equal(TextAlignment.Center, wpfParas[1].TextAlignment);
        Assert.Equal(TextAlignment.Right, wpfParas[2].TextAlignment);
        Assert.Equal(TextAlignment.Justify, wpfParas[3].TextAlignment);
    }

    [Fact]
    public void Build_BulletListBecomesWpfListWithDiscMarker()
    {
        var doc = new PolyDonkyument();
        var section = new Section();
        doc.Sections.Add(section);
        foreach (var text in new[] { "a", "b" })
        {
            var p = new Paragraph();
            p.Style.ListMarker = new ListMarker { Kind = ListKind.Bullet };
            p.AddText(text);
            section.Blocks.Add(p);
        }

        var fd = FlowDocumentBuilder.Build(doc);
        var list = (Wpf.List)fd.Blocks.First();

        Assert.Equal(TextMarkerStyle.Disc, list.MarkerStyle);
        Assert.Equal(2, list.ListItems.Count);
    }

    [Fact]
    public void Build_NegativeFirstLineIndent_WithoutLeftIndent_IsClampedToAvoidLeftClipping()
    {
        // DOC Heading 의 sprmPDxaLeft1 가 음수(−7.64mm)인데 좌측 들여쓰기가 0 이면
        // 첫 줄이 본문 왼쪽 경계 밖으로 나가 글자가 잘린다 (CORE-3399PRO-JD4.doc "설치 드라이버").
        var p = new Paragraph();
        p.Style.Outline           = OutlineLevel.H3;
        p.Style.IndentLeftMm      = 0;
        p.Style.IndentFirstLineMm = -7.64;
        p.AddText("설치 드라이버");

        var fd = FlowDocumentBuilder.Build(WrapInDocument(p));
        var wpfPara = (Wpf.Paragraph)fd.Blocks.First();

        // 첫 줄이 페이지 본문 왼쪽 경계(x=0) 밖으로 나가지 않게 클램프.
        Assert.True(wpfPara.TextIndent >= 0, $"TextIndent={wpfPara.TextIndent} 는 음수면 안 됨");
    }

    [Fact]
    public void Build_ProperHangingIndent_IsPreserved()
    {
        // 좌측 +X, 첫 줄 −X = 정상 행잉 인덴트 — 클램프되지 않고 보존되어야 한다.
        var p = new Paragraph();
        p.Style.IndentLeftMm      = 7.64;
        p.Style.IndentFirstLineMm = -7.64;
        p.AddText("hang");

        var fd = FlowDocumentBuilder.Build(WrapInDocument(p));
        var wpfPara = (Wpf.Paragraph)fd.Blocks.First();

        Assert.True(wpfPara.TextIndent < 0, "행잉 인덴트의 음수 TextIndent 가 보존되어야 함");
        Assert.Equal(FlowDocumentBuilder.MmToDip(7.64), wpfPara.Margin.Left, precision: 1);
    }

    [Fact]
    public void Build_TableWithoutColumns_SynthesizesContentProportionalStarColumns()
    {
        // sprmTDefTable 부재 DOC 표처럼 열 정의가 없는 표 — WPF Auto 회피를 위해 Star 열을 합성하고,
        // Word AutoFit 근사로 내용이 긴 열에 더 큰 폭(Star 가중치)을 부여해야 한다.
        var table = new Table();
        var trow = new TableRow();
        var shortCell = new TableCell(); shortCell.Blocks.Add(Paragraph.Of("예"));
        var longCell  = new TableCell(); longCell.Blocks.Add(Paragraph.Of(
            "높은 전송 속도 통신, 우수한 안정성 및 내구성 지원하는 매우 긴 평가 문장입니다"));
        trow.Cells.Add(shortCell);
        trow.Cells.Add(longCell);
        table.Rows.Add(trow);

        var doc = new PolyDonkyument();
        var sec = new Section();
        sec.Blocks.Add(table);
        doc.Sections.Add(sec);

        var fd = FlowDocumentBuilder.Build(doc);
        var wtable = fd.Blocks.OfType<Wpf.Table>().First();

        Assert.Equal(2, wtable.Columns.Count);
        Assert.All(wtable.Columns, col => Assert.Equal(GridUnitType.Star, col.Width.GridUnitType));
        // 내용이 긴 열의 Star 가중치가 짧은 열보다 커야 한다 (균등 분배가 아님).
        Assert.True(wtable.Columns[1].Width.Value > wtable.Columns[0].Width.Value,
            $"긴 열({wtable.Columns[1].Width.Value}) 이 짧은 열({wtable.Columns[0].Width.Value}) 보다 넓어야 함");
    }

    [Fact]
    public void Build_TableWithDefaultBorderThickness_DrawsCellBorders()
    {
        // DOC 리더가 "맨 표" 에 부여하는 기본 0.5pt 테두리가 셀 보더로 렌더되는지 검증.
        var table = new Table { BorderThicknessPt = 0.5, BorderColor = "#808080" };
        var trow = new TableRow();
        for (int c = 0; c < 2; c++)
        {
            var cell = new TableCell();
            cell.Blocks.Add(Paragraph.Of($"c{c}"));
            trow.Cells.Add(cell);
        }
        table.Rows.Add(trow);
        var doc = new PolyDonkyument();
        var sec = new Section();
        sec.Blocks.Add(table);
        doc.Sections.Add(sec);

        var fd = FlowDocumentBuilder.Build(doc);
        var wtable = fd.Blocks.OfType<Wpf.Table>().First();
        var firstCell = wtable.RowGroups.First().Rows.First().Cells.First();

        // 0.5pt 공통 두께 → 적어도 오른쪽/아래 면은 그려져야 함 (collapse 정책상 내부 top/left 는 0 가능).
        var t = firstCell.BorderThickness;
        Assert.True(t.Right > 0 || t.Bottom > 0 || t.Left > 0 || t.Top > 0,
            $"기본 테두리가 그려지지 않음: {t}");
    }

    [Fact]
    public void RoundTrip_PreservesBoldAndColor()
    {
        var doc = SingleParagraph(text: "강조", style: new RunStyle
        {
            Bold = true,
            Foreground = Color.FromHex("#0066CC"),
        });

        var fd = FlowDocumentBuilder.Build(doc);
        var rebuilt = FlowDocumentParser.Parse(fd, originalForMerge: doc);

        var run = rebuilt.EnumerateParagraphs().Single().Runs.Single();
        Assert.True(run.Style.Bold);
        Assert.NotNull(run.Style.Foreground);
        Assert.Equal(0x00, run.Style.Foreground!.Value.R);
        Assert.Equal(0x66, run.Style.Foreground!.Value.G);
        Assert.Equal(0xCC, run.Style.Foreground!.Value.B);
    }

    [Fact]
    public void RoundTrip_PreservesAlignmentAndOutlineLevel()
    {
        var p = new Paragraph
        {
            Style = { Outline = OutlineLevel.H2, Alignment = Alignment.Center },
        };
        p.AddText("부제목");
        var doc = WrapInDocument(p);

        var fd = FlowDocumentBuilder.Build(doc);
        var rebuilt = FlowDocumentParser.Parse(fd, originalForMerge: doc);

        var rebuiltP = rebuilt.EnumerateParagraphs().Single();
        Assert.Equal(OutlineLevel.H2, rebuiltP.Style.Outline);
        Assert.Equal(Alignment.Center, rebuiltP.Style.Alignment);
    }

    [Fact]
    public void RoundTrip_PreservesKoreanTypographyExtrasViaMergeBase()
    {
        // FlowDocument 가 표현 못 하는 한글 조판 속성(WidthPercent, LetterSpacingPx)은
        // originalForMerge 를 통해 비파괴 보존되어야 한다.
        // WidthPercent != 100 / LetterSpacingPx != 0 인 Run 은 BuildScaledContainer → per-char
        // TextBlock 생성 경로를 거치는데 WPF FrameworkElement 는 STA 스레드에서만 만들 수 있다.
        RunOnSta(() =>
        {
            var p = new Paragraph();
            p.AddText("장평 90", new RunStyle { WidthPercent = 90, LetterSpacingPx = 1.5 });
            var doc = WrapInDocument(p);

            var fd = FlowDocumentBuilder.Build(doc);
            var rebuilt = FlowDocumentParser.Parse(fd, originalForMerge: doc);

            var run = rebuilt.EnumerateParagraphs().Single().Runs.Single();
            Assert.Equal(90, run.Style.WidthPercent);
            Assert.Equal(1.5, run.Style.LetterSpacingPx);
        });
    }

    /// <summary>WPF FrameworkElement (TextBlock 등) 가 등장하는 테스트 본문을 STA 스레드에서 실행.
    /// PaginationTests.RunOnSta 와 동일 패턴 — 향후 공통 헬퍼로 통합 가능.</summary>
    private static void RunOnSta(System.Action action)
    {
        System.Exception? caught = null;
        var t = new Thread(() =>
        {
            try { action(); }
            catch (System.Exception ex) { caught = ex; }
        });
        t.SetApartmentState(ApartmentState.STA);
        t.Start();
        t.Join();
        if (caught is not null)
            System.Runtime.ExceptionServices.ExceptionDispatchInfo.Capture(caught).Throw();
    }

    private static PolyDonkyument SingleParagraph(string text, RunStyle style)
    {
        var p = new Paragraph();
        p.AddText(text, style);
        return WrapInDocument(p);
    }

    private static PolyDonkyument WrapInDocument(Paragraph p)
    {
        var doc = new PolyDonkyument();
        var section = new Section();
        section.Blocks.Add(p);
        doc.Sections.Add(section);
        return doc;
    }

    private static Wpf.Run FirstWpfRun(Wpf.FlowDocument fd)
    {
        var para = fd.Blocks.OfType<Wpf.Paragraph>().First();
        return para.Inlines.OfType<Wpf.Run>().First();
    }
}
