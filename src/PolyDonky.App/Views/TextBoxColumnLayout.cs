using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using PolyDonky.App.Services;
using PolyDonky.Core;
using WpfDocs = System.Windows.Documents;

namespace PolyDonky.App.Views;

/// <summary>
/// 글상자 내부 다단 레이아웃 헬퍼. 텍스트 박스의 <see cref="TextBoxObject.Content"/>
/// 블록들을 N 개 단(slot)으로 분배하고 단별 FlowDocument 를 만든다.
/// 본문 다단(<see cref="Pagination.FlowDocumentPaginationAdapter"/>) 과 같은 원리지만
/// "단일 페이지 × 다단" 으로 단순화 (페이지 개념 없음).
/// </summary>
public static class TextBoxColumnLayout
{
    /// <summary>
    /// 빈 단 placeholder 단락의 Id 마커. SyncMultiColContentToModel 에서 빈 placeholder 만
    /// 골라 제거하기 위함 — 사용자 입력으로 채워지면 Id 가 클리어된다.
    /// </summary>
    public const string PlaceholderId = "§tbph";

    /// <summary>분배 결과 한 단의 정보.</summary>
    public sealed class ColumnSlice
    {
        public required int                  ColumnIndex   { get; init; }
        public required double               WidthDip      { get; init; }
        public required double               HeightDip     { get; init; }
        public required double               XOffsetDip    { get; init; }
        public required IReadOnlyList<Block> Blocks        { get; init; }
        public required WpfDocs.FlowDocument FlowDocument  { get; init; }
    }

    /// <summary>
    /// <paramref name="content"/> 를 <paramref name="columnCount"/> 개 단으로 분배한다.
    /// 본문 다단의 <c>MapBodyBlocksToPages</c> + <c>SplitCoreParagraph</c> 와 동일한 원리:
    /// 단 폭을 측정 폭으로 잡은 오프스크린 RTB 에서 블록의 Y 좌표를 잰 뒤
    /// <c>slotIdx = floor(topY / colHeight)</c> 로 단 인덱스를 정한다.
    /// 경계를 가로지르는 단락은 줄 단위로 분할(§f0/§f1 접미사) — 본문 다단과 동일.
    /// </summary>
    public static IReadOnlyList<ColumnSlice> Distribute(
        IList<Block> content,
        int          columnCount,
        double       innerWidthDip,
        double       innerHeightDip,
        double       columnGapDip,
        IList<double>? columnWidthsDip = null)
    {
        ArgumentNullException.ThrowIfNull(content);
        if (columnCount < 1) columnCount = 1;
        if (innerWidthDip  < 1) innerWidthDip  = 1;
        if (innerHeightDip < 1) innerHeightDip = 1;

        // 단 너비 결정 (균등 또는 사용자 지정)
        double[] colWidths;
        if (columnCount > 1
            && columnWidthsDip is { Count: > 0 } wList
            && wList.Count == columnCount)
        {
            colWidths = wList.Select(w => Math.Max(10.0, w)).ToArray();
        }
        else if (columnCount > 1)
        {
            double equalW = Math.Max(10.0,
                (innerWidthDip - columnGapDip * (columnCount - 1)) / columnCount);
            colWidths = Enumerable.Repeat(equalW, columnCount).ToArray();
        }
        else
        {
            colWidths = new[] { innerWidthDip };
        }

        double[] colXOffsets = new double[columnCount];
        double offset = 0;
        for (int i = 0; i < columnCount; i++)
        {
            colXOffsets[i] = offset;
            offset += colWidths[i] + (columnCount > 1 ? columnGapDip : 0);
        }

        // 단일 단이거나 콘텐츠가 없으면 단순 처리
        if (columnCount == 1)
        {
            var fdSingle = FlowDocumentBuilder.BuildFromBlocks(content, page: null);
            fdSingle.PageWidth   = colWidths[0];
            fdSingle.PagePadding = new Thickness(0);
            return new[]
            {
                new ColumnSlice
                {
                    ColumnIndex  = 0,
                    WidthDip     = colWidths[0],
                    HeightDip    = innerHeightDip,
                    XOffsetDip   = 0,
                    Blocks       = content.ToArray(),
                    FlowDocument = fdSingle,
                }
            };
        }

        // 오프스크린 RTB 측정 — 단 폭 = 첫 단 너비 (서로 다른 너비 시 근사치).
        // 본문 다단의 측정 방식과 동일.
        var measureFd = FlowDocumentBuilder.BuildFromBlocks(content, page: null);
        measureFd.PageWidth   = colWidths[0];
        measureFd.PagePadding = new Thickness(0);

        var measureRtb = new RichTextBox
        {
            Document          = measureFd,
            Padding           = new Thickness(0),
            BorderThickness   = new Thickness(0),
            VerticalAlignment = VerticalAlignment.Top,
        };
        measureRtb.Measure(new Size(colWidths[0], double.PositiveInfinity));
        measureRtb.Arrange(new Rect(measureRtb.DesiredSize));
        measureRtb.UpdateLayout();

        // 본문과 동일한 mm→DIP 반올림 오차 흡수용 허용치
        const double BoundaryTol = 2.0;

        var slotBlocks = new List<List<Block>>();
        for (int i = 0; i < columnCount; i++) slotBlocks.Add(new List<Block>());
        var slotFill = new Dictionary<int, double>();

        foreach (var wpfBlock in FlattenBlocks(measureFd.Blocks))
        {
            if (wpfBlock.Tag is not Block coreBlock) continue;

            double topY    = TryGetTopY(wpfBlock);
            double bottomY = TryGetBottomY(wpfBlock);
            if (double.IsNaN(topY))
            {
                slotBlocks[0].Add(coreBlock);
                continue;
            }

            double blockH  = (!double.IsNaN(bottomY) && bottomY > topY) ? (bottomY - topY) : 0.0;
            int    slotTop = Math.Max(0, (int)(topY / innerHeightDip));

            // 줄 단위 분할 — 본문과 동일한 조건/원리.
            if (coreBlock is Paragraph corePara
                && corePara.Style.ListMarker == null
                && blockH > 0 && blockH < innerHeightDip
                && wpfBlock is WpfDocs.Paragraph wpfPara
                && slotTop < columnCount)
            {
                double slotBoundaryY = (slotTop + 1) * innerHeightDip;
                bool   crossesBoundary = !double.IsNaN(bottomY) && bottomY > slotBoundaryY + BoundaryTol;
                bool   fillOverflow    = slotFill.GetValueOrDefault(slotTop, 0.0) + blockH > innerHeightDip + BoundaryTol;

                if (crossesBoundary || fillOverflow)
                {
                    double splitY = crossesBoundary
                        ? slotBoundaryY
                        : topY + Math.Max(0.0, innerHeightDip - slotFill.GetValueOrDefault(slotTop, 0.0));

                    int splitCharOffset = Pagination.FlowDocumentPaginationAdapter
                        .FindSplitCharOffsetPublic(wpfPara, splitY);
                    int totalChars = corePara.Runs.Sum(r => r.Text.Length);

                    if (splitCharOffset > 0 && splitCharOffset < totalChars)
                    {
                        var (frag1, frag2) = Pagination.FlowDocumentPaginationAdapter
                            .SplitCoreParagraphPublic(corePara, splitCharOffset);

                        double frag1H = Math.Max(0.0, splitY - topY);
                        slotFill[slotTop] = Math.Min(innerHeightDip,
                            slotFill.GetValueOrDefault(slotTop, 0.0) + frag1H);
                        if (slotTop < columnCount) slotBlocks[slotTop].Add(frag1);

                        int    nextSlot = slotTop + 1;
                        double frag2H   = blockH - frag1H;
                        if (frag2H > 0 && frag2H < innerHeightDip)
                        {
                            while (nextSlot < columnCount
                                && slotFill.GetValueOrDefault(nextSlot, 0.0) + frag2H > innerHeightDip + BoundaryTol)
                                nextSlot++;
                        }
                        if (frag2H > 0 && nextSlot < columnCount)
                            slotFill[nextSlot] = Math.Min(innerHeightDip,
                                slotFill.GetValueOrDefault(nextSlot, 0.0) + Math.Min(frag2H, innerHeightDip));
                        if (nextSlot < columnCount)
                            slotBlocks[nextSlot].Add(frag2);
                        else
                            // 마지막 단을 넘어가는 frag2 는 마지막 단에 클립 (= 콘텐츠 잘림 — 사용자가 박스 키우거나 단 수 늘려야)
                            slotBlocks[columnCount - 1].Add(frag2);
                        continue;
                    }
                }
            }

            // 단일 블록 배정
            if (!double.IsNaN(bottomY)
                && bottomY > (slotTop + 1) * innerHeightDip + BoundaryTol
                && blockH < innerHeightDip)
            {
                slotTop += 1;
            }
            if (blockH > 0 && blockH < innerHeightDip)
            {
                while (slotTop < columnCount - 1
                    && slotFill.GetValueOrDefault(slotTop, 0.0) + blockH > innerHeightDip + BoundaryTol)
                    slotTop += 1;
            }
            if (slotTop >= columnCount) slotTop = columnCount - 1;

            if (blockH > 0)
                slotFill[slotTop] = Math.Min(innerHeightDip,
                    slotFill.GetValueOrDefault(slotTop, 0.0) + Math.Min(blockH, innerHeightDip));
            slotBlocks[slotTop].Add(coreBlock);
        }

        // 측정용 RTB 분리
        measureRtb.Document = new WpfDocs.FlowDocument();

        // 단별 FlowDocument 생성
        var slices = new ColumnSlice[columnCount];
        for (int col = 0; col < columnCount; col++)
        {
            var blocks = slotBlocks[col];
            // 빈 단도 편집 가능하도록 placeholder 단락 1개 추가. PlaceholderId 마커로
            // SyncMultiColContentToModel 이 사용자 미입력 placeholder 를 골라 제거할 수 있게 함.
            if (blocks.Count == 0) blocks.Add(new Paragraph { Id = PlaceholderId });
            var fd = FlowDocumentBuilder.BuildFromBlocks(blocks, page: null);
            fd.PageWidth   = colWidths[col];
            fd.PagePadding = new Thickness(0);
            slices[col] = new ColumnSlice
            {
                ColumnIndex  = col,
                WidthDip     = colWidths[col],
                HeightDip    = innerHeightDip,
                XOffsetDip   = colXOffsets[col],
                Blocks       = blocks.ToArray(),
                FlowDocument = fd,
            };
        }
        return slices;
    }

    private static IEnumerable<WpfDocs.Block> FlattenBlocks(WpfDocs.BlockCollection blocks)
    {
        foreach (var b in blocks)
        {
            yield return b;
            if (b is WpfDocs.List list)
                foreach (var li in list.ListItems)
                    foreach (var nested in FlattenBlocks(li.Blocks))
                        yield return nested;
        }
    }

    private static double TryGetTopY(WpfDocs.Block block)
    {
        try
        {
            var r = block.ContentStart.GetCharacterRect(WpfDocs.LogicalDirection.Forward);
            if (r == Rect.Empty || double.IsInfinity(r.Y) || double.IsNaN(r.Y)) return double.NaN;
            return r.Y;
        }
        catch { return double.NaN; }
    }

    private static double TryGetBottomY(WpfDocs.Block block)
    {
        try
        {
            var r = block.ContentEnd.GetCharacterRect(WpfDocs.LogicalDirection.Backward);
            if (r == Rect.Empty || double.IsNaN(r.Bottom) || double.IsInfinity(r.Bottom))
                return double.NaN;
            return r.Bottom;
        }
        catch { return double.NaN; }
    }
}
