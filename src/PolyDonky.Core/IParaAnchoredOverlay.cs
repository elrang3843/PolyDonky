namespace PolyDonky.Core;

/// <summary>
/// HWP VertRel=paragraph(단락 기준) 앵커링이 해소되기 전 상태의 오버레이 개체가
/// 구현하는 인터페이스.
///
/// <see cref="AnchorParagraphBlockIndex"/> 가 0 이상인 동안에는 <c>OverlayYMm</c> 이
/// 미확정이며, <c>FlowDocumentPaginationAdapter.Paginate()</c> 가 레이아웃 직후
/// 실제 페이지 Y 로 대체한 뒤 <c>-1</c> 로 초기화한다.
/// </summary>
public interface IParaAnchoredOverlay : IOverlayAnchored
{
    /// <summary>
    /// 앵커 단락의 <c>Section.Blocks</c> 인덱스.
    /// <list type="bullet">
    ///   <item><c>-1</c> — 단락 앵커링 없음(해소 완료 또는 해당 없음).</item>
    ///   <item><c>-2</c> — 보류 sentinel — 다음 GSO-anchor 단락이 확정되면 실제 인덱스로 교체된다.</item>
    ///   <item><c>≥ 0</c> — 해소 대기 중인 앵커 인덱스.</item>
    /// </list>
    /// </summary>
    int AnchorParagraphBlockIndex { get; set; }

    /// <summary>앵커 단락 상단으로부터의 Y 오프셋 (mm). HWP raw 좌표 그대로 저장한다.</summary>
    double AnchorRelativeYMm { get; set; }
}
