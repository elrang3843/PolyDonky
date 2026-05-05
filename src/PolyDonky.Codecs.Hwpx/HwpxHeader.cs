using PolyDonky.Core;

namespace PolyDonky.Codecs.Hwpx;

/// <summary>
/// HWPX 의 header.xml 을 파싱한 결과. section.xml 의
/// paraPrIDRef / charPrIDRef / styleIDRef 가 가리키는 정의를 ID → PolyDonky 모델로 보관한다.
/// 한컴 오피스가 만든 hwpx 는 우리 codec 의 0~5 약속 ID 와 다른 임의 ID 를 쓰므로,
/// 이 컨텍스트를 통해 본문 paragraph/run 이 실제 서식을 회수한다.
/// </summary>
internal sealed class HwpxHeader
{
    public Dictionary<int, string> Fonts { get; } = new();
    public Dictionary<int, RunStyle> CharProperties { get; } = new();
    public Dictionary<int, ParagraphStyle> ParaProperties { get; } = new();
    public Dictionary<int, HwpxStyleDef> Styles { get; } = new();

    public RunStyle GetRunStyle(int? charPrId)
    {
        if (charPrId is { } id && CharProperties.TryGetValue(id, out var src))
        {
            return CloneRunStyle(src);
        }
        return new RunStyle();
    }

    public ParagraphStyle? GetParagraphStyle(int? paraPrId)
    {
        if (paraPrId is { } id && ParaProperties.TryGetValue(id, out var src))
        {
            return CloneParagraphStyle(src);
        }
        return null;
    }

    public HwpxStyleDef? GetStyle(int? styleId)
        => styleId is { } id && Styles.TryGetValue(id, out var s) ? s : null;

    // Core 의 정식 RunStyle.Clone() 사용 — 시그니처는 backward-compat 유지.
    public static RunStyle CloneRunStyle(RunStyle s) => s.Clone();

    public static ParagraphStyle CloneParagraphStyle(ParagraphStyle s) => s.Clone();
}

/// <summary>HWPX style 정의 — paragraph 의 styleIDRef 가 가리킨다.</summary>
internal sealed class HwpxStyleDef
{
    public int? ParaPrIdRef { get; init; }
    public int? CharPrIdRef { get; init; }
    public string? Name { get; init; }
    public string? EngName { get; init; }
    public OutlineLevel Outline { get; init; } = OutlineLevel.Body;
}
