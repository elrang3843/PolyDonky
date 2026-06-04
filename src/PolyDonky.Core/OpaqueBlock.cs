namespace PolyDonky.Core;

/// <summary>
/// "보존 섬(opaque island)" — <b>임시 퇴피 상태</b>. 영구 상태가 아니다.
///
/// <para>
/// 아직 1급 Core Block 서브클래스로 끌어올리지 못한 외부 포맷 요소(예: DOCX SDT,
/// 미인식 도형, OLE 개체 등)를 원본 조각 그대로 보관해 라운드트립 시 재출력한다.
/// 에디터에서는 read-only 자리표시자로 표시한다.
/// </para>
///
/// <para>
/// <b>원칙(CLAUDE.md §4)</b>: 구현 역량이 확보되는 시점에 적절한 1급 Block 서브클래스로
/// 승격(promote)해야 한다. 새 OpaqueBlock 종류가 생기면 WORK_PLAN.md 의
/// "IWPF 전완전성 추적" 백로그에 반드시 등록한다.
/// </para>
/// </summary>
public sealed class OpaqueBlock : Block
{
    /// <summary>예: "docx", "hwpx", "doc-binary", "html". 원본 포맷 식별자.</summary>
    public string Format { get; set; } = string.Empty;

    /// <summary>원본 형식 안에서의 하위 종류 (예: "drawing", "sdt", "altChunk").</summary>
    public string? Kind { get; set; }

    /// <summary>XML 기반 포맷이면 원본 OuterXml. 비어 있으면 <see cref="Bytes"/> 사용.</summary>
    public string? Xml { get; set; }

    /// <summary>바이너리 포맷의 원본 조각.</summary>
    public byte[]? Bytes { get; set; }

    /// <summary>편집기에 표시할 사용자 친화적 라벨. 예: "[표]", "[그림]", "[도형]".</summary>
    public string DisplayLabel { get; set; } = "[보존된 개체]";

    public OpaqueBlock()
    {
        Status = NodeStatus.Opaque;
    }
}
