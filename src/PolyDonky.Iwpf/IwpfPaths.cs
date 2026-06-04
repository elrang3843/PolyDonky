namespace PolyDonky.Iwpf;

/// <summary>IWPF 패키지 내 표준 파트 경로. 슬래시는 ZIP 표준에 맞춰 항상 forward slash.</summary>
public static class IwpfPaths
{
    public const string Manifest = "manifest.json";
    public const string DocumentJson = "content/document.json";
    public const string StylesJson = "content/styles.json";
    public const string ProvenanceJson = "provenance/source-map.json";

    public const string ResourcesDir = "resources/";
    public const string ImagesDir = "resources/images/";
    public const string OleDir = "resources/ole/";
    public const string FontsDir = "resources/fonts/";

    /// <summary>
    /// import 시점 외부 포맷 파일 스냅샷 디렉터리.
    /// 이후 사용자 편집으로 IWPF 콘텐츠와 내용이 달라질 수 있음(stale 가능).
    /// 편집 데이터의 권위 있는 출처가 아님 — 변환 검증·하이브리드 export 참고용으로만 사용.
    /// </summary>
    public const string FidelityImportSnapshotDir = "fidelity/import-snapshot/";
    /// <summary>이전 경로. 기존 .iwpf 파일 호환 읽기용으로만 유지. 신규 쓰기에 사용 금지.</summary>
    [Obsolete("Use FidelityImportSnapshotDir. This path name implies the stored data is authoritative, which is incorrect after editing.")]
    public const string FidelityOriginalDir = "fidelity/original/";
    public const string FidelityCapsulesDir = "fidelity/capsules/";

    public const string RenderDir = "render/";
    public const string SignaturesDir = "signatures/";

    public const string SecurityDir       = "security/";
    public const string SecurityEnvelope  = "security/envelope.json";
    public const string SecurityPayload   = "security/payload.bin";
    public const string SecurityWriteLock = "security/write-lock.json";
}

public static class IwpfMediaTypes
{
    public const string Manifest = "application/vnd.polydoc.iwpf.manifest+json";
    public const string Document = "application/vnd.polydoc.iwpf.document+json";
    public const string Styles = "application/vnd.polydoc.iwpf.styles+json";
    public const string Provenance = "application/vnd.polydoc.iwpf.provenance+json";
    public const string Json = "application/json";
    public const string OctetStream = "application/octet-stream";
}
