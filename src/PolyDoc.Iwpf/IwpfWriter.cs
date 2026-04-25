using System.IO.Compression;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using PolyDoc.Core;

namespace PolyDoc.Iwpf;

/// <summary>
/// PolyDocument 를 IWPF (ZIP+JSON) 패키지로 직렬화한다.
/// 매니페스트는 모든 파트의 SHA-256 해시를 포함하므로 본문 파트를 먼저
/// ZIP 에 기록한 뒤 마지막에 매니페스트를 추가한다.
/// </summary>
public sealed class IwpfWriter : IDocumentWriter
{
    public string FormatId => "iwpf";

    public void Write(PolyDocument document, Stream output)
    {
        ArgumentNullException.ThrowIfNull(document);
        ArgumentNullException.ThrowIfNull(output);

        using var archive = new ZipArchive(output, ZipArchiveMode.Create, leaveOpen: true);

        var manifest = new IwpfManifest();
        manifest.Created = document.Metadata.Created;
        manifest.Modified = document.Metadata.Modified;

        // content/document.json — 문서 본문 (스타일은 별도 파트로 분리)
        var documentBytes = SerializeDocument(document);
        AddPart(archive, manifest, IwpfPaths.DocumentJson, IwpfMediaTypes.Document, documentBytes);

        // content/styles.json — 스타일 시트
        var stylesBytes = JsonSerializer.SerializeToUtf8Bytes(document.Styles, JsonDefaults.Options);
        AddPart(archive, manifest, IwpfPaths.StylesJson, IwpfMediaTypes.Styles, stylesBytes);

        // provenance/source-map.json — 노드 ↔ 원본 위치 매핑
        if (document.Provenance.NodeAnchors.Count > 0)
        {
            var provenanceBytes = JsonSerializer.SerializeToUtf8Bytes(document.Provenance, JsonDefaults.Options);
            AddPart(archive, manifest, IwpfPaths.ProvenanceJson, IwpfMediaTypes.Provenance, provenanceBytes);
        }

        // manifest.json — 항상 마지막. 본문 파트의 해시가 모두 결정된 뒤에 작성한다.
        var manifestBytes = JsonSerializer.SerializeToUtf8Bytes(manifest, JsonDefaults.Options);
        WriteEntry(archive, IwpfPaths.Manifest, manifestBytes);
    }

    private static byte[] SerializeDocument(PolyDocument document)
    {
        // 패키지에서는 styles 와 provenance 를 별도 파트로 분리해 저장하므로
        // document.json 본문에는 그 둘을 비워서 직렬화한다.
        var detached = new PolyDocument
        {
            Metadata = document.Metadata,
            Sections = document.Sections,
            Styles = new StyleSheet(),
            Provenance = new Provenance(),
        };
        return JsonSerializer.SerializeToUtf8Bytes(detached, JsonDefaults.Options);
    }

    private static void AddPart(ZipArchive archive, IwpfManifest manifest, string path, string mediaType, byte[] payload)
    {
        WriteEntry(archive, path, payload);
        manifest.Parts[path] = new IwpfManifestEntry
        {
            Path = path,
            MediaType = mediaType,
            Size = payload.LongLength,
            Sha256 = Sha256Hex(payload),
        };
    }

    private static void WriteEntry(ZipArchive archive, string path, byte[] payload)
    {
        var entry = archive.CreateEntry(path, CompressionLevel.Optimal);
        using var entryStream = entry.Open();
        entryStream.Write(payload, 0, payload.Length);
    }

    private static string Sha256Hex(ReadOnlySpan<byte> data)
    {
        Span<byte> hash = stackalloc byte[SHA256.HashSizeInBytes];
        SHA256.HashData(data, hash);
        return Convert.ToHexStringLower(hash);
    }

    /// <summary>편의 메서드 — 텍스트 디버깅용으로 매니페스트만 별도 직렬화.</summary>
    public static string SerializeManifest(IwpfManifest manifest)
        => Encoding.UTF8.GetString(JsonSerializer.SerializeToUtf8Bytes(manifest, JsonDefaults.Options));
}
