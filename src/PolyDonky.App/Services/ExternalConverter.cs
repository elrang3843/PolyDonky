using System;
using System.Diagnostics;
using System.IO;
using System.Threading.Tasks;

namespace PolyDonky.App.Services;

/// <summary>
/// CLAUDE.md §3 의 외부 변환 모듈 분리 원칙 — IWPF/MD/TXT 가 아닌 포맷은 별도 CLI 실행 파일을
/// spawn 해 IWPF 로 변환한 뒤 메인 앱이 읽는다(역방향: 저장 시 메인 앱이 IWPF 임시파일을 만들고
/// CLI 가 대상 포맷으로 변환).
///
/// 등록된 변환기:
///   .html / .htm / .xml / .xhtml — `PolyDonky.Convert.Html.exe` / `PolyDonky.Convert.Xml.exe`
/// </summary>
public static class ExternalConverter
{
    /// <summary>이 확장자가 외부 CLI 변환을 거쳐야 하는지 검사.</summary>
    public static bool IsExternalFormat(string path) => GetConverter(path) is not null;

    /// <summary>지정 확장자에 대응하는 CLI 실행 파일 절대 경로 (없으면 null).</summary>
    public static string? GetConverter(string path)
    {
        var ext = Path.GetExtension(path).TrimStart('.').ToLowerInvariant();
        var name = ext switch
        {
            "html" or "htm"   => "PolyDonky.Convert.Html",
            "xml"  or "xhtml" => "PolyDonky.Convert.Xml",
            _                 => null,
        };
        if (name is null) return null;

        var baseDir = AppContext.BaseDirectory;
        // 같은 디렉터리에 .exe 또는 .dll(framework-dependent) 가 배치된다고 가정 — 빌드 타깃이 복사함.
        foreach (var candidate in new[]
        {
            Path.Combine(baseDir, name + ".exe"),
            Path.Combine(baseDir, name + ".dll"),
        })
        {
            if (File.Exists(candidate)) return candidate;
        }
        return null;
    }

    /// <summary>
    /// CLI 변환기를 spawn 해 input → output 변환을 수행한다. 백그라운드 스레드에서 호출 권장.
    /// 실패 시 Exception throw — 호출측이 ReportError 등으로 처리.
    /// </summary>
    public static async Task ConvertAsync(string converterPath, string inputPath, string outputPath)
    {
        ArgumentNullException.ThrowIfNull(converterPath);
        ArgumentNullException.ThrowIfNull(inputPath);
        ArgumentNullException.ThrowIfNull(outputPath);

        bool isDll = converterPath.EndsWith(".dll", StringComparison.OrdinalIgnoreCase);
        var psi = new ProcessStartInfo
        {
            FileName  = isDll ? "dotnet" : converterPath,
            UseShellExecute        = false,
            RedirectStandardOutput = true,
            RedirectStandardError  = true,
            CreateNoWindow         = true,
        };
        if (isDll) psi.ArgumentList.Add(converterPath);
        psi.ArgumentList.Add(inputPath);
        psi.ArgumentList.Add(outputPath);

        using var proc = Process.Start(psi)
            ?? throw new InvalidOperationException($"외부 변환기 실행 실패: {converterPath}");

        // 출력은 흡수 — 안 흡수하면 PIPE 가 가득 차서 block 될 수 있음.
        var stdoutTask = proc.StandardOutput.ReadToEndAsync();
        var stderrTask = proc.StandardError.ReadToEndAsync();
        await proc.WaitForExitAsync().ConfigureAwait(false);
        var stderr = await stderrTask.ConfigureAwait(false);
        _ = await stdoutTask.ConfigureAwait(false);

        if (proc.ExitCode != 0)
        {
            throw new InvalidOperationException(
                $"외부 변환 실패 (종료 코드 {proc.ExitCode}): {converterPath}\n{stderr}");
        }
    }

    /// <summary>임시 IWPF 파일 경로 — Path.GetTempPath() 아래 고유한 이름.</summary>
    public static string CreateTempIwpfPath()
        => Path.Combine(Path.GetTempPath(), $"polydonky-{Guid.NewGuid():N}.iwpf");
}
