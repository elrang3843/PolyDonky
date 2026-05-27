using System.Text;
using PolyDonky.Convert.Common;
using PolyDonky.Convert.Doc;
using PolyDonky.Core;
using PolyDonky.Iwpf;

// PolyDonky.Convert.Doc — IWPF ↔ RTF / DOC (Word 97-2003 binary) 변환 콘솔 도구.
// CLAUDE.md §3 의 외부 변환 모듈 분리 원칙: 메인 앱은 IWPF/MD/TXT 만 직접 처리하고
// RTF / DOC 는 이 CLI 가 처리한다.
//
// 변환 파이프라인:
//   *.rtf  → *.iwpf : DocReader       → IwpfWriter
//   *.doc  → *.iwpf : DocBinaryReader → IwpfWriter   (Phase 1a — 텍스트·단락; 서식은 향후 PR)
//   *.iwpf → *.rtf  : IwpfReader      → DocWriter
//
// 참고: .doc → .rtf 또는 .doc 쓰기는 미구현. DOC export 는 호환성 비용이 커서 보류.
//
// 사용법:
//   PolyDonky.Convert.Doc <input> <output>
//   PolyDonky.Convert.Doc --version | -v
//   PolyDonky.Convert.Doc --help    | -h | /?
//
// 종료 코드:
//   0 성공  2 인자 오류  3 지원하지 않는 변환 쌍
//   4 입출력 실패  5 변환 실패
// (상수는 PolyDonky.Convert.Common.ConverterExitCodes 에 정의됨)
//
// RTF 형식 특징:
//   - 배경색(하이라이트) 완벽 지원
//   - Word 97 이상에서 100% 호환
//   - 텍스트 기반 형식, 가볍고 안정적

try { Console.OutputEncoding = Encoding.UTF8; } catch { }

// ── 공통 옵션 파싱 ──────────────────────────────────────────────────
var parsed     = ConverterArgs.Parse(args);
var positional = parsed.Positional;
if (parsed.DebugLog)
    Console.Error.WriteLine("[DEBUG] 진단 로그 활성화");

if (positional.Length == 1 && (positional[0] is "--version" or "-v"))
{
    Console.WriteLine("PolyDonky.Convert.Doc 1.0");
    return ConverterExitCodes.Ok;
}

if (positional.Length == 1 && (positional[0] is "--help" or "-h" or "/?"))
{
    PrintHelp();
    return ConverterExitCodes.Ok;
}

if (positional.Length != 2)
{
    Console.Error.WriteLine("Usage: PolyDonky.Convert.Doc <input> <output> [--debug]");
    Console.Error.WriteLine("  Supported: .rtf → .iwpf  (import)");
    Console.Error.WriteLine("             .doc → .iwpf  (import, Word 97-2003 OLE2 binary, Phase 1a — 텍스트·단락)");
    Console.Error.WriteLine("             .iwpf → .rtf  (export)");
    return ConverterExitCodes.BadArgs;
}

string inPath, outPath;
try
{
    inPath  = Path.GetFullPath(positional[0]);
    outPath = Path.GetFullPath(positional[1]);
}
catch (Exception ex)
{
    Console.Error.WriteLine($"경로 해석 실패: {ex.Message}");
    return ConverterExitCodes.BadArgs;
}

string Ext(string p) => Path.GetExtension(p).TrimStart('.').ToLowerInvariant();
string inExt  = Ext(inPath);
string outExt = Ext(outPath);

if (string.Equals(inPath, outPath, StringComparison.OrdinalIgnoreCase))
{
    Console.Error.WriteLine("입력과 출력 경로가 같습니다.");
    return ConverterExitCodes.BadArgs;
}

bool isRtfImport = inExt == "rtf" && outExt == "iwpf";
bool isDocImport = inExt == "doc" && outExt == "iwpf";
bool isExport    = inExt == "iwpf" && outExt == "rtf";
bool isImport    = isRtfImport || isDocImport;
if (!isImport && !isExport)
{
    Console.Error.WriteLine($"지원하지 않는 변환: .{inExt} → .{outExt}");
    Console.Error.WriteLine("  지원: .rtf → .iwpf  (import)");
    Console.Error.WriteLine("        .doc → .iwpf  (import)");
    Console.Error.WriteLine("        .iwpf → .rtf  (export)");
    return ConverterExitCodes.UnsupportedOp;
}

if (!File.Exists(inPath))
{
    Console.Error.WriteLine($"입력 파일이 없습니다: {inPath}");
    return ConverterExitCodes.IoError;
}

if (new FileInfo(inPath).Length == 0)
{
    Console.Error.WriteLine($"입력 파일이 비어 있습니다(0 byte): {inPath}");
    return ConverterExitCodes.IoError;
}

var outDir = Path.GetDirectoryName(outPath);
if (!string.IsNullOrEmpty(outDir) && !Directory.Exists(outDir))
{
    try { Directory.CreateDirectory(outDir); }
    catch (Exception ex)
    {
        Console.Error.WriteLine($"출력 디렉터리 생성 실패: {outDir}\n  → {ex.Message}");
        return ConverterExitCodes.IoError;
    }
}

try
{
    PolyDonkyument doc;
    if (isRtfImport)
    {
        ConverterProgress.Write(0, "RTF 읽는 중");
        using (var fs = File.OpenRead(inPath))
            doc = new DocReader().Read(fs);

        ConverterProgress.Write(80, "IWPF 로 변환 중");
        using (var ofs = File.Create(outPath))
            new IwpfWriter().Write(doc, ofs);
    }
    else if (isDocImport)
    {
        ConverterProgress.Write(0, "DOC (Word 97-2003 binary) 읽는 중");
        var docReader = new DocBinaryReader();
        using (var fs = File.OpenRead(inPath))
            doc = docReader.Read(fs);

        // Phase 3n/3n-2/3n-3 — fidelity capsule 에 매크로 / 디지털 서명 / 미인식 root storage 적재.
        //   IWPF 패키지의 fidelity/capsules/msdoc/ 아래로 저장됨.
        AddDocFidelityCapsules(doc, docReader);

        ConverterProgress.Write(80, "IWPF 로 변환 중");
        using (var ofs = File.Create(outPath))
            new IwpfWriter().Write(doc, ofs);
    }
    else
    {
        ConverterProgress.Write(0, "IWPF 읽는 중");
        using (var fs = File.OpenRead(inPath))
            doc = new IwpfReader().Read(fs);

        ConverterProgress.Write(50, "RTF 로 변환 중");
        using (var ofs = File.Create(outPath))
            new DocWriter().Write(doc, ofs);
    }

    ConverterProgress.Write(100, "완료");
    Console.WriteLine($"OK: {Path.GetFileName(inPath)} → {Path.GetFileName(outPath)}");
    Console.Out.Flush();
    return ConverterExitCodes.Ok;
}
catch (FileNotFoundException ex)
{
    Console.Error.WriteLine($"파일을 찾을 수 없습니다: {ex.FileName ?? inPath}");
    if (parsed.DebugLog) Console.Error.WriteLine(ex.StackTrace);
    return ConverterExitCodes.IoError;
}
catch (IOException ex)
{
    Console.Error.WriteLine($"I/O 실패: {ex.Message}");
    if (parsed.DebugLog) Console.Error.WriteLine(ex.StackTrace);
    return ConverterExitCodes.IoError;
}
catch (Exception ex)
{
    Console.Error.WriteLine($"변환 실패: {ex.GetType().Name}: {ex.Message}");
    Console.Error.WriteLine(ex.StackTrace);
    return ConverterExitCodes.ConvertError;
}
finally
{
    Console.Error.Flush();
    Console.Out.Flush();
}

static void PrintHelp()
{
    Console.WriteLine("PolyDonky.Convert.Doc — IWPF ↔ RTF 변환기");
    Console.WriteLine();
    Console.WriteLine("사용법:");
    Console.WriteLine("  PolyDonky.Convert.Doc <input> <output> [--debug]");
    Console.WriteLine();
    Console.WriteLine("옵션:");
    Console.WriteLine("  --debug | -d | DEBUG  예외 스택 트레이스 등 상세 진단 출력");
    Console.WriteLine();
    Console.WriteLine("변환 쌍:");
    Console.WriteLine("  *.rtf  → *.iwpf : import (텍스트·서식 지원)");
    Console.WriteLine("  *.doc  → *.iwpf : import (Word 97-2003 OLE2 binary; Phase 1a — 텍스트·단락)");
    Console.WriteLine("  *.iwpf → *.rtf  : export (텍스트·서식·배경색 지원)");
    Console.WriteLine();
    Console.WriteLine("향후 단계:");
    Console.WriteLine("  *.doc  → *.iwpf : Phase 1b — 단락 정렬·굵게/이탤릭/밑줄/폰트크기");
    Console.WriteLine("  *.doc  → *.iwpf : Phase 2  — 표·이미지·필드·헤더/푸터·섹션");
    Console.WriteLine("  *.iwpf → *.doc  : export (보류 — 호환성 비용 큼)");
    Console.WriteLine();
    Console.WriteLine("종료 코드:");
    Console.WriteLine("  0  성공");
    Console.WriteLine("  2  인자 오류");
    Console.WriteLine("  3  지원하지 않는 변환 쌍");
    Console.WriteLine("  4  입출력 실패");
    Console.WriteLine("  5  변환 실패");
}

// Phase 3n + 3n-2 + 3n-3 — DOC 의 활성/미인식 콘텐츠를 IWPF fidelity/capsules/msdoc/ 아래로 매핑.
//   Macros          → msdoc/macros/<storageName>/<path>
//   DigitalSignature → msdoc/signatures/<storageName>/<path>
//   PreservedRootStorages → msdoc/preserved/<storageName>/<path>
// 절대 파싱·실행하지 않으며 raw bytes 그대로 ZIP 안에 격리 저장.
static void AddDocFidelityCapsules(PolyDonkyument doc, DocBinaryReader reader)
{
    if (reader.MacroProject is { } macros)
    {
        foreach (var kv in macros.Streams)
            doc.FidelityCapsules[$"msdoc/macros/{macros.StorageName}/{kv.Key}"] = kv.Value;
    }
    if (reader.DigitalSignature is { } sig)
    {
        foreach (var kv in sig.Streams)
            doc.FidelityCapsules[$"msdoc/signatures/{sig.StorageName}/{kv.Key}"] = kv.Value;
    }
    foreach (var storage in reader.PreservedRootStorages)
    {
        foreach (var kv in storage.Streams)
            doc.FidelityCapsules[$"msdoc/preserved/{storage.Name}/{kv.Key}"] = kv.Value;
    }
}
