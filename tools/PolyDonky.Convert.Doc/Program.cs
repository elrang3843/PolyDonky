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
//   *.doc  → *.iwpf : DocBinaryReader → IwpfWriter
//   *.iwpf → *.rtf  : IwpfReader      → DocWriter
//   *.iwpf → *.doc  : IwpfReader      → DocBinaryWriter (Phase F1-W2a — 본문 텍스트 골격)
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
    Console.Error.WriteLine("             .doc → .iwpf  (import, Word 97-2003 OLE2 binary)");
    Console.Error.WriteLine("             .iwpf → .rtf  (export)");
    Console.Error.WriteLine("             .iwpf → .doc  (export, Phase F1-W2a — 본문 텍스트 골격)");
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

bool isRtfImport = inExt == "rtf"  && outExt == "iwpf";
bool isDocImport = inExt == "doc"  && outExt == "iwpf";
bool isRtfExport = inExt == "iwpf" && outExt == "rtf";
bool isDocExport = inExt == "iwpf" && outExt == "doc";
bool isImport    = isRtfImport || isDocImport;
bool isExport    = isRtfExport || isDocExport;
if (!isImport && !isExport)
{
    Console.Error.WriteLine($"지원하지 않는 변환: .{inExt} → .{outExt}");
    Console.Error.WriteLine("  지원: .rtf → .iwpf  (import)");
    Console.Error.WriteLine("        .doc → .iwpf  (import)");
    Console.Error.WriteLine("        .iwpf → .rtf  (export)");
    Console.Error.WriteLine("        .iwpf → .doc  (export)");
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

// 다른 프로세스(워드/한글 등)가 파일을 잡고 있어도 읽기로 통과시키는 헬퍼.
//   File.OpenRead 기본값은 FileShare.Read 라 다른 프로세스의 쓰기 잠금에 차단되지만,
//   변환기는 어차피 읽기 전용이므로 ReadWrite|Delete 공유를 허용해도 안전하다.
static FileStream OpenReadShared(string path)
    => new FileStream(path, FileMode.Open, FileAccess.Read,
                      FileShare.ReadWrite | FileShare.Delete);

try
{
    PolyDonkyument doc;
    if (isRtfImport)
    {
        ConverterProgress.Write(0, "RTF 읽는 중");
        using (var fs = OpenReadShared(inPath))
            doc = new DocReader().Read(fs);

        ConverterProgress.Write(80, "IWPF 로 변환 중");
        using (var ofs = File.Create(outPath))
            new IwpfWriter().Write(doc, ofs);
    }
    else if (isDocImport)
    {
        ConverterProgress.Write(0, "DOC (Word 97-2003 binary) 읽는 중");
        var docReader = new DocBinaryReader();
        using (var fs = OpenReadShared(inPath))
            doc = docReader.Read(fs);

        // Phase 3n/3n-2/3n-3 — fidelity capsule 에 매크로 / 디지털 서명 / 미인식 root storage 적재.
        //   IWPF 패키지의 fidelity/capsules/msdoc/ 아래로 저장됨.
        AddDocFidelityCapsules(doc, docReader);

        if (parsed.DebugLog)
        {
            Console.Error.WriteLine("[DEBUG] === STTB FFN 폰트 목록 ===");
            for (int i = 0; i < docReader.DiagFontNames.Count; i++)
                Console.Error.WriteLine($"  [{i}] {repr(docReader.DiagFontNames[i])}");
            Console.Error.WriteLine("[DEBUG] === STSH 스타일 목록 ===");
            foreach (var (istd, stk, sti, name) in docReader.DiagStyleNames)
                Console.Error.WriteLine($"  [{istd:D3}] stk={stk} sti={sti:D3}  {repr(name)}");
        }

        ConverterProgress.Write(80, "IWPF 로 변환 중");
        using (var ofs = File.Create(outPath))
            new IwpfWriter().Write(doc, ofs);
    }
    else if (isRtfExport)
    {
        ConverterProgress.Write(0, "IWPF 읽는 중");
        using (var fs = OpenReadShared(inPath))
            doc = new IwpfReader().Read(fs);

        ConverterProgress.Write(50, "RTF 로 변환 중");
        using (var ofs = File.Create(outPath))
            new DocWriter().Write(doc, ofs);
    }
    else // isDocExport
    {
        ConverterProgress.Write(0, "IWPF 읽는 중");
        using (var fs = OpenReadShared(inPath))
            doc = new IwpfReader().Read(fs);

        // .doc 출력은 RTF 내용을 .doc 확장자로 쓴다.
        //   한컴 한글이 .doc 저장 시 실제로 생성하는 형식 (RTF + \kis94 확장)과 동일한 전략이고,
        //   Word / 한글 / Doc Viewer / Office 365 모두 .doc 확장자라도 {\rtf1 시그니처를 보고
        //   RTF 로 정상 인식한다. Word 97-2003 OLE2 바이너리(DocBinaryWriter)는 strict parser
        //   호환을 100% 맞추기 어려워, 실용적으로 검증된 RTF 경로를 .doc export 기본값으로 사용.
        //   (DocBinaryWriter 는 .doc ingest 라운드트립 검증·향후용으로 코드베이스에 유지.)
        ConverterProgress.Write(50, "DOC (RTF 호환) 로 변환 중");
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

static string repr(string? s) => s is null ? "(null)" : $"\"{s}\" [len={s.Length}]";

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
    Console.WriteLine("  *.doc  → *.iwpf : import (Word 97-2003 OLE2 binary, 매크로/서명 fidelity 보존)");
    Console.WriteLine("  *.iwpf → *.rtf  : export (텍스트·서식·표·이미지·도형·각주·필드·헤더/푸터)");
    Console.WriteLine("  *.iwpf → *.doc  : export (RTF 내용을 .doc 로 — 한글이 .doc 저장 시 쓰는 방식과 동일.");
    Console.WriteLine("                            Word/한글/Doc Viewer 가 {\\rtf1 시그니처로 인식해 정상 오픈)");
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
