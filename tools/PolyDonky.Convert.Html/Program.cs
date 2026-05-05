using PolyDonky.Codecs.Html;
using PolyDonky.Core;
using PolyDonky.Iwpf;

// PolyDonky.Convert.Html — HTML ↔ IWPF 변환 전용 콘솔 도구.
// CLAUDE.md §3 의 외부 변환 모듈 분리 원칙에 따라 메인 앱이 HTML 을 직접 처리하지 않고
// 이 CLI 를 spawn 한다.
//
// 사용법:
//   PolyDonky.Convert.Html <input> <output>
//     - input  *.html|*.htm, output *.iwpf  → import (HTML → IWPF)
//     - input  *.iwpf,        output *.html|*.htm → export (IWPF → HTML)
//   PolyDonky.Convert.Html --version
//
// 종료 코드:
//   0 성공, 2 인자 오류, 3 지원 안 함, 4 입출력 실패, 5 변환 실패.

if (args.Length == 1 && (args[0] == "--version" || args[0] == "-v"))
{
    Console.WriteLine("PolyDonky.Convert.Html 1.0");
    return 0;
}

if (args.Length != 2)
{
    Console.Error.WriteLine("Usage: PolyDonky.Convert.Html <input> <output>");
    Console.Error.WriteLine("  Supported pairs: .html|.htm <-> .iwpf");
    return 2;
}

string inPath  = args[0];
string outPath = args[1];

string Ext(string p) => Path.GetExtension(p).TrimStart('.').ToLowerInvariant();
string inExt  = Ext(inPath);
string outExt = Ext(outPath);

if (!File.Exists(inPath))
{
    Console.Error.WriteLine($"입력 파일이 없습니다: {inPath}");
    return 4;
}

try
{
    if ((inExt is "html" or "htm") && outExt == "iwpf")
    {
        // import: HTML → IWPF
        WriteProgress(0, "HTML 읽는 중");
        var reader = new HtmlReader { MaxBlocks = 0 };
        PolyDonkyument doc;
        using (var fs = File.OpenRead(inPath))
            doc = reader.Read(fs);

        WriteProgress(60, "IWPF 로 변환 중");
        using var ofs = File.Create(outPath);
        new IwpfWriter().Write(doc, ofs);
        WriteProgress(100, "완료");
        Console.WriteLine($"OK: {Path.GetFileName(inPath)} → {Path.GetFileName(outPath)}");
        return 0;
    }

    if (inExt == "iwpf" && (outExt is "html" or "htm"))
    {
        // export: IWPF → HTML
        WriteProgress(0, "IWPF 읽는 중");
        PolyDonkyument doc;
        using (var fs = File.OpenRead(inPath))
            doc = new IwpfReader().Read(fs);

        WriteProgress(60, "HTML 로 변환 중");
        using var ofs = File.Create(outPath);
        new HtmlWriter().Write(doc, ofs);
        WriteProgress(100, "완료");
        Console.WriteLine($"OK: {Path.GetFileName(inPath)} → {Path.GetFileName(outPath)}");
        return 0;
    }

    Console.Error.WriteLine($"지원하지 않는 변환: .{inExt} → .{outExt}");
    return 3;
}
catch (FileNotFoundException ex)
{
    Console.Error.WriteLine($"파일 없음: {ex.FileName}");
    return 4;
}
catch (UnauthorizedAccessException ex)
{
    Console.Error.WriteLine($"권한 거부: {ex.Message}");
    return 4;
}
catch (Exception ex)
{
    Console.Error.WriteLine($"변환 실패: {ex.GetType().Name}: {ex.Message}");
    return 5;
}

// ── 진행 막대 출력 ────────────────────────────────────────────────────
// 메인 앱(Services/ExternalConverter) 가 stdout 을 한 줄씩 읽고 "PROGRESS:" 접두로
// 시작하는 줄을 파싱해 IProgress 로 보고한다. 콘솔에서 직접 실행해도 사람이 읽을 수 있게
// 메시지는 평문 그대로 — 메인 앱이 ":" 로 split 해 percent 와 message 를 분리.
static void WriteProgress(int percent, string message)
{
    Console.WriteLine($"PROGRESS:{percent}:{message}");
    Console.Out.Flush();
}
