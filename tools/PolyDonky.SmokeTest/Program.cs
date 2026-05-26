using System.Text;
using PolyDonky.Codecs.Docx;
using PolyDonky.Codecs.Html;
using PolyDonky.Codecs.Hwpx;
using PolyDonky.Codecs.Markdown;
using PolyDonky.Codecs.Text;
using PolyDonky.Convert.Doc;
using PolyDonky.Core;
using PolyDonky.Iwpf;
using PdXmlReader = PolyDonky.Codecs.Xml.XmlReader;
using PdXmlWriter = PolyDonky.Codecs.Xml.XmlWriter;

// PolyDonky.SmokeTest — BCL 만으로 동작하는 자체 스모크 러너.
// xUnit/NuGet 차단 환경에서도 핵심 라운드트립을 검증하기 위한 임시 도구.
// 정식 테스트는 tests/PolyDonky.*.Tests 의 xUnit 프로젝트에서 수행한다.

var harness = new SmokeHarness();

harness.Run("Plain text round-trip", PlainTextRoundTrip);
harness.Run("Markdown round-trip (headers + emphasis + lists)", MarkdownRoundTrip);
harness.Run("IWPF round-trip (manifest + integrity)", IwpfRoundTrip);
harness.Run("IWPF tampering detection", IwpfTamperingDetection);
harness.Run("DOCX round-trip (headings + emphasis)", DocxRoundTrip);
harness.Run("HWPX round-trip (KS X 6101 self interop)", HwpxRoundTrip);
harness.Run("HTML round-trip (HTML5 + tables + links)", HtmlRoundTrip);
harness.Run("XML/XHTML round-trip (well-formed XHTML5)", XmlRoundTrip);
harness.Run("DOC (RTF) round-trip — encoding/escape/empty-para", DocRtfRoundTrip);
harness.Run("DOC (RTF) read — CP949 한글 \\'XX + \\uc1 fallback", DocRtfKoreanAnsi);
harness.Run("DOC (Word 97-2003 binary) — 합성 OLE2 → 텍스트·단락 추출", DocBinaryMinimalRead);
harness.Run("DOC (Word 97-2003 binary) — MS-OFFCRYPTO 암호화 감지 후 거부", DocBinaryEncryptedRejected);
harness.Run("DOC (Word 97-2003 binary) — 비-OLE2 입력은 친절한 한국어 에러로 거부", DocBinaryNonOleRejected);
harness.Run("DOC (Word 97-2003 binary) — Phase 1b PAPX 단락 정렬 + CHPX 굵게·크기 적용", DocBinaryPapxChpxFormatting);
harness.Run("DOC (Word 97-2003 binary) — Phase 1c PAPX 들여쓰기·간격·줄간격 + CHPX 색·하이라이트", DocBinaryPhase1cFormatting);
harness.Run("DOC (Word 97-2003 binary) — Phase 1d STTB FFN + sprmCRgFtc0 폰트 패밀리", DocBinaryPhase1dFontFamily);
harness.Run("DOC (Word 97-2003 binary) — Phase 1e STSH istd → OutlineLevel (Heading 1)", DocBinaryPhase1eStshHeading);
harness.Run("DOC (Word 97-2003 binary) — Phase 1f STD grpprl 상속 (Heading 의 Bold+크기 자동)", DocBinaryPhase1fStdInheritance);
harness.Run("DOC (Word 97-2003 binary) — Phase 1g STD istdBase 체이닝 (Heading 1 ← Normal)", DocBinaryPhase1gStdChain);

return harness.Finish();


static void PlainTextRoundTrip()
{
    const string sample = "첫 번째 줄\n두 번째 줄입니다.\n세 번째 — 한글 조판 테스트.";
    var doc = PlainTextReader.FromText(sample);
    SmokeHarness.Equal(3, doc.Sections[0].Blocks.Count, "block count");

    var roundTripped = PlainTextWriter.ToText(doc);
    SmokeHarness.Equal(sample, roundTripped, "plain text body");
}

static void MarkdownRoundTrip()
{
    const string source =
        "# 제목 1\n" +
        "\n" +
        "본문 단락은 **굵게** 와 *기울임* 을 함께 가진다.\n" +
        "\n" +
        "## 부제목\n" +
        "\n" +
        "- 첫 번째 항목\n" +
        "- 두 번째 항목\n" +
        "\n" +
        "1. 순서 항목 A\n" +
        "2. 순서 항목 B\n";

    var doc = MarkdownReader.FromMarkdown(source);

    var paragraphs = doc.EnumerateParagraphs().ToList();
    SmokeHarness.Equal(7, paragraphs.Count, "paragraph count");
    SmokeHarness.Equal(OutlineLevel.H1, paragraphs[0].Style.Outline, "first outline level");
    SmokeHarness.Equal(OutlineLevel.H2, paragraphs[2].Style.Outline, "third outline level");
    SmokeHarness.Equal(ListKind.Bullet, paragraphs[3].Style.ListMarker!.Kind, "bullet list kind");
    SmokeHarness.Equal(ListKind.OrderedDecimal, paragraphs[5].Style.ListMarker!.Kind, "ordered list kind");

    // bold/italic Run 분리 확인
    var bodyRuns = paragraphs[1].Runs;
    var hasBold = bodyRuns.Any(r => r.Style.Bold && r.Text == "굵게");
    var hasItalic = bodyRuns.Any(r => r.Style.Italic && r.Text == "기울임");
    SmokeHarness.True(hasBold, "bold run present");
    SmokeHarness.True(hasItalic, "italic run present");

    // 라운드트립 후 다시 파싱했을 때 헤더/리스트 구조가 보존되는지
    var rendered = MarkdownWriter.ToMarkdown(doc);
    var reparsed = MarkdownReader.FromMarkdown(rendered);
    var reparsedParagraphs = reparsed.EnumerateParagraphs().ToList();
    SmokeHarness.Equal(paragraphs.Count, reparsedParagraphs.Count, "reparsed paragraph count");
    SmokeHarness.Equal(OutlineLevel.H1, reparsedParagraphs[0].Style.Outline, "reparsed first outline level");
    SmokeHarness.Equal(ListKind.Bullet, reparsedParagraphs[3].Style.ListMarker!.Kind, "reparsed bullet kind");
}

static void IwpfRoundTrip()
{
    var doc = new PolyDonkyument();
    doc.Metadata.Title = "스모크 테스트 문서";
    doc.Metadata.Author = "Noh JinMoon";

    var section = new Section();
    doc.Sections.Add(section);

    var heading = new Paragraph();
    heading.Style.Outline = OutlineLevel.H1;
    heading.AddText("PolyDonky IWPF 라운드트립");
    section.Blocks.Add(heading);

    var body = new Paragraph();
    body.AddText("이것은 IWPF 패키지의 ", new RunStyle());
    body.AddText("핵심", new RunStyle { Bold = true });
    body.AddText(" 라운드트립을 검증합니다.", new RunStyle());
    section.Blocks.Add(body);

    using var ms = new MemoryStream();
    new IwpfWriter().Write(doc, ms);
    SmokeHarness.True(ms.Length > 100, $"package size > 100 bytes (got {ms.Length})");

    ms.Position = 0;
    var read = new IwpfReader().Read(ms);

    SmokeHarness.Equal("스모크 테스트 문서", read.Metadata.Title!, "metadata.title");
    SmokeHarness.Equal("Noh JinMoon", read.Metadata.Author!, "metadata.author");

    var roundParagraphs = read.EnumerateParagraphs().ToList();
    SmokeHarness.Equal(2, roundParagraphs.Count, "paragraph count after read");
    SmokeHarness.Equal(OutlineLevel.H1, roundParagraphs[0].Style.Outline, "heading outline preserved");
    SmokeHarness.Equal("PolyDonky IWPF 라운드트립", roundParagraphs[0].GetPlainText(), "heading text");
    SmokeHarness.True(roundParagraphs[1].Runs.Any(r => r.Style.Bold && r.Text == "핵심"), "bold run preserved");
}

static void IwpfTamperingDetection()
{
    var doc = new PolyDonkyument();
    doc.Sections.Add(new Section());
    doc.Sections[0].Blocks.Add(Paragraph.Of("위변조 검출 테스트"));

    using var ms = new MemoryStream();
    new IwpfWriter().Write(doc, ms);

    // ZIP 내부의 document.json 페이로드를 추출/변조해 다시 ZIP 으로 묶는다.
    var tampered = TamperDocumentJson(ms.ToArray());

    using var ts = new MemoryStream(tampered);
    var caught = false;
    try
    {
        new IwpfReader().Read(ts);
    }
    catch (InvalidDataException)
    {
        caught = true;
    }
    SmokeHarness.True(caught, "tampered package was rejected");
}

static void DocxRoundTrip()
{
    var doc = new PolyDonkyument();
    doc.Metadata.Title = "DOCX 스모크";
    doc.Metadata.Author = "Noh JinMoon";
    var section = new Section();
    doc.Sections.Add(section);

    var heading = new Paragraph { Style = { Outline = OutlineLevel.H1 } };
    heading.AddText("DOCX 1급 시민");
    section.Blocks.Add(heading);

    var body = new Paragraph();
    body.AddText("OpenXml 기반 ");
    body.AddText("DOCX", new RunStyle { Bold = true });
    body.AddText(" 라운드트립을 검증합니다.");
    section.Blocks.Add(body);

    using var ms = new MemoryStream();
    new DocxWriter().Write(doc, ms);
    SmokeHarness.True(ms.Length > 1000, $"DOCX size > 1 KB (got {ms.Length})");

    ms.Position = 0;
    var read = new DocxReader().Read(ms);

    SmokeHarness.Equal("DOCX 스모크", read.Metadata.Title!, "DOCX metadata.title");
    SmokeHarness.Equal("Noh JinMoon", read.Metadata.Author!, "DOCX metadata.author");

    var paragraphs = read.EnumerateParagraphs().ToList();
    SmokeHarness.Equal(2, paragraphs.Count, "DOCX paragraph count");
    SmokeHarness.Equal(OutlineLevel.H1, paragraphs[0].Style.Outline, "DOCX heading outline");
    SmokeHarness.True(paragraphs[1].Runs.Any(r => r.Style.Bold && r.Text == "DOCX"), "DOCX bold run preserved");
}

static void HwpxRoundTrip()
{
    var doc = new PolyDonkyument();
    doc.Metadata.Title = "HWPX 스모크";
    doc.Metadata.Author = "Noh JinMoon";
    var section = new Section();
    doc.Sections.Add(section);

    var heading = new Paragraph { Style = { Outline = OutlineLevel.H2 } };
    heading.AddText("KS X 6101 자체 라운드트립");
    section.Blocks.Add(heading);

    var body = new Paragraph { Style = { Alignment = Alignment.Center } };
    body.AddText("한글 ");
    body.AddText("굵게", new RunStyle { Bold = true });
    body.AddText(" 가운데 정렬.");
    section.Blocks.Add(body);

    using var ms = new MemoryStream();
    new HwpxWriter().Write(doc, ms);
    SmokeHarness.True(ms.Length > 500, $"HWPX size > 500 bytes (got {ms.Length})");

    ms.Position = 0;
    var read = new HwpxReader().Read(ms);

    SmokeHarness.Equal("HWPX 스모크", read.Metadata.Title!, "HWPX metadata.title");
    SmokeHarness.Equal("Noh JinMoon", read.Metadata.Author!, "HWPX metadata.author");

    var paragraphs = read.EnumerateParagraphs().ToList();
    SmokeHarness.Equal(2, paragraphs.Count, "HWPX paragraph count");
    SmokeHarness.Equal(OutlineLevel.H2, paragraphs[0].Style.Outline, "HWPX heading outline");
    SmokeHarness.Equal(Alignment.Center, paragraphs[1].Style.Alignment, "HWPX center alignment");
    SmokeHarness.True(paragraphs[1].Runs.Any(r => r.Style.Bold && r.Text == "굵게"), "HWPX bold run preserved");
}

static void HtmlRoundTrip()
{
    const string source =
        "<!DOCTYPE html><html><body>" +
        "<h1>HTML 스모크</h1>" +
        "<p>본문 <strong>굵게</strong> + <a href=\"https://x\">링크</a></p>" +
        "<ul><li><input type=\"checkbox\" checked> 한 일</li>" +
        "<li><input type=\"checkbox\"> 할 일</li></ul>" +
        "<blockquote><p>인용</p></blockquote>" +
        "<pre><code class=\"language-py\">print(1)</code></pre>" +
        "<hr>" +
        "<table><thead><tr><th>A</th><th>B</th></tr></thead>" +
        "<tbody><tr><td>1</td><td style=\"text-align:right\">2</td></tr></tbody></table>" +
        "</body></html>";

    var doc = HtmlReader.FromHtml(source);

    var ps = doc.EnumerateParagraphs().ToList();
    SmokeHarness.Equal(OutlineLevel.H1, ps[0].Style.Outline, "HTML H1");
    SmokeHarness.True(ps.Any(p => p.Style.QuoteLevel >= 1),       "blockquote level recorded");
    SmokeHarness.True(doc.Sections[0].Blocks.OfType<ThematicBreakBlock>().Any(), "hr → thematic break");
    SmokeHarness.True(ps.Any(p => p.Style.CodeLanguage == "py"),  "code language preserved");
    SmokeHarness.True(ps.Any(p => p.Style.ListMarker?.Checked == true),  "task list checked");
    SmokeHarness.True(ps.Any(p => p.Style.ListMarker?.Checked == false), "task list unchecked");
    SmokeHarness.Equal(1, doc.Sections[0].Blocks.OfType<PolyDonky.Core.Table>().Count(), "table count");

    var rendered = HtmlWriter.ToHtml(doc, fullDocument: false);
    SmokeHarness.True(rendered.Contains("<h1>"),                  "writer emits h1");
    SmokeHarness.True(rendered.Contains("<a href=\"https://x\">"), "writer emits anchor");
    SmokeHarness.True(rendered.Contains("language-py"),           "writer emits code language");
    SmokeHarness.True(rendered.Contains("<input type=\"checkbox\""), "writer emits task checkbox");
    SmokeHarness.True(rendered.Contains("<thead>"),               "writer emits thead");

    var reread = HtmlReader.FromHtml(rendered);
    SmokeHarness.Equal(OutlineLevel.H1, reread.EnumerateParagraphs().First().Style.Outline, "round-trip H1");
}

static void XmlRoundTrip()
{
    var doc = new PolyDonkyument();
    doc.Sections.Add(new Section());
    var h = new Paragraph { Style = { Outline = OutlineLevel.H1 } }; h.AddText("XML 스모크");
    doc.Sections[0].Blocks.Add(h);
    var p = new Paragraph(); p.AddText("본문 ");
    p.AddText("굵게", new RunStyle { Bold = true });
    p.AddText(" + ");
    p.Runs.Add(new Run { Text = "링크", Style = new RunStyle(), Url = "https://x" });
    doc.Sections[0].Blocks.Add(p);
    doc.Sections[0].Blocks.Add(new ThematicBreakBlock());

    var xml = PdXmlWriter.ToXml(doc);
    SmokeHarness.True(xml.StartsWith("<?xml "),                       "writer emits XML declaration");
    SmokeHarness.True(xml.Contains("xmlns=\"http://www.w3.org/1999/xhtml\""), "writer emits xhtml namespace");
    SmokeHarness.True(xml.Contains("<hr/>"),                          "writer self-closes hr");
    SmokeHarness.True(xml.Contains("<meta charset=\"utf-8\"/>"),      "writer self-closes meta");

    // System.Xml 로 well-formed 확인.
    var settings = new System.Xml.XmlReaderSettings
    {
        DtdProcessing = System.Xml.DtdProcessing.Parse,
        XmlResolver   = null,
    };
    using (var sr = new StringReader(xml))
    using (var xr = System.Xml.XmlReader.Create(sr, settings))
    {
        while (xr.Read()) { }  // 형식 오류 시 예외.
    }

    var read = PdXmlReader.FromXml(xml);
    var rps  = read.EnumerateParagraphs().ToList();
    SmokeHarness.Equal(OutlineLevel.H1, rps[0].Style.Outline,        "round-trip H1");
    SmokeHarness.True(read.Sections[0].Blocks.OfType<ThematicBreakBlock>().Any(), "round-trip thematic break");
    SmokeHarness.True(rps.SelectMany(rp => rp.Runs).Any(r => r.Url == "https://x"), "round-trip link URL");
}

static void DocRtfRoundTrip()
{
    // DocWriter 가 RTF 로 직렬화한 결과를 DocReader 로 다시 읽어 의미가 보존되는지 검증.
    var doc = new PolyDonkyument();
    doc.Metadata.Title  = "DOC RTF 스모크";
    doc.Metadata.Author = "Noh JinMoon";
    var section = new Section();
    doc.Sections.Add(section);

    var heading = new Paragraph { Style = { Alignment = Alignment.Center } };
    heading.AddText("RTF 라운드트립", new RunStyle { Bold = true, FontSizePt = 16 });
    section.Blocks.Add(heading);

    var body = new Paragraph();
    body.AddText("본문 ");
    body.AddText("굵게", new RunStyle { Bold = true });
    body.AddText(" + ");
    body.AddText("기울임", new RunStyle { Italic = true });
    body.AddText(" — 한글 조판 테스트.");
    section.Blocks.Add(body);

    // 빈 단락 (사용자가 본 빈 줄) 이 보존되는지
    section.Blocks.Add(new Paragraph());

    var tail = new Paragraph();
    tail.AddText("끝 단락.");
    section.Blocks.Add(tail);

    using var ms = new MemoryStream();
    new DocWriter().Write(doc, ms);
    SmokeHarness.True(ms.Length > 200, $"RTF size > 200 bytes (got {ms.Length})");

    ms.Position = 0;
    var read = new DocReader().Read(ms);
    SmokeHarness.Equal("DOC RTF 스모크", read.Metadata.Title!, "RTF metadata.title");
    SmokeHarness.Equal("Noh JinMoon",   read.Metadata.Author!, "RTF metadata.author");

    var paragraphs = read.EnumerateParagraphs().ToList();
    // 헤딩 + 본문 + 빈 단락 + 꼬리 = 4
    SmokeHarness.Equal(4, paragraphs.Count, "RTF paragraph count (empty para preserved)");
    SmokeHarness.Equal(Alignment.Center, paragraphs[0].Style.Alignment, "RTF heading alignment");
    SmokeHarness.True(paragraphs[1].Runs.Any(r => r.Style.Bold   && r.Text.Contains("굵게")),   "RTF bold run preserved");
    SmokeHarness.True(paragraphs[1].Runs.Any(r => r.Style.Italic && r.Text.Contains("기울임")), "RTF italic run preserved");
    SmokeHarness.True(paragraphs[3].GetPlainText().Contains("끝 단락"), "RTF tail paragraph text");
}

static void DocRtfKoreanAnsi()
{
    // 한글 CP949 RTF 의 \'XX + \uc1 fallback 처리 검증.
    // 워드패드/한글 등이 저장하는 실제 RTF 와 유사한 헤더로 합성.
    // CP949: 안=BEC8 녕=B3E7 하=C7CF 세=BCBC 요=BFE4. '한' U+D55C = 54620.
    const string rtf =
        @"{\rtf1\ansi\ansicpg949\deff0" +
        @"{\fonttbl{\f0\fnil\fcharset129 Malgun Gothic;}}" +
        @"{\colortbl;\red0\green0\blue0;}" +
        @"\viewkind4\uc1\pard\f0\fs22 " +
        @"\'be\'c8\'b3\'e7\'c7\'cf\'bc\'bc\'bf\'e4" +     // '안녕하세요' (CP949 \'XX)
        @"\par " +
        // '한' (U+D55C = 54620) + ANSI fallback '?' (1글자) → 디코딩 결과 "한 END"
        @"\" + "u54620 ? END" + @"\par}";

    using var ms = new MemoryStream(Encoding.GetEncoding(28591).GetBytes(rtf));
    var read = new DocReader().Read(ms);

    var text = string.Join("\n", read.EnumerateParagraphs().Select(p => p.GetPlainText()));
    SmokeHarness.True(text.Contains("안녕하세요"), $"CP949 \\'XX 디코딩 (got: {text})");
    SmokeHarness.True(text.Contains("한"),         $"\\u54620 → '한' 디코딩 (got: {text})");
    SmokeHarness.True(text.Contains("END"),
        $"\\uc1 fallback '?' 소비 후 다음 텍스트 살아 있음 (got: {text})");
    SmokeHarness.True(!text.Contains("?"),
        $"fallback '?' 자체는 출력되지 않아야 함 (got: {text})");
}

static void DocBinaryMinimalRead()
{
    // Word 97-2003 binary 파일을 합성해 DocBinaryReader 가 텍스트·단락을 정확히 추출하는지 검증.
    // 단락 마커 '\r' 기준 분리, UTF-16LE piece 디코딩, 한글 round-trip 까지 확인.
    const string mainText = "Hello\rWorld\r한글 단락\r";  // 3 단락

    byte[] wordDocBytes;
    byte[] tableBytes;
    BuildMinimalDocStreams(mainText, out wordDocBytes, out tableBytes);

    // OLE2 컨테이너 생성 (임시 파일)
    var tmp = Path.Combine(Path.GetTempPath(), $"polydonky-smoke-{Guid.NewGuid():N}.doc");
    try
    {
        using (var root = OpenMcdf.RootStorage.Create(tmp))
        {
            using (var wd = root.CreateStream("WordDocument"))
                wd.Write(wordDocBytes);
            using (var tb = root.CreateStream("0Table"))
                tb.Write(tableBytes);
            // non-transacted 모드 — Commit 없이 Dispose 시 flush.
        }

        using var fs = File.OpenRead(tmp);
        var doc = new PolyDonky.Convert.Doc.DocBinaryReader().Read(fs);

        var paragraphs = doc.EnumerateParagraphs().ToList();
        // mainText 의 마지막 '\r' 후 빈 단락을 BuildDocument 가 만들어내므로 단락 수는 4.
        SmokeHarness.True(paragraphs.Count >= 3,
            $"단락 수 >= 3 (got {paragraphs.Count})");
        SmokeHarness.Equal("Hello",     paragraphs[0].GetPlainText(), "1st paragraph text");
        SmokeHarness.Equal("World",     paragraphs[1].GetPlainText(), "2nd paragraph text");
        SmokeHarness.Equal("한글 단락", paragraphs[2].GetPlainText(), "3rd paragraph text (Korean)");
    }
    finally
    {
        try { File.Delete(tmp); } catch { }
    }
}

static void DocBinaryEncryptedRejected()
{
    // 같은 minimal 합성에 FIB flags 의 fEncrypted (bit 8 = 0x0100) 를 켜고 던지면
    // 명세 [MS-OFFCRYPTO] 에 따라 본문 진입 전에 거부되어야 한다.
    const string mainText = "ignored\r";
    BuildMinimalDocStreams(mainText, out var wordDoc, out var table);
    // FIB flags @ 0x000A 의 fEncrypted 비트 켜기
    ushort flags = BitConverter.ToUInt16(wordDoc, 0x0A);
    flags |= 0x0100;
    BitConverter.TryWriteBytes(wordDoc.AsSpan(0x0A), flags);

    var tmp = Path.Combine(Path.GetTempPath(), $"polydonky-smoke-{Guid.NewGuid():N}.doc");
    bool rejected = false;
    string? msg = null;
    try
    {
        using (var root = OpenMcdf.RootStorage.Create(tmp))
        {
            using (var wd = root.CreateStream("WordDocument")) wd.Write(wordDoc);
            using (var tb = root.CreateStream("0Table"))       tb.Write(table);
        }
        using var fs = File.OpenRead(tmp);
        try { new PolyDonky.Convert.Doc.DocBinaryReader().Read(fs); }
        catch (InvalidOperationException ex) { rejected = true; msg = ex.Message; }
    }
    finally { try { File.Delete(tmp); } catch { } }

    SmokeHarness.True(rejected, "fEncrypted 비트 켜진 .doc 는 거부되어야 함");
    SmokeHarness.True(msg?.Contains("암호화") == true,
        $"거부 메시지는 '암호화' 안내를 포함해야 함 (got: {msg})");
}

static void DocBinaryNonOleRejected()
{
    // OLE2 가 아닌 임의 byte 배열을 .doc 로 던지면 CFB 단계에서 친절한 한국어 에러로 거부되어야 한다.
    var bogus = Encoding.UTF8.GetBytes("이건 그냥 텍스트 파일인데 누가 .doc 라고 우긴 경우.");
    using var ms = new MemoryStream(bogus);
    bool rejected = false;
    string? msg = null;
    try { new PolyDonky.Convert.Doc.DocBinaryReader().Read(ms); }
    catch (InvalidOperationException ex) { rejected = true; msg = ex.Message; }
    SmokeHarness.True(rejected, "OLE2 가 아닌 입력은 거부되어야 함");
    SmokeHarness.True(msg?.Contains("Word 97-2003") == true || msg?.Contains("OLE2") == true,
        $"거부 메시지는 형식 안내를 포함해야 함 (got: {msg})");
}

static void DocBinaryPapxChpxFormatting()
{
    // 두 단락의 합성 .doc:
    //   - 단락 1: "Bold-X"  → PAPX 정렬 Center, CHPX 처음 4글자 굵게 + 14pt
    //   - 단락 2: "Plain"   → PAPX 정렬 Right, CHPX 기본
    // FKP 들을 직접 합성해 PAPX/CHPX 운영이 실제 sprm 을 읽어 Style 에 반영하는지 검증.
    const string text = "Bold-X\rPlain\r";  // 13 chars; ccpText = 13
    int ccp = text.Length;
    var textBytes = Encoding.Unicode.GetBytes(text);
    int fcText  = 0x200;
    int fcEnd   = fcText + textBytes.Length;  // 0x21A
    int pnPapx  = 4;  // FKP page 4 → byte 0x800
    int pnChpx  = 5;  // FKP page 5 → byte 0xA00
    int fcPapxFkp = pnPapx * 512;
    int fcChpxFkp = pnChpx * 512;
    int wdSize    = fcChpxFkp + 512;  // 0xC00

    var wd = new byte[wdSize];

    // ── text ───────────────────────────────────────────────────
    Buffer.BlockCopy(textBytes, 0, wd, fcText, textBytes.Length);

    // ── PAPX FKP ───────────────────────────────────────────────
    // 단락 1 의 \r 은 CP 6 → fc = 0x200 + 6*2 = 0x20C
    // 단락 2 의 \r 은 CP 12 → fc = 0x200 + 12*2 = 0x218
    // PapxFkp rgfc 구간: [0x200, 0x20E) → 단락 1; [0x20E, 0x21A) → 단락 2
    int cpara = 2;
    // rgfc[3]
    BitConverter.TryWriteBytes(wd.AsSpan(fcPapxFkp + 0),  (int)0x200);
    BitConverter.TryWriteBytes(wd.AsSpan(fcPapxFkp + 4),  (int)0x20E);
    BitConverter.TryWriteBytes(wd.AsSpan(fcPapxFkp + 8),  (int)0x21A);
    // BXPap[0] 첫 byte = bOffset (2-byte 단위), 나머지 12 byte PHE 는 0
    int papx0Off = 256;            // FKP-local byte offset
    int papx1Off = 280;
    wd[fcPapxFkp + 12 + 0 * 13]    = (byte)(papx0Off / 2);  // bOffset 128
    wd[fcPapxFkp + 12 + 1 * 13]    = (byte)(papx1Off / 2);  // bOffset 140
    // PapxInFkp[0] @ FKP+256: cb=3 (size = 3*2 - 1 = 5), istd(2) + sprmPJc80(2) + op(1)
    wd[fcPapxFkp + papx0Off + 0]   = 3;
    wd[fcPapxFkp + papx0Off + 1]   = 0; wd[fcPapxFkp + papx0Off + 2] = 0;   // istd
    BitConverter.TryWriteBytes(wd.AsSpan(fcPapxFkp + papx0Off + 3), (ushort)0x2461);
    wd[fcPapxFkp + papx0Off + 5]   = 1;  // center
    // PapxInFkp[1] @ FKP+280
    wd[fcPapxFkp + papx1Off + 0]   = 3;
    wd[fcPapxFkp + papx1Off + 1]   = 0; wd[fcPapxFkp + papx1Off + 2] = 0;
    BitConverter.TryWriteBytes(wd.AsSpan(fcPapxFkp + papx1Off + 3), (ushort)0x2461);
    wd[fcPapxFkp + papx1Off + 5]   = 2;  // right
    wd[fcPapxFkp + 511]            = (byte)cpara;

    // ── CHPX FKP ───────────────────────────────────────────────
    // run 0: fc [0x200, 0x208) — 4 chars "Bold" — Bold ON + 14pt (=28 halfPt)
    // run 1: fc [0x208, 0x20A) — 1 char "-"    — Bold OFF, 기본 크기
    // run 2: fc [0x20A, 0x21A) — 나머지 "X\rPlain\r" — 기본
    int crun = 3;
    BitConverter.TryWriteBytes(wd.AsSpan(fcChpxFkp + 0),  (int)0x200);
    BitConverter.TryWriteBytes(wd.AsSpan(fcChpxFkp + 4),  (int)0x208);
    BitConverter.TryWriteBytes(wd.AsSpan(fcChpxFkp + 8),  (int)0x20A);
    BitConverter.TryWriteBytes(wd.AsSpan(fcChpxFkp + 12), (int)0x21A);
    int chpx0Off = 64, chpx1Off = 80, chpx2Off = 96;
    int rgbBase  = 4 * (crun + 1);  // = 16
    wd[fcChpxFkp + rgbBase + 0] = (byte)(chpx0Off / 2);  // 32
    wd[fcChpxFkp + rgbBase + 1] = (byte)(chpx1Off / 2);  // 40
    wd[fcChpxFkp + rgbBase + 2] = (byte)(chpx2Off / 2);  // 48
    // ChpxInFkp[0]: cb=7 (sprmCFBold 2+op 1 + sprmCHps 2+op 2 = 7 byte grpprl)
    wd[fcChpxFkp + chpx0Off + 0] = 7;
    BitConverter.TryWriteBytes(wd.AsSpan(fcChpxFkp + chpx0Off + 1), (ushort)0x0835);
    wd[fcChpxFkp + chpx0Off + 3] = 1;
    BitConverter.TryWriteBytes(wd.AsSpan(fcChpxFkp + chpx0Off + 4), (ushort)0x4A43);
    BitConverter.TryWriteBytes(wd.AsSpan(fcChpxFkp + chpx0Off + 6), (ushort)28);  // 14pt
    // ChpxInFkp[1], [2]: cb=0 (no sprms)
    wd[fcChpxFkp + chpx1Off + 0] = 0;
    wd[fcChpxFkp + chpx2Off + 0] = 0;
    wd[fcChpxFkp + 511]          = (byte)crun;

    // ── Table stream: CLX + PlcBtePapx + PlcBteChpx ───────────
    var tblMs = new MemoryStream();
    Span<byte> b4 = stackalloc byte[4];
    // CLX (PCDT 0x02 + lcb + PlcPcd, 1 piece)
    tblMs.WriteByte(0x02);
    BitConverter.TryWriteBytes(b4, (uint)16); tblMs.Write(b4);
    BitConverter.TryWriteBytes(b4, (uint)0);   tblMs.Write(b4);   // aCP[0]
    BitConverter.TryWriteBytes(b4, (uint)ccp); tblMs.Write(b4);   // aCP[1]
    tblMs.WriteByte(0); tblMs.WriteByte(0);                       // PCD flags
    BitConverter.TryWriteBytes(b4, (uint)fcText); tblMs.Write(b4);// PCD fc
    tblMs.WriteByte(0); tblMs.WriteByte(0);                       // PCD prm
    int clxEnd = (int)tblMs.Position;

    // PlcBtePapx: aFC[2] + aPnFkp[1]
    int papxBteStart = clxEnd;
    BitConverter.TryWriteBytes(b4, (int)0x200); tblMs.Write(b4);
    BitConverter.TryWriteBytes(b4, (int)0x300); tblMs.Write(b4);  // upper bound
    BitConverter.TryWriteBytes(b4, (int)pnPapx); tblMs.Write(b4);
    int papxBteLen = (int)tblMs.Position - papxBteStart;

    // PlcBteChpx: aFC[2] + aPnFkp[1]
    int chpxBteStart = (int)tblMs.Position;
    BitConverter.TryWriteBytes(b4, (int)0x200); tblMs.Write(b4);
    BitConverter.TryWriteBytes(b4, (int)0x300); tblMs.Write(b4);
    BitConverter.TryWriteBytes(b4, (int)pnChpx); tblMs.Write(b4);
    int chpxBteLen = (int)tblMs.Position - chpxBteStart;

    var table = tblMs.ToArray();

    // ── FIB ───────────────────────────────────────────────────
    BitConverter.TryWriteBytes(wd.AsSpan(0x00),   (ushort)0xA5EC);
    BitConverter.TryWriteBytes(wd.AsSpan(0x02),   (ushort)193);
    BitConverter.TryWriteBytes(wd.AsSpan(0x0A),   (ushort)0x0000);  // 0Table, no encryption
    BitConverter.TryWriteBytes(wd.AsSpan(0x18),   (uint)fcText);
    BitConverter.TryWriteBytes(wd.AsSpan(0x4C),   (uint)ccp);
    BitConverter.TryWriteBytes(wd.AsSpan(0x01A2), (uint)0);                  // fcClx
    BitConverter.TryWriteBytes(wd.AsSpan(0x01A6), (uint)clxEnd);             // lcbClx
    BitConverter.TryWriteBytes(wd.AsSpan(0x00FA), (uint)chpxBteStart);       // fcPlcfBteChpx
    BitConverter.TryWriteBytes(wd.AsSpan(0x00FE), (uint)chpxBteLen);         // lcbPlcfBteChpx
    BitConverter.TryWriteBytes(wd.AsSpan(0x0102), (uint)papxBteStart);       // fcPlcfBtePapx
    BitConverter.TryWriteBytes(wd.AsSpan(0x0106), (uint)papxBteLen);         // lcbPlcfBtePapx

    var tmp = Path.Combine(Path.GetTempPath(), $"polydonky-smoke-{Guid.NewGuid():N}.doc");
    try
    {
        using (var root = OpenMcdf.RootStorage.Create(tmp))
        {
            using (var s = root.CreateStream("WordDocument")) s.Write(wd);
            using (var s = root.CreateStream("0Table"))       s.Write(table);
        }

        using var fs = File.OpenRead(tmp);
        var doc = new PolyDonky.Convert.Doc.DocBinaryReader().Read(fs);
        var paragraphs = doc.EnumerateParagraphs().ToList();

        SmokeHarness.True(paragraphs.Count >= 2,
            $"단락 수 >= 2 (got {paragraphs.Count})");
        SmokeHarness.Equal(Alignment.Center, paragraphs[0].Style.Alignment, "단락 1 PAPX 정렬 Center");
        SmokeHarness.Equal(Alignment.Right,  paragraphs[1].Style.Alignment, "단락 2 PAPX 정렬 Right");

        // 단락 1 의 첫 번째 Run 은 "Bold" 이고 Bold + 14pt 여야 함.
        var firstRun = paragraphs[0].Runs[0];
        SmokeHarness.True(firstRun.Style.Bold,
            $"단락 1 첫 Run 은 굵게 (got Bold={firstRun.Style.Bold}, text='{firstRun.Text}')");
        SmokeHarness.Equal(14.0, firstRun.Style.FontSizePt, "단락 1 첫 Run 글자 크기 14pt");
        SmokeHarness.Equal("Bold", firstRun.Text, "단락 1 첫 Run 텍스트 = 'Bold'");

        // 단락 1 의 마지막 Run 은 "-X" 이고 굵지 않아야 함.
        var lastRun1 = paragraphs[0].Runs[^1];
        SmokeHarness.True(!lastRun1.Style.Bold,
            $"단락 1 마지막 Run 은 일반 (got Bold={lastRun1.Style.Bold}, text='{lastRun1.Text}')");
    }
    finally { try { File.Delete(tmp); } catch { } }
}

static void DocBinaryPhase1cFormatting()
{
    // 1 단락 "Styled" 에 PAPX 들여쓰기 720 twips (=0.5 inch ≈ 12.7 mm), 단락-앞 간격 240 twips (=12 pt),
    // 줄 간격 LSPD(dyaLine=360, fMult=1 → 1.5x), CHPX sprmCCv = #FF0000 (red),
    // sprmCHighlight = 7 (yellow). 단일 piece + 단일 PAPX + 단일 CHPX 합성.
    const string text = "Styled\r";   // 7 chars; ccpText = 7
    int ccp = text.Length;
    var textBytes = Encoding.Unicode.GetBytes(text);
    int fcText  = 0x200;
    int pnPapx  = 4;
    int pnChpx  = 5;
    int fcPapxFkp = pnPapx * 512;
    int fcChpxFkp = pnChpx * 512;
    int wdSize    = fcChpxFkp + 512;

    var wd = new byte[wdSize];
    Buffer.BlockCopy(textBytes, 0, wd, fcText, textBytes.Length);

    // ── PAPX FKP ───────────────────────────────────────────────
    // cpara=1, rgfc[0]=0x200, rgfc[1]=0x210 (단락이 다루는 fc 상한)
    int cpara = 1;
    BitConverter.TryWriteBytes(wd.AsSpan(fcPapxFkp + 0), (int)0x200);
    BitConverter.TryWriteBytes(wd.AsSpan(fcPapxFkp + 4), (int)0x210);
    int papx0Off = 200;
    wd[fcPapxFkp + 8 + 0 * 13] = (byte)(papx0Off / 2);  // bOffset = 100

    // PapxInFkp: cb=0 형식 → grpprlSize = cb'*2 = 16. istd(2) + 4 sprm-pair(14) = 16.
    int p = fcPapxFkp + papx0Off;
    wd[p++] = 0;   // cb=0 (variable form)
    wd[p++] = 8;   // cb' = 8 → grpprl size = 16 byte
    // istd (2 byte) = 0
    p += 2;
    // sprmPDxaLeft (0x845D, 2-byte signed): 720 twips
    BitConverter.TryWriteBytes(wd.AsSpan(p), (ushort)0x845D); p += 2;
    BitConverter.TryWriteBytes(wd.AsSpan(p), (short)720);     p += 2;
    // sprmPDyaBefore (0xA413, 2-byte unsigned): 240 twips
    BitConverter.TryWriteBytes(wd.AsSpan(p), (ushort)0xA413); p += 2;
    BitConverter.TryWriteBytes(wd.AsSpan(p), (ushort)240);    p += 2;
    // sprmPDyaLine (0x6412, 4-byte LSPD)
    BitConverter.TryWriteBytes(wd.AsSpan(p), (ushort)0x6412); p += 2;
    BitConverter.TryWriteBytes(wd.AsSpan(p), (short)360);     p += 2;  // dyaLine
    BitConverter.TryWriteBytes(wd.AsSpan(p), (ushort)1);      p += 2;  // fMultLinespace = 1
    wd[fcPapxFkp + 511] = (byte)cpara;

    // ── CHPX FKP ───────────────────────────────────────────────
    int crun = 1;
    BitConverter.TryWriteBytes(wd.AsSpan(fcChpxFkp + 0), (int)0x200);
    BitConverter.TryWriteBytes(wd.AsSpan(fcChpxFkp + 4), (int)0x20E);   // 7 chars * 2 byte = 14 (0xE)
    int chpx0Off = 64;
    int rgbBase  = 4 * (crun + 1);
    wd[fcChpxFkp + rgbBase + 0] = (byte)(chpx0Off / 2);
    // ChpxInFkp: cb = 9 (sprmCCv 6 + sprmCHighlight 3)
    int c = fcChpxFkp + chpx0Off;
    wd[c++] = 9;
    // sprmCCv (0x6870, 4-byte: R, G, B, fAuto)
    BitConverter.TryWriteBytes(wd.AsSpan(c), (ushort)0x6870); c += 2;
    wd[c++] = 0xFF; wd[c++] = 0x00; wd[c++] = 0x00; wd[c++] = 0x00;  // red, not auto
    // sprmCHighlight (0x2A0C, 1-byte palette index)
    BitConverter.TryWriteBytes(wd.AsSpan(c), (ushort)0x2A0C); c += 2;
    wd[c++] = 7;  // yellow
    wd[fcChpxFkp + 511] = (byte)crun;

    // ── Table stream ───────────────────────────────────────────
    var tblMs = new MemoryStream();
    Span<byte> b4 = stackalloc byte[4];
    tblMs.WriteByte(0x02);
    BitConverter.TryWriteBytes(b4, (uint)16); tblMs.Write(b4);
    BitConverter.TryWriteBytes(b4, (uint)0);   tblMs.Write(b4);
    BitConverter.TryWriteBytes(b4, (uint)ccp); tblMs.Write(b4);
    tblMs.WriteByte(0); tblMs.WriteByte(0);
    BitConverter.TryWriteBytes(b4, (uint)fcText); tblMs.Write(b4);
    tblMs.WriteByte(0); tblMs.WriteByte(0);
    int clxEnd = (int)tblMs.Position;

    int papxBteStart = clxEnd;
    BitConverter.TryWriteBytes(b4, (int)0x200);  tblMs.Write(b4);
    BitConverter.TryWriteBytes(b4, (int)0x300);  tblMs.Write(b4);
    BitConverter.TryWriteBytes(b4, (int)pnPapx); tblMs.Write(b4);
    int papxBteLen = (int)tblMs.Position - papxBteStart;
    int chpxBteStart = (int)tblMs.Position;
    BitConverter.TryWriteBytes(b4, (int)0x200);  tblMs.Write(b4);
    BitConverter.TryWriteBytes(b4, (int)0x300);  tblMs.Write(b4);
    BitConverter.TryWriteBytes(b4, (int)pnChpx); tblMs.Write(b4);
    int chpxBteLen = (int)tblMs.Position - chpxBteStart;
    var table = tblMs.ToArray();

    BitConverter.TryWriteBytes(wd.AsSpan(0x00),   (ushort)0xA5EC);
    BitConverter.TryWriteBytes(wd.AsSpan(0x02),   (ushort)193);
    BitConverter.TryWriteBytes(wd.AsSpan(0x0A),   (ushort)0x0000);
    BitConverter.TryWriteBytes(wd.AsSpan(0x18),   (uint)fcText);
    BitConverter.TryWriteBytes(wd.AsSpan(0x4C),   (uint)ccp);
    BitConverter.TryWriteBytes(wd.AsSpan(0x01A2), (uint)0);
    BitConverter.TryWriteBytes(wd.AsSpan(0x01A6), (uint)clxEnd);
    BitConverter.TryWriteBytes(wd.AsSpan(0x00FA), (uint)chpxBteStart);
    BitConverter.TryWriteBytes(wd.AsSpan(0x00FE), (uint)chpxBteLen);
    BitConverter.TryWriteBytes(wd.AsSpan(0x0102), (uint)papxBteStart);
    BitConverter.TryWriteBytes(wd.AsSpan(0x0106), (uint)papxBteLen);

    var tmp = Path.Combine(Path.GetTempPath(), $"polydonky-smoke-{Guid.NewGuid():N}.doc");
    try
    {
        using (var root = OpenMcdf.RootStorage.Create(tmp))
        {
            using (var s = root.CreateStream("WordDocument")) s.Write(wd);
            using (var s = root.CreateStream("0Table"))       s.Write(table);
        }
        using var fs = File.OpenRead(tmp);
        var doc = new PolyDonky.Convert.Doc.DocBinaryReader().Read(fs);
        var ps  = doc.EnumerateParagraphs().First().Style;
        var run = doc.EnumerateParagraphs().First().Runs[0];

        // PAPX
        SmokeHarness.True(Math.Abs(ps.IndentLeftMm - 720.0 / 56.692) < 0.01,
            $"PAPX 들여쓰기 12.7 mm (got {ps.IndentLeftMm:F2})");
        SmokeHarness.Equal(12.0, ps.SpaceBeforePt, "PAPX 단락 앞 간격 12pt");
        SmokeHarness.Equal(1.5,  ps.LineHeightFactor, "PAPX 줄 간격 1.5x");

        // CHPX
        SmokeHarness.True(run.Style.Foreground.HasValue, "CHPX sprmCCv 전경색이 설정됨");
        SmokeHarness.Equal((byte)255, run.Style.Foreground!.Value.R, "전경 R = 255 (red)");
        SmokeHarness.Equal((byte)0,   run.Style.Foreground!.Value.G, "전경 G = 0");
        SmokeHarness.True(run.Style.Background.HasValue, "CHPX sprmCHighlight 배경색이 설정됨");
        SmokeHarness.Equal((byte)255, run.Style.Background!.Value.R, "하이라이트 R = 255");
        SmokeHarness.Equal((byte)255, run.Style.Background!.Value.G, "하이라이트 G = 255 (yellow)");
        SmokeHarness.Equal((byte)0,   run.Style.Background!.Value.B, "하이라이트 B = 0");
    }
    finally { try { File.Delete(tmp); } catch { } }
}

static void DocBinaryPhase1dFontFamily()
{
    // STTB FFN 에 폰트명 "Malgun Gothic" 1 개 등록, CHPX sprmCRgFtc0 = 0 으로 Run 에 적용.
    const string fontName = "Malgun Gothic";
    const string text = "FontTest\r";   // 9 chars; ccpText = 9
    int ccp = text.Length;
    var textBytes = Encoding.Unicode.GetBytes(text);
    int fcText  = 0x200;
    int pnPapx  = 4;
    int pnChpx  = 5;
    int fcPapxFkp = pnPapx * 512;
    int fcChpxFkp = pnChpx * 512;
    int wdSize    = fcChpxFkp + 512;

    var wd = new byte[wdSize];
    Buffer.BlockCopy(textBytes, 0, wd, fcText, textBytes.Length);

    // ── PAPX FKP (정렬만; FontFamily 와 무관) ──────────────────
    int cpara = 1;
    BitConverter.TryWriteBytes(wd.AsSpan(fcPapxFkp + 0), (int)0x200);
    BitConverter.TryWriteBytes(wd.AsSpan(fcPapxFkp + 4), (int)0x214);
    int papx0Off = 200;
    wd[fcPapxFkp + 8 + 0 * 13] = (byte)(papx0Off / 2);
    int p = fcPapxFkp + papx0Off;
    wd[p++] = 3;   // cb=3 (grpprl size = 5)
    p += 2;        // istd
    BitConverter.TryWriteBytes(wd.AsSpan(p), (ushort)0x2461); p += 2;
    wd[p++] = 0;   // left
    wd[fcPapxFkp + 511] = (byte)cpara;

    // ── CHPX FKP (sprmCRgFtc0 = 0 → 첫 폰트) ──────────────────
    int crun = 1;
    BitConverter.TryWriteBytes(wd.AsSpan(fcChpxFkp + 0), (int)0x200);
    BitConverter.TryWriteBytes(wd.AsSpan(fcChpxFkp + 4), (int)0x212);  // 9 chars * 2 byte
    int chpx0Off = 64;
    int rgbBase  = 4 * (crun + 1);
    wd[fcChpxFkp + rgbBase + 0] = (byte)(chpx0Off / 2);
    int c = fcChpxFkp + chpx0Off;
    wd[c++] = 4;  // cb = 4 (sprmCRgFtc0 2 + operand 2)
    BitConverter.TryWriteBytes(wd.AsSpan(c), (ushort)0x4A4F); c += 2;
    BitConverter.TryWriteBytes(wd.AsSpan(c), (ushort)0);      c += 2;  // ftc = 0
    wd[fcChpxFkp + 511] = (byte)crun;

    // ── Table stream: CLX + PlcBtePapx + PlcBteChpx + STTB FFN ─
    var tblMs = new MemoryStream();
    Span<byte> b4 = stackalloc byte[4];
    // CLX
    tblMs.WriteByte(0x02);
    BitConverter.TryWriteBytes(b4, (uint)16); tblMs.Write(b4);
    BitConverter.TryWriteBytes(b4, (uint)0);   tblMs.Write(b4);
    BitConverter.TryWriteBytes(b4, (uint)ccp); tblMs.Write(b4);
    tblMs.WriteByte(0); tblMs.WriteByte(0);
    BitConverter.TryWriteBytes(b4, (uint)fcText); tblMs.Write(b4);
    tblMs.WriteByte(0); tblMs.WriteByte(0);
    int clxEnd = (int)tblMs.Position;
    // PlcBtePapx
    int papxBteStart = clxEnd;
    BitConverter.TryWriteBytes(b4, (int)0x200);  tblMs.Write(b4);
    BitConverter.TryWriteBytes(b4, (int)0x300);  tblMs.Write(b4);
    BitConverter.TryWriteBytes(b4, (int)pnPapx); tblMs.Write(b4);
    int papxBteLen = (int)tblMs.Position - papxBteStart;
    // PlcBteChpx
    int chpxBteStart = (int)tblMs.Position;
    BitConverter.TryWriteBytes(b4, (int)0x200);  tblMs.Write(b4);
    BitConverter.TryWriteBytes(b4, (int)0x300);  tblMs.Write(b4);
    BitConverter.TryWriteBytes(b4, (int)pnChpx); tblMs.Write(b4);
    int chpxBteLen = (int)tblMs.Position - chpxBteStart;

    // STTB FFN (extended, 1 폰트)
    // Header: 0xFFFF, cData=1, cbExtra=0
    int sttbStart = (int)tblMs.Position;
    BitConverter.TryWriteBytes(b4.Slice(0, 2), (ushort)0xFFFF); tblMs.Write(b4.Slice(0, 2));
    BitConverter.TryWriteBytes(b4.Slice(0, 2), (ushort)1);      tblMs.Write(b4.Slice(0, 2));
    BitConverter.TryWriteBytes(b4.Slice(0, 2), (ushort)0);      tblMs.Write(b4.Slice(0, 2));
    // Entry 1: cchData(2) + FFN
    // FFN: 40 byte header + xszFfn (UTF-16LE null-terminated)
    var nameBytes = Encoding.Unicode.GetBytes(fontName + "\0");  // null-term
    int ffnSize = 40 + nameBytes.Length;
    // cchData = ffnSize / 2 (wide-chars)
    BitConverter.TryWriteBytes(b4.Slice(0, 2), (ushort)(ffnSize / 2)); tblMs.Write(b4.Slice(0, 2));
    // FFN header 40 byte (대부분 0 으로 채움)
    var ffnHdr = new byte[40];
    ffnHdr[0] = (byte)(ffnSize - 1);  // cbFfnM1
    ffnHdr[1] = 0;                    // flags
    BitConverter.TryWriteBytes(ffnHdr.AsSpan(2), (ushort)400);  // wWeight
    tblMs.Write(ffnHdr);
    // xszFfn
    tblMs.Write(nameBytes);
    int sttbLen = (int)tblMs.Position - sttbStart;

    var table = tblMs.ToArray();

    // ── FIB ───────────────────────────────────────────────────
    BitConverter.TryWriteBytes(wd.AsSpan(0x00),   (ushort)0xA5EC);
    BitConverter.TryWriteBytes(wd.AsSpan(0x02),   (ushort)193);
    BitConverter.TryWriteBytes(wd.AsSpan(0x0A),   (ushort)0x0000);
    BitConverter.TryWriteBytes(wd.AsSpan(0x18),   (uint)fcText);
    BitConverter.TryWriteBytes(wd.AsSpan(0x4C),   (uint)ccp);
    BitConverter.TryWriteBytes(wd.AsSpan(0x01A2), (uint)0);
    BitConverter.TryWriteBytes(wd.AsSpan(0x01A6), (uint)clxEnd);
    BitConverter.TryWriteBytes(wd.AsSpan(0x00FA), (uint)chpxBteStart);
    BitConverter.TryWriteBytes(wd.AsSpan(0x00FE), (uint)chpxBteLen);
    BitConverter.TryWriteBytes(wd.AsSpan(0x0102), (uint)papxBteStart);
    BitConverter.TryWriteBytes(wd.AsSpan(0x0106), (uint)papxBteLen);
    BitConverter.TryWriteBytes(wd.AsSpan(0x0112), (uint)sttbStart);
    BitConverter.TryWriteBytes(wd.AsSpan(0x0116), (uint)sttbLen);

    var tmp = Path.Combine(Path.GetTempPath(), $"polydonky-smoke-{Guid.NewGuid():N}.doc");
    try
    {
        using (var root = OpenMcdf.RootStorage.Create(tmp))
        {
            using (var s = root.CreateStream("WordDocument")) s.Write(wd);
            using (var s = root.CreateStream("0Table"))       s.Write(table);
        }
        using var fs = File.OpenRead(tmp);
        var doc = new PolyDonky.Convert.Doc.DocBinaryReader().Read(fs);
        var run = doc.EnumerateParagraphs().First().Runs[0];

        SmokeHarness.Equal(fontName, run.Style.FontFamily ?? "", "CHPX sprmCRgFtc0 → STTB FFN 폰트명");
        SmokeHarness.Equal("FontTest", run.Text, "본문 텍스트 보존");
    }
    finally { try { File.Delete(tmp); } catch { } }
}

static void DocBinaryPhase1eStshHeading()
{
    // STSH 에 STD[0] = (sti=1 Heading 1, stk=1 paragraph) 등록, PAPX 의 GrpPrlAndIstd 의
    // istd = 0 으로 설정. 결과: 단락의 OutlineLevel = H1.
    const string text = "Title\r";  // 6 chars; ccpText = 6
    int ccp = text.Length;
    var textBytes = Encoding.Unicode.GetBytes(text);
    int fcText  = 0x200;
    int pnPapx  = 4;
    int pnChpx  = 5;
    int fcPapxFkp = pnPapx * 512;
    int fcChpxFkp = pnChpx * 512;
    int wdSize    = fcChpxFkp + 512;

    var wd = new byte[wdSize];
    Buffer.BlockCopy(textBytes, 0, wd, fcText, textBytes.Length);

    // ── PAPX FKP — istd=0 ────────────────────────────────────────
    int cpara = 1;
    BitConverter.TryWriteBytes(wd.AsSpan(fcPapxFkp + 0), (int)0x200);
    BitConverter.TryWriteBytes(wd.AsSpan(fcPapxFkp + 4), (int)0x20E);
    int papx0Off = 200;
    wd[fcPapxFkp + 8 + 0 * 13] = (byte)(papx0Off / 2);
    int p = fcPapxFkp + papx0Off;
    // cb=1 (grpprl size = 1*2 - 1 = 1)... 문제는 istd 만 2 byte 인데 sprm 없이도 istd 가 필요.
    // GrpPrlAndIstd 의 최소 크기는 istd 만 2 byte → grpprl = 2 byte → cb*2 - 1 = 2 ⇒ cb=1.5 안 됨
    // cb=0 형식 사용: cb=0, cb'=1, grpprlSize=2 → istd(2) 만 박는다.
    wd[p++] = 0;
    wd[p++] = 1;        // cb' = 1 → grpprl size = 2 byte (= istd 만)
    BitConverter.TryWriteBytes(wd.AsSpan(p), (ushort)0);  // istd = 0
    wd[fcPapxFkp + 511] = (byte)cpara;

    // ── CHPX FKP — 비어 있음 (run 1, 빈 grpprl) ──────────────────
    int crun = 1;
    BitConverter.TryWriteBytes(wd.AsSpan(fcChpxFkp + 0), (int)0x200);
    BitConverter.TryWriteBytes(wd.AsSpan(fcChpxFkp + 4), (int)0x20C);  // 6 chars * 2 byte
    int chpx0Off = 64;
    wd[fcChpxFkp + 4 * (crun + 1) + 0] = (byte)(chpx0Off / 2);
    wd[fcChpxFkp + chpx0Off] = 0;  // cb=0
    wd[fcChpxFkp + 511] = (byte)crun;

    // ── Table stream: CLX + PlcBtePapx + PlcBteChpx + STSH ───────
    var tblMs = new MemoryStream();
    Span<byte> b4 = stackalloc byte[4];
    tblMs.WriteByte(0x02);
    BitConverter.TryWriteBytes(b4, (uint)16); tblMs.Write(b4);
    BitConverter.TryWriteBytes(b4, (uint)0);   tblMs.Write(b4);
    BitConverter.TryWriteBytes(b4, (uint)ccp); tblMs.Write(b4);
    tblMs.WriteByte(0); tblMs.WriteByte(0);
    BitConverter.TryWriteBytes(b4, (uint)fcText); tblMs.Write(b4);
    tblMs.WriteByte(0); tblMs.WriteByte(0);
    int clxEnd = (int)tblMs.Position;

    int papxBteStart = clxEnd;
    BitConverter.TryWriteBytes(b4, (int)0x200);  tblMs.Write(b4);
    BitConverter.TryWriteBytes(b4, (int)0x300);  tblMs.Write(b4);
    BitConverter.TryWriteBytes(b4, (int)pnPapx); tblMs.Write(b4);
    int papxBteLen = (int)tblMs.Position - papxBteStart;

    int chpxBteStart = (int)tblMs.Position;
    BitConverter.TryWriteBytes(b4, (int)0x200);  tblMs.Write(b4);
    BitConverter.TryWriteBytes(b4, (int)0x300);  tblMs.Write(b4);
    BitConverter.TryWriteBytes(b4, (int)pnChpx); tblMs.Write(b4);
    int chpxBteLen = (int)tblMs.Position - chpxBteStart;

    // STSH — LPStshi (4 byte: cbStshi=4 + Stshi minimal: cstd=1, cbSTDBaseInFile=10)
    //      + rgLPStd[1]:
    //          LPStd[0]: cbStd(2)=10 + STD (10 byte: stdfBase only)
    //                    stdfBase word0 = sti(12)=1, others 0 → 0x0001
    //                    stdfBase word1 = stk(4)=1, istdBase(12)=0 → 0x0001
    int stshStart = (int)tblMs.Position;
    // LPStshi
    BitConverter.TryWriteBytes(b4.Slice(0, 2), (ushort)4); tblMs.Write(b4.Slice(0, 2));  // cbStshi = 4
    BitConverter.TryWriteBytes(b4.Slice(0, 2), (ushort)1); tblMs.Write(b4.Slice(0, 2));  // cstd = 1
    BitConverter.TryWriteBytes(b4.Slice(0, 2), (ushort)10); tblMs.Write(b4.Slice(0, 2)); // cbSTDBaseInFile = 10
    // LPStd[0]
    BitConverter.TryWriteBytes(b4.Slice(0, 2), (ushort)10); tblMs.Write(b4.Slice(0, 2)); // cbStd = 10
    BitConverter.TryWriteBytes(b4.Slice(0, 2), (ushort)0x0001); tblMs.Write(b4.Slice(0, 2));  // sti=1
    BitConverter.TryWriteBytes(b4.Slice(0, 2), (ushort)0x0001); tblMs.Write(b4.Slice(0, 2));  // stk=1
    BitConverter.TryWriteBytes(b4.Slice(0, 2), (ushort)0); tblMs.Write(b4.Slice(0, 2));  // cupx/istdNext
    BitConverter.TryWriteBytes(b4.Slice(0, 2), (ushort)0); tblMs.Write(b4.Slice(0, 2));  // bchUpe
    BitConverter.TryWriteBytes(b4.Slice(0, 2), (ushort)0); tblMs.Write(b4.Slice(0, 2));  // grfstd
    int stshLen = (int)tblMs.Position - stshStart;

    var table = tblMs.ToArray();

    BitConverter.TryWriteBytes(wd.AsSpan(0x00),   (ushort)0xA5EC);
    BitConverter.TryWriteBytes(wd.AsSpan(0x02),   (ushort)193);
    BitConverter.TryWriteBytes(wd.AsSpan(0x0A),   (ushort)0x0000);
    BitConverter.TryWriteBytes(wd.AsSpan(0x18),   (uint)fcText);
    BitConverter.TryWriteBytes(wd.AsSpan(0x4C),   (uint)ccp);
    BitConverter.TryWriteBytes(wd.AsSpan(0x01A2), (uint)0);
    BitConverter.TryWriteBytes(wd.AsSpan(0x01A6), (uint)clxEnd);
    BitConverter.TryWriteBytes(wd.AsSpan(0x00FA), (uint)chpxBteStart);
    BitConverter.TryWriteBytes(wd.AsSpan(0x00FE), (uint)chpxBteLen);
    BitConverter.TryWriteBytes(wd.AsSpan(0x0102), (uint)papxBteStart);
    BitConverter.TryWriteBytes(wd.AsSpan(0x0106), (uint)papxBteLen);
    BitConverter.TryWriteBytes(wd.AsSpan(0x00A2), (uint)stshStart);
    BitConverter.TryWriteBytes(wd.AsSpan(0x00A6), (uint)stshLen);

    var tmp = Path.Combine(Path.GetTempPath(), $"polydonky-smoke-{Guid.NewGuid():N}.doc");
    try
    {
        using (var root = OpenMcdf.RootStorage.Create(tmp))
        {
            using (var s = root.CreateStream("WordDocument")) s.Write(wd);
            using (var s = root.CreateStream("0Table"))       s.Write(table);
        }
        using var fs = File.OpenRead(tmp);
        var doc = new PolyDonky.Convert.Doc.DocBinaryReader().Read(fs);
        var ps  = doc.EnumerateParagraphs().First().Style;

        SmokeHarness.Equal(OutlineLevel.H1, ps.Outline,
            $"STSH sti=1 (Heading 1) → OutlineLevel.H1 (got {ps.Outline})");
        SmokeHarness.Equal("Title", doc.EnumerateParagraphs().First().GetPlainText(),
            "본문 텍스트 보존");
    }
    finally { try { File.Delete(tmp); } catch { } }
}

static void DocBinaryPhase1fStdInheritance()
{
    // STSH 의 STD[0] = Heading 1 (sti=1, stk=1, cupx=2) 에 LPUpx[1]=CHPX 로 sprmCFBold + sprmCHps=24 박음.
    // 단락 PAPX 의 istd=0, 직접 CHPX 없음 → STD 의 CHPX 가 상속되어 Run 은 Bold + 12pt 가 되어야 함.
    const string text = "Title\r";  // 6 chars
    int ccp = text.Length;
    var textBytes = Encoding.Unicode.GetBytes(text);
    int fcText = 0x200;
    int pnPapx = 4, pnChpx = 5;
    int fcPapxFkp = pnPapx * 512;
    int fcChpxFkp = pnChpx * 512;
    int wdSize = fcChpxFkp + 512;

    var wd = new byte[wdSize];
    Buffer.BlockCopy(textBytes, 0, wd, fcText, textBytes.Length);

    // ── PAPX FKP — istd=0, sprm 없음 ─────────────────────────────
    int cpara = 1;
    BitConverter.TryWriteBytes(wd.AsSpan(fcPapxFkp + 0), (int)0x200);
    BitConverter.TryWriteBytes(wd.AsSpan(fcPapxFkp + 4), (int)0x20E);
    int papx0Off = 200;
    wd[fcPapxFkp + 8 + 0 * 13] = (byte)(papx0Off / 2);
    int p = fcPapxFkp + papx0Off;
    wd[p++] = 0;
    wd[p++] = 1;        // cb'=1 → grpprl size = 2 (istd 만)
    BitConverter.TryWriteBytes(wd.AsSpan(p), (ushort)0);  // istd=0
    wd[fcPapxFkp + 511] = (byte)cpara;

    // ── CHPX FKP — 비어 있음 ─────────────────────────────────────
    int crun = 1;
    BitConverter.TryWriteBytes(wd.AsSpan(fcChpxFkp + 0), (int)0x200);
    BitConverter.TryWriteBytes(wd.AsSpan(fcChpxFkp + 4), (int)0x20C);
    int chpx0Off = 64;
    wd[fcChpxFkp + 4 * (crun + 1) + 0] = (byte)(chpx0Off / 2);
    wd[fcChpxFkp + chpx0Off] = 0;  // cb=0 (직접 sprm 없음)
    wd[fcChpxFkp + 511] = (byte)crun;

    // ── Table stream ────────────────────────────────────────────
    var tblMs = new MemoryStream();
    Span<byte> b4 = stackalloc byte[4];
    // CLX
    tblMs.WriteByte(0x02);
    BitConverter.TryWriteBytes(b4, (uint)16); tblMs.Write(b4);
    BitConverter.TryWriteBytes(b4, (uint)0);   tblMs.Write(b4);
    BitConverter.TryWriteBytes(b4, (uint)ccp); tblMs.Write(b4);
    tblMs.WriteByte(0); tblMs.WriteByte(0);
    BitConverter.TryWriteBytes(b4, (uint)fcText); tblMs.Write(b4);
    tblMs.WriteByte(0); tblMs.WriteByte(0);
    int clxEnd = (int)tblMs.Position;
    int papxBteStart = clxEnd;
    BitConverter.TryWriteBytes(b4, (int)0x200);  tblMs.Write(b4);
    BitConverter.TryWriteBytes(b4, (int)0x300);  tblMs.Write(b4);
    BitConverter.TryWriteBytes(b4, (int)pnPapx); tblMs.Write(b4);
    int papxBteLen = (int)tblMs.Position - papxBteStart;
    int chpxBteStart = (int)tblMs.Position;
    BitConverter.TryWriteBytes(b4, (int)0x200);  tblMs.Write(b4);
    BitConverter.TryWriteBytes(b4, (int)0x300);  tblMs.Write(b4);
    BitConverter.TryWriteBytes(b4, (int)pnChpx); tblMs.Write(b4);
    int chpxBteLen = (int)tblMs.Position - chpxBteStart;

    // STSH — LPStshi(cbStshi=4 + Stshi cstd=1 cbSTDBaseInFile=10) + LPStd[0](cbStd + STD)
    // STD layout (Heading 1 with CHPX Bold + 12pt):
    //   stdfBase (10 byte): sti=1, stk=1, cupx=2, ...
    //   xstzName: cchData=0(2) + null(2) = 4 byte
    //   LPUpx[0] PAPX: cbUpx=2(2) + UPX{istd=0(2)} = 4 byte
    //   LPUpx[1] CHPX: cbUpx=7(2) + UPX{sprmCFBold 3 + sprmCHps 4} = 9 byte (홀수 → 1 byte pad)
    //   총 = 10 + 4 + 4 + 10 = 28 byte
    int stshStart = (int)tblMs.Position;
    // LPStshi
    BitConverter.TryWriteBytes(b4.Slice(0, 2), (ushort)4); tblMs.Write(b4.Slice(0, 2));
    BitConverter.TryWriteBytes(b4.Slice(0, 2), (ushort)1); tblMs.Write(b4.Slice(0, 2));  // cstd
    BitConverter.TryWriteBytes(b4.Slice(0, 2), (ushort)10); tblMs.Write(b4.Slice(0, 2)); // cbSTDBaseInFile
    // LPStd[0]
    int cbStd = 28;
    BitConverter.TryWriteBytes(b4.Slice(0, 2), (ushort)cbStd); tblMs.Write(b4.Slice(0, 2));
    // STD stdfBase
    BitConverter.TryWriteBytes(b4.Slice(0, 2), (ushort)0x0001); tblMs.Write(b4.Slice(0, 2));  // sti=1
    BitConverter.TryWriteBytes(b4.Slice(0, 2), (ushort)((0xFFF << 4) | 1)); tblMs.Write(b4.Slice(0, 2));  // stk=1, istdBase=nil
    BitConverter.TryWriteBytes(b4.Slice(0, 2), (ushort)((0xFFF << 4) | 2)); tblMs.Write(b4.Slice(0, 2));  // cupx=2, istdNext=nil
    BitConverter.TryWriteBytes(b4.Slice(0, 2), (ushort)0); tblMs.Write(b4.Slice(0, 2));  // bchUpe
    BitConverter.TryWriteBytes(b4.Slice(0, 2), (ushort)0); tblMs.Write(b4.Slice(0, 2));  // grfstd
    // xstzName: cchData=0, null
    BitConverter.TryWriteBytes(b4.Slice(0, 2), (ushort)0); tblMs.Write(b4.Slice(0, 2));
    BitConverter.TryWriteBytes(b4.Slice(0, 2), (ushort)0); tblMs.Write(b4.Slice(0, 2));
    // LPUpx[0] PAPX: cbUpx=2 + istd=0
    BitConverter.TryWriteBytes(b4.Slice(0, 2), (ushort)2); tblMs.Write(b4.Slice(0, 2));
    BitConverter.TryWriteBytes(b4.Slice(0, 2), (ushort)0); tblMs.Write(b4.Slice(0, 2));
    // LPUpx[1] CHPX: cbUpx=7, sprmCFBold(2)+op(1) + sprmCHps(2)+op(2) = 7 byte
    BitConverter.TryWriteBytes(b4.Slice(0, 2), (ushort)7); tblMs.Write(b4.Slice(0, 2));
    BitConverter.TryWriteBytes(b4.Slice(0, 2), (ushort)0x0835); tblMs.Write(b4.Slice(0, 2));  // sprmCFBold
    tblMs.WriteByte(1);                                                                      // Bold ON
    BitConverter.TryWriteBytes(b4.Slice(0, 2), (ushort)0x4A43); tblMs.Write(b4.Slice(0, 2));  // sprmCHps
    BitConverter.TryWriteBytes(b4.Slice(0, 2), (ushort)24); tblMs.Write(b4.Slice(0, 2));      // 12pt
    // 1 byte padding to make total STD = 28 byte
    tblMs.WriteByte(0);
    int stshLen = (int)tblMs.Position - stshStart;

    var table = tblMs.ToArray();

    BitConverter.TryWriteBytes(wd.AsSpan(0x00),   (ushort)0xA5EC);
    BitConverter.TryWriteBytes(wd.AsSpan(0x02),   (ushort)193);
    BitConverter.TryWriteBytes(wd.AsSpan(0x0A),   (ushort)0x0000);
    BitConverter.TryWriteBytes(wd.AsSpan(0x18),   (uint)fcText);
    BitConverter.TryWriteBytes(wd.AsSpan(0x4C),   (uint)ccp);
    BitConverter.TryWriteBytes(wd.AsSpan(0x01A2), (uint)0);
    BitConverter.TryWriteBytes(wd.AsSpan(0x01A6), (uint)clxEnd);
    BitConverter.TryWriteBytes(wd.AsSpan(0x00FA), (uint)chpxBteStart);
    BitConverter.TryWriteBytes(wd.AsSpan(0x00FE), (uint)chpxBteLen);
    BitConverter.TryWriteBytes(wd.AsSpan(0x0102), (uint)papxBteStart);
    BitConverter.TryWriteBytes(wd.AsSpan(0x0106), (uint)papxBteLen);
    BitConverter.TryWriteBytes(wd.AsSpan(0x00A2), (uint)stshStart);
    BitConverter.TryWriteBytes(wd.AsSpan(0x00A6), (uint)stshLen);

    var tmp = Path.Combine(Path.GetTempPath(), $"polydonky-smoke-{Guid.NewGuid():N}.doc");
    try
    {
        using (var root = OpenMcdf.RootStorage.Create(tmp))
        {
            using (var s = root.CreateStream("WordDocument")) s.Write(wd);
            using (var s = root.CreateStream("0Table"))       s.Write(table);
        }
        using var fs = File.OpenRead(tmp);
        var doc = new PolyDonky.Convert.Doc.DocBinaryReader().Read(fs);
        var para = doc.EnumerateParagraphs().First();
        var run  = para.Runs[0];

        SmokeHarness.Equal(OutlineLevel.H1, para.Style.Outline, "Heading 1 sti → Outline.H1");
        SmokeHarness.True(run.Style.Bold,
            $"STD CHPX sprmCFBold 상속 → Run.Bold=true (got {run.Style.Bold})");
        SmokeHarness.Equal(12.0, run.Style.FontSizePt,
            "STD CHPX sprmCHps=24(halfPt) 상속 → 12pt");
        SmokeHarness.Equal("Title", para.GetPlainText(), "본문 보존");
    }
    finally { try { File.Delete(tmp); } catch { } }
}

static void DocBinaryPhase1gStdChain()
{
    // STSH 의 STD 두 개:
    //   STD[0] Normal     (sti=0, stk=1, istdBase=nil): CHPX sprmCRgFtc0=0 → "ChainArial" 폰트
    //   STD[1] Heading 1  (sti=1, stk=1, istdBase=0):    CHPX sprmCFBold ON
    // 단락 PAPX istd=1. Phase 1g 체이닝으로 Run 은 폰트는 Normal 에서, Bold 는 Heading 1 에서 상속.
    const string fontName = "ChainArial";
    const string text = "Chain\r";  // 6 chars
    int ccp = text.Length;
    var textBytes = Encoding.Unicode.GetBytes(text);
    int fcText = 0x200;
    int pnPapx = 4, pnChpx = 5;
    int fcPapxFkp = pnPapx * 512;
    int fcChpxFkp = pnChpx * 512;
    int wdSize = fcChpxFkp + 512;

    var wd = new byte[wdSize];
    Buffer.BlockCopy(textBytes, 0, wd, fcText, textBytes.Length);

    // ── PAPX FKP — istd=1 (Heading 1), 직접 sprm 없음 ───────────
    int cpara = 1;
    BitConverter.TryWriteBytes(wd.AsSpan(fcPapxFkp + 0), (int)0x200);
    BitConverter.TryWriteBytes(wd.AsSpan(fcPapxFkp + 4), (int)0x20E);
    int papx0Off = 200;
    wd[fcPapxFkp + 8 + 0 * 13] = (byte)(papx0Off / 2);
    int p = fcPapxFkp + papx0Off;
    wd[p++] = 0;
    wd[p++] = 1;   // cb'=1 → grpprl size=2 (istd 만)
    BitConverter.TryWriteBytes(wd.AsSpan(p), (ushort)1);  // istd=1 (Heading 1)
    wd[fcPapxFkp + 511] = (byte)cpara;

    // ── CHPX FKP — 비어 있음 ─────────────────────────────────────
    int crun = 1;
    BitConverter.TryWriteBytes(wd.AsSpan(fcChpxFkp + 0), (int)0x200);
    BitConverter.TryWriteBytes(wd.AsSpan(fcChpxFkp + 4), (int)0x20C);
    int chpx0Off = 64;
    wd[fcChpxFkp + 4 * (crun + 1) + 0] = (byte)(chpx0Off / 2);
    wd[fcChpxFkp + chpx0Off] = 0;
    wd[fcChpxFkp + 511] = (byte)crun;

    // ── Table stream ────────────────────────────────────────────
    var tblMs = new MemoryStream();
    Span<byte> b4 = stackalloc byte[4];
    tblMs.WriteByte(0x02);
    BitConverter.TryWriteBytes(b4, (uint)16); tblMs.Write(b4);
    BitConverter.TryWriteBytes(b4, (uint)0);   tblMs.Write(b4);
    BitConverter.TryWriteBytes(b4, (uint)ccp); tblMs.Write(b4);
    tblMs.WriteByte(0); tblMs.WriteByte(0);
    BitConverter.TryWriteBytes(b4, (uint)fcText); tblMs.Write(b4);
    tblMs.WriteByte(0); tblMs.WriteByte(0);
    int clxEnd = (int)tblMs.Position;
    int papxBteStart = clxEnd;
    BitConverter.TryWriteBytes(b4, (int)0x200);  tblMs.Write(b4);
    BitConverter.TryWriteBytes(b4, (int)0x300);  tblMs.Write(b4);
    BitConverter.TryWriteBytes(b4, (int)pnPapx); tblMs.Write(b4);
    int papxBteLen = (int)tblMs.Position - papxBteStart;
    int chpxBteStart = (int)tblMs.Position;
    BitConverter.TryWriteBytes(b4, (int)0x200);  tblMs.Write(b4);
    BitConverter.TryWriteBytes(b4, (int)0x300);  tblMs.Write(b4);
    BitConverter.TryWriteBytes(b4, (int)pnChpx); tblMs.Write(b4);
    int chpxBteLen = (int)tblMs.Position - chpxBteStart;

    // STSH ─────────────────────────────────────────────────────
    // LPStshi: cbStshi=4 + Stshi(cstd=2, cbSTDBaseInFile=10)
    // LPStd[0] Normal:    cbStd=24 + STD(10+4+4+6)
    // LPStd[1] Heading 1: cbStd=24 + STD(10+4+4+5+1pad)
    int stshStart = (int)tblMs.Position;
    BitConverter.TryWriteBytes(b4.Slice(0, 2), (ushort)4);  tblMs.Write(b4.Slice(0, 2));
    BitConverter.TryWriteBytes(b4.Slice(0, 2), (ushort)2);  tblMs.Write(b4.Slice(0, 2));  // cstd=2
    BitConverter.TryWriteBytes(b4.Slice(0, 2), (ushort)10); tblMs.Write(b4.Slice(0, 2)); // cbSTDBaseInFile

    void WriteStd(int sti, int istdBase, byte[] chpxUpx)
    {
        // stdfBase: sti(2)+stk+istdBase(2)+cupx+istdNext(2)+bchUpe(2)+grfstd(2) = 10
        int cbStd = 10 + 4 + 4 + (2 + chpxUpx.Length) + (chpxUpx.Length % 2 == 0 ? 0 : 1);
        void W16(ushort v)
        {
            var buf = new byte[2];
            BitConverter.TryWriteBytes(buf.AsSpan(), v);
            tblMs.Write(buf);
        }
        W16((ushort)cbStd);
        W16((ushort)sti);
        W16((ushort)((istdBase << 4) | 1));        // stk=1
        W16((ushort)((0xFFF << 4) | 2));           // cupx=2, istdNext=nil
        W16(0);                                    // bchUpe
        W16(0);                                    // grfstd
        // xstzName: 빈 이름 (cchData=0, null)
        W16(0); W16(0);
        // LPUpx[0] PAPX: cbUpx=2 + istd=sti
        W16(2);
        W16((ushort)sti);
        // LPUpx[1] CHPX: cbUpx + grpprl
        W16((ushort)chpxUpx.Length);
        tblMs.Write(chpxUpx);
        if (chpxUpx.Length % 2 != 0) tblMs.WriteByte(0);
    }

    // Normal STD: CHPX = sprmCRgFtc0(2) + ftc=0(2) = 4 byte
    var normalChpx = new byte[4];
    BitConverter.TryWriteBytes(normalChpx.AsSpan(0), (ushort)0x4A4F);
    BitConverter.TryWriteBytes(normalChpx.AsSpan(2), (ushort)0);
    WriteStd(sti: 0, istdBase: 0xFFF, chpxUpx: normalChpx);

    // Heading 1 STD: CHPX = sprmCFBold(2) + 1(1) = 3 byte
    var headingChpx = new byte[] { 0x35, 0x08, 0x01 };
    WriteStd(sti: 1, istdBase: 0, chpxUpx: headingChpx);

    int stshLen = (int)tblMs.Position - stshStart;

    // STTB FFN: 1 폰트
    int sttbStart = (int)tblMs.Position;
    BitConverter.TryWriteBytes(b4.Slice(0, 2), (ushort)0xFFFF); tblMs.Write(b4.Slice(0, 2));
    BitConverter.TryWriteBytes(b4.Slice(0, 2), (ushort)1);      tblMs.Write(b4.Slice(0, 2));
    BitConverter.TryWriteBytes(b4.Slice(0, 2), (ushort)0);      tblMs.Write(b4.Slice(0, 2));
    var nameBytes = Encoding.Unicode.GetBytes(fontName + "\0");
    int ffnSize = 40 + nameBytes.Length;
    BitConverter.TryWriteBytes(b4.Slice(0, 2), (ushort)(ffnSize / 2)); tblMs.Write(b4.Slice(0, 2));
    var ffnHdr = new byte[40];
    ffnHdr[0] = (byte)(ffnSize - 1);
    BitConverter.TryWriteBytes(ffnHdr.AsSpan(2), (ushort)400);
    tblMs.Write(ffnHdr);
    tblMs.Write(nameBytes);
    int sttbLen = (int)tblMs.Position - sttbStart;

    var table = tblMs.ToArray();

    BitConverter.TryWriteBytes(wd.AsSpan(0x00),   (ushort)0xA5EC);
    BitConverter.TryWriteBytes(wd.AsSpan(0x02),   (ushort)193);
    BitConverter.TryWriteBytes(wd.AsSpan(0x0A),   (ushort)0x0000);
    BitConverter.TryWriteBytes(wd.AsSpan(0x18),   (uint)fcText);
    BitConverter.TryWriteBytes(wd.AsSpan(0x4C),   (uint)ccp);
    BitConverter.TryWriteBytes(wd.AsSpan(0x01A2), (uint)0);
    BitConverter.TryWriteBytes(wd.AsSpan(0x01A6), (uint)clxEnd);
    BitConverter.TryWriteBytes(wd.AsSpan(0x00FA), (uint)chpxBteStart);
    BitConverter.TryWriteBytes(wd.AsSpan(0x00FE), (uint)chpxBteLen);
    BitConverter.TryWriteBytes(wd.AsSpan(0x0102), (uint)papxBteStart);
    BitConverter.TryWriteBytes(wd.AsSpan(0x0106), (uint)papxBteLen);
    BitConverter.TryWriteBytes(wd.AsSpan(0x00A2), (uint)stshStart);
    BitConverter.TryWriteBytes(wd.AsSpan(0x00A6), (uint)stshLen);
    BitConverter.TryWriteBytes(wd.AsSpan(0x0112), (uint)sttbStart);
    BitConverter.TryWriteBytes(wd.AsSpan(0x0116), (uint)sttbLen);

    var tmp = Path.Combine(Path.GetTempPath(), $"polydonky-smoke-{Guid.NewGuid():N}.doc");
    try
    {
        using (var root = OpenMcdf.RootStorage.Create(tmp))
        {
            using (var s = root.CreateStream("WordDocument")) s.Write(wd);
            using (var s = root.CreateStream("0Table"))       s.Write(table);
        }
        using var fs = File.OpenRead(tmp);
        var doc = new PolyDonky.Convert.Doc.DocBinaryReader().Read(fs);
        var para = doc.EnumerateParagraphs().First();
        var run  = para.Runs[0];

        SmokeHarness.Equal(OutlineLevel.H1, para.Style.Outline, "Heading 1 sti → H1");
        SmokeHarness.True(run.Style.Bold,
            $"Heading 1 STD CHPX Bold (직접 정의) → Run.Bold (got {run.Style.Bold})");
        SmokeHarness.Equal(fontName, run.Style.FontFamily ?? "",
            "Normal STD CHPX 폰트 → Heading 1 단락에 체이닝 상속");
    }
    finally { try { File.Delete(tmp); } catch { } }
}

// minimal Word 97-2003 WordDocument + 0Table stream bytes 합성.
// FIB(0x200) + UTF-16LE text  /  0Table = CLX(PCDT only, single piece)
static void BuildMinimalDocStreams(string mainText, out byte[] wordDoc, out byte[] table)
{
    byte[] textBytes = Encoding.Unicode.GetBytes(mainText);
    int fcTextStart  = 0x200;
    int ccpText      = mainText.Length;

    // CLX: 0x02 + lcb(4) + PlcPcd { aCP[2] + aPcd[1] = 8+8 = 16 bytes payload, lcb=16 }
    var clx = new MemoryStream();
    clx.WriteByte(0x02);
    Span<byte> buf4 = stackalloc byte[4];
    BitConverter.TryWriteBytes(buf4, (uint)16); clx.Write(buf4);
    BitConverter.TryWriteBytes(buf4, (uint)0);          clx.Write(buf4);   // aCP[0] = 0
    BitConverter.TryWriteBytes(buf4, (uint)ccpText);    clx.Write(buf4);   // aCP[1] = ccpText
    // PCD: flags(2) + fc(4) + prm(2). fc 의 bit 30=0 이면 unicode, fc = byte offset.
    clx.WriteByte(0); clx.WriteByte(0);                                    // flags
    BitConverter.TryWriteBytes(buf4, (uint)fcTextStart); clx.Write(buf4);  // fc = 0x200
    clx.WriteByte(0); clx.WriteByte(0);                                    // prm
    table = clx.ToArray();

    // FIB
    var fib = new byte[0x200];
    BitConverter.TryWriteBytes(fib.AsSpan(0x00),   (ushort)0xA5EC);  // magic
    BitConverter.TryWriteBytes(fib.AsSpan(0x02),   (ushort)193);     // nFib
    BitConverter.TryWriteBytes(fib.AsSpan(0x0A),   (ushort)0x0000);  // flags — fWhichTblStm = 0 → "0Table"
    BitConverter.TryWriteBytes(fib.AsSpan(0x18),   (uint)fcTextStart); // fcMin
    BitConverter.TryWriteBytes(fib.AsSpan(0x4C),   (uint)ccpText);   // ccpText
    BitConverter.TryWriteBytes(fib.AsSpan(0x01A2), (uint)0);         // fcClx (Table stream 시작)
    BitConverter.TryWriteBytes(fib.AsSpan(0x01A6), (uint)table.Length); // lcbClx

    wordDoc = new byte[fcTextStart + textBytes.Length];
    Buffer.BlockCopy(fib,       0, wordDoc, 0,             fib.Length);
    Buffer.BlockCopy(textBytes, 0, wordDoc, fcTextStart,   textBytes.Length);
}

static byte[] TamperDocumentJson(byte[] original)
{
    using var input = new MemoryStream(original);
    using var output = new MemoryStream();
    using (var inputZip = new System.IO.Compression.ZipArchive(input, System.IO.Compression.ZipArchiveMode.Read, leaveOpen: true))
    using (var outputZip = new System.IO.Compression.ZipArchive(output, System.IO.Compression.ZipArchiveMode.Create, leaveOpen: true))
    {
        foreach (var entry in inputZip.Entries)
        {
            byte[] payload;
            using (var es = entry.Open())
            using (var ms = new MemoryStream())
            {
                es.CopyTo(ms);
                payload = ms.ToArray();
            }

            if (entry.FullName == "content/document.json")
            {
                // 한 바이트 뒤집어 해시를 깬다.
                payload[^2] ^= 0x55;
            }

            var newEntry = outputZip.CreateEntry(entry.FullName, System.IO.Compression.CompressionLevel.Optimal);
            using var ws = newEntry.Open();
            ws.Write(payload, 0, payload.Length);
        }
    }
    return output.ToArray();
}


internal sealed class SmokeHarness
{
    private int _passed;
    private int _failed;
    private readonly List<string> _failures = new();

    public void Run(string name, Action body)
    {
        try
        {
            body();
            _passed++;
            Console.WriteLine($"  PASS  {name}");
        }
        catch (Exception ex)
        {
            _failed++;
            _failures.Add($"{name}: {ex.Message}");
            Console.WriteLine($"  FAIL  {name}");
            Console.WriteLine($"        {ex.GetType().Name}: {ex.Message}");
        }
    }

    public int Finish()
    {
        Console.WriteLine();
        Console.WriteLine(new string('-', 60));
        Console.WriteLine($"PolyDonky.SmokeTest: {_passed} passed, {_failed} failed");
        if (_failed > 0)
        {
            foreach (var f in _failures)
            {
                Console.WriteLine($"  - {f}");
            }
            return 1;
        }
        return 0;
    }

    public static void Equal<T>(T expected, T actual, string what)
    {
        if (!EqualityComparer<T>.Default.Equals(expected, actual))
        {
            throw new InvalidOperationException($"{what}: expected={expected}, actual={actual}");
        }
    }

    public static void True(bool condition, string what)
    {
        if (!condition)
        {
            throw new InvalidOperationException($"assertion failed: {what}");
        }
    }
}
