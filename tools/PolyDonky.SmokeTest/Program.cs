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
harness.Run("DOC (Word 97-2003 binary) — Phase 2a 표 인식 (sprmPFInTable + 0x07 셀)", DocBinaryPhase2aTable);
harness.Run("DOC (Word 97-2003 binary) — Phase 2b 셀 너비 (sprmTDefTable rgdxaCenter)", DocBinaryPhase2bTableWidths);
harness.Run("DOC (Word 97-2003 binary) — Phase 2c 셀 테두리 (TC97 BRC)", DocBinaryPhase2cCellBorders);
harness.Run("DOC (Word 97-2003 binary) — Phase 2d 셀 배경 (sprmTSetShd cvBack)", DocBinaryPhase2dCellShading);
harness.Run("DOC (Word 97-2003 binary) — Phase 2e 가로 셀 병합 (TC97 fFirstMerged/fMerged)", DocBinaryPhase2eHorizontalMerge);
harness.Run("DOC (Word 97-2003 binary) — Phase 2f 세로 셀 병합 (TC97 fVertMerge/fVertRestart)", DocBinaryPhase2fVerticalMerge);
harness.Run("DOC (Word 97-2003 binary) — Phase 2g 셀 안 멀티-단락", DocBinaryPhase2gMultiParaCell);
harness.Run("DOC (Word 97-2003 binary) — Phase 2h 중첩 표 (sprmPItap)", DocBinaryPhase2hNestedTable);
harness.Run("DOC (Word 97-2003 binary) — Phase 3a 필드 결과 보존 (0x13/0x14/0x15)", DocBinaryPhase3aFieldResult);
harness.Run("DOC (Word 97-2003 binary) — Phase 3b 페이지 break (sprmPFPageBreakBefore)", DocBinaryPhase3bPageBreak);
harness.Run("DOC (Word 97-2003 binary) — Phase 3c 섹션 분할 (PlcfSed)", DocBinaryPhase3cSections);
harness.Run("DOC (Word 97-2003 binary) — Phase 3d 헤더·푸터 (PlcfHdd + subdoc text)", DocBinaryPhase3dHeaderFooter);
harness.Run("DOC (Word 97-2003 binary) — Phase 3c-2 SEPX 페이지 크기·여백", DocBinaryPhase3c2SepxPageProps);
harness.Run("DOC (Word 97-2003 binary) — Phase 3e 이미지 placeholder (0x01)", DocBinaryPhase3eImagePlaceholder);
harness.Run("DOC (Word 97-2003 binary) — Phase 3e-2 이미지 PNG 데이터 추출 (PICF + Data stream)", DocBinaryPhase3e2ImageData);
harness.Run("DOC (Word 97-2003 binary) — Phase 3e-3 WMF placeable header 인식", DocBinaryPhase3e3WmfImage);
harness.Run("DOC (Word 97-2003 binary) — Phase 3e-3 DIB BITMAPINFOHEADER 인식", DocBinaryPhase3e3DibImage);
harness.Run("DOC (Word 97-2003 binary) — Phase 3e-4 표 셀 안 inline 이미지", DocBinaryPhase3e4ImageInTableCell);
harness.Run("DOC (Word 97-2003 binary) — Phase 3e-5 EMF EMR_HEADER 인식", DocBinaryPhase3e5EmfImage);
harness.Run("DOC (Word 97-2003 binary) — Phase 3e-6 OfficeArt BLIP_PNG (SpContainer)", DocBinaryPhase3e6BlipPng);
harness.Run("DOC (Word 97-2003 binary) — Phase 3e-6 OfficeArt BLIP_JPEG (atom)", DocBinaryPhase3e6BlipJpeg);
harness.Run("DOC (Word 97-2003 binary) — Phase 3e-7 OfficeArt BLIP_DIB (atom)", DocBinaryPhase3e7BlipDib);
harness.Run("DOC (Word 97-2003 binary) — Phase 3e-7 OfficeArt BLIP_WMF (uncompressed)", DocBinaryPhase3e7BlipWmfUncompressed);
harness.Run("DOC (Word 97-2003 binary) — Phase 3e-7 OfficeArt BLIP_EMF (zlib deflate)", DocBinaryPhase3e7BlipEmfDeflate);
harness.Run("DOC (Word 97-2003 binary) — Phase 3f OfficeArtBStoreContainer FBSE → BLIP 인덱스", DocBinaryPhase3fBStoreFbse);
harness.Run("DOC (Word 97-2003 binary) — Phase 3f-2 PICF FOPT pib → BStore lookup", DocBinaryPhase3f2PicfPibReferencesBStore);
harness.Run("DOC (Word 97-2003 binary) — Phase 3a-2 필드 instr (HYPERLINK / PAGE)", DocBinaryPhase3a2FieldInstr);
harness.Run("DOC (Word 97-2003 binary) — Phase 3g 각주/미주 텍스트 추출", DocBinaryPhase3gFootnotesEndnotes);
harness.Run("DOC (Word 97-2003 binary) — Phase 3a-3 필드 instr 추가 (NUMCHARS/FILENAME/SUBJECT/KEYWORDS/COMMENTS)", DocBinaryPhase3a3ExtraFieldInstr);
harness.Run("DOC (Word 97-2003 binary) — Phase 3a-4 필드 instr 추가 (SEQ/REF/STYLEREF/INCLUDETEXT/IF)", DocBinaryPhase3a4ReferenceFieldInstr);
harness.Run("DOC (Word 97-2003 binary) — Phase 3a-5 필드 인스트 인자 (Run.FieldArg)", DocBinaryPhase3a5FieldArg);
harness.Run("DOC (Word 97-2003 binary) — Phase 3f-3 FBSE foDelay → Data stream BLIP", DocBinaryPhase3f3FbseFoDelay);
harness.Run("DOC (Word 97-2003 binary) — Phase 3f-4 PlcSpaMom (FSPA) 파싱", DocBinaryPhase3f4PlcSpaMom);
harness.Run("DOC (Word 97-2003 binary) — Phase 3f-5 SpContainer spid → pib (BStore index) 맵", DocBinaryPhase3f5SpidPibMap);
harness.Run("DOC (Word 97-2003 binary) — Phase 3f-6 Floating shape ImageBlock 합성", DocBinaryPhase3f6FloatingShapeImages);
harness.Run("DOC (Word 97-2003 binary) — Phase 3g-2 0x02 footnote/endnote ref 자동 연결", DocBinaryPhase3g2FootnoteRefAutoLink);
harness.Run("DOC (Word 97-2003 binary) — Phase 3h 변경추적 (sprmCFRMarkIns/Del)", DocBinaryPhase3hRevisionMarks);
harness.Run("DOC (Word 97-2003 binary) — Phase 3h-2 변경추적 메타 (Author + DTTM)", DocBinaryPhase3h2RevisionMeta);
harness.Run("DOC (Word 97-2003 binary) — Phase 3h-3 단락 마크 변경추적", DocBinaryPhase3h3ParagraphRevision);
harness.Run("DOC (Word 97-2003 binary) — Phase 3i 책갈피 (SttbfBkmk + PlcfBkf + PlcfBkl)", DocBinaryPhase3iBookmarks);
harness.Run("DOC (Word 97-2003 binary) — Phase 3i-2 책갈피 Run marker 자동 삽입", DocBinaryPhase3i2BookmarkRunMarkers);
harness.Run("DOC (Word 97-2003 binary) — Phase 3j 주석 (annotations)", DocBinaryPhase3jComments);

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

static void DocBinaryPhase2aTable()
{
    // 1행 2셀 표 합성: text = "A\x07B\x07\rEnd\r"
    //   - 단락 1 (paraEndFc 의 PAPX 에 sprmPFInTable=1): "A\x07B\x07" — 0x07 으로 분리 → 셀 ["A", "B"]
    //   - 단락 2 (PAPX 비어 있음): "End" — 비-InTable → 표 마감
    // 결과: section.Blocks = [Table(1행 2셀), Paragraph("End")]
    const string text = "AB\rEnd\r";  // 9 chars
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

    // ── PAPX FKP — 2 단락 ──────────────────────────────────────
    // 단락 1: fc 범위 [0x200, 0x20A) — \r 는 CP 4 (fc=0x208)
    // 단락 2: fc 범위 [0x20A, 0x214) — \r 는 CP 8 (fc=0x210)
    int cpara = 2;
    BitConverter.TryWriteBytes(wd.AsSpan(fcPapxFkp + 0),  (int)0x200);
    BitConverter.TryWriteBytes(wd.AsSpan(fcPapxFkp + 4),  (int)0x20A);
    BitConverter.TryWriteBytes(wd.AsSpan(fcPapxFkp + 8),  (int)0x214);
    int papx0Off = 200, papx1Off = 220;
    wd[fcPapxFkp + 12 + 0 * 13] = (byte)(papx0Off / 2);
    wd[fcPapxFkp + 12 + 1 * 13] = (byte)(papx1Off / 2);

    // PapxInFkp[0]: cb=3 → grpprl size = 5 byte. istd(2)=0 + sprmPFInTable(2)=0x2416 + op(1)=1
    int p = fcPapxFkp + papx0Off;
    wd[p++] = 3;
    BitConverter.TryWriteBytes(wd.AsSpan(p), (ushort)0); p += 2;          // istd = 0
    BitConverter.TryWriteBytes(wd.AsSpan(p), (ushort)0x2416); p += 2;     // sprmPFInTable
    wd[p++] = 1;                                                           // InTable ON

    // PapxInFkp[1]: cb=0 형식, cb'=1 → grpprl size = 2 byte = istd 만 = 0
    p = fcPapxFkp + papx1Off;
    wd[p++] = 0;
    wd[p++] = 1;
    BitConverter.TryWriteBytes(wd.AsSpan(p), (ushort)0);                  // istd = 0

    wd[fcPapxFkp + 511] = (byte)cpara;

    // ── CHPX FKP — 비어 있음 ───────────────────────────────────
    int crun = 1;
    BitConverter.TryWriteBytes(wd.AsSpan(fcChpxFkp + 0), (int)0x200);
    BitConverter.TryWriteBytes(wd.AsSpan(fcChpxFkp + 4), (int)0x214);
    int chpx0Off = 64;
    wd[fcChpxFkp + 4 * (crun + 1) + 0] = (byte)(chpx0Off / 2);
    wd[fcChpxFkp + chpx0Off] = 0;
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
        var blocks = doc.Sections[0].Blocks;

        SmokeHarness.True(blocks.Count >= 2,
            $"section.Blocks.Count >= 2 (got {blocks.Count})");
        SmokeHarness.True(blocks[0] is PolyDonky.Core.Table,
            $"첫 블록은 Table (got {blocks[0].GetType().Name})");
        var tbl = (PolyDonky.Core.Table)blocks[0];
        SmokeHarness.Equal(1, tbl.Rows.Count, "행 수 1");
        SmokeHarness.Equal(2, tbl.Rows[0].Cells.Count, "셀 수 2");
        SmokeHarness.Equal("A",
            ((Paragraph)tbl.Rows[0].Cells[0].Blocks[0]).GetPlainText(),
            "셀 1 텍스트");
        SmokeHarness.Equal("B",
            ((Paragraph)tbl.Rows[0].Cells[1].Blocks[0]).GetPlainText(),
            "셀 2 텍스트");
        SmokeHarness.True(blocks[1] is Paragraph,
            $"두 번째 블록은 Paragraph (got {blocks[1].GetType().Name})");
        SmokeHarness.Equal("End", ((Paragraph)blocks[1]).GetPlainText(),
            "표 다음 단락 텍스트");
    }
    finally { try { File.Delete(tmp); } catch { } }
}

static void DocBinaryPhase2bTableWidths()
{
    // 1행 2셀 표 + TTP 단락에 sprmTDefTable 박음 — 셀 너비 1440/2880 twips.
    // 단락 1 (CP 0~4): "A\x07B\x07" + \r, InTable, !IsTtp
    // 단락 2 (CP 5):   "" + \r, InTable, IsTtp, sprmTDefTable
    // 단락 3 (CP 6~9): "End" + \r, 비-InTable
    const string text = "AB\r\rEnd\r";  // 10 chars
    int ccp = text.Length;
    var textBytes = Encoding.Unicode.GetBytes(text);
    int fcText = 0x200;
    int pnPapx = 4, pnChpx = 5;
    int fcPapxFkp = pnPapx * 512;
    int fcChpxFkp = pnChpx * 512;
    int wdSize = fcChpxFkp + 512;

    var wd = new byte[wdSize];
    Buffer.BlockCopy(textBytes, 0, wd, fcText, textBytes.Length);

    // ── PAPX FKP — 3 단락 ─────────────────────────────────────
    int cpara = 3;
    BitConverter.TryWriteBytes(wd.AsSpan(fcPapxFkp + 0),  (int)0x200);  // 단락 1 시작
    BitConverter.TryWriteBytes(wd.AsSpan(fcPapxFkp + 4),  (int)0x20A);  // 단락 2 시작
    BitConverter.TryWriteBytes(wd.AsSpan(fcPapxFkp + 8),  (int)0x20C);  // 단락 3 시작
    BitConverter.TryWriteBytes(wd.AsSpan(fcPapxFkp + 12), (int)0x214);  // 단락 3 끝+2
    int papx0Off = 64, papx1Off = 128, papx2Off = 300;
    wd[fcPapxFkp + 16 + 0 * 13] = (byte)(papx0Off / 2);
    wd[fcPapxFkp + 16 + 1 * 13] = (byte)(papx1Off / 2);
    wd[fcPapxFkp + 16 + 2 * 13] = (byte)(papx2Off / 2);

    // PapxInFkp[0]: cb=3 → istd(2)=0 + sprmPFInTable(2)=0x2416 + op(1)=1
    int p = fcPapxFkp + papx0Off;
    wd[p++] = 3;
    BitConverter.TryWriteBytes(wd.AsSpan(p), (ushort)0); p += 2;
    BitConverter.TryWriteBytes(wd.AsSpan(p), (ushort)0x2416); p += 2;
    wd[p++] = 1;

    // PapxInFkp[1]: cb=30 → grpprl 59 byte = istd(2)=0 + sprms(57)
    //   sprms = sprmPFInTable(2)+op(1) + sprmPFTtp(2)+op(1) + sprmTDefTable(2)+cb(2)+operand(47)
    //   operand = itcMac(1)=2 + rgdxaCenter[3](6) + rgTc[2](40) = 47 byte
    p = fcPapxFkp + papx1Off;
    wd[p++] = 30;
    BitConverter.TryWriteBytes(wd.AsSpan(p), (ushort)0); p += 2;          // istd=0
    BitConverter.TryWriteBytes(wd.AsSpan(p), (ushort)0x2416); p += 2;     // sprmPFInTable
    wd[p++] = 1;
    BitConverter.TryWriteBytes(wd.AsSpan(p), (ushort)0x2417); p += 2;     // sprmPFTtp
    wd[p++] = 1;
    BitConverter.TryWriteBytes(wd.AsSpan(p), (ushort)0xD608); p += 2;     // sprmTDefTable
    BitConverter.TryWriteBytes(wd.AsSpan(p), (ushort)47);     p += 2;     // operand cb
    wd[p++] = 2;                                                           // itcMac=2
    BitConverter.TryWriteBytes(wd.AsSpan(p), (short)0);     p += 2;       // rgdxaCenter[0]
    BitConverter.TryWriteBytes(wd.AsSpan(p), (short)1440);  p += 2;       // rgdxaCenter[1]
    BitConverter.TryWriteBytes(wd.AsSpan(p), (short)4320);  p += 2;       // rgdxaCenter[2]
    // rgTc[2] = 40 byte zeros (이미 0)
    p += 40;

    // PapxInFkp[2]: cb=0, cb'=1, istd 만
    p = fcPapxFkp + papx2Off;
    wd[p++] = 0;
    wd[p++] = 1;
    BitConverter.TryWriteBytes(wd.AsSpan(p), (ushort)0);

    wd[fcPapxFkp + 511] = (byte)cpara;

    // ── CHPX FKP — 비어 있음 ──────────────────────────────────
    int crun = 1;
    BitConverter.TryWriteBytes(wd.AsSpan(fcChpxFkp + 0), (int)0x200);
    BitConverter.TryWriteBytes(wd.AsSpan(fcChpxFkp + 4), (int)0x214);
    int chpx0Off = 64;
    wd[fcChpxFkp + 4 * (crun + 1) + 0] = (byte)(chpx0Off / 2);
    wd[fcChpxFkp + chpx0Off] = 0;
    wd[fcChpxFkp + 511] = (byte)crun;

    // ── Table stream ──────────────────────────────────────────
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

    var tblBytes = tblMs.ToArray();

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
            using (var s = root.CreateStream("0Table"))       s.Write(tblBytes);
        }
        using var fs = File.OpenRead(tmp);
        var doc = new PolyDonky.Convert.Doc.DocBinaryReader().Read(fs);
        var blocks = doc.Sections[0].Blocks;

        SmokeHarness.True(blocks.Count >= 1, "최소 1 블록");
        var tbl = blocks[0] as PolyDonky.Core.Table;
        SmokeHarness.True(tbl is not null, "첫 블록은 Table");
        SmokeHarness.Equal(2, tbl!.Columns.Count, "컬럼 수 = 2");
        // 1440 twips = 25.4 mm, 2880 twips = 50.8 mm
        SmokeHarness.True(Math.Abs(tbl.Columns[0].WidthMm - 25.4) < 0.1,
            $"컬럼 0 너비 ≈ 25.4 mm (got {tbl.Columns[0].WidthMm:F2})");
        SmokeHarness.True(Math.Abs(tbl.Columns[1].WidthMm - 50.8) < 0.1,
            $"컬럼 1 너비 ≈ 50.8 mm (got {tbl.Columns[1].WidthMm:F2})");
        SmokeHarness.Equal(1, tbl.Rows.Count, "행 수 = 1");
        SmokeHarness.Equal(2, tbl.Rows[0].Cells.Count, "셀 수 = 2");
    }
    finally { try { File.Delete(tmp); } catch { } }
}

static void DocBinaryPhase2cCellBorders()
{
    // 1행 2셀 표에 sprmTDefTable 의 rgTc 로 셀별 BRC 박음.
    // 셀 0 BorderTop: dpt=8 (1pt), brcType=1 (Solid), ico=6 (red)
    // 셀 1 BorderTop: dpt=16 (2pt), brcType=3 (Double), ico=2 (blue)
    const string text = "AB\r\rEnd\r";  // 10 chars (Phase 2b 와 동일)
    int ccp = text.Length;
    var textBytes = Encoding.Unicode.GetBytes(text);
    int fcText = 0x200;
    int pnPapx = 4, pnChpx = 5;
    int fcPapxFkp = pnPapx * 512;
    int fcChpxFkp = pnChpx * 512;
    int wdSize = fcChpxFkp + 512;

    var wd = new byte[wdSize];
    Buffer.BlockCopy(textBytes, 0, wd, fcText, textBytes.Length);

    // ── PAPX FKP — 3 단락 ────────────────────────────────────
    int cpara = 3;
    BitConverter.TryWriteBytes(wd.AsSpan(fcPapxFkp + 0),  (int)0x200);
    BitConverter.TryWriteBytes(wd.AsSpan(fcPapxFkp + 4),  (int)0x20A);
    BitConverter.TryWriteBytes(wd.AsSpan(fcPapxFkp + 8),  (int)0x20C);
    BitConverter.TryWriteBytes(wd.AsSpan(fcPapxFkp + 12), (int)0x214);
    int papx0Off = 64, papx1Off = 128, papx2Off = 300;
    wd[fcPapxFkp + 16 + 0 * 13] = (byte)(papx0Off / 2);
    wd[fcPapxFkp + 16 + 1 * 13] = (byte)(papx1Off / 2);
    wd[fcPapxFkp + 16 + 2 * 13] = (byte)(papx2Off / 2);

    // PapxInFkp[0]: cb=3 → istd(2)=0 + sprmPFInTable(2)+op(1)
    int p = fcPapxFkp + papx0Off;
    wd[p++] = 3;
    BitConverter.TryWriteBytes(wd.AsSpan(p), (ushort)0); p += 2;
    BitConverter.TryWriteBytes(wd.AsSpan(p), (ushort)0x2416); p += 2;
    wd[p++] = 1;

    // PapxInFkp[1]: cb=30 → grpprl 59 byte = istd(2)+sprms(57)
    //   sprms = sprmPFInTable(3) + sprmPFTtp(3) + sprmTDefTable(2+2+47) = 57 byte
    //   operand = itcMac(1)=2 + rgdxaCenter[3](6) + rgTc[2](40) = 47 byte
    p = fcPapxFkp + papx1Off;
    wd[p++] = 30;
    BitConverter.TryWriteBytes(wd.AsSpan(p), (ushort)0); p += 2;
    BitConverter.TryWriteBytes(wd.AsSpan(p), (ushort)0x2416); p += 2; wd[p++] = 1;
    BitConverter.TryWriteBytes(wd.AsSpan(p), (ushort)0x2417); p += 2; wd[p++] = 1;
    BitConverter.TryWriteBytes(wd.AsSpan(p), (ushort)0xD608); p += 2;
    BitConverter.TryWriteBytes(wd.AsSpan(p), (ushort)47);     p += 2;
    int operandStart = p;
    wd[p++] = 2;                                                           // itcMac=2
    BitConverter.TryWriteBytes(wd.AsSpan(p), (short)0);    p += 2;
    BitConverter.TryWriteBytes(wd.AsSpan(p), (short)1440); p += 2;
    BitConverter.TryWriteBytes(wd.AsSpan(p), (short)4320); p += 2;
    // rgTc[0]: 20 byte — bf(2)+wUnused(2)+brcTop(4)+brcLeft(4)+brcBottom(4)+brcRight(4)
    p += 4;  // bf + wUnused
    // brcTop: dpt=8 (1pt), brcType=1 (Solid), ico=6 (red)
    wd[p++] = 8; wd[p++] = 1; wd[p++] = 6; wd[p++] = 0;
    // brcLeft / brcBottom / brcRight: 모두 동일 (dpt=8, brcType=1, ico=1=Black)
    for (int k = 0; k < 3; k++)
    {
        wd[p++] = 8; wd[p++] = 1; wd[p++] = 1; wd[p++] = 0;
    }
    // rgTc[1]: brcTop dpt=16 (2pt), brcType=3 (Double), ico=2 (blue)
    p += 4;
    wd[p++] = 16; wd[p++] = 3; wd[p++] = 2; wd[p++] = 0;
    for (int k = 0; k < 3; k++)
    {
        wd[p++] = 8; wd[p++] = 1; wd[p++] = 1; wd[p++] = 0;
    }

    // PapxInFkp[2]: cb=0, cb'=1, istd 만
    p = fcPapxFkp + papx2Off;
    wd[p++] = 0;
    wd[p++] = 1;
    BitConverter.TryWriteBytes(wd.AsSpan(p), (ushort)0);
    wd[fcPapxFkp + 511] = (byte)cpara;

    // ── CHPX FKP — 비어 있음 ─────────────────────────────────
    int crun = 1;
    BitConverter.TryWriteBytes(wd.AsSpan(fcChpxFkp + 0), (int)0x200);
    BitConverter.TryWriteBytes(wd.AsSpan(fcChpxFkp + 4), (int)0x214);
    int chpx0Off = 64;
    wd[fcChpxFkp + 4 * (crun + 1) + 0] = (byte)(chpx0Off / 2);
    wd[fcChpxFkp + chpx0Off] = 0;
    wd[fcChpxFkp + 511] = (byte)crun;

    // ── Table stream ─────────────────────────────────────────
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
    var tblBytes = tblMs.ToArray();

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
            using (var s = root.CreateStream("0Table"))       s.Write(tblBytes);
        }
        using var fs = File.OpenRead(tmp);
        var doc = new PolyDonky.Convert.Doc.DocBinaryReader().Read(fs);
        var tbl = (PolyDonky.Core.Table)doc.Sections[0].Blocks[0];

        // 셀 0 BorderTop
        var b0 = tbl.Rows[0].Cells[0].BorderTop;
        SmokeHarness.True(b0.HasValue, "셀 0 BorderTop 설정됨");
        SmokeHarness.Equal(1.0, b0!.Value.ThicknessPt, "셀 0 BorderTop 두께 1pt (dpt=8)");
        SmokeHarness.Equal(BorderLineStyle.Solid, b0.Value.LineStyle, "셀 0 BorderTop 스타일 Solid");
        SmokeHarness.Equal("#FF0000", b0.Value.Color ?? "", "셀 0 BorderTop 색 빨강 (ico=6)");

        // 셀 1 BorderTop
        var b1 = tbl.Rows[0].Cells[1].BorderTop;
        SmokeHarness.True(b1.HasValue, "셀 1 BorderTop 설정됨");
        SmokeHarness.Equal(2.0, b1!.Value.ThicknessPt, "셀 1 BorderTop 두께 2pt (dpt=16)");
        SmokeHarness.Equal(BorderLineStyle.Double, b1.Value.LineStyle, "셀 1 BorderTop 스타일 Double");
        SmokeHarness.Equal("#0000FF", b1.Value.Color ?? "", "셀 1 BorderTop 색 파랑 (ico=2)");
    }
    finally { try { File.Delete(tmp); } catch { } }
}

static void DocBinaryPhase2dCellShading()
{
    // 1행 2셀 표 + TTP 의 PAPX 에 sprmTDefTable + sprmTSetShd.
    // sprmTSetShd: itcFirst=0, itcLim=2, cvBack=#FFFF00 (yellow), ipat=1 (solid).
    // 결과: 두 셀 모두 BackgroundColor="#FFFF00".
    const string text = "AB\r\rEnd\r";  // 10 chars; ccpText = 10
    int ccp = text.Length;
    var textBytes = Encoding.Unicode.GetBytes(text);
    int fcText = 0x200;
    int pnPapx = 4, pnChpx = 5;
    int fcPapxFkp = pnPapx * 512;
    int fcChpxFkp = pnChpx * 512;
    int wdSize = fcChpxFkp + 512;
    var wd = new byte[wdSize];
    Buffer.BlockCopy(textBytes, 0, wd, fcText, textBytes.Length);

    // PAPX FKP — 3 단락
    int cpara = 3;
    BitConverter.TryWriteBytes(wd.AsSpan(fcPapxFkp + 0),  (int)0x200);
    BitConverter.TryWriteBytes(wd.AsSpan(fcPapxFkp + 4),  (int)0x20A);
    BitConverter.TryWriteBytes(wd.AsSpan(fcPapxFkp + 8),  (int)0x20C);
    BitConverter.TryWriteBytes(wd.AsSpan(fcPapxFkp + 12), (int)0x214);
    int papx0Off = 64, papx1Off = 128, papx2Off = 320;
    wd[fcPapxFkp + 16 + 0 * 13] = (byte)(papx0Off / 2);
    wd[fcPapxFkp + 16 + 1 * 13] = (byte)(papx1Off / 2);
    wd[fcPapxFkp + 16 + 2 * 13] = (byte)(papx2Off / 2);

    // PapxInFkp[0]: sprmPFInTable=1
    int p = fcPapxFkp + papx0Off;
    wd[p++] = 3;
    BitConverter.TryWriteBytes(wd.AsSpan(p), (ushort)0); p += 2;
    BitConverter.TryWriteBytes(wd.AsSpan(p), (ushort)0x2416); p += 2;
    wd[p++] = 1;

    // PapxInFkp[1]: sprms = sprmPFInTable(3) + sprmPFTtp(3) + sprmTDefTable(51) + sprmTSetShd(15) = 72 byte
    //               grpprl = istd(2) + sprms(72) = 74 byte. cb=0 형식, cb'=37 (=74)
    p = fcPapxFkp + papx1Off;
    wd[p++] = 0;
    wd[p++] = 37;
    BitConverter.TryWriteBytes(wd.AsSpan(p), (ushort)0); p += 2;          // istd
    BitConverter.TryWriteBytes(wd.AsSpan(p), (ushort)0x2416); p += 2; wd[p++] = 1;
    BitConverter.TryWriteBytes(wd.AsSpan(p), (ushort)0x2417); p += 2; wd[p++] = 1;
    BitConverter.TryWriteBytes(wd.AsSpan(p), (ushort)0xD608); p += 2;
    BitConverter.TryWriteBytes(wd.AsSpan(p), (ushort)47);     p += 2;
    wd[p++] = 2;                                                           // itcMac
    BitConverter.TryWriteBytes(wd.AsSpan(p), (short)0);    p += 2;
    BitConverter.TryWriteBytes(wd.AsSpan(p), (short)1440); p += 2;
    BitConverter.TryWriteBytes(wd.AsSpan(p), (short)4320); p += 2;
    p += 4;  // TC0 bf+wUnused
    wd[p++] = 8; wd[p++] = 1; wd[p++] = 1; wd[p++] = 0;
    for (int k = 0; k < 3; k++) { wd[p++] = 8; wd[p++] = 1; wd[p++] = 1; wd[p++] = 0; }
    p += 4;  // TC1
    wd[p++] = 8; wd[p++] = 1; wd[p++] = 1; wd[p++] = 0;
    for (int k = 0; k < 3; k++) { wd[p++] = 8; wd[p++] = 1; wd[p++] = 1; wd[p++] = 0; }
    // sprmTSetShd
    BitConverter.TryWriteBytes(wd.AsSpan(p), (ushort)0xD612); p += 2;
    wd[p++] = 12;                                                          // cb
    wd[p++] = 0;                                                           // itcFirst
    wd[p++] = 2;                                                           // itcLim
    wd[p++] = 0; wd[p++] = 0; wd[p++] = 0; wd[p++] = 0xFF;                 // cvFore auto
    wd[p++] = 0xFF; wd[p++] = 0xFF; wd[p++] = 0x00; wd[p++] = 0x00;        // cvBack #FFFF00
    BitConverter.TryWriteBytes(wd.AsSpan(p), (ushort)1); p += 2;           // ipat solid

    // PapxInFkp[2]: 빈 sprm
    p = fcPapxFkp + papx2Off;
    wd[p++] = 0;
    wd[p++] = 1;
    BitConverter.TryWriteBytes(wd.AsSpan(p), (ushort)0);
    wd[fcPapxFkp + 511] = (byte)cpara;

    // CHPX FKP — 비어 있음
    int crun = 1;
    BitConverter.TryWriteBytes(wd.AsSpan(fcChpxFkp + 0), (int)0x200);
    BitConverter.TryWriteBytes(wd.AsSpan(fcChpxFkp + 4), (int)0x214);
    int chpx0Off = 64;
    wd[fcChpxFkp + 4 * (crun + 1) + 0] = (byte)(chpx0Off / 2);
    wd[fcChpxFkp + chpx0Off] = 0;
    wd[fcChpxFkp + 511] = (byte)crun;

    // Table stream
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
    var tblBytes = tblMs.ToArray();

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
            using (var s = root.CreateStream("0Table"))       s.Write(tblBytes);
        }
        using var fs = File.OpenRead(tmp);
        var doc = new PolyDonky.Convert.Doc.DocBinaryReader().Read(fs);
        var tbl = (PolyDonky.Core.Table)doc.Sections[0].Blocks[0];

        SmokeHarness.Equal("#FFFF00", tbl.Rows[0].Cells[0].BackgroundColor ?? "",
            "셀 0 배경색 = #FFFF00 (yellow)");
        SmokeHarness.Equal("#FFFF00", tbl.Rows[0].Cells[1].BackgroundColor ?? "",
            "셀 1 배경색 = #FFFF00");
    }
    finally { try { File.Delete(tmp); } catch { } }
}

static void DocBinaryPhase2eHorizontalMerge()
{
    // 1행 3셀 표:
    //   TC0: fFirstMerged=1, fMerged=1 (병합 시작)
    //   TC1: fFirstMerged=0, fMerged=1 (병합 흡수)
    //   TC2: 일반
    // 결과: row.Cells.Count == 2, Cells[0].ColumnSpan == 2 (셀 0+1 병합), Cells[1].ColumnSpan == 1
    const string text = "ABC\r\rEnd\r";  // A 0x07 B 0x07 C 0x07 \r \r E n d \r = 11 chars
    int ccp = text.Length;
    var textBytes = Encoding.Unicode.GetBytes(text);
    int fcText = 0x200;
    int pnPapx = 4, pnChpx = 5;
    int fcPapxFkp = pnPapx * 512;
    int fcChpxFkp = pnChpx * 512;
    int wdSize = fcChpxFkp + 512;
    var wd = new byte[wdSize];
    Buffer.BlockCopy(textBytes, 0, wd, fcText, textBytes.Length);

    // PAPX FKP — 3 단락
    // 단락 1 (CP 0~6, \r at CP 6, fc=0x20C): InTable, !IsTtp
    // 단락 2 (CP 7, \r at CP 7, fc=0x20E): IsTtp, sprmTDefTable
    // 단락 3 (CP 8~10, \r at CP 10, fc=0x214): 비-InTable
    int cpara = 3;
    BitConverter.TryWriteBytes(wd.AsSpan(fcPapxFkp + 0),  (int)0x200);
    BitConverter.TryWriteBytes(wd.AsSpan(fcPapxFkp + 4),  (int)0x20E);
    BitConverter.TryWriteBytes(wd.AsSpan(fcPapxFkp + 8),  (int)0x210);
    BitConverter.TryWriteBytes(wd.AsSpan(fcPapxFkp + 12), (int)0x216);
    int papx0Off = 64, papx1Off = 128, papx2Off = 340;
    wd[fcPapxFkp + 16 + 0 * 13] = (byte)(papx0Off / 2);
    wd[fcPapxFkp + 16 + 1 * 13] = (byte)(papx1Off / 2);
    wd[fcPapxFkp + 16 + 2 * 13] = (byte)(papx2Off / 2);

    // PapxInFkp[0]: sprmPFInTable=1
    int p = fcPapxFkp + papx0Off;
    wd[p++] = 3;
    BitConverter.TryWriteBytes(wd.AsSpan(p), (ushort)0); p += 2;
    BitConverter.TryWriteBytes(wd.AsSpan(p), (ushort)0x2416); p += 2;
    wd[p++] = 1;

    // PapxInFkp[1]: sprms = sprmPFInTable(3) + sprmPFTtp(3) + sprmTDefTable(2+2+1+8+60) = 79 byte
    //   sprmTDefTable operand: itcMac(1)=3 + rgdxaCenter[4](8) + rgTc[3](60) = 69 byte
    //   sprm header(2) + cb(2) + 69 = 73 byte sprmTDefTable
    //   sprms = 3 + 3 + 73 = 79 byte
    //   grpprl = istd(2) + sprms(79) = 81 byte. cb=41 → 41*2-1=81. OK
    p = fcPapxFkp + papx1Off;
    wd[p++] = 41;
    BitConverter.TryWriteBytes(wd.AsSpan(p), (ushort)0); p += 2;          // istd
    BitConverter.TryWriteBytes(wd.AsSpan(p), (ushort)0x2416); p += 2; wd[p++] = 1;
    BitConverter.TryWriteBytes(wd.AsSpan(p), (ushort)0x2417); p += 2; wd[p++] = 1;
    BitConverter.TryWriteBytes(wd.AsSpan(p), (ushort)0xD608); p += 2;
    BitConverter.TryWriteBytes(wd.AsSpan(p), (ushort)69);     p += 2;
    wd[p++] = 3;                                                           // itcMac
    BitConverter.TryWriteBytes(wd.AsSpan(p), (short)0);    p += 2;
    BitConverter.TryWriteBytes(wd.AsSpan(p), (short)1440); p += 2;
    BitConverter.TryWriteBytes(wd.AsSpan(p), (short)2880); p += 2;
    BitConverter.TryWriteBytes(wd.AsSpan(p), (short)4320); p += 2;
    // TC0: bf bit 0+1 = 0x03 (fFirstMerged+fMerged)
    BitConverter.TryWriteBytes(wd.AsSpan(p), (ushort)0x0003); p += 2;
    p += 2;  // wUnused
    for (int k = 0; k < 4; k++) { wd[p++] = 0; wd[p++] = 0; wd[p++] = 0; wd[p++] = 0; }  // 4 BRC zeros
    // TC1: bf bit 1 = 0x02 (fMerged only)
    BitConverter.TryWriteBytes(wd.AsSpan(p), (ushort)0x0002); p += 2;
    p += 2;
    for (int k = 0; k < 4; k++) { wd[p++] = 0; wd[p++] = 0; wd[p++] = 0; wd[p++] = 0; }
    // TC2: bf = 0 (일반 셀)
    p += 4;  // bf+wUnused
    for (int k = 0; k < 4; k++) { wd[p++] = 0; wd[p++] = 0; wd[p++] = 0; wd[p++] = 0; }

    // PapxInFkp[2]
    p = fcPapxFkp + papx2Off;
    wd[p++] = 0;
    wd[p++] = 1;
    BitConverter.TryWriteBytes(wd.AsSpan(p), (ushort)0);
    wd[fcPapxFkp + 511] = (byte)cpara;

    // CHPX FKP — 비어 있음
    int crun = 1;
    BitConverter.TryWriteBytes(wd.AsSpan(fcChpxFkp + 0), (int)0x200);
    BitConverter.TryWriteBytes(wd.AsSpan(fcChpxFkp + 4), (int)0x216);
    int chpx0Off = 64;
    wd[fcChpxFkp + 4 * (crun + 1) + 0] = (byte)(chpx0Off / 2);
    wd[fcChpxFkp + chpx0Off] = 0;
    wd[fcChpxFkp + 511] = (byte)crun;

    // Table stream
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
    var tblBytes = tblMs.ToArray();

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
            using (var s = root.CreateStream("0Table"))       s.Write(tblBytes);
        }
        using var fs = File.OpenRead(tmp);
        var doc = new PolyDonky.Convert.Doc.DocBinaryReader().Read(fs);
        var tbl = (PolyDonky.Core.Table)doc.Sections[0].Blocks[0];

        SmokeHarness.Equal(1, tbl.Rows.Count, "행 수 = 1");
        SmokeHarness.Equal(2, tbl.Rows[0].Cells.Count,
            "셀 수 = 2 (TC0+TC1 가로 병합, TC2 별개)");
        SmokeHarness.Equal(2, tbl.Rows[0].Cells[0].ColumnSpan,
            "Cells[0].ColumnSpan = 2 (TC0+TC1 흡수)");
        SmokeHarness.Equal(1, tbl.Rows[0].Cells[1].ColumnSpan,
            "Cells[1].ColumnSpan = 1 (TC2 일반)");
    }
    finally { try { File.Delete(tmp); } catch { } }
}

static void DocBinaryPhase2fVerticalMerge()
{
    // 2행 2셀 표, col 0 만 세로 병합 (TC97 bf bit 5=fVertMerge, bit 6=fVertRestart).
    //   행 1 TC0: bf=0x60 (VertMerge+VertRestart) → 병합 시작
    //   행 1 TC1: 일반
    //   행 2 TC0: bf=0x20 (VertMerge only) → 흡수
    //   행 2 TC1: 일반
    // 결과: Rows[0].Cells = [A(RowSpan=2), B], Rows[1].Cells = [D] (col 0 흡수됨, col 1 살아남음).
    const string text = "AB\r\rCD\r\rEnd\r";  // A 0x07 B 0x07 \r \r C 0x07 D 0x07 \r \r E n d \r = 16 chars
    int ccp = text.Length;
    var textBytes = Encoding.Unicode.GetBytes(text);
    int fcText = 0x200;
    int pnPapx = 4, pnChpx = 5;
    int fcPapxFkp = pnPapx * 512;
    int fcChpxFkp = pnChpx * 512;
    int wdSize = fcChpxFkp + 512;
    var wd = new byte[wdSize];
    Buffer.BlockCopy(textBytes, 0, wd, fcText, textBytes.Length);

    // PAPX FKP — 5 단락
    // rgfc[6]: 단락별 시작 fc + 끝
    int cpara = 5;
    BitConverter.TryWriteBytes(wd.AsSpan(fcPapxFkp + 0),  (int)0x200);
    BitConverter.TryWriteBytes(wd.AsSpan(fcPapxFkp + 4),  (int)0x20A);
    BitConverter.TryWriteBytes(wd.AsSpan(fcPapxFkp + 8),  (int)0x20C);
    BitConverter.TryWriteBytes(wd.AsSpan(fcPapxFkp + 12), (int)0x216);
    BitConverter.TryWriteBytes(wd.AsSpan(fcPapxFkp + 16), (int)0x218);
    BitConverter.TryWriteBytes(wd.AsSpan(fcPapxFkp + 20), (int)0x220);

    int papx0Off = 100, papx1Off = 110, papx2Off = 200, papx3Off = 210, papx4Off = 300;
    wd[fcPapxFkp + 24 + 0 * 13] = (byte)(papx0Off / 2);
    wd[fcPapxFkp + 24 + 1 * 13] = (byte)(papx1Off / 2);
    wd[fcPapxFkp + 24 + 2 * 13] = (byte)(papx2Off / 2);
    wd[fcPapxFkp + 24 + 3 * 13] = (byte)(papx3Off / 2);
    wd[fcPapxFkp + 24 + 4 * 13] = (byte)(papx4Off / 2);

    // Helper for writing PAPX[1] / PAPX[3] (sprmPFInTable + sprmPFTtp + sprmTDefTable)
    void WriteTtpPapx(int off, ushort tc0bf)
    {
        int q = fcPapxFkp + off;
        wd[q++] = 30;  // cb=30 → grpprl 59 byte
        BitConverter.TryWriteBytes(wd.AsSpan(q), (ushort)0); q += 2;          // istd
        BitConverter.TryWriteBytes(wd.AsSpan(q), (ushort)0x2416); q += 2; wd[q++] = 1;
        BitConverter.TryWriteBytes(wd.AsSpan(q), (ushort)0x2417); q += 2; wd[q++] = 1;
        BitConverter.TryWriteBytes(wd.AsSpan(q), (ushort)0xD608); q += 2;
        BitConverter.TryWriteBytes(wd.AsSpan(q), (ushort)47);     q += 2;
        wd[q++] = 2;                                                          // itcMac=2
        BitConverter.TryWriteBytes(wd.AsSpan(q), (short)0);    q += 2;
        BitConverter.TryWriteBytes(wd.AsSpan(q), (short)1440); q += 2;
        BitConverter.TryWriteBytes(wd.AsSpan(q), (short)2880); q += 2;
        // TC0: bf = tc0bf, 나머지 0
        BitConverter.TryWriteBytes(wd.AsSpan(q), tc0bf); q += 2;
        q += 2;  // wUnused
        for (int k = 0; k < 4; k++) { q += 4; }  // 4 BRC zeros (already 0)
        // TC1: bf=0, 나머지 0
        q += 20;
    }

    // PapxInFkp[0]: sprmPFInTable=1 (3 byte sprm + istd 2 byte → cb=3 → grpprl 5 byte)
    int p = fcPapxFkp + papx0Off;
    wd[p++] = 3;
    BitConverter.TryWriteBytes(wd.AsSpan(p), (ushort)0); p += 2;
    BitConverter.TryWriteBytes(wd.AsSpan(p), (ushort)0x2416); p += 2;
    wd[p++] = 1;

    // PapxInFkp[1]: TTP, TC0 bf = 0x60 (VertMerge + VertRestart)
    WriteTtpPapx(papx1Off, 0x0060);

    // PapxInFkp[2]: 같은 sprmPFInTable=1
    p = fcPapxFkp + papx2Off;
    wd[p++] = 3;
    BitConverter.TryWriteBytes(wd.AsSpan(p), (ushort)0); p += 2;
    BitConverter.TryWriteBytes(wd.AsSpan(p), (ushort)0x2416); p += 2;
    wd[p++] = 1;

    // PapxInFkp[3]: TTP, TC0 bf = 0x20 (VertMerge only — 흡수)
    WriteTtpPapx(papx3Off, 0x0020);

    // PapxInFkp[4]: 빈 sprm
    p = fcPapxFkp + papx4Off;
    wd[p++] = 0;
    wd[p++] = 1;
    BitConverter.TryWriteBytes(wd.AsSpan(p), (ushort)0);
    wd[fcPapxFkp + 511] = (byte)cpara;

    // CHPX FKP — 비어 있음
    int crun = 1;
    BitConverter.TryWriteBytes(wd.AsSpan(fcChpxFkp + 0), (int)0x200);
    BitConverter.TryWriteBytes(wd.AsSpan(fcChpxFkp + 4), (int)0x220);
    int chpx0Off = 64;
    wd[fcChpxFkp + 4 * (crun + 1) + 0] = (byte)(chpx0Off / 2);
    wd[fcChpxFkp + chpx0Off] = 0;
    wd[fcChpxFkp + 511] = (byte)crun;

    // Table stream
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
    var tblBytes = tblMs.ToArray();

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
            using (var s = root.CreateStream("0Table"))       s.Write(tblBytes);
        }
        using var fs = File.OpenRead(tmp);
        var doc = new PolyDonky.Convert.Doc.DocBinaryReader().Read(fs);
        var tbl = (PolyDonky.Core.Table)doc.Sections[0].Blocks[0];

        SmokeHarness.Equal(2, tbl.Rows.Count, "행 수 = 2");
        SmokeHarness.Equal(2, tbl.Rows[0].Cells.Count, "Rows[0].Cells.Count = 2");
        SmokeHarness.Equal(2, tbl.Rows[0].Cells[0].RowSpan,
            "Rows[0].Cells[0].RowSpan = 2 (세로 병합 시작)");
        SmokeHarness.Equal(1, tbl.Rows[1].Cells.Count,
            "Rows[1].Cells.Count = 1 (col 0 흡수, col 1 살아남음)");
        SmokeHarness.Equal("D", ((Paragraph)tbl.Rows[1].Cells[0].Blocks[0]).GetPlainText(),
            "Rows[1].Cells[0] = 'D' (col 1 의 셀)");
    }
    finally { try { File.Delete(tmp); } catch { } }
}

static void DocBinaryPhase2gMultiParaCell()
{
    // 1행 2셀 표, 셀 0 가 두 단락 보유.
    //   단락 1 "Line1\r" InTable, !IsTtp — 0x07 없음 → 셀 0 의 첫 단락
    //   단락 2 "Line2\x07OtherCell\x07\r" InTable, !IsTtp → 첫 0x07 에서 셀 0 finalize([Line1, Line2]),
    //                                                       두 번째 0x07 에서 셀 1 finalize([OtherCell])
    //   단락 3 "\r" TTP — 행 종료
    //   단락 4 "End\r" 비-InTable — 표 마감
    const string text = "Line1\rLine2OtherCell\r\rEnd\r";  // 28 chars (0x07 2 곳)
    int ccp = text.Length;
    var textBytes = Encoding.Unicode.GetBytes(text);
    int fcText = 0x200;
    int pnPapx = 4, pnChpx = 5;
    int fcPapxFkp = pnPapx * 512;
    int fcChpxFkp = pnChpx * 512;
    int wdSize = fcChpxFkp + 512;
    var wd = new byte[wdSize];
    Buffer.BlockCopy(textBytes, 0, wd, fcText, textBytes.Length);

    // PAPX FKP — 4 단락
    // 단락 \r 위치 CP 5, 22, 23, 27 → fc 0x20A, 0x22C, 0x22E, 0x236
    int cpara = 4;
    BitConverter.TryWriteBytes(wd.AsSpan(fcPapxFkp + 0),  (int)0x200);
    BitConverter.TryWriteBytes(wd.AsSpan(fcPapxFkp + 4),  (int)0x20C);   // 단락 1 끝+2
    BitConverter.TryWriteBytes(wd.AsSpan(fcPapxFkp + 8),  (int)0x22E);   // 단락 2 끝+2
    BitConverter.TryWriteBytes(wd.AsSpan(fcPapxFkp + 12), (int)0x230);   // 단락 3 (TTP) 끝+2
    BitConverter.TryWriteBytes(wd.AsSpan(fcPapxFkp + 16), (int)0x238);   // 단락 4 끝+2

    int papx0Off = 80, papx1Off = 90, papx2Off = 100, papx3Off = 200;
    wd[fcPapxFkp + 20 + 0 * 13] = (byte)(papx0Off / 2);
    wd[fcPapxFkp + 20 + 1 * 13] = (byte)(papx1Off / 2);
    wd[fcPapxFkp + 20 + 2 * 13] = (byte)(papx2Off / 2);
    wd[fcPapxFkp + 20 + 3 * 13] = (byte)(papx3Off / 2);

    // PapxInFkp[0] / [1]: sprmPFInTable=1 (cb=3)
    int p = fcPapxFkp + papx0Off;
    wd[p++] = 3;
    BitConverter.TryWriteBytes(wd.AsSpan(p), (ushort)0); p += 2;
    BitConverter.TryWriteBytes(wd.AsSpan(p), (ushort)0x2416); p += 2;
    wd[p++] = 1;

    p = fcPapxFkp + papx1Off;
    wd[p++] = 3;
    BitConverter.TryWriteBytes(wd.AsSpan(p), (ushort)0); p += 2;
    BitConverter.TryWriteBytes(wd.AsSpan(p), (ushort)0x2416); p += 2;
    wd[p++] = 1;

    // PapxInFkp[2]: TTP + sprmTDefTable (itcMac=2, 모두 일반)
    p = fcPapxFkp + papx2Off;
    wd[p++] = 30;
    BitConverter.TryWriteBytes(wd.AsSpan(p), (ushort)0); p += 2;
    BitConverter.TryWriteBytes(wd.AsSpan(p), (ushort)0x2416); p += 2; wd[p++] = 1;
    BitConverter.TryWriteBytes(wd.AsSpan(p), (ushort)0x2417); p += 2; wd[p++] = 1;
    BitConverter.TryWriteBytes(wd.AsSpan(p), (ushort)0xD608); p += 2;
    BitConverter.TryWriteBytes(wd.AsSpan(p), (ushort)47);     p += 2;
    wd[p++] = 2;  // itcMac
    BitConverter.TryWriteBytes(wd.AsSpan(p), (short)0);    p += 2;
    BitConverter.TryWriteBytes(wd.AsSpan(p), (short)1440); p += 2;
    BitConverter.TryWriteBytes(wd.AsSpan(p), (short)2880); p += 2;
    p += 40;  // TC[0] + TC[1] 모두 0

    // PapxInFkp[3]: 빈 sprm
    p = fcPapxFkp + papx3Off;
    wd[p++] = 0;
    wd[p++] = 1;
    BitConverter.TryWriteBytes(wd.AsSpan(p), (ushort)0);
    wd[fcPapxFkp + 511] = (byte)cpara;

    // CHPX FKP — 비어 있음
    int crun = 1;
    BitConverter.TryWriteBytes(wd.AsSpan(fcChpxFkp + 0), (int)0x200);
    BitConverter.TryWriteBytes(wd.AsSpan(fcChpxFkp + 4), (int)0x238);
    int chpx0Off = 64;
    wd[fcChpxFkp + 4 * (crun + 1) + 0] = (byte)(chpx0Off / 2);
    wd[fcChpxFkp + chpx0Off] = 0;
    wd[fcChpxFkp + 511] = (byte)crun;

    // Table stream
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
    var tblBytes = tblMs.ToArray();

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
            using (var s = root.CreateStream("0Table"))       s.Write(tblBytes);
        }
        using var fs = File.OpenRead(tmp);
        var doc = new PolyDonky.Convert.Doc.DocBinaryReader().Read(fs);
        var tbl = (PolyDonky.Core.Table)doc.Sections[0].Blocks[0];

        SmokeHarness.Equal(1, tbl.Rows.Count, "행 수 = 1");
        SmokeHarness.Equal(2, tbl.Rows[0].Cells.Count, "Cells.Count = 2");
        SmokeHarness.Equal(2, tbl.Rows[0].Cells[0].Blocks.Count,
            "Cells[0].Blocks.Count = 2 (멀티-단락 셀)");
        SmokeHarness.Equal("Line1",
            ((Paragraph)tbl.Rows[0].Cells[0].Blocks[0]).GetPlainText(),
            "Cells[0].Blocks[0] = 'Line1'");
        SmokeHarness.Equal("Line2",
            ((Paragraph)tbl.Rows[0].Cells[0].Blocks[1]).GetPlainText(),
            "Cells[0].Blocks[1] = 'Line2'");
        SmokeHarness.Equal(1, tbl.Rows[0].Cells[1].Blocks.Count,
            "Cells[1].Blocks.Count = 1");
        SmokeHarness.Equal("OtherCell",
            ((Paragraph)tbl.Rows[0].Cells[1].Blocks[0]).GetPlainText(),
            "Cells[1] = 'OtherCell'");
    }
    finally { try { File.Delete(tmp); } catch { } }
}

static void DocBinaryPhase2hNestedTable()
{
    // 1행 1셀 외곽 표 안에 1행 1셀 내부 표.
    //   단락 1 "Outer\r"           itap=1, InTable (외곽 셀의 첫 단락)
    //   단락 2 "Inner\x07\r"       itap=2, InTable (내부 셀, 0x07 으로 셀 종료)
    //   단락 3 "\r"                itap=2, IsTtp + sprmTDefTable (내부 TTP)
    //   단락 4 "\x07\r"            itap=1, InTable (외곽 셀의 cell mark — 내부 표 다음 셀 종료)
    //   단락 5 "\r"                itap=1, IsTtp + sprmTDefTable (외곽 TTP)
    //   단락 6 "End\r"             itap=0, 비-InTable
    // 결과:
    //   section.Blocks = [Table(outer), Paragraph("End")]
    //   outerTable.Rows[0].Cells[0].Blocks = [Para("Outer"), Table(inner)]
    //   innerTable.Rows[0].Cells[0].Blocks = [Para("Inner")]
    const string text = "Outer\rInner\r\r\r\rEnd\r";  // 0x07 두 곳 박은 후 21 chars
    int ccp = text.Length;
    var textBytes = Encoding.Unicode.GetBytes(text);
    int fcText = 0x200;
    int pnPapx = 4, pnChpx = 5;
    int fcPapxFkp = pnPapx * 512;
    int fcChpxFkp = pnChpx * 512;
    int wdSize = fcChpxFkp + 512;
    var wd = new byte[wdSize];
    Buffer.BlockCopy(textBytes, 0, wd, fcText, textBytes.Length);

    // PAPX FKP — 6 단락. rgfc[7]:
    int cpara = 6;
    BitConverter.TryWriteBytes(wd.AsSpan(fcPapxFkp + 0),  (int)0x200);  // 단락 1 시작
    BitConverter.TryWriteBytes(wd.AsSpan(fcPapxFkp + 4),  (int)0x20C);  // 단락 2 시작 (= "Outer\r" 다음)
    BitConverter.TryWriteBytes(wd.AsSpan(fcPapxFkp + 8),  (int)0x21A);  // 단락 3 (= "Inner\x07\r" 다음)
    BitConverter.TryWriteBytes(wd.AsSpan(fcPapxFkp + 12), (int)0x21C);  // 단락 4 (= inner TTP "\r" 다음)
    BitConverter.TryWriteBytes(wd.AsSpan(fcPapxFkp + 16), (int)0x220);  // 단락 5 (= "\x07\r" 다음)
    BitConverter.TryWriteBytes(wd.AsSpan(fcPapxFkp + 20), (int)0x222);  // 단락 6 (= outer TTP "\r" 다음)
    BitConverter.TryWriteBytes(wd.AsSpan(fcPapxFkp + 24), (int)0x22A);  // 단락 6 끝+2 ("End\r" 다음)

    int papx0Off = 120, papx1Off = 140, papx2Off = 160, papx3Off = 210, papx4Off = 230, papx5Off = 290;
    wd[fcPapxFkp + 28 + 0 * 13] = (byte)(papx0Off / 2);
    wd[fcPapxFkp + 28 + 1 * 13] = (byte)(papx1Off / 2);
    wd[fcPapxFkp + 28 + 2 * 13] = (byte)(papx2Off / 2);
    wd[fcPapxFkp + 28 + 3 * 13] = (byte)(papx3Off / 2);
    wd[fcPapxFkp + 28 + 4 * 13] = (byte)(papx4Off / 2);
    wd[fcPapxFkp + 28 + 5 * 13] = (byte)(papx5Off / 2);

    // Helper — sprmPItap (0x6649, spra=3, 4-byte signed) + sprmPFInTable(1) 박기
    void WritePapxLevel(int off, int itap, bool ttp)
    {
        // grpprl = istd(2) + sprmPItap(2+4) + sprmPFInTable(2+1) [+ sprmPFTtp(2+1) + sprmTDefTable(2+2+25)]
        int q = fcPapxFkp + off;
        int sprmsLen = 6 + 3;  // sprmPItap + sprmPFInTable
        if (ttp) sprmsLen += 3 + 29;  // sprmPFTtp + sprmTDefTable(header+cb+operand 25)
        int grpprlLen = 2 + sprmsLen;
        // cb = (grpprlLen + 1) / 2 — cb*2-1 = grpprlLen
        int cb = (grpprlLen + 1) / 2;
        // OK only if cb*2-1 == grpprlLen. grpprlLen 가 홀수 (11, 43) 일 때만 가능. 짝수면 cb=0 형식 사용.
        if (grpprlLen % 2 == 0)
        {
            wd[q++] = 0;
            wd[q++] = (byte)(grpprlLen / 2);
        }
        else
        {
            wd[q++] = (byte)cb;
        }
        BitConverter.TryWriteBytes(wd.AsSpan(q), (ushort)0); q += 2;  // istd
        // sprmPItap (0x6649, 4-byte operand)
        BitConverter.TryWriteBytes(wd.AsSpan(q), (ushort)0x6649); q += 2;
        BitConverter.TryWriteBytes(wd.AsSpan(q), itap); q += 4;
        // sprmPFInTable (1-byte)
        BitConverter.TryWriteBytes(wd.AsSpan(q), (ushort)0x2416); q += 2;
        wd[q++] = 1;
        if (ttp)
        {
            BitConverter.TryWriteBytes(wd.AsSpan(q), (ushort)0x2417); q += 2;
            wd[q++] = 1;
            // sprmTDefTable (1행 1셀): itcMac=1 + rgdxa[2](4) + rgTc[1](20) = 25 byte operand
            BitConverter.TryWriteBytes(wd.AsSpan(q), (ushort)0xD608); q += 2;
            BitConverter.TryWriteBytes(wd.AsSpan(q), (ushort)25);     q += 2;
            wd[q++] = 1;  // itcMac
            BitConverter.TryWriteBytes(wd.AsSpan(q), (short)0);    q += 2;
            BitConverter.TryWriteBytes(wd.AsSpan(q), (short)1440); q += 2;
            q += 20;  // TC[0] = all zeros
        }
    }

    // PapxInFkp[0..5]
    WritePapxLevel(papx0Off, itap: 1, ttp: false);
    WritePapxLevel(papx1Off, itap: 2, ttp: false);
    WritePapxLevel(papx2Off, itap: 2, ttp: true);
    WritePapxLevel(papx3Off, itap: 1, ttp: false);
    WritePapxLevel(papx4Off, itap: 1, ttp: true);
    // PapxInFkp[5]: 빈 (itap=0)
    int p5 = fcPapxFkp + papx5Off;
    wd[p5++] = 0;
    wd[p5++] = 1;
    BitConverter.TryWriteBytes(wd.AsSpan(p5), (ushort)0);
    wd[fcPapxFkp + 511] = (byte)cpara;

    // CHPX FKP — 비어 있음
    int crun = 1;
    BitConverter.TryWriteBytes(wd.AsSpan(fcChpxFkp + 0), (int)0x200);
    BitConverter.TryWriteBytes(wd.AsSpan(fcChpxFkp + 4), (int)0x22A);
    int chpx0Off = 64;
    wd[fcChpxFkp + 4 * (crun + 1) + 0] = (byte)(chpx0Off / 2);
    wd[fcChpxFkp + chpx0Off] = 0;
    wd[fcChpxFkp + 511] = (byte)crun;

    // Table stream
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
    var tblBytes = tblMs.ToArray();

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
            using (var s = root.CreateStream("0Table"))       s.Write(tblBytes);
        }
        using var fs = File.OpenRead(tmp);
        var doc = new PolyDonky.Convert.Doc.DocBinaryReader().Read(fs);
        var blocks = doc.Sections[0].Blocks;

        SmokeHarness.True(blocks.Count >= 2,
            $"section.Blocks.Count >= 2 (got {blocks.Count})");
        SmokeHarness.True(blocks[0] is PolyDonky.Core.Table,
            $"첫 블록은 외곽 Table (got {blocks[0].GetType().Name})");
        var outer = (PolyDonky.Core.Table)blocks[0];
        SmokeHarness.Equal(1, outer.Rows.Count, "외곽 행 수 1");
        SmokeHarness.Equal(1, outer.Rows[0].Cells.Count, "외곽 셀 수 1");

        // 외곽 셀 안: Para("Outer") + Inner Table
        var outerCellBlocks = outer.Rows[0].Cells[0].Blocks;
        SmokeHarness.True(outerCellBlocks.Count >= 2,
            $"외곽 셀.Blocks.Count >= 2 (got {outerCellBlocks.Count})");
        SmokeHarness.Equal("Outer",
            ((Paragraph)outerCellBlocks[0]).GetPlainText(),
            "외곽 셀의 첫 단락 = 'Outer'");
        SmokeHarness.True(outerCellBlocks[1] is PolyDonky.Core.Table,
            $"외곽 셀의 두 번째 블록은 Table (got {outerCellBlocks[1].GetType().Name})");
        var inner = (PolyDonky.Core.Table)outerCellBlocks[1];
        SmokeHarness.Equal(1, inner.Rows.Count, "내부 행 수 1");
        SmokeHarness.Equal(1, inner.Rows[0].Cells.Count, "내부 셀 수 1");
        SmokeHarness.Equal("Inner",
            ((Paragraph)inner.Rows[0].Cells[0].Blocks[0]).GetPlainText(),
            "내부 셀의 단락 = 'Inner'");

        SmokeHarness.True(blocks[1] is Paragraph,
            $"두 번째 블록은 Paragraph (got {blocks[1].GetType().Name})");
        SmokeHarness.Equal("End", ((Paragraph)blocks[1]).GetPlainText(),
            "본문 단락 'End'");
    }
    finally { try { File.Delete(tmp); } catch { } }
}

static void DocBinaryPhase3aFieldResult()
{
    // 필드: "Before\x13PAGE\x14Result\x15After\r" — 25 chars.
    //   0x13 (field begin) ~ 0x14 (separator): field code "PAGE" (폐기)
    //   0x14 ~ 0x15 (end): field result "Result" (포함)
    // 결과: 단락 텍스트 = "BeforeResultAfter"
    const string text = "BeforePAGEResultAfter\r";  // 0x13/0x14/0x15 박은 후 25 chars
    int ccp = text.Length;
    var textBytes = Encoding.Unicode.GetBytes(text);
    int fcText = 0x200;
    int pnPapx = 4, pnChpx = 5;
    int fcPapxFkp = pnPapx * 512;
    int fcChpxFkp = pnChpx * 512;
    int wdSize = fcChpxFkp + 512;
    var wd = new byte[wdSize];
    Buffer.BlockCopy(textBytes, 0, wd, fcText, textBytes.Length);

    // PAPX FKP — 1 단락, 빈 sprm
    int cpara = 1;
    BitConverter.TryWriteBytes(wd.AsSpan(fcPapxFkp + 0), (int)0x200);
    BitConverter.TryWriteBytes(wd.AsSpan(fcPapxFkp + 4), (int)(0x200 + ccp * 2));
    int papx0Off = 64;
    wd[fcPapxFkp + 8 + 0 * 13] = (byte)(papx0Off / 2);
    int p = fcPapxFkp + papx0Off;
    wd[p++] = 0;
    wd[p++] = 1;
    BitConverter.TryWriteBytes(wd.AsSpan(p), (ushort)0);  // istd
    wd[fcPapxFkp + 511] = (byte)cpara;

    // CHPX FKP — 비어 있음
    int crun = 1;
    BitConverter.TryWriteBytes(wd.AsSpan(fcChpxFkp + 0), (int)0x200);
    BitConverter.TryWriteBytes(wd.AsSpan(fcChpxFkp + 4), (int)(0x200 + ccp * 2));
    int chpx0Off = 64;
    wd[fcChpxFkp + 4 * (crun + 1) + 0] = (byte)(chpx0Off / 2);
    wd[fcChpxFkp + chpx0Off] = 0;
    wd[fcChpxFkp + 511] = (byte)crun;

    // Table stream
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
    var tblBytes = tblMs.ToArray();

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
            using (var s = root.CreateStream("0Table"))       s.Write(tblBytes);
        }
        using var fs = File.OpenRead(tmp);
        var doc = new PolyDonky.Convert.Doc.DocBinaryReader().Read(fs);
        var paragraphs = doc.EnumerateParagraphs().ToList();

        SmokeHarness.True(paragraphs.Count >= 1, $"단락 수 >= 1 (got {paragraphs.Count})");
        SmokeHarness.Equal("BeforeResultAfter", paragraphs[0].GetPlainText(),
            "필드 코드 폐기 + 결과 포함 + 주변 텍스트 보존");
    }
    finally { try { File.Delete(tmp); } catch { } }
}

static void DocBinaryPhase3bPageBreak()
{
    // 두 단락: 첫 단락 일반, 두 번째 단락의 PAPX 에 sprmPFPageBreakBefore=1.
    //   결과: paragraphs[1].Style.ForcePageBreakBefore == true.
    const string text = "First\rSecond\r";  // 13 chars
    int ccp = text.Length;
    var textBytes = Encoding.Unicode.GetBytes(text);
    int fcText = 0x200;
    int pnPapx = 4, pnChpx = 5;
    int fcPapxFkp = pnPapx * 512;
    int fcChpxFkp = pnChpx * 512;
    int wdSize = fcChpxFkp + 512;
    var wd = new byte[wdSize];
    Buffer.BlockCopy(textBytes, 0, wd, fcText, textBytes.Length);

    // PAPX FKP — 2 단락
    int cpara = 2;
    BitConverter.TryWriteBytes(wd.AsSpan(fcPapxFkp + 0),  (int)0x200);
    BitConverter.TryWriteBytes(wd.AsSpan(fcPapxFkp + 4),  (int)0x20C);  // 단락 1 끝+2 ("First\r" 다음)
    BitConverter.TryWriteBytes(wd.AsSpan(fcPapxFkp + 8),  (int)(0x200 + ccp * 2));  // 단락 2 끝+2
    int papx0Off = 100, papx1Off = 110;
    wd[fcPapxFkp + 12 + 0 * 13] = (byte)(papx0Off / 2);
    wd[fcPapxFkp + 12 + 1 * 13] = (byte)(papx1Off / 2);

    // PapxInFkp[0]: 빈 sprm
    int p = fcPapxFkp + papx0Off;
    wd[p++] = 0;
    wd[p++] = 1;
    BitConverter.TryWriteBytes(wd.AsSpan(p), (ushort)0);

    // PapxInFkp[1]: sprmPFPageBreakBefore=1
    //   istd(2) + sprm(2)+op(1) = 5 byte → cb=3
    p = fcPapxFkp + papx1Off;
    wd[p++] = 3;
    BitConverter.TryWriteBytes(wd.AsSpan(p), (ushort)0); p += 2;       // istd
    BitConverter.TryWriteBytes(wd.AsSpan(p), (ushort)0x2407); p += 2;  // sprmPFPageBreakBefore
    wd[p++] = 1;
    wd[fcPapxFkp + 511] = (byte)cpara;

    // CHPX FKP — 비어 있음
    int crun = 1;
    BitConverter.TryWriteBytes(wd.AsSpan(fcChpxFkp + 0), (int)0x200);
    BitConverter.TryWriteBytes(wd.AsSpan(fcChpxFkp + 4), (int)(0x200 + ccp * 2));
    int chpx0Off = 64;
    wd[fcChpxFkp + 4 * (crun + 1) + 0] = (byte)(chpx0Off / 2);
    wd[fcChpxFkp + chpx0Off] = 0;
    wd[fcChpxFkp + 511] = (byte)crun;

    // Table stream
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
    var tblBytes = tblMs.ToArray();

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
            using (var s = root.CreateStream("0Table"))       s.Write(tblBytes);
        }
        using var fs = File.OpenRead(tmp);
        var doc = new PolyDonky.Convert.Doc.DocBinaryReader().Read(fs);
        var paragraphs = doc.EnumerateParagraphs().ToList();

        SmokeHarness.True(paragraphs.Count >= 2, $"단락 수 >= 2 (got {paragraphs.Count})");
        SmokeHarness.True(!paragraphs[0].Style.ForcePageBreakBefore,
            "단락 0 의 ForcePageBreakBefore = false (빈 PAPX)");
        SmokeHarness.True(paragraphs[1].Style.ForcePageBreakBefore,
            "단락 1 의 ForcePageBreakBefore = true (sprmPFPageBreakBefore=1)");
    }
    finally { try { File.Delete(tmp); } catch { } }
}

static void DocBinaryPhase3cSections()
{
    // 6 chars 본문 ("AB\rCD\r") + PlcfSed cps = [0, 3, 6] (두 섹션).
    //   첫 섹션 = CP [0..3) → "AB\r"
    //   두 번째 섹션 = CP [3..6) → "CD\r"
    // 결과: doc.Sections.Count == 2, 각 섹션의 첫 단락이 "AB" / "CD".
    const string text = "AB\rCD\r";  // 6 chars
    int ccp = text.Length;
    var textBytes = Encoding.Unicode.GetBytes(text);
    int fcText = 0x200;
    int pnPapx = 4, pnChpx = 5;
    int fcPapxFkp = pnPapx * 512;
    int fcChpxFkp = pnChpx * 512;
    int wdSize = fcChpxFkp + 512;
    var wd = new byte[wdSize];
    Buffer.BlockCopy(textBytes, 0, wd, fcText, textBytes.Length);

    // PAPX FKP — 2 단락
    int cpara = 2;
    BitConverter.TryWriteBytes(wd.AsSpan(fcPapxFkp + 0), (int)0x200);
    BitConverter.TryWriteBytes(wd.AsSpan(fcPapxFkp + 4), (int)0x206);  // 단락 1 끝+2 ("AB\r" 다음)
    BitConverter.TryWriteBytes(wd.AsSpan(fcPapxFkp + 8), (int)(0x200 + ccp * 2));
    int papx0Off = 100, papx1Off = 110;
    wd[fcPapxFkp + 12 + 0 * 13] = (byte)(papx0Off / 2);
    wd[fcPapxFkp + 12 + 1 * 13] = (byte)(papx1Off / 2);
    int p = fcPapxFkp + papx0Off;
    wd[p++] = 0; wd[p++] = 1; BitConverter.TryWriteBytes(wd.AsSpan(p), (ushort)0);
    p = fcPapxFkp + papx1Off;
    wd[p++] = 0; wd[p++] = 1; BitConverter.TryWriteBytes(wd.AsSpan(p), (ushort)0);
    wd[fcPapxFkp + 511] = (byte)cpara;

    // CHPX FKP — 비어 있음
    int crun = 1;
    BitConverter.TryWriteBytes(wd.AsSpan(fcChpxFkp + 0), (int)0x200);
    BitConverter.TryWriteBytes(wd.AsSpan(fcChpxFkp + 4), (int)(0x200 + ccp * 2));
    int chpx0Off = 64;
    wd[fcChpxFkp + 4 * (crun + 1) + 0] = (byte)(chpx0Off / 2);
    wd[fcChpxFkp + chpx0Off] = 0;
    wd[fcChpxFkp + 511] = (byte)crun;

    // Table stream
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

    // PlcfSed: aCP[3] = [0, 3, 6] + aSed[2] (각 12 byte = 0)
    int plcfSedStart = (int)tblMs.Position;
    BitConverter.TryWriteBytes(b4, (int)0); tblMs.Write(b4);   // cps[0] = 0
    BitConverter.TryWriteBytes(b4, (int)3); tblMs.Write(b4);   // cps[1] = 3 (첫 섹션 끝)
    BitConverter.TryWriteBytes(b4, (int)6); tblMs.Write(b4);   // cps[2] = 6 (본문 끝)
    var emptySed = new byte[24];   // 2 SED × 12 byte = 24 byte (zeros)
    tblMs.Write(emptySed);
    int plcfSedLen = (int)tblMs.Position - plcfSedStart;

    var tblBytes = tblMs.ToArray();

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
    BitConverter.TryWriteBytes(wd.AsSpan(0x00CA), (uint)plcfSedStart);
    BitConverter.TryWriteBytes(wd.AsSpan(0x00CE), (uint)plcfSedLen);

    var tmp = Path.Combine(Path.GetTempPath(), $"polydonky-smoke-{Guid.NewGuid():N}.doc");
    try
    {
        using (var root = OpenMcdf.RootStorage.Create(tmp))
        {
            using (var s = root.CreateStream("WordDocument")) s.Write(wd);
            using (var s = root.CreateStream("0Table"))       s.Write(tblBytes);
        }
        using var fs = File.OpenRead(tmp);
        var doc = new PolyDonky.Convert.Doc.DocBinaryReader().Read(fs);

        SmokeHarness.Equal(2, doc.Sections.Count, "doc.Sections.Count = 2");
        SmokeHarness.True(doc.Sections[0].Blocks.Count >= 1, "섹션 0 단락 >= 1");
        SmokeHarness.True(doc.Sections[1].Blocks.Count >= 1, "섹션 1 단락 >= 1");
        SmokeHarness.Equal("AB",
            ((Paragraph)doc.Sections[0].Blocks[0]).GetPlainText(),
            "섹션 0 단락 0 텍스트 'AB'");
        SmokeHarness.Equal("CD",
            ((Paragraph)doc.Sections[1].Blocks[0]).GetPlainText(),
            "섹션 1 단락 0 텍스트 'CD'");
    }
    finally { try { File.Delete(tmp); } catch { } }
}

static void DocBinaryPhase3dHeaderFooter()
{
    // 본문 "Body\r" + 헤더/푸터 영역 (4 sub-stories): evenFtr/oddFtr/evenHdr/oddHdr.
    //   PlcfHdd 인덱스 매핑 (§2.8.7): 0=evenFtr, 1=oddFtr, 2=evenHdr, 3=oddHdr.
    //   합성: evenFtr="\r" 비어있음, oddFtr="Ftr\r", evenHdr="\r" 비어있음, oddHdr="Hdr\r".
    //   ccpText = 5, ccpFtn = 0, ccpHdd = 10.  총 piece = 15 chars.
    //   PlcfHdd aCP[5] = [0, 1, 5, 6, 10] — 영역 내 offset.
    const string text = "Body\r\rFtr\r\rHdr\r";  // 15 chars
    int ccpText = 5;
    int ccpHdd  = 10;
    int totalCcp = text.Length;  // 15
    var textBytes = Encoding.Unicode.GetBytes(text);
    int fcText = 0x200;
    int pnPapx = 4, pnChpx = 5;
    int fcPapxFkp = pnPapx * 512;
    int fcChpxFkp = pnChpx * 512;
    int wdSize = fcChpxFkp + 512;
    var wd = new byte[wdSize];
    Buffer.BlockCopy(textBytes, 0, wd, fcText, textBytes.Length);

    // PAPX FKP — main 단락 1 개만 인식 (헤더/푸터 단락은 ApplyHeaderFooter 가 별도 처리)
    int cpara = 1;
    BitConverter.TryWriteBytes(wd.AsSpan(fcPapxFkp + 0), (int)0x200);
    BitConverter.TryWriteBytes(wd.AsSpan(fcPapxFkp + 4), (int)(0x200 + totalCcp * 2));
    int papx0Off = 64;
    wd[fcPapxFkp + 8 + 0 * 13] = (byte)(papx0Off / 2);
    int p = fcPapxFkp + papx0Off;
    wd[p++] = 0; wd[p++] = 1; BitConverter.TryWriteBytes(wd.AsSpan(p), (ushort)0);
    wd[fcPapxFkp + 511] = (byte)cpara;

    // CHPX FKP — 비어 있음
    int crun = 1;
    BitConverter.TryWriteBytes(wd.AsSpan(fcChpxFkp + 0), (int)0x200);
    BitConverter.TryWriteBytes(wd.AsSpan(fcChpxFkp + 4), (int)(0x200 + totalCcp * 2));
    int chpx0Off = 64;
    wd[fcChpxFkp + 4 * (crun + 1) + 0] = (byte)(chpx0Off / 2);
    wd[fcChpxFkp + chpx0Off] = 0;
    wd[fcChpxFkp + 511] = (byte)crun;

    // Table stream: CLX (piece covers all chars 0..totalCcp) + PAPX/CHPX BTE + PlcfHdd
    var tblMs = new MemoryStream();
    Span<byte> b4 = stackalloc byte[4];
    tblMs.WriteByte(0x02);
    BitConverter.TryWriteBytes(b4, (uint)16); tblMs.Write(b4);
    BitConverter.TryWriteBytes(b4, (uint)0);   tblMs.Write(b4);
    BitConverter.TryWriteBytes(b4, (uint)totalCcp); tblMs.Write(b4);
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

    // PlcfHdd aCP[5] = [0, 1, 5, 6, 10] — 영역 내 offset (4 sub-stories)
    int plcfHddStart = (int)tblMs.Position;
    BitConverter.TryWriteBytes(b4, (int)0);  tblMs.Write(b4);
    BitConverter.TryWriteBytes(b4, (int)1);  tblMs.Write(b4);
    BitConverter.TryWriteBytes(b4, (int)5);  tblMs.Write(b4);
    BitConverter.TryWriteBytes(b4, (int)6);  tblMs.Write(b4);
    BitConverter.TryWriteBytes(b4, (int)10); tblMs.Write(b4);
    int plcfHddLen = (int)tblMs.Position - plcfHddStart;

    var tblBytes = tblMs.ToArray();

    BitConverter.TryWriteBytes(wd.AsSpan(0x00),   (ushort)0xA5EC);
    BitConverter.TryWriteBytes(wd.AsSpan(0x02),   (ushort)193);
    BitConverter.TryWriteBytes(wd.AsSpan(0x0A),   (ushort)0x0000);
    BitConverter.TryWriteBytes(wd.AsSpan(0x18),   (uint)fcText);
    BitConverter.TryWriteBytes(wd.AsSpan(0x4C),   (uint)ccpText);    // main text 길이 = 5
    BitConverter.TryWriteBytes(wd.AsSpan(0x0050), (uint)0);          // ccpFtn = 0
    BitConverter.TryWriteBytes(wd.AsSpan(0x0054), (uint)ccpHdd);     // ccpHdd = 8
    BitConverter.TryWriteBytes(wd.AsSpan(0x01A2), (uint)0);
    BitConverter.TryWriteBytes(wd.AsSpan(0x01A6), (uint)clxEnd);
    BitConverter.TryWriteBytes(wd.AsSpan(0x00FA), (uint)chpxBteStart);
    BitConverter.TryWriteBytes(wd.AsSpan(0x00FE), (uint)chpxBteLen);
    BitConverter.TryWriteBytes(wd.AsSpan(0x0102), (uint)papxBteStart);
    BitConverter.TryWriteBytes(wd.AsSpan(0x0106), (uint)papxBteLen);
    BitConverter.TryWriteBytes(wd.AsSpan(0x00F2), (uint)plcfHddStart);
    BitConverter.TryWriteBytes(wd.AsSpan(0x00F6), (uint)plcfHddLen);

    var tmp = Path.Combine(Path.GetTempPath(), $"polydonky-smoke-{Guid.NewGuid():N}.doc");
    try
    {
        using (var root = OpenMcdf.RootStorage.Create(tmp))
        {
            using (var s = root.CreateStream("WordDocument")) s.Write(wd);
            using (var s = root.CreateStream("0Table"))       s.Write(tblBytes);
        }
        using var fs = File.OpenRead(tmp);
        var doc = new PolyDonky.Convert.Doc.DocBinaryReader().Read(fs);

        // 본문 단락: "Body"
        SmokeHarness.Equal("Body",
            doc.EnumerateParagraphs().First().GetPlainText(),
            "본문 단락 = 'Body'");

        var header = doc.Sections[0].Page.Header.Center;
        SmokeHarness.True(header.Paragraphs.Count >= 1, "Header.Center 단락 >= 1");
        SmokeHarness.Equal("Hdr", header.Paragraphs[0].GetPlainText(),
            "Header.Center.Paragraphs[0] = 'Hdr'");

        var footer = doc.Sections[0].Page.Footer.Center;
        SmokeHarness.True(footer.Paragraphs.Count >= 1, "Footer.Center 단락 >= 1");
        SmokeHarness.Equal("Ftr", footer.Paragraphs[0].GetPlainText(),
            "Footer.Center.Paragraphs[0] = 'Ftr'");
    }
    finally { try { File.Delete(tmp); } catch { } }
}

static void DocBinaryPhase3c2SepxPageProps()
{
    // 1 section + SEPX 박음. SEPX 에 페이지 크기·여백 sprm:
    //   sprmSDxaPage  (0xB016) = 12240 twips (8.5 inch = 215.9 mm)
    //   sprmSDyaPage  (0xB017) = 15840 twips (11 inch = 279.4 mm)
    //   sprmSDxaLeft  (0xB02C) = 1440 twips (1 inch = 25.4 mm)
    //   sprmSDxaRight (0xB02D) = 1440 twips
    //   sprmSDyaTop   (0xB02E) = 720 twips (0.5 inch = 12.7 mm)
    //   sprmSDyaBottom(0xB02F) = 720 twips
    const string text = "Body\r";  // 5 chars
    int ccp = text.Length;
    var textBytes = Encoding.Unicode.GetBytes(text);
    int fcText = 0x200;
    int pnPapx = 4, pnChpx = 5;
    int fcPapxFkp = pnPapx * 512;
    int fcChpxFkp = pnChpx * 512;
    int wdSize = fcChpxFkp + 512;
    var wd = new byte[wdSize];
    Buffer.BlockCopy(textBytes, 0, wd, fcText, textBytes.Length);

    // PAPX FKP — 1 단락 빈 sprm
    int cpara = 1;
    BitConverter.TryWriteBytes(wd.AsSpan(fcPapxFkp + 0), (int)0x200);
    BitConverter.TryWriteBytes(wd.AsSpan(fcPapxFkp + 4), (int)(0x200 + ccp * 2));
    int papx0Off = 64;
    wd[fcPapxFkp + 8 + 0 * 13] = (byte)(papx0Off / 2);
    int p = fcPapxFkp + papx0Off;
    wd[p++] = 0; wd[p++] = 1; BitConverter.TryWriteBytes(wd.AsSpan(p), (ushort)0);
    wd[fcPapxFkp + 511] = (byte)cpara;

    // CHPX FKP — 비어 있음
    int crun = 1;
    BitConverter.TryWriteBytes(wd.AsSpan(fcChpxFkp + 0), (int)0x200);
    BitConverter.TryWriteBytes(wd.AsSpan(fcChpxFkp + 4), (int)(0x200 + ccp * 2));
    int chpx0Off = 64;
    wd[fcChpxFkp + 4 * (crun + 1) + 0] = (byte)(chpx0Off / 2);
    wd[fcChpxFkp + chpx0Off] = 0;
    wd[fcChpxFkp + 511] = (byte)crun;

    // Table stream
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

    // SEPX: cb(2) + grpprl 6 sprm × 4 byte each = 24 byte. cb=24.
    int sepxStart = (int)tblMs.Position;
    BitConverter.TryWriteBytes(b4.Slice(0, 2), (ushort)24); tblMs.Write(b4.Slice(0, 2));
    void Sprm2(ushort sprm, ushort op)
    {
        var buf = new byte[4];
        BitConverter.TryWriteBytes(buf.AsSpan(0, 2), sprm);
        BitConverter.TryWriteBytes(buf.AsSpan(2, 2), op);
        tblMs.Write(buf);
    }
    Sprm2(0xB016, 12240);  // page width
    Sprm2(0xB017, 15840);  // page height
    Sprm2(0xB02C, 1440);   // margin left
    Sprm2(0xB02D, 1440);   // margin right
    Sprm2(0xB02E, 720);    // margin top
    Sprm2(0xB02F, 720);    // margin bottom

    // PlcfSed: aCP[2] = [0, ccp] (1 section) + aSed[1] = SED with fcSepx
    int plcfSedStart = (int)tblMs.Position;
    BitConverter.TryWriteBytes(b4, (int)0);   tblMs.Write(b4);   // cps[0]
    BitConverter.TryWriteBytes(b4, (int)ccp); tblMs.Write(b4);   // cps[1] = ccpText
    // SED: fn(2) + fcSepx(4) + fnMpr(2) + fcMpr(4)
    BitConverter.TryWriteBytes(b4.Slice(0, 2), (ushort)0); tblMs.Write(b4.Slice(0, 2));
    BitConverter.TryWriteBytes(b4, (int)sepxStart); tblMs.Write(b4);
    BitConverter.TryWriteBytes(b4.Slice(0, 2), (ushort)0); tblMs.Write(b4.Slice(0, 2));
    BitConverter.TryWriteBytes(b4, (int)0); tblMs.Write(b4);
    int plcfSedLen = (int)tblMs.Position - plcfSedStart;

    var tblBytes = tblMs.ToArray();

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
    BitConverter.TryWriteBytes(wd.AsSpan(0x00CA), (uint)plcfSedStart);
    BitConverter.TryWriteBytes(wd.AsSpan(0x00CE), (uint)plcfSedLen);

    var tmp = Path.Combine(Path.GetTempPath(), $"polydonky-smoke-{Guid.NewGuid():N}.doc");
    try
    {
        using (var root = OpenMcdf.RootStorage.Create(tmp))
        {
            using (var s = root.CreateStream("WordDocument")) s.Write(wd);
            using (var s = root.CreateStream("0Table"))       s.Write(tblBytes);
        }
        using var fs = File.OpenRead(tmp);
        var doc = new PolyDonky.Convert.Doc.DocBinaryReader().Read(fs);
        var page = doc.Sections[0].Page;

        SmokeHarness.True(Math.Abs(page.WidthMm  - 12240 / 56.692) < 0.05,
            $"WidthMm ≈ 215.9 (got {page.WidthMm:F2})");
        SmokeHarness.True(Math.Abs(page.HeightMm - 15840 / 56.692) < 0.05,
            $"HeightMm ≈ 279.4 (got {page.HeightMm:F2})");
        SmokeHarness.True(Math.Abs(page.MarginLeftMm   - 1440 / 56.692) < 0.05,
            $"MarginLeftMm ≈ 25.4 (got {page.MarginLeftMm:F2})");
        SmokeHarness.True(Math.Abs(page.MarginRightMm  - 1440 / 56.692) < 0.05,
            $"MarginRightMm ≈ 25.4 (got {page.MarginRightMm:F2})");
        SmokeHarness.True(Math.Abs(page.MarginTopMm    - 720 / 56.692) < 0.05,
            $"MarginTopMm ≈ 12.7 (got {page.MarginTopMm:F2})");
        SmokeHarness.True(Math.Abs(page.MarginBottomMm - 720 / 56.692) < 0.05,
            $"MarginBottomMm ≈ 12.7 (got {page.MarginBottomMm:F2})");
    }
    finally { try { File.Delete(tmp); } catch { } }
}

static void DocBinaryPhase3eImagePlaceholder()
{
    // 본문 "Before\x01After\r" — \x01 picture marker 가 단락을 끊고 ImageBlock placeholder 삽입.
    //   결과: section.Blocks = [Para("Before"), ImageBlock, Para("After")]
    const string text = "BeforeAfter\r";  // 0x01 박은 후 12 chars
    int ccp = text.Length;
    var textBytes = Encoding.Unicode.GetBytes(text);
    int fcText = 0x200;
    int pnPapx = 4, pnChpx = 5;
    int fcPapxFkp = pnPapx * 512;
    int fcChpxFkp = pnChpx * 512;
    int wdSize = fcChpxFkp + 512;
    var wd = new byte[wdSize];
    Buffer.BlockCopy(textBytes, 0, wd, fcText, textBytes.Length);

    // PAPX FKP — 1 단락 빈 sprm
    int cpara = 1;
    BitConverter.TryWriteBytes(wd.AsSpan(fcPapxFkp + 0), (int)0x200);
    BitConverter.TryWriteBytes(wd.AsSpan(fcPapxFkp + 4), (int)(0x200 + ccp * 2));
    int papx0Off = 64;
    wd[fcPapxFkp + 8 + 0 * 13] = (byte)(papx0Off / 2);
    int p = fcPapxFkp + papx0Off;
    wd[p++] = 0; wd[p++] = 1; BitConverter.TryWriteBytes(wd.AsSpan(p), (ushort)0);
    wd[fcPapxFkp + 511] = (byte)cpara;

    // CHPX FKP — 비어 있음
    int crun = 1;
    BitConverter.TryWriteBytes(wd.AsSpan(fcChpxFkp + 0), (int)0x200);
    BitConverter.TryWriteBytes(wd.AsSpan(fcChpxFkp + 4), (int)(0x200 + ccp * 2));
    int chpx0Off = 64;
    wd[fcChpxFkp + 4 * (crun + 1) + 0] = (byte)(chpx0Off / 2);
    wd[fcChpxFkp + chpx0Off] = 0;
    wd[fcChpxFkp + 511] = (byte)crun;

    // Table stream
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
    var tblBytes = tblMs.ToArray();

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
            using (var s = root.CreateStream("0Table"))       s.Write(tblBytes);
        }
        using var fs = File.OpenRead(tmp);
        var doc = new PolyDonky.Convert.Doc.DocBinaryReader().Read(fs);
        var blocks = doc.Sections[0].Blocks;

        SmokeHarness.True(blocks.Count >= 3,
            $"section.Blocks.Count >= 3 (got {blocks.Count})");
        SmokeHarness.True(blocks[0] is Paragraph,
            $"blocks[0] = Paragraph (got {blocks[0].GetType().Name})");
        SmokeHarness.Equal("Before", ((Paragraph)blocks[0]).GetPlainText(),
            "blocks[0] = 'Before'");
        SmokeHarness.True(blocks[1] is ImageBlock,
            $"blocks[1] = ImageBlock (got {blocks[1].GetType().Name})");
        SmokeHarness.True(blocks[2] is Paragraph,
            $"blocks[2] = Paragraph (got {blocks[2].GetType().Name})");
        SmokeHarness.Equal("After", ((Paragraph)blocks[2]).GetPlainText(),
            "blocks[2] = 'After'");
    }
    finally { try { File.Delete(tmp); } catch { } }
}

static void DocBinaryPhase3e2ImageData()
{
    // 본문 "Before\x01After\r" + Data stream 의 PICF + PNG signature.
    //   CHPX FKP 의 \x01 char run 에 sprmCPicLocation = fcPic (Data stream 내 PICF 위치).
    //   PICF: lcb=20, body 16 byte = PNG signature(8) + dummy(8).
    //   결과: ImageBlock.MediaType = "image/png", Data 의 처음 8 byte = PNG signature.
    const string text = "BeforeAfter\r";  // 0x01 박은 후 12 chars
    int ccp = text.Length;
    var textBytes = Encoding.Unicode.GetBytes(text);
    int fcText = 0x200;
    int pnPapx = 4, pnChpx = 5;
    int fcPapxFkp = pnPapx * 512;
    int fcChpxFkp = pnChpx * 512;
    int wdSize = fcChpxFkp + 512;
    var wd = new byte[wdSize];
    Buffer.BlockCopy(textBytes, 0, wd, fcText, textBytes.Length);

    // PAPX FKP — 1 단락
    int cpara = 1;
    BitConverter.TryWriteBytes(wd.AsSpan(fcPapxFkp + 0), (int)0x200);
    BitConverter.TryWriteBytes(wd.AsSpan(fcPapxFkp + 4), (int)(0x200 + ccp * 2));
    int papx0Off = 64;
    wd[fcPapxFkp + 8 + 0 * 13] = (byte)(papx0Off / 2);
    int p = fcPapxFkp + papx0Off;
    wd[p++] = 0; wd[p++] = 1; BitConverter.TryWriteBytes(wd.AsSpan(p), (ushort)0);
    wd[fcPapxFkp + 511] = (byte)cpara;

    // CHPX FKP — 3 runs:
    //   run 0: [0x200, 0x20C) "Before" — 빈 CHPX
    //   run 1: [0x20C, 0x20E) "\x01"   — sprmCPicLocation = fcPic = 0
    //   run 2: [0x20E, 0x218) "After\r" — 빈 CHPX
    int crun = 3;
    BitConverter.TryWriteBytes(wd.AsSpan(fcChpxFkp + 0),  (int)0x200);
    BitConverter.TryWriteBytes(wd.AsSpan(fcChpxFkp + 4),  (int)0x20C);
    BitConverter.TryWriteBytes(wd.AsSpan(fcChpxFkp + 8),  (int)0x20E);
    BitConverter.TryWriteBytes(wd.AsSpan(fcChpxFkp + 12), (int)0x218);
    int chpx0Off = 64, chpx1Off = 80, chpx2Off = 96;
    int rgbBase  = 4 * (crun + 1);
    wd[fcChpxFkp + rgbBase + 0] = (byte)(chpx0Off / 2);
    wd[fcChpxFkp + rgbBase + 1] = (byte)(chpx1Off / 2);
    wd[fcChpxFkp + rgbBase + 2] = (byte)(chpx2Off / 2);
    // ChpxInFkp[0]: cb=0
    wd[fcChpxFkp + chpx0Off + 0] = 0;
    // ChpxInFkp[1]: cb=6, sprmCPicLocation(2)=0x6A03 + op(4)=0
    int c = fcChpxFkp + chpx1Off;
    wd[c++] = 6;
    BitConverter.TryWriteBytes(wd.AsSpan(c), (ushort)0x6A03); c += 2;
    BitConverter.TryWriteBytes(wd.AsSpan(c), (int)0); c += 4;  // fcPic = 0 (Data stream 내)
    // ChpxInFkp[2]: cb=0
    wd[fcChpxFkp + chpx2Off + 0] = 0;
    wd[fcChpxFkp + 511] = (byte)crun;

    // Data stream: PICF (lcb=20 + 16 byte body), 처음 8 byte = PNG signature
    var dataStream = new byte[20];
    BitConverter.TryWriteBytes(dataStream.AsSpan(0), (int)20);  // lcb
    // PICF body 시작 byte 4. PNG signature 박음.
    byte[] pngSig = { 0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A };
    Buffer.BlockCopy(pngSig, 0, dataStream, 4, 8);
    // 나머지 dataStream[12..19] = 0 (dummy)

    // Table stream
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
    var tblBytes = tblMs.ToArray();

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
            using (var s = root.CreateStream("0Table"))       s.Write(tblBytes);
            using (var s = root.CreateStream("Data"))         s.Write(dataStream);
        }
        using var fs = File.OpenRead(tmp);
        var doc = new PolyDonky.Convert.Doc.DocBinaryReader().Read(fs);
        var blocks = doc.Sections[0].Blocks;

        SmokeHarness.True(blocks.Count >= 3, $"blocks.Count >= 3 (got {blocks.Count})");
        SmokeHarness.True(blocks[1] is ImageBlock, "blocks[1] = ImageBlock");
        var img = (ImageBlock)blocks[1];
        SmokeHarness.Equal("image/png", img.MediaType, "MediaType = image/png");
        SmokeHarness.True(img.Data.Length >= 8, $"Data.Length >= 8 (got {img.Data.Length})");
        SmokeHarness.True(
            img.Data[0] == 0x89 && img.Data[1] == 0x50 && img.Data[2] == 0x4E && img.Data[3] == 0x47,
            $"Data 의 처음 4 byte = PNG signature (got {img.Data[0]:X2}{img.Data[1]:X2}{img.Data[2]:X2}{img.Data[3]:X2})");
    }
    finally { try { File.Delete(tmp); } catch { } }
}

static void DocBinaryPhase3e3WmfImage()
{
    // Phase 3e-2 와 동일 구조 — Data stream 의 PICF 안에 PNG 대신 WMF placeable header(D7 CD C6 9A) 박는다.
    //   결과: ImageBlock.MediaType = "image/wmf", Data 의 처음 4 byte = D7 CD C6 9A.
    const string text = "BeforeAfter\r";
    int ccp = text.Length;
    var textBytes = Encoding.Unicode.GetBytes(text);
    int fcText = 0x200;
    int pnPapx = 4, pnChpx = 5;
    int fcPapxFkp = pnPapx * 512;
    int fcChpxFkp = pnChpx * 512;
    int wdSize = fcChpxFkp + 512;
    var wd = new byte[wdSize];
    Buffer.BlockCopy(textBytes, 0, wd, fcText, textBytes.Length);

    // PAPX FKP — 1 단락
    int cpara = 1;
    BitConverter.TryWriteBytes(wd.AsSpan(fcPapxFkp + 0), (int)0x200);
    BitConverter.TryWriteBytes(wd.AsSpan(fcPapxFkp + 4), (int)(0x200 + ccp * 2));
    int papx0Off = 64;
    wd[fcPapxFkp + 8 + 0 * 13] = (byte)(papx0Off / 2);
    int p = fcPapxFkp + papx0Off;
    wd[p++] = 0; wd[p++] = 1; BitConverter.TryWriteBytes(wd.AsSpan(p), (ushort)0);
    wd[fcPapxFkp + 511] = (byte)cpara;

    // CHPX FKP — 3 runs (Phase 3e-2 와 동일).
    int crun = 3;
    BitConverter.TryWriteBytes(wd.AsSpan(fcChpxFkp + 0),  (int)0x200);
    BitConverter.TryWriteBytes(wd.AsSpan(fcChpxFkp + 4),  (int)0x20C);
    BitConverter.TryWriteBytes(wd.AsSpan(fcChpxFkp + 8),  (int)0x20E);
    BitConverter.TryWriteBytes(wd.AsSpan(fcChpxFkp + 12), (int)0x218);
    int chpx0Off = 64, chpx1Off = 80, chpx2Off = 96;
    int rgbBase  = 4 * (crun + 1);
    wd[fcChpxFkp + rgbBase + 0] = (byte)(chpx0Off / 2);
    wd[fcChpxFkp + rgbBase + 1] = (byte)(chpx1Off / 2);
    wd[fcChpxFkp + rgbBase + 2] = (byte)(chpx2Off / 2);
    wd[fcChpxFkp + chpx0Off + 0] = 0;
    int c = fcChpxFkp + chpx1Off;
    wd[c++] = 6;
    BitConverter.TryWriteBytes(wd.AsSpan(c), (ushort)0x6A03); c += 2;
    BitConverter.TryWriteBytes(wd.AsSpan(c), (int)0); c += 4;
    wd[fcChpxFkp + chpx2Off + 0] = 0;
    wd[fcChpxFkp + 511] = (byte)crun;

    // Data stream: PICF (lcb=32 + 32 byte body). offset 4 부터 WMF placeable header.
    var dataStream = new byte[36];
    BitConverter.TryWriteBytes(dataStream.AsSpan(0), (int)32);
    byte[] wmfSig = { 0xD7, 0xCD, 0xC6, 0x9A };
    Buffer.BlockCopy(wmfSig, 0, dataStream, 4, 4);
    // 나머지 8..31 = 0 (dummy placeable header bytes)

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
    var tblBytes = tblMs.ToArray();

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
            using (var s = root.CreateStream("0Table"))       s.Write(tblBytes);
            using (var s = root.CreateStream("Data"))         s.Write(dataStream);
        }
        using var fs = File.OpenRead(tmp);
        var doc = new PolyDonky.Convert.Doc.DocBinaryReader().Read(fs);
        var blocks = doc.Sections[0].Blocks;

        SmokeHarness.True(blocks.Count >= 3, $"blocks.Count >= 3 (got {blocks.Count})");
        SmokeHarness.True(blocks[1] is ImageBlock, "blocks[1] = ImageBlock");
        var img = (ImageBlock)blocks[1];
        SmokeHarness.Equal("image/wmf", img.MediaType, "MediaType = image/wmf");
        SmokeHarness.True(img.Data.Length >= 4, $"Data.Length >= 4 (got {img.Data.Length})");
        SmokeHarness.True(
            img.Data[0] == 0xD7 && img.Data[1] == 0xCD && img.Data[2] == 0xC6 && img.Data[3] == 0x9A,
            $"Data 의 처음 4 byte = WMF placeable signature (got {img.Data[0]:X2}{img.Data[1]:X2}{img.Data[2]:X2}{img.Data[3]:X2})");
    }
    finally { try { File.Delete(tmp); } catch { } }
}

static void DocBinaryPhase3e3DibImage()
{
    // Phase 3e-2 와 동일 구조 — Data stream 의 PICF 안에 DIB BITMAPINFOHEADER 박는다.
    //   biSize=40, biPlanes=1, biBitCount=24.
    //   결과: ImageBlock.MediaType = "image/x-dib", Data[0] = 0x28.
    const string text = "BeforeAfter\r";
    int ccp = text.Length;
    var textBytes = Encoding.Unicode.GetBytes(text);
    int fcText = 0x200;
    int pnPapx = 4, pnChpx = 5;
    int fcPapxFkp = pnPapx * 512;
    int fcChpxFkp = pnChpx * 512;
    int wdSize = fcChpxFkp + 512;
    var wd = new byte[wdSize];
    Buffer.BlockCopy(textBytes, 0, wd, fcText, textBytes.Length);

    int cpara = 1;
    BitConverter.TryWriteBytes(wd.AsSpan(fcPapxFkp + 0), (int)0x200);
    BitConverter.TryWriteBytes(wd.AsSpan(fcPapxFkp + 4), (int)(0x200 + ccp * 2));
    int papx0Off = 64;
    wd[fcPapxFkp + 8 + 0 * 13] = (byte)(papx0Off / 2);
    int p = fcPapxFkp + papx0Off;
    wd[p++] = 0; wd[p++] = 1; BitConverter.TryWriteBytes(wd.AsSpan(p), (ushort)0);
    wd[fcPapxFkp + 511] = (byte)cpara;

    int crun = 3;
    BitConverter.TryWriteBytes(wd.AsSpan(fcChpxFkp + 0),  (int)0x200);
    BitConverter.TryWriteBytes(wd.AsSpan(fcChpxFkp + 4),  (int)0x20C);
    BitConverter.TryWriteBytes(wd.AsSpan(fcChpxFkp + 8),  (int)0x20E);
    BitConverter.TryWriteBytes(wd.AsSpan(fcChpxFkp + 12), (int)0x218);
    int chpx0Off = 64, chpx1Off = 80, chpx2Off = 96;
    int rgbBase  = 4 * (crun + 1);
    wd[fcChpxFkp + rgbBase + 0] = (byte)(chpx0Off / 2);
    wd[fcChpxFkp + rgbBase + 1] = (byte)(chpx1Off / 2);
    wd[fcChpxFkp + rgbBase + 2] = (byte)(chpx2Off / 2);
    wd[fcChpxFkp + chpx0Off + 0] = 0;
    int c = fcChpxFkp + chpx1Off;
    wd[c++] = 6;
    BitConverter.TryWriteBytes(wd.AsSpan(c), (ushort)0x6A03); c += 2;
    BitConverter.TryWriteBytes(wd.AsSpan(c), (int)0); c += 4;
    wd[fcChpxFkp + chpx2Off + 0] = 0;
    wd[fcChpxFkp + 511] = (byte)crun;

    // Data stream: lcb=48 + 48 byte body, offset 4 부터 BITMAPINFOHEADER 박는다.
    //   bytes 4..7 = 0x28 0 0 0 (biSize=40)
    //   bytes 8..11 = biWidth = 10
    //   bytes 12..15 = biHeight = 10
    //   bytes 16..17 = biPlanes = 1
    //   bytes 18..19 = biBitCount = 24
    var dataStream = new byte[52];
    BitConverter.TryWriteBytes(dataStream.AsSpan(0), (int)48);
    dataStream[4] = 0x28;
    BitConverter.TryWriteBytes(dataStream.AsSpan(8),  (int)10);
    BitConverter.TryWriteBytes(dataStream.AsSpan(12), (int)10);
    BitConverter.TryWriteBytes(dataStream.AsSpan(16), (ushort)1);
    BitConverter.TryWriteBytes(dataStream.AsSpan(18), (ushort)24);

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
    var tblBytes = tblMs.ToArray();

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
            using (var s = root.CreateStream("0Table"))       s.Write(tblBytes);
            using (var s = root.CreateStream("Data"))         s.Write(dataStream);
        }
        using var fs = File.OpenRead(tmp);
        var doc = new PolyDonky.Convert.Doc.DocBinaryReader().Read(fs);
        var blocks = doc.Sections[0].Blocks;

        SmokeHarness.True(blocks.Count >= 3, $"blocks.Count >= 3 (got {blocks.Count})");
        SmokeHarness.True(blocks[1] is ImageBlock, "blocks[1] = ImageBlock");
        var img = (ImageBlock)blocks[1];
        SmokeHarness.Equal("image/x-dib", img.MediaType, "MediaType = image/x-dib");
        SmokeHarness.True(img.Data.Length >= 16, $"Data.Length >= 16 (got {img.Data.Length})");
        SmokeHarness.True(
            img.Data[0] == 0x28 && img.Data[1] == 0x00 && img.Data[2] == 0x00 && img.Data[3] == 0x00,
            $"Data 의 처음 4 byte = BITMAPINFOHEADER biSize=40 (got {img.Data[0]:X2}{img.Data[1]:X2}{img.Data[2]:X2}{img.Data[3]:X2})");
    }
    finally { try { File.Delete(tmp); } catch { } }
}

static void DocBinaryPhase3e4ImageInTableCell()
{
    // 1행 1셀 표 + 셀 안 inline 이미지 (Phase 2a 표 + Phase 3e-2 PNG 추출 합성).
    //   본문 텍스트 (이상): "X\x01Y\x07\rEnd\r" — 9 chars
    //   파라 1 (InTable=1, !TTP): "X\x01Y\x07" + \r → 셀 1개 ["X", Img, "Y"]
    //   파라 2: "End"
    const string text = "XY\rEnd\r";  // 0x01·0x07·\r·E·n·d·\r — 9 chars
    int ccp = text.Length;
    var textBytes = Encoding.Unicode.GetBytes(text);
    int fcText = 0x200;
    int pnPapx = 4, pnChpx = 5;
    int fcPapxFkp = pnPapx * 512;
    int fcChpxFkp = pnChpx * 512;
    int wdSize = fcChpxFkp + 512;
    var wd = new byte[wdSize];
    Buffer.BlockCopy(textBytes, 0, wd, fcText, textBytes.Length);

    // PAPX FKP — 2 단락.
    //   Para 1: fc [0x200, 0x20A), istd=0, sprmPFInTable=1
    //   Para 2: fc [0x20A, 0x212), istd=0
    int cpara = 2;
    BitConverter.TryWriteBytes(wd.AsSpan(fcPapxFkp + 0), (int)0x200);
    BitConverter.TryWriteBytes(wd.AsSpan(fcPapxFkp + 4), (int)0x20A);
    BitConverter.TryWriteBytes(wd.AsSpan(fcPapxFkp + 8), (int)0x212);
    int papx0Off = 200, papx1Off = 220;
    wd[fcPapxFkp + 12 + 0 * 13] = (byte)(papx0Off / 2);
    wd[fcPapxFkp + 12 + 1 * 13] = (byte)(papx1Off / 2);
    int p = fcPapxFkp + papx0Off;
    wd[p++] = 3;
    BitConverter.TryWriteBytes(wd.AsSpan(p), (ushort)0); p += 2;
    BitConverter.TryWriteBytes(wd.AsSpan(p), (ushort)0x2416); p += 2;  // sprmPFInTable
    wd[p++] = 1;
    p = fcPapxFkp + papx1Off;
    wd[p++] = 0; wd[p++] = 1; BitConverter.TryWriteBytes(wd.AsSpan(p), (ushort)0);
    wd[fcPapxFkp + 511] = (byte)cpara;

    // CHPX FKP — 3 runs:
    //   run 0: fc [0x200, 0x202) "X" — 빈 CHPX
    //   run 1: fc [0x202, 0x204) "\x01" — sprmCPicLocation = fcPic = 0
    //   run 2: fc [0x204, 0x212) "Y\x07\rEnd\r" — 빈 CHPX
    int crun = 3;
    BitConverter.TryWriteBytes(wd.AsSpan(fcChpxFkp + 0),  (int)0x200);
    BitConverter.TryWriteBytes(wd.AsSpan(fcChpxFkp + 4),  (int)0x202);
    BitConverter.TryWriteBytes(wd.AsSpan(fcChpxFkp + 8),  (int)0x204);
    BitConverter.TryWriteBytes(wd.AsSpan(fcChpxFkp + 12), (int)0x212);
    int chpx0Off = 64, chpx1Off = 80, chpx2Off = 96;
    int rgbBase  = 4 * (crun + 1);
    wd[fcChpxFkp + rgbBase + 0] = (byte)(chpx0Off / 2);
    wd[fcChpxFkp + rgbBase + 1] = (byte)(chpx1Off / 2);
    wd[fcChpxFkp + rgbBase + 2] = (byte)(chpx2Off / 2);
    wd[fcChpxFkp + chpx0Off + 0] = 0;
    int c = fcChpxFkp + chpx1Off;
    wd[c++] = 6;
    BitConverter.TryWriteBytes(wd.AsSpan(c), (ushort)0x6A03); c += 2;
    BitConverter.TryWriteBytes(wd.AsSpan(c), (int)0); c += 4;
    wd[fcChpxFkp + chpx2Off + 0] = 0;
    wd[fcChpxFkp + 511] = (byte)crun;

    // Data stream: PICF lcb=20, PNG signature at offset 4.
    var dataStream = new byte[20];
    BitConverter.TryWriteBytes(dataStream.AsSpan(0), (int)20);
    byte[] pngSig = { 0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A };
    Buffer.BlockCopy(pngSig, 0, dataStream, 4, 8);

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
    var tblBytes = tblMs.ToArray();

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
            using (var s = root.CreateStream("0Table"))       s.Write(tblBytes);
            using (var s = root.CreateStream("Data"))         s.Write(dataStream);
        }
        using var fs = File.OpenRead(tmp);
        var doc = new PolyDonky.Convert.Doc.DocBinaryReader().Read(fs);
        var blocks = doc.Sections[0].Blocks;

        SmokeHarness.True(blocks.Count >= 2, $"blocks.Count >= 2 (got {blocks.Count})");
        SmokeHarness.True(blocks[0] is PolyDonky.Core.Table, $"blocks[0] = Table (got {blocks[0].GetType().Name})");
        var tbl = (PolyDonky.Core.Table)blocks[0];
        SmokeHarness.Equal(1, tbl.Rows.Count, "Rows.Count = 1");
        SmokeHarness.Equal(1, tbl.Rows[0].Cells.Count, "Cells.Count = 1");
        var cellBlocks = tbl.Rows[0].Cells[0].Blocks;
        SmokeHarness.Equal(3, cellBlocks.Count, $"cell.Blocks.Count = 3 (got {cellBlocks.Count})");
        SmokeHarness.True(cellBlocks[0] is Paragraph, $"cell[0] = Paragraph (got {cellBlocks[0].GetType().Name})");
        SmokeHarness.True(cellBlocks[1] is ImageBlock,  $"cell[1] = ImageBlock (got {cellBlocks[1].GetType().Name})");
        SmokeHarness.True(cellBlocks[2] is Paragraph, $"cell[2] = Paragraph (got {cellBlocks[2].GetType().Name})");
        SmokeHarness.Equal("X", ((Paragraph)cellBlocks[0]).GetPlainText(), "cell[0] 텍스트 = X");
        SmokeHarness.Equal("Y", ((Paragraph)cellBlocks[2]).GetPlainText(), "cell[2] 텍스트 = Y");
        var img = (ImageBlock)cellBlocks[1];
        SmokeHarness.Equal("image/png", img.MediaType, "셀 안 이미지 MediaType = image/png");
    }
    finally { try { File.Delete(tmp); } catch { } }
}

static void DocBinaryPhase3e5EmfImage()
{
    // Phase 3e-2 와 동일 구조 — Data stream 의 PICF 안에 EMF EMR_HEADER 박는다.
    //   iType=1 (01 00 00 00) + nSize + zero RECTL + dSignature " EMF" at offset 40.
    //   결과: ImageBlock.MediaType = "image/emf".
    const string text = "BeforeAfter\r";
    int ccp = text.Length;
    var textBytes = Encoding.Unicode.GetBytes(text);
    int fcText = 0x200;
    int pnPapx = 4, pnChpx = 5;
    int fcPapxFkp = pnPapx * 512;
    int fcChpxFkp = pnChpx * 512;
    int wdSize = fcChpxFkp + 512;
    var wd = new byte[wdSize];
    Buffer.BlockCopy(textBytes, 0, wd, fcText, textBytes.Length);

    int cpara = 1;
    BitConverter.TryWriteBytes(wd.AsSpan(fcPapxFkp + 0), (int)0x200);
    BitConverter.TryWriteBytes(wd.AsSpan(fcPapxFkp + 4), (int)(0x200 + ccp * 2));
    int papx0Off = 64;
    wd[fcPapxFkp + 8 + 0 * 13] = (byte)(papx0Off / 2);
    int p = fcPapxFkp + papx0Off;
    wd[p++] = 0; wd[p++] = 1; BitConverter.TryWriteBytes(wd.AsSpan(p), (ushort)0);
    wd[fcPapxFkp + 511] = (byte)cpara;

    int crun = 3;
    BitConverter.TryWriteBytes(wd.AsSpan(fcChpxFkp + 0),  (int)0x200);
    BitConverter.TryWriteBytes(wd.AsSpan(fcChpxFkp + 4),  (int)0x20C);
    BitConverter.TryWriteBytes(wd.AsSpan(fcChpxFkp + 8),  (int)0x20E);
    BitConverter.TryWriteBytes(wd.AsSpan(fcChpxFkp + 12), (int)0x218);
    int chpx0Off = 64, chpx1Off = 80, chpx2Off = 96;
    int rgbBase  = 4 * (crun + 1);
    wd[fcChpxFkp + rgbBase + 0] = (byte)(chpx0Off / 2);
    wd[fcChpxFkp + rgbBase + 1] = (byte)(chpx1Off / 2);
    wd[fcChpxFkp + rgbBase + 2] = (byte)(chpx2Off / 2);
    wd[fcChpxFkp + chpx0Off + 0] = 0;
    int c = fcChpxFkp + chpx1Off;
    wd[c++] = 6;
    BitConverter.TryWriteBytes(wd.AsSpan(c), (ushort)0x6A03); c += 2;
    BitConverter.TryWriteBytes(wd.AsSpan(c), (int)0); c += 4;
    wd[fcChpxFkp + chpx2Off + 0] = 0;
    wd[fcChpxFkp + 511] = (byte)crun;

    // Data stream: lcb=60, EMR_HEADER 시작.
    //   offset 4..7 : iType = 01 00 00 00
    //   offset 8..11: nSize = 88 (any)
    //   offset 12..43: RECTL × 2 (zeros)
    //   offset 44..47: dSignature " EMF"
    //   offset 48..63: filler
    var dataStream = new byte[64];
    BitConverter.TryWriteBytes(dataStream.AsSpan(0), (int)60);
    BitConverter.TryWriteBytes(dataStream.AsSpan(4), (int)1);     // iType = 1 (EMR_HEADER)
    BitConverter.TryWriteBytes(dataStream.AsSpan(8), (int)88);    // nSize
    dataStream[44] = 0x20; dataStream[45] = 0x45; dataStream[46] = 0x4D; dataStream[47] = 0x46;  // " EMF"

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
    var tblBytes = tblMs.ToArray();

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
            using (var s = root.CreateStream("0Table"))       s.Write(tblBytes);
            using (var s = root.CreateStream("Data"))         s.Write(dataStream);
        }
        using var fs = File.OpenRead(tmp);
        var doc = new PolyDonky.Convert.Doc.DocBinaryReader().Read(fs);
        var blocks = doc.Sections[0].Blocks;

        SmokeHarness.True(blocks.Count >= 3, $"blocks.Count >= 3 (got {blocks.Count})");
        SmokeHarness.True(blocks[1] is ImageBlock, "blocks[1] = ImageBlock");
        var img = (ImageBlock)blocks[1];
        SmokeHarness.Equal("image/emf", img.MediaType, "MediaType = image/emf");
        SmokeHarness.True(img.Data.Length >= 44, $"Data.Length >= 44 (got {img.Data.Length})");
        SmokeHarness.True(
            img.Data[0] == 0x01 && img.Data[1] == 0 && img.Data[2] == 0 && img.Data[3] == 0 &&
            img.Data[40] == 0x20 && img.Data[41] == 0x45 && img.Data[42] == 0x4D && img.Data[43] == 0x46,
            "Data 의 EMR_HEADER iType=1 + dSignature \" EMF\" @ offset 40");
    }
    finally { try { File.Delete(tmp); } catch { } }
}

// Phase 3e-6 — OfficeArt 컨테이너 정식 파싱.
//
// 두 케이스 모두 signature 스캔으로는 못 찾는 image bytes (raw signature 아님) 를 BLIP body 에 박아
// OfficeArt path 가 실제로 사용됨을 검증한다.

static void DocBinaryPhase3e6BlipPng()
{
    // OfficeArtSpContainer (0xF004, recVer=0xF) 안에 BLIP_PNG (0xF01E) atom.
    //   image data = AB CD EF 12 (PNG signature 아님 — OfficeArt path 만 추출 가능).
    DocBinaryRunOfficeArtPicCase(BuildSpContainerPngStream(), "image/png",
        new byte[] { 0xAB, 0xCD, 0xEF, 0x12 }, "BLIP_PNG in SpContainer");
}

static void DocBinaryPhase3e6BlipJpeg()
{
    // BLIP_JPEG (0xF01D) atom 을 PICF 최상위에 직접.
    //   image data = 11 22 33 44 (JPEG signature 아님).
    DocBinaryRunOfficeArtPicCase(BuildBlipJpegStream(), "image/jpeg",
        new byte[] { 0x11, 0x22, 0x33, 0x44 }, "BLIP_JPEG atom");
}

static byte[] BuildSpContainerPngStream()
{
    // Layout (41 byte):
    //   [0..3]   lcb = 41
    //   [4..11]  SpContainer header: verInst=0x000F, recType=0xF004, recLen=29
    //   [12..19] BLIP_PNG header: verInst=0x6E00, recType=0xF01E, recLen=21
    //   [20..35] UID = 16 zeros
    //   [36]     tag = 0xFF
    //   [37..40] image data = AB CD EF 12
    var d = new byte[41];
    BitConverter.TryWriteBytes(d.AsSpan(0),  (int)41);
    // SpContainer
    BitConverter.TryWriteBytes(d.AsSpan(4),  (ushort)0x000F);
    BitConverter.TryWriteBytes(d.AsSpan(6),  (ushort)0xF004);
    BitConverter.TryWriteBytes(d.AsSpan(8),  (uint)29);
    // BLIP_PNG
    BitConverter.TryWriteBytes(d.AsSpan(12), (ushort)0x6E00);
    BitConverter.TryWriteBytes(d.AsSpan(14), (ushort)0xF01E);
    BitConverter.TryWriteBytes(d.AsSpan(16), (uint)21);
    // UID = zeros
    d[36] = 0xFF;
    d[37] = 0xAB; d[38] = 0xCD; d[39] = 0xEF; d[40] = 0x12;
    return d;
}

static byte[] BuildBlipJpegStream()
{
    // Layout (33 byte):
    //   [0..3]   lcb = 33
    //   [4..11]  BLIP_JPEG header: verInst=0x46A0, recType=0xF01D, recLen=21
    //   [12..27] UID = 16 zeros
    //   [28]     tag = 0xFF
    //   [29..32] image data = 11 22 33 44
    var d = new byte[33];
    BitConverter.TryWriteBytes(d.AsSpan(0),  (int)33);
    BitConverter.TryWriteBytes(d.AsSpan(4),  (ushort)0x46A0);
    BitConverter.TryWriteBytes(d.AsSpan(6),  (ushort)0xF01D);
    BitConverter.TryWriteBytes(d.AsSpan(8),  (uint)21);
    d[28] = 0xFF;
    d[29] = 0x11; d[30] = 0x22; d[31] = 0x33; d[32] = 0x44;
    return d;
}

static void DocBinaryRunOfficeArtPicCase(
    byte[] dataStream, string expectedMime, byte[] expectedBytes, string label)
{
    const string text = "BeforeAfter\r";
    int ccp = text.Length;
    var textBytes = Encoding.Unicode.GetBytes(text);
    int fcText = 0x200;
    int pnPapx = 4, pnChpx = 5;
    int fcPapxFkp = pnPapx * 512;
    int fcChpxFkp = pnChpx * 512;
    int wdSize = fcChpxFkp + 512;
    var wd = new byte[wdSize];
    Buffer.BlockCopy(textBytes, 0, wd, fcText, textBytes.Length);

    int cpara = 1;
    BitConverter.TryWriteBytes(wd.AsSpan(fcPapxFkp + 0), (int)0x200);
    BitConverter.TryWriteBytes(wd.AsSpan(fcPapxFkp + 4), (int)(0x200 + ccp * 2));
    int papx0Off = 64;
    wd[fcPapxFkp + 8 + 0 * 13] = (byte)(papx0Off / 2);
    int p = fcPapxFkp + papx0Off;
    wd[p++] = 0; wd[p++] = 1; BitConverter.TryWriteBytes(wd.AsSpan(p), (ushort)0);
    wd[fcPapxFkp + 511] = (byte)cpara;

    int crun = 3;
    BitConverter.TryWriteBytes(wd.AsSpan(fcChpxFkp + 0),  (int)0x200);
    BitConverter.TryWriteBytes(wd.AsSpan(fcChpxFkp + 4),  (int)0x20C);
    BitConverter.TryWriteBytes(wd.AsSpan(fcChpxFkp + 8),  (int)0x20E);
    BitConverter.TryWriteBytes(wd.AsSpan(fcChpxFkp + 12), (int)0x218);
    int chpx0Off = 64, chpx1Off = 80, chpx2Off = 96;
    int rgbBase  = 4 * (crun + 1);
    wd[fcChpxFkp + rgbBase + 0] = (byte)(chpx0Off / 2);
    wd[fcChpxFkp + rgbBase + 1] = (byte)(chpx1Off / 2);
    wd[fcChpxFkp + rgbBase + 2] = (byte)(chpx2Off / 2);
    wd[fcChpxFkp + chpx0Off + 0] = 0;
    int c = fcChpxFkp + chpx1Off;
    wd[c++] = 6;
    BitConverter.TryWriteBytes(wd.AsSpan(c), (ushort)0x6A03); c += 2;
    BitConverter.TryWriteBytes(wd.AsSpan(c), (int)0); c += 4;
    wd[fcChpxFkp + chpx2Off + 0] = 0;
    wd[fcChpxFkp + 511] = (byte)crun;

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
    var tblBytes = tblMs.ToArray();

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
            using (var s = root.CreateStream("0Table"))       s.Write(tblBytes);
            using (var s = root.CreateStream("Data"))         s.Write(dataStream);
        }
        using var fs = File.OpenRead(tmp);
        var doc = new PolyDonky.Convert.Doc.DocBinaryReader().Read(fs);
        var blocks = doc.Sections[0].Blocks;

        SmokeHarness.True(blocks.Count >= 3, $"[{label}] blocks.Count >= 3 (got {blocks.Count})");
        SmokeHarness.True(blocks[1] is ImageBlock, $"[{label}] blocks[1] = ImageBlock");
        var img = (ImageBlock)blocks[1];
        SmokeHarness.Equal(expectedMime, img.MediaType, $"[{label}] MediaType");
        SmokeHarness.Equal(expectedBytes.Length, img.Data.Length, $"[{label}] Data.Length");
        for (int i = 0; i < expectedBytes.Length; i++)
            SmokeHarness.Equal(expectedBytes[i], img.Data[i], $"[{label}] Data[{i}]");
    }
    finally { try { File.Delete(tmp); } catch { } }
}

// Phase 3e-7 — WMF/EMF/DIB BLIP. Bitmap (DIB) 은 PNG/JPEG 와 같은 레이아웃, Metafile
// (WMF/EMF) 은 UID 다음에 34-byte OfficeArtMetafileHeader 가 추가됨. compression=0xFE
// 는 raw, 0x00 은 zlib — ZLibStream 으로 풀어 원본 복원.

static void DocBinaryPhase3e7BlipDib()
{
    // BLIP_DIB (0xF01F) atom — bitmap BLIP, layout = UID(16) + tag(1) + image bytes.
    //   image data = 99 88 77 66 (DIB BITMAPINFOHEADER signature 아님).
    DocBinaryRunOfficeArtPicCase(BuildBlipDibStream(), "image/x-dib",
        new byte[] { 0x99, 0x88, 0x77, 0x66 }, "BLIP_DIB atom");
}

static void DocBinaryPhase3e7BlipWmfUncompressed()
{
    // BLIP_WMF (0xF01B) atom, compression=0xFE (raw).
    //   image data = 55 44 33 22 (WMF placeable signature 아님).
    DocBinaryRunOfficeArtPicCase(BuildBlipWmfRawStream(), "image/wmf",
        new byte[] { 0x55, 0x44, 0x33, 0x22 }, "BLIP_WMF uncompressed");
}

static void DocBinaryPhase3e7BlipEmfDeflate()
{
    // BLIP_EMF (0xF01A) atom, compression=0x00 (zlib).
    //   uncompressed payload = "EMF_RAW!" (8 byte) — extractor 가 ZLibStream 으로 풀어야 일치.
    byte[] payload = { 0x45, 0x4D, 0x46, 0x5F, 0x52, 0x41, 0x57, 0x21 };
    DocBinaryRunOfficeArtPicCase(BuildBlipEmfDeflateStream(payload), "image/emf",
        payload, "BLIP_EMF deflate");
}

static byte[] BuildBlipDibStream()
{
    // 33 byte: lcb(4) + BLIP header(8) + UID(16) + tag(1) + image(4)
    var d = new byte[33];
    BitConverter.TryWriteBytes(d.AsSpan(0),  (int)33);
    BitConverter.TryWriteBytes(d.AsSpan(4),  (ushort)0x7A80);   // recInst=0x7A8, recVer=0
    BitConverter.TryWriteBytes(d.AsSpan(6),  (ushort)0xF01F);   // BLIP_DIB
    BitConverter.TryWriteBytes(d.AsSpan(8),  (uint)21);         // recLen = 16 + 1 + 4
    d[28] = 0xFF;
    d[29] = 0x99; d[30] = 0x88; d[31] = 0x77; d[32] = 0x66;
    return d;
}

static byte[] BuildBlipWmfRawStream()
{
    // 66 byte: lcb(4) + BLIP header(8) + UID(16) + MetafileHeader(34) + image(4)
    var d = new byte[66];
    BitConverter.TryWriteBytes(d.AsSpan(0),  (int)66);
    BitConverter.TryWriteBytes(d.AsSpan(4),  (ushort)0x2160);   // recInst=0x216, recVer=0
    BitConverter.TryWriteBytes(d.AsSpan(6),  (ushort)0xF01B);   // BLIP_WMF
    BitConverter.TryWriteBytes(d.AsSpan(8),  (uint)54);         // recLen = 16 + 34 + 4
    // [12..27] UID = zeros
    // [28..61] OfficeArtMetafileHeader (34 byte)
    BitConverter.TryWriteBytes(d.AsSpan(28), (uint)4);          // cbSize = 4 (uncompressed)
    // [32..47] rcBounds = zeros (16)
    // [48..55] ptSize  = zeros (8)
    BitConverter.TryWriteBytes(d.AsSpan(56), (uint)4);          // cbSave = 4
    d[60] = 0xFE;                                                // compression = none
    d[61] = 0xFE;                                                // filter
    // [62..65] image data
    d[62] = 0x55; d[63] = 0x44; d[64] = 0x33; d[65] = 0x22;
    return d;
}

static byte[] BuildBlipEmfDeflateStream(byte[] payload)
{
    // payload 를 zlib 으로 압축한 뒤 BLIP_EMF body 안에 넣는다.
    byte[] compressed;
    using (var msComp = new MemoryStream())
    {
        using (var zs = new System.IO.Compression.ZLibStream(
            msComp, System.IO.Compression.CompressionLevel.Optimal, leaveOpen: true))
        {
            zs.Write(payload, 0, payload.Length);
        }
        compressed = msComp.ToArray();
    }

    // lcb(4) + BLIP header(8) + UID(16) + MetafileHeader(34) + compressed(N)
    int total = 4 + 8 + 16 + 34 + compressed.Length;
    var d = new byte[total];
    BitConverter.TryWriteBytes(d.AsSpan(0),  (int)total);
    BitConverter.TryWriteBytes(d.AsSpan(4),  (ushort)0x3D40);   // recInst=0x3D4, recVer=0
    BitConverter.TryWriteBytes(d.AsSpan(6),  (ushort)0xF01A);   // BLIP_EMF
    BitConverter.TryWriteBytes(d.AsSpan(8),  (uint)(16 + 34 + compressed.Length));
    // [12..27] UID = zeros
    BitConverter.TryWriteBytes(d.AsSpan(28), (uint)payload.Length);     // cbSize
    BitConverter.TryWriteBytes(d.AsSpan(56), (uint)compressed.Length);  // cbSave
    d[60] = 0x00;  // compression = DEFLATE
    d[61] = 0xFE;  // filter
    Buffer.BlockCopy(compressed, 0, d, 62, compressed.Length);
    return d;
}

// Phase 3f — OfficeArtDggContainer / BStoreContainer / FBSE 들에서 공유 BLIP 추출 검증.
//
// Table stream 안에 다음 구조 박는다:
//   DggContainer (0xF000, container)
//     BStoreContainer (0xF001, container, recInst = cBlips = 2)
//       FBSE #1 (0xF007, atom) — embedded BLIP_PNG, image data = AA BB CC DD
//       FBSE #2 (0xF007, atom) — embedded BLIP_PNG, image data = 11 22 33 44
// FIB.fcDggInfo / lcbDggInfo 가 이 DggContainer 영역을 가리킨다.
//
// 결과: reader.BStoreImages = [(image/png, AABBCCDD), (image/png, 11223344)].
static void DocBinaryPhase3fBStoreFbse()
{
    const string text = "Body\r";
    int ccp = text.Length;
    var textBytes = Encoding.Unicode.GetBytes(text);
    int fcText = 0x200;
    int pnPapx = 4, pnChpx = 5;
    int fcPapxFkp = pnPapx * 512;
    int fcChpxFkp = pnChpx * 512;
    int wdSize = fcChpxFkp + 512;
    var wd = new byte[wdSize];
    Buffer.BlockCopy(textBytes, 0, wd, fcText, textBytes.Length);

    // PAPX FKP — 1 단락.
    int cpara = 1;
    BitConverter.TryWriteBytes(wd.AsSpan(fcPapxFkp + 0), (int)0x200);
    BitConverter.TryWriteBytes(wd.AsSpan(fcPapxFkp + 4), (int)(0x200 + ccp * 2));
    int papx0Off = 64;
    wd[fcPapxFkp + 8 + 0 * 13] = (byte)(papx0Off / 2);
    int p = fcPapxFkp + papx0Off;
    wd[p++] = 0; wd[p++] = 1; BitConverter.TryWriteBytes(wd.AsSpan(p), (ushort)0);
    wd[fcPapxFkp + 511] = (byte)cpara;

    // CHPX FKP — 빈.
    int crun = 1;
    BitConverter.TryWriteBytes(wd.AsSpan(fcChpxFkp + 0), (int)0x200);
    BitConverter.TryWriteBytes(wd.AsSpan(fcChpxFkp + 4), (int)(0x200 + ccp * 2));
    int chpx0Off = 64;
    wd[fcChpxFkp + 4 * (crun + 1) + 0] = (byte)(chpx0Off / 2);
    wd[fcChpxFkp + chpx0Off] = 0;
    wd[fcChpxFkp + 511] = (byte)crun;

    // Table stream — CLX, BTE, 그리고 DggContainer.
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

    // ── DggContainer 합성 ─────────────────────────────────────────────
    // FBSE(73) = header(8) + body(36) + embedded BLIP_PNG(29)
    //   embedded BLIP_PNG(29) = header(8) + UID(16) + tag(1) + image(4)
    // BStoreContainer body = 2 * 73 = 146 → total = 154
    // DggContainer body = 154 → total = 162
    int dggStart = (int)tblMs.Position;

    Span<byte> b2 = stackalloc byte[2];

    // DggContainer header
    BitConverter.TryWriteBytes(b2, (ushort)0x000F); tblMs.Write(b2);   // recVer=0xF (container), recInst=0
    BitConverter.TryWriteBytes(b2, (ushort)0xF000); tblMs.Write(b2);
    BitConverter.TryWriteBytes(b4, (uint)154);      tblMs.Write(b4);

    // BStoreContainer header
    BitConverter.TryWriteBytes(b2, (ushort)0x002F); tblMs.Write(b2);   // recInst=2 (cBlips), recVer=0xF
    BitConverter.TryWriteBytes(b2, (ushort)0xF001); tblMs.Write(b2);
    BitConverter.TryWriteBytes(b4, (uint)146);      tblMs.Write(b4);

    static byte[] BuildFbseWithPng(byte[] imageBytes)
    {
        var f = new byte[8 + 36 + 29];
        // FBSE header
        BitConverter.TryWriteBytes(f.AsSpan(0), (ushort)0x0062);   // recVer=2, recInst=msoblipPNG(6)
        BitConverter.TryWriteBytes(f.AsSpan(2), (ushort)0xF007);
        BitConverter.TryWriteBytes(f.AsSpan(4), (uint)(36 + 29));  // recLen = 65
        // [8..43] FBSE body 36 byte: 모두 0 (btWin32, btMacOS, rgbUid, tag, size, cRef, foDelay,
        //                                    unused1, cbName=0, unused2, unused3).
        // [44..51] 임베드 BLIP_PNG header (8 byte)
        BitConverter.TryWriteBytes(f.AsSpan(44), (ushort)0x6E00);  // recInst=0x6E0, recVer=0
        BitConverter.TryWriteBytes(f.AsSpan(46), (ushort)0xF01E);
        BitConverter.TryWriteBytes(f.AsSpan(48), (uint)21);
        // [52..67] UID = 16 zeros
        f[68] = 0xFF;                                              // tag
        Buffer.BlockCopy(imageBytes, 0, f, 69, 4);
        return f;
    }

    var fbse1 = BuildFbseWithPng(new byte[] { 0xAA, 0xBB, 0xCC, 0xDD });
    var fbse2 = BuildFbseWithPng(new byte[] { 0x11, 0x22, 0x33, 0x44 });
    tblMs.Write(fbse1, 0, fbse1.Length);
    tblMs.Write(fbse2, 0, fbse2.Length);

    int dggLen = (int)tblMs.Position - dggStart;
    var tblBytes = tblMs.ToArray();

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
    // Phase 3f — FIB fcDggInfo / lcbDggInfo
    BitConverter.TryWriteBytes(wd.AsSpan(0x0312), (uint)dggStart);
    BitConverter.TryWriteBytes(wd.AsSpan(0x0316), (uint)dggLen);

    var tmp = Path.Combine(Path.GetTempPath(), $"polydonky-smoke-{Guid.NewGuid():N}.doc");
    try
    {
        using (var root = OpenMcdf.RootStorage.Create(tmp))
        {
            using (var s = root.CreateStream("WordDocument")) s.Write(wd);
            using (var s = root.CreateStream("0Table"))       s.Write(tblBytes);
        }
        using var fs = File.OpenRead(tmp);
        var reader = new PolyDonky.Convert.Doc.DocBinaryReader();
        reader.Read(fs);

        SmokeHarness.Equal(2, reader.BStoreImages.Count,
            $"BStoreImages.Count = 2 (got {reader.BStoreImages.Count})");
        SmokeHarness.Equal("image/png", reader.BStoreImages[0].MediaType, "BLIP #1 MediaType");
        SmokeHarness.Equal("image/png", reader.BStoreImages[1].MediaType, "BLIP #2 MediaType");
        SmokeHarness.Equal(4, reader.BStoreImages[0].Data.Length, "BLIP #1 Data.Length");
        SmokeHarness.Equal(4, reader.BStoreImages[1].Data.Length, "BLIP #2 Data.Length");
        SmokeHarness.True(
            reader.BStoreImages[0].Data[0] == 0xAA && reader.BStoreImages[0].Data[1] == 0xBB &&
            reader.BStoreImages[0].Data[2] == 0xCC && reader.BStoreImages[0].Data[3] == 0xDD,
            "BLIP #1 bytes = AA BB CC DD");
        SmokeHarness.True(
            reader.BStoreImages[1].Data[0] == 0x11 && reader.BStoreImages[1].Data[1] == 0x22 &&
            reader.BStoreImages[1].Data[2] == 0x33 && reader.BStoreImages[1].Data[3] == 0x44,
            "BLIP #2 bytes = 11 22 33 44");
    }
    finally { try { File.Delete(tmp); } catch { } }
}

// Phase 3f-2 — PICF body 가 인라인 BLIP 없이 OfficeArtFOPT pib (property ID 260, fBid=1) 만 가지고
// BStore index 를 가리키는 케이스. 본문 picture marker 의 ImageBlock 이 BStore[0] 의 데이터로 채워져야 함.
static void DocBinaryPhase3f2PicfPibReferencesBStore()
{
    // 본문 텍스트 "Before\x01After\r" (Phase 3e-2 와 동일 구조).
    const string text = "BeforeAfter\r";
    int ccp = text.Length;
    var textBytes = Encoding.Unicode.GetBytes(text);
    int fcText = 0x200;
    int pnPapx = 4, pnChpx = 5;
    int fcPapxFkp = pnPapx * 512;
    int fcChpxFkp = pnChpx * 512;
    int wdSize = fcChpxFkp + 512;
    var wd = new byte[wdSize];
    Buffer.BlockCopy(textBytes, 0, wd, fcText, textBytes.Length);

    // PAPX FKP — 1 단락.
    int cpara = 1;
    BitConverter.TryWriteBytes(wd.AsSpan(fcPapxFkp + 0), (int)0x200);
    BitConverter.TryWriteBytes(wd.AsSpan(fcPapxFkp + 4), (int)(0x200 + ccp * 2));
    int papx0Off = 64;
    wd[fcPapxFkp + 8 + 0 * 13] = (byte)(papx0Off / 2);
    int p = fcPapxFkp + papx0Off;
    wd[p++] = 0; wd[p++] = 1; BitConverter.TryWriteBytes(wd.AsSpan(p), (ushort)0);
    wd[fcPapxFkp + 511] = (byte)cpara;

    // CHPX FKP — 3 runs ("Before", "\x01", "After\r"). \x01 의 CHPX 에 sprmCPicLocation = 0.
    int crun = 3;
    BitConverter.TryWriteBytes(wd.AsSpan(fcChpxFkp + 0),  (int)0x200);
    BitConverter.TryWriteBytes(wd.AsSpan(fcChpxFkp + 4),  (int)0x20C);
    BitConverter.TryWriteBytes(wd.AsSpan(fcChpxFkp + 8),  (int)0x20E);
    BitConverter.TryWriteBytes(wd.AsSpan(fcChpxFkp + 12), (int)0x218);
    int chpx0Off = 64, chpx1Off = 80, chpx2Off = 96;
    int rgbBase  = 4 * (crun + 1);
    wd[fcChpxFkp + rgbBase + 0] = (byte)(chpx0Off / 2);
    wd[fcChpxFkp + rgbBase + 1] = (byte)(chpx1Off / 2);
    wd[fcChpxFkp + rgbBase + 2] = (byte)(chpx2Off / 2);
    wd[fcChpxFkp + chpx0Off + 0] = 0;
    int c = fcChpxFkp + chpx1Off;
    wd[c++] = 6;
    BitConverter.TryWriteBytes(wd.AsSpan(c), (ushort)0x6A03); c += 2;
    BitConverter.TryWriteBytes(wd.AsSpan(c), (int)0); c += 4;
    wd[fcChpxFkp + chpx2Off + 0] = 0;
    wd[fcChpxFkp + 511] = (byte)crun;

    // Data stream PICF — 인라인 BLIP 없이 OfficeArtFOPT(0xF00B) 만.
    //   bytes [0..3]  lcb = 18
    //   bytes [4..11] OfficeArtFOPT header (verInst=(propCount=1<<4)|recVer=3 = 0x13, recType=0xF00B, recLen=6)
    //   bytes [12..17] pib property: opid=(260|0x4000)=0x4104, op=1 (1-based BStore index)
    var dataStream = new byte[18];
    BitConverter.TryWriteBytes(dataStream.AsSpan(0),  (int)18);
    BitConverter.TryWriteBytes(dataStream.AsSpan(4),  (ushort)0x0013);  // recInst=1, recVer=3
    BitConverter.TryWriteBytes(dataStream.AsSpan(6),  (ushort)0xF00B);
    BitConverter.TryWriteBytes(dataStream.AsSpan(8),  (uint)6);
    BitConverter.TryWriteBytes(dataStream.AsSpan(12), (ushort)0x4104); // propId=260, fBid=1
    BitConverter.TryWriteBytes(dataStream.AsSpan(14), (uint)1);        // BStore index 1

    // Table stream — CLX + BTE + DggContainer/BStoreContainer/FBSE × 1 (image bytes = CC DD EE FF).
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

    // DggContainer — BStoreContainer 안 FBSE × 1, embedded BLIP_PNG, image data = CC DD EE FF.
    int dggStart = (int)tblMs.Position;
    Span<byte> b2 = stackalloc byte[2];
    // DggContainer
    BitConverter.TryWriteBytes(b2, (ushort)0x000F); tblMs.Write(b2);
    BitConverter.TryWriteBytes(b2, (ushort)0xF000); tblMs.Write(b2);
    BitConverter.TryWriteBytes(b4, (uint)81);       tblMs.Write(b4);      // BStoreContainer(8) + FBSE(73) = 81
    // BStoreContainer
    BitConverter.TryWriteBytes(b2, (ushort)0x001F); tblMs.Write(b2);      // cBlips=1
    BitConverter.TryWriteBytes(b2, (ushort)0xF001); tblMs.Write(b2);
    BitConverter.TryWriteBytes(b4, (uint)73);       tblMs.Write(b4);
    // FBSE
    BitConverter.TryWriteBytes(b2, (ushort)0x0062); tblMs.Write(b2);      // recVer=2, recInst=6 (PNG)
    BitConverter.TryWriteBytes(b2, (ushort)0xF007); tblMs.Write(b2);
    BitConverter.TryWriteBytes(b4, (uint)(36 + 29)); tblMs.Write(b4);
    tblMs.Write(new byte[36]);                                            // FBSE body fixed 36 byte
    // Embedded BLIP_PNG
    BitConverter.TryWriteBytes(b2, (ushort)0x6E00); tblMs.Write(b2);
    BitConverter.TryWriteBytes(b2, (ushort)0xF01E); tblMs.Write(b2);
    BitConverter.TryWriteBytes(b4, (uint)21);       tblMs.Write(b4);
    tblMs.Write(new byte[16]);                                            // UID
    tblMs.WriteByte(0xFF);                                                // tag
    tblMs.Write(new byte[] { 0xCC, 0xDD, 0xEE, 0xFF }, 0, 4);
    int dggLen = (int)tblMs.Position - dggStart;
    var tblBytes = tblMs.ToArray();

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
    BitConverter.TryWriteBytes(wd.AsSpan(0x0312), (uint)dggStart);
    BitConverter.TryWriteBytes(wd.AsSpan(0x0316), (uint)dggLen);

    var tmp = Path.Combine(Path.GetTempPath(), $"polydonky-smoke-{Guid.NewGuid():N}.doc");
    try
    {
        using (var root = OpenMcdf.RootStorage.Create(tmp))
        {
            using (var s = root.CreateStream("WordDocument")) s.Write(wd);
            using (var s = root.CreateStream("0Table"))       s.Write(tblBytes);
            using (var s = root.CreateStream("Data"))         s.Write(dataStream);
        }
        using var fs = File.OpenRead(tmp);
        var doc = new PolyDonky.Convert.Doc.DocBinaryReader().Read(fs);
        var blocks = doc.Sections[0].Blocks;

        SmokeHarness.True(blocks.Count >= 3, $"blocks.Count >= 3 (got {blocks.Count})");
        SmokeHarness.True(blocks[1] is ImageBlock, "blocks[1] = ImageBlock");
        var img = (ImageBlock)blocks[1];
        SmokeHarness.Equal("image/png", img.MediaType, "MediaType = image/png (BStore[0])");
        SmokeHarness.Equal(4, img.Data.Length, "Data.Length = 4");
        SmokeHarness.True(
            img.Data[0] == 0xCC && img.Data[1] == 0xDD &&
            img.Data[2] == 0xEE && img.Data[3] == 0xFF,
            "Data bytes = CC DD EE FF (BStore 인덱스 1 의 BLIP)");
    }
    finally { try { File.Delete(tmp); } catch { } }
}

// Phase 3a-2 — 필드 instr 파싱. HYPERLINK → Run.Url, PAGE/NUMPAGES/DATE/... → Run.Field.
//   본문: "X<0x13>HYPERLINK \"https://example.com\"<0x14>here<0x15>Y<0x13>PAGE<0x14>3<0x15>\r"
//   결과 단락: [Run("X"), Run("here", Url=URL), Run("Y"), Run("3", Field=Page)]
static void DocBinaryPhase3a2FieldInstr()
{
    const string instr1 = "HYPERLINK \"https://example.com\"";
    const string instr2 = "PAGE";
    // text 구성:
    //   "X" + 0x13 + instr1 + 0x14 + "here" + 0x15 +
    //   "Y" + 0x13 + instr2 + 0x14 + "3"    + 0x15 + "\r"
    string text =
        "X" +
        "" + instr1 + "" + "here" + "" +
        "Y" +
        "" + instr2 + "" + "3"    + "" +
        "\r";
    int ccp = text.Length;
    var textBytes = Encoding.Unicode.GetBytes(text);
    int fcText = 0x200;
    int pnPapx = 4, pnChpx = 5;
    int fcPapxFkp = pnPapx * 512;
    int fcChpxFkp = pnChpx * 512;
    int wdSize = fcChpxFkp + 512;
    var wd = new byte[wdSize];
    Buffer.BlockCopy(textBytes, 0, wd, fcText, textBytes.Length);

    // PAPX FKP — 1 단락 ([0x200, 0x200 + ccp*2)).
    int cpara = 1;
    int paraEndFc = 0x200 + ccp * 2;
    BitConverter.TryWriteBytes(wd.AsSpan(fcPapxFkp + 0), (int)0x200);
    BitConverter.TryWriteBytes(wd.AsSpan(fcPapxFkp + 4), (int)paraEndFc);
    int papx0Off = 64;
    wd[fcPapxFkp + 8 + 0 * 13] = (byte)(papx0Off / 2);
    int p = fcPapxFkp + papx0Off;
    wd[p++] = 0; wd[p++] = 1; BitConverter.TryWriteBytes(wd.AsSpan(p), (ushort)0);
    wd[fcPapxFkp + 511] = (byte)cpara;

    // CHPX FKP — 빈 (1 run 전체 paragraph 범위).
    int crun = 1;
    BitConverter.TryWriteBytes(wd.AsSpan(fcChpxFkp + 0), (int)0x200);
    BitConverter.TryWriteBytes(wd.AsSpan(fcChpxFkp + 4), (int)paraEndFc);
    int chpx0Off = 64;
    wd[fcChpxFkp + 4 * (crun + 1) + 0] = (byte)(chpx0Off / 2);
    wd[fcChpxFkp + chpx0Off] = 0;
    wd[fcChpxFkp + 511] = (byte)crun;

    // Table stream.
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
    BitConverter.TryWriteBytes(b4, (int)paraEndFc); tblMs.Write(b4);
    BitConverter.TryWriteBytes(b4, (int)pnPapx); tblMs.Write(b4);
    int papxBteLen = (int)tblMs.Position - papxBteStart;
    int chpxBteStart = (int)tblMs.Position;
    BitConverter.TryWriteBytes(b4, (int)0x200);  tblMs.Write(b4);
    BitConverter.TryWriteBytes(b4, (int)paraEndFc); tblMs.Write(b4);
    BitConverter.TryWriteBytes(b4, (int)pnChpx); tblMs.Write(b4);
    int chpxBteLen = (int)tblMs.Position - chpxBteStart;
    var tblBytes = tblMs.ToArray();

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
            using (var s = root.CreateStream("0Table"))       s.Write(tblBytes);
        }
        using var fs = File.OpenRead(tmp);
        var doc = new PolyDonky.Convert.Doc.DocBinaryReader().Read(fs);
        var blocks = doc.Sections[0].Blocks;
        SmokeHarness.True(blocks.Count >= 1, $"blocks.Count >= 1 (got {blocks.Count})");
        SmokeHarness.True(blocks[0] is Paragraph, $"blocks[0] = Paragraph");
        var para = (Paragraph)blocks[0];

        // 결과 텍스트 = "XhereY3" (필드 instr 폐기됨, 결과만 보존).
        SmokeHarness.Equal("XhereY3", para.GetPlainText(), "단락 plain text");

        // Run 분할은 (스타일 + Url + Field) 경계마다. 같은 스타일 + 같은 (Url=null/Field=null) 인
        // "X" / "Y" 는 인접한 다른 (Url 또는 Field) 의 Run 과만 break. URL/Field 가 바뀌는 지점에서 break.
        // 예상: X | here(Url) | Y | 3(Field=Page)
        SmokeHarness.Equal(4, para.Runs.Count, $"Runs.Count = 4 (got {para.Runs.Count})");
        SmokeHarness.Equal("X",    para.Runs[0].Text, "Run[0].Text = X");
        SmokeHarness.True(para.Runs[0].Url is null && para.Runs[0].Field is null, "Run[0] 일반");

        SmokeHarness.Equal("here", para.Runs[1].Text, "Run[1].Text = here");
        SmokeHarness.Equal("https://example.com", para.Runs[1].Url, "Run[1].Url = URL");

        SmokeHarness.Equal("Y",    para.Runs[2].Text, "Run[2].Text = Y");
        SmokeHarness.True(para.Runs[2].Url is null && para.Runs[2].Field is null, "Run[2] 일반");

        SmokeHarness.Equal("3",    para.Runs[3].Text, "Run[3].Text = 3");
        SmokeHarness.Equal(FieldType.Page, para.Runs[3].Field, "Run[3].Field = Page");
    }
    finally { try { File.Delete(tmp); } catch { } }
}

// Phase 3g — 각주/미주 sub-document 추출. WordDocument 의 [ccpText, +ccpFtn) 와
// [+ccpHdd+ccpAtn, +ccpEdn) 영역의 텍스트를 PlcfFndTxt / PlcfEdnTxt 로 분할해
// doc.Footnotes / doc.Endnotes 로 노출.
static void DocBinaryPhase3gFootnotesEndnotes()
{
    // 텍스트 영역 구성:
    //   main:     "Body\r"             — 5 chars  (ccpText)
    //   footnote: " \rFootnote\r"      — 11 chars (ccpFtn)  separator + 1 footnote
    //   header:   (none)               — 0       (ccpHdd)
    //   comment:  (none)               — 0       (ccpAtn)
    //   endnote:  " \rEndnote\r"       — 10 chars (ccpEdn)  separator + 1 endnote
    const string text = "Body\r \rFootnote\r \rEndnote\r";
    int ccpMain  = 5;
    int ccpFtn   = 11;
    int ccpEdn   = 10;
    int totalCcp = ccpMain + ccpFtn + ccpEdn;  // = 26
    SmokeHarness.Equal(totalCcp, text.Length, "text 길이 검증");

    var textBytes = Encoding.Unicode.GetBytes(text);
    int fcText = 0x200;
    int pnPapx = 4, pnChpx = 5;
    int fcPapxFkp = pnPapx * 512;
    int fcChpxFkp = pnChpx * 512;
    int wdSize = fcChpxFkp + 512;
    var wd = new byte[wdSize];
    Buffer.BlockCopy(textBytes, 0, wd, fcText, textBytes.Length);

    // PAPX FKP — 1 단락 (본문 한 줄만 — footnote/endnote 영역의 PAPX 는 본 테스트 비검증).
    int cpara = 1;
    int paraEndFc = fcText + ccpMain * 2;
    BitConverter.TryWriteBytes(wd.AsSpan(fcPapxFkp + 0), (int)fcText);
    BitConverter.TryWriteBytes(wd.AsSpan(fcPapxFkp + 4), (int)paraEndFc);
    int papx0Off = 64;
    wd[fcPapxFkp + 8 + 0 * 13] = (byte)(papx0Off / 2);
    int p = fcPapxFkp + papx0Off;
    wd[p++] = 0; wd[p++] = 1; BitConverter.TryWriteBytes(wd.AsSpan(p), (ushort)0);
    wd[fcPapxFkp + 511] = (byte)cpara;

    // CHPX FKP — 빈.
    int crun = 1;
    BitConverter.TryWriteBytes(wd.AsSpan(fcChpxFkp + 0), (int)fcText);
    BitConverter.TryWriteBytes(wd.AsSpan(fcChpxFkp + 4), (int)paraEndFc);
    int chpx0Off = 64;
    wd[fcChpxFkp + 4 * (crun + 1) + 0] = (byte)(chpx0Off / 2);
    wd[fcChpxFkp + chpx0Off] = 0;
    wd[fcChpxFkp + 511] = (byte)crun;

    // Table stream — CLX (전체 텍스트 한 piece 로 커버), BTE, PlcfFndTxt, PlcfEdnTxt.
    var tblMs = new MemoryStream();
    Span<byte> b4 = stackalloc byte[4];
    tblMs.WriteByte(0x02);
    BitConverter.TryWriteBytes(b4, (uint)16); tblMs.Write(b4);
    BitConverter.TryWriteBytes(b4, (uint)0);          tblMs.Write(b4);  // CP[0] = 0
    BitConverter.TryWriteBytes(b4, (uint)totalCcp);   tblMs.Write(b4);  // CP[1] = totalCcp
    tblMs.WriteByte(0); tblMs.WriteByte(0);
    BitConverter.TryWriteBytes(b4, (uint)fcText);     tblMs.Write(b4);
    tblMs.WriteByte(0); tblMs.WriteByte(0);
    int clxEnd = (int)tblMs.Position;
    int papxBteStart = clxEnd;
    BitConverter.TryWriteBytes(b4, (int)fcText);      tblMs.Write(b4);
    BitConverter.TryWriteBytes(b4, (int)(fcText + totalCcp * 2 + 0x100)); tblMs.Write(b4);
    BitConverter.TryWriteBytes(b4, (int)pnPapx);      tblMs.Write(b4);
    int papxBteLen = (int)tblMs.Position - papxBteStart;
    int chpxBteStart = (int)tblMs.Position;
    BitConverter.TryWriteBytes(b4, (int)fcText);      tblMs.Write(b4);
    BitConverter.TryWriteBytes(b4, (int)(fcText + totalCcp * 2 + 0x100)); tblMs.Write(b4);
    BitConverter.TryWriteBytes(b4, (int)pnChpx);      tblMs.Write(b4);
    int chpxBteLen = (int)tblMs.Position - chpxBteStart;

    // PlcfFndTxt: aCP = [0, 2, 11] (separator: 2 chars, footnote: 9 chars)
    int fndStart = (int)tblMs.Position;
    BitConverter.TryWriteBytes(b4, (int)0);  tblMs.Write(b4);
    BitConverter.TryWriteBytes(b4, (int)2);  tblMs.Write(b4);
    BitConverter.TryWriteBytes(b4, (int)11); tblMs.Write(b4);
    int fndLen = (int)tblMs.Position - fndStart;

    // PlcfEdnTxt: aCP = [0, 2, 10]
    int ednStart = (int)tblMs.Position;
    BitConverter.TryWriteBytes(b4, (int)0);  tblMs.Write(b4);
    BitConverter.TryWriteBytes(b4, (int)2);  tblMs.Write(b4);
    BitConverter.TryWriteBytes(b4, (int)10); tblMs.Write(b4);
    int ednLen = (int)tblMs.Position - ednStart;

    var tblBytes = tblMs.ToArray();

    BitConverter.TryWriteBytes(wd.AsSpan(0x00),   (ushort)0xA5EC);
    BitConverter.TryWriteBytes(wd.AsSpan(0x02),   (ushort)193);
    BitConverter.TryWriteBytes(wd.AsSpan(0x0A),   (ushort)0x0000);
    BitConverter.TryWriteBytes(wd.AsSpan(0x18),   (uint)fcText);
    BitConverter.TryWriteBytes(wd.AsSpan(0x4C),   (uint)ccpMain);   // ccpText
    BitConverter.TryWriteBytes(wd.AsSpan(0x50),   (uint)ccpFtn);    // ccpFtn
    BitConverter.TryWriteBytes(wd.AsSpan(0x54),   (uint)0);         // ccpHdd
    BitConverter.TryWriteBytes(wd.AsSpan(0x5C),   (uint)0);         // ccpAtn
    BitConverter.TryWriteBytes(wd.AsSpan(0x60),   (uint)ccpEdn);    // ccpEdn
    BitConverter.TryWriteBytes(wd.AsSpan(0xB2),   (uint)fndStart);  // fcPlcffndTxt
    BitConverter.TryWriteBytes(wd.AsSpan(0xB6),   (uint)fndLen);    // lcbPlcffndTxt
    BitConverter.TryWriteBytes(wd.AsSpan(0x152),  (uint)ednStart);  // fcPlcfendTxt
    BitConverter.TryWriteBytes(wd.AsSpan(0x156),  (uint)ednLen);    // lcbPlcfendTxt
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
            using (var s = root.CreateStream("0Table"))       s.Write(tblBytes);
        }
        using var fs = File.OpenRead(tmp);
        var doc = new PolyDonky.Convert.Doc.DocBinaryReader().Read(fs);

        SmokeHarness.Equal(1, doc.Footnotes.Count, $"Footnotes.Count = 1 (got {doc.Footnotes.Count})");
        SmokeHarness.Equal("fn1", doc.Footnotes[0].Id, "Footnote Id = fn1");
        var fnPara = (Paragraph)doc.Footnotes[0].Blocks[0];
        SmokeHarness.Equal("Footnote", fnPara.GetPlainText(), "Footnote 본문");

        SmokeHarness.Equal(1, doc.Endnotes.Count, $"Endnotes.Count = 1 (got {doc.Endnotes.Count})");
        SmokeHarness.Equal("en1", doc.Endnotes[0].Id, "Endnote Id = en1");
        var enPara = (Paragraph)doc.Endnotes[0].Blocks[0];
        SmokeHarness.Equal("Endnote", enPara.GetPlainText(), "Endnote 본문");
    }
    finally { try { File.Delete(tmp); } catch { } }
}

// Phase 3a-3 — 필드 instr 추가 (NUMCHARS / FILENAME / SUBJECT / KEYWORDS / COMMENTS).
//   본문에 5 개의 인접 필드를 박아 각 result Run 에 알맞은 FieldType 이 매핑되는지 검증.
static void DocBinaryPhase3a3ExtraFieldInstr()
{
    // text = <0x13>NUMCHARS<0x14>a<0x15><0x13>FILENAME<0x14>b<0x15>
    //        <0x13>SUBJECT<0x14>c<0x15><0x13>KEYWORDS<0x14>d<0x15>
    //        <0x13>COMMENTS<0x14>e<0x15><\r>
    string text =
        "" + "NUMCHARS" + "" + "a" + "" +
        "" + "FILENAME" + "" + "b" + "" +
        "" + "SUBJECT"  + "" + "c" + "" +
        "" + "KEYWORDS" + "" + "d" + "" +
        "" + "COMMENTS" + "" + "e" + "" +
        "\r";
    int ccp = text.Length;
    var textBytes = Encoding.Unicode.GetBytes(text);
    int fcText = 0x200;
    int pnPapx = 4, pnChpx = 5;
    int fcPapxFkp = pnPapx * 512;
    int fcChpxFkp = pnChpx * 512;
    int wdSize = fcChpxFkp + 512;
    var wd = new byte[wdSize];
    Buffer.BlockCopy(textBytes, 0, wd, fcText, textBytes.Length);

    int cpara = 1;
    int paraEndFc = 0x200 + ccp * 2;
    BitConverter.TryWriteBytes(wd.AsSpan(fcPapxFkp + 0), (int)0x200);
    BitConverter.TryWriteBytes(wd.AsSpan(fcPapxFkp + 4), (int)paraEndFc);
    int papx0Off = 64;
    wd[fcPapxFkp + 8 + 0 * 13] = (byte)(papx0Off / 2);
    int p = fcPapxFkp + papx0Off;
    wd[p++] = 0; wd[p++] = 1; BitConverter.TryWriteBytes(wd.AsSpan(p), (ushort)0);
    wd[fcPapxFkp + 511] = (byte)cpara;

    int crun = 1;
    BitConverter.TryWriteBytes(wd.AsSpan(fcChpxFkp + 0), (int)0x200);
    BitConverter.TryWriteBytes(wd.AsSpan(fcChpxFkp + 4), (int)paraEndFc);
    int chpx0Off = 64;
    wd[fcChpxFkp + 4 * (crun + 1) + 0] = (byte)(chpx0Off / 2);
    wd[fcChpxFkp + chpx0Off] = 0;
    wd[fcChpxFkp + 511] = (byte)crun;

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
    BitConverter.TryWriteBytes(b4, (int)0x200);     tblMs.Write(b4);
    BitConverter.TryWriteBytes(b4, (int)paraEndFc); tblMs.Write(b4);
    BitConverter.TryWriteBytes(b4, (int)pnPapx);    tblMs.Write(b4);
    int papxBteLen = (int)tblMs.Position - papxBteStart;
    int chpxBteStart = (int)tblMs.Position;
    BitConverter.TryWriteBytes(b4, (int)0x200);     tblMs.Write(b4);
    BitConverter.TryWriteBytes(b4, (int)paraEndFc); tblMs.Write(b4);
    BitConverter.TryWriteBytes(b4, (int)pnChpx);    tblMs.Write(b4);
    int chpxBteLen = (int)tblMs.Position - chpxBteStart;
    var tblBytes = tblMs.ToArray();

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
            using (var s = root.CreateStream("0Table"))       s.Write(tblBytes);
        }
        using var fs = File.OpenRead(tmp);
        var doc = new PolyDonky.Convert.Doc.DocBinaryReader().Read(fs);
        var blocks = doc.Sections[0].Blocks;
        SmokeHarness.True(blocks.Count >= 1, $"blocks.Count >= 1 (got {blocks.Count})");
        var para = (Paragraph)blocks[0];

        SmokeHarness.Equal("abcde", para.GetPlainText(), "plain text = abcde");
        SmokeHarness.Equal(5, para.Runs.Count, $"Runs.Count = 5 (got {para.Runs.Count})");

        SmokeHarness.Equal("a", para.Runs[0].Text, "Run[0].Text");
        SmokeHarness.Equal(FieldType.NumChars, para.Runs[0].Field, "Run[0].Field = NumChars");
        SmokeHarness.Equal("b", para.Runs[1].Text, "Run[1].Text");
        SmokeHarness.Equal(FieldType.FileName, para.Runs[1].Field, "Run[1].Field = FileName");
        SmokeHarness.Equal("c", para.Runs[2].Text, "Run[2].Text");
        SmokeHarness.Equal(FieldType.Subject,  para.Runs[2].Field, "Run[2].Field = Subject");
        SmokeHarness.Equal("d", para.Runs[3].Text, "Run[3].Text");
        SmokeHarness.Equal(FieldType.Keywords, para.Runs[3].Field, "Run[3].Field = Keywords");
        SmokeHarness.Equal("e", para.Runs[4].Text, "Run[4].Text");
        SmokeHarness.Equal(FieldType.Comments, para.Runs[4].Field, "Run[4].Field = Comments");
    }
    finally { try { File.Delete(tmp); } catch { } }
}

// Phase 3a-4 — 참조 / 인클루드 / 조건 계열 필드 (SEQ / REF / STYLEREF / INCLUDETEXT / IF).
//   인스트의 첫 토큰만 파싱 — 인자 (카테고리/책갈피명/스타일명/경로) 는 현재 비보존.
//   결과 텍스트와 FieldType 만 검증.
static void DocBinaryPhase3a4ReferenceFieldInstr()
{
    // 5 필드 인접 — 각 instr 에 인자도 함께 박아 첫 토큰만 추출됨을 확인.
    //   "<0x13>SEQ Figure<0x14>1<0x15>
    //    <0x13>REF MyBookmark<0x14>2<0x15>
    //    <0x13>STYLEREF \"Heading 1\"<0x14>3<0x15>
    //    <0x13>INCLUDETEXT \"file.doc\"<0x14>4<0x15>
    //    <0x13>IF { PAGE } = 1<0x14>5<0x15>\r"
    string text =
        "" + "SEQ Figure"               + "" + "1" + "" +
        "" + "REF MyBookmark"           + "" + "2" + "" +
        "" + "STYLEREF \"Heading 1\""   + "" + "3" + "" +
        "" + "INCLUDETEXT \"file.doc\"" + "" + "4" + "" +
        "" + "IF { PAGE } = 1"          + "" + "5" + "" +
        "\r";
    int ccp = text.Length;
    var textBytes = Encoding.Unicode.GetBytes(text);
    int fcText = 0x200;
    int pnPapx = 4, pnChpx = 5;
    int fcPapxFkp = pnPapx * 512;
    int fcChpxFkp = pnChpx * 512;
    int wdSize = fcChpxFkp + 512;
    var wd = new byte[wdSize];
    Buffer.BlockCopy(textBytes, 0, wd, fcText, textBytes.Length);

    int cpara = 1;
    int paraEndFc = 0x200 + ccp * 2;
    BitConverter.TryWriteBytes(wd.AsSpan(fcPapxFkp + 0), (int)0x200);
    BitConverter.TryWriteBytes(wd.AsSpan(fcPapxFkp + 4), (int)paraEndFc);
    int papx0Off = 64;
    wd[fcPapxFkp + 8 + 0 * 13] = (byte)(papx0Off / 2);
    int p = fcPapxFkp + papx0Off;
    wd[p++] = 0; wd[p++] = 1; BitConverter.TryWriteBytes(wd.AsSpan(p), (ushort)0);
    wd[fcPapxFkp + 511] = (byte)cpara;

    int crun = 1;
    BitConverter.TryWriteBytes(wd.AsSpan(fcChpxFkp + 0), (int)0x200);
    BitConverter.TryWriteBytes(wd.AsSpan(fcChpxFkp + 4), (int)paraEndFc);
    int chpx0Off = 64;
    wd[fcChpxFkp + 4 * (crun + 1) + 0] = (byte)(chpx0Off / 2);
    wd[fcChpxFkp + chpx0Off] = 0;
    wd[fcChpxFkp + 511] = (byte)crun;

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
    BitConverter.TryWriteBytes(b4, (int)0x200);     tblMs.Write(b4);
    BitConverter.TryWriteBytes(b4, (int)paraEndFc); tblMs.Write(b4);
    BitConverter.TryWriteBytes(b4, (int)pnPapx);    tblMs.Write(b4);
    int papxBteLen = (int)tblMs.Position - papxBteStart;
    int chpxBteStart = (int)tblMs.Position;
    BitConverter.TryWriteBytes(b4, (int)0x200);     tblMs.Write(b4);
    BitConverter.TryWriteBytes(b4, (int)paraEndFc); tblMs.Write(b4);
    BitConverter.TryWriteBytes(b4, (int)pnChpx);    tblMs.Write(b4);
    int chpxBteLen = (int)tblMs.Position - chpxBteStart;
    var tblBytes = tblMs.ToArray();

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
            using (var s = root.CreateStream("0Table"))       s.Write(tblBytes);
        }
        using var fs = File.OpenRead(tmp);
        var doc = new PolyDonky.Convert.Doc.DocBinaryReader().Read(fs);
        var blocks = doc.Sections[0].Blocks;
        SmokeHarness.True(blocks.Count >= 1, $"blocks.Count >= 1 (got {blocks.Count})");
        var para = (Paragraph)blocks[0];

        SmokeHarness.Equal("12345", para.GetPlainText(), "plain text = 12345");
        SmokeHarness.Equal(5, para.Runs.Count, $"Runs.Count = 5 (got {para.Runs.Count})");

        SmokeHarness.Equal("1", para.Runs[0].Text, "Run[0].Text");
        SmokeHarness.Equal(FieldType.Seq,         para.Runs[0].Field, "Run[0].Field = Seq");
        SmokeHarness.Equal("2", para.Runs[1].Text, "Run[1].Text");
        SmokeHarness.Equal(FieldType.Ref,         para.Runs[1].Field, "Run[1].Field = Ref");
        SmokeHarness.Equal("3", para.Runs[2].Text, "Run[2].Text");
        SmokeHarness.Equal(FieldType.StyleRef,    para.Runs[2].Field, "Run[2].Field = StyleRef");
        SmokeHarness.Equal("4", para.Runs[3].Text, "Run[3].Text");
        SmokeHarness.Equal(FieldType.IncludeText, para.Runs[3].Field, "Run[3].Field = IncludeText");
        SmokeHarness.Equal("5", para.Runs[4].Text, "Run[4].Text");
        SmokeHarness.Equal(FieldType.If,          para.Runs[4].Field, "Run[4].Field = If");
    }
    finally { try { File.Delete(tmp); } catch { } }
}

// Phase 3a-5 — 필드 인스트의 핵심 인자를 Run.FieldArg 에 보존.
//   SEQ Figure / REF MyBookmark / STYLEREF "Heading 1" / INCLUDETEXT "file.doc" 의 첫 인자를
//   각 Run.FieldArg 로 추출. 따옴표 안쪽은 따옴표 빼고, 일반 토큰은 공백/스위치까지.
static void DocBinaryPhase3a5FieldArg()
{
    // 4 인접 필드 — 각 instr 인자 패턴이 다름.
    string text =
        "" + "SEQ Figure \\* ARABIC"     + "" + "1" + "" +
        "" + "REF MyBookmark \\h"        + "" + "2" + "" +
        "" + "STYLEREF \"Heading 1\""    + "" + "3" + "" +
        "" + "INCLUDETEXT \"C:\\file.doc\"" + "" + "4" + "" +
        "\r";
    int ccp = text.Length;
    var textBytes = Encoding.Unicode.GetBytes(text);
    int fcText = 0x200;
    int pnPapx = 4, pnChpx = 5;
    int fcPapxFkp = pnPapx * 512;
    int fcChpxFkp = pnChpx * 512;
    int wdSize = fcChpxFkp + 512;
    var wd = new byte[wdSize];
    Buffer.BlockCopy(textBytes, 0, wd, fcText, textBytes.Length);

    int cpara = 1;
    int paraEndFc = 0x200 + ccp * 2;
    BitConverter.TryWriteBytes(wd.AsSpan(fcPapxFkp + 0), (int)0x200);
    BitConverter.TryWriteBytes(wd.AsSpan(fcPapxFkp + 4), (int)paraEndFc);
    int papx0Off = 64;
    wd[fcPapxFkp + 8 + 0 * 13] = (byte)(papx0Off / 2);
    int p = fcPapxFkp + papx0Off;
    wd[p++] = 0; wd[p++] = 1; BitConverter.TryWriteBytes(wd.AsSpan(p), (ushort)0);
    wd[fcPapxFkp + 511] = (byte)cpara;

    int crun = 1;
    BitConverter.TryWriteBytes(wd.AsSpan(fcChpxFkp + 0), (int)0x200);
    BitConverter.TryWriteBytes(wd.AsSpan(fcChpxFkp + 4), (int)paraEndFc);
    int chpx0Off = 64;
    wd[fcChpxFkp + 4 * (crun + 1) + 0] = (byte)(chpx0Off / 2);
    wd[fcChpxFkp + chpx0Off] = 0;
    wd[fcChpxFkp + 511] = (byte)crun;

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
    BitConverter.TryWriteBytes(b4, (int)0x200);     tblMs.Write(b4);
    BitConverter.TryWriteBytes(b4, (int)paraEndFc); tblMs.Write(b4);
    BitConverter.TryWriteBytes(b4, (int)pnPapx);    tblMs.Write(b4);
    int papxBteLen = (int)tblMs.Position - papxBteStart;
    int chpxBteStart = (int)tblMs.Position;
    BitConverter.TryWriteBytes(b4, (int)0x200);     tblMs.Write(b4);
    BitConverter.TryWriteBytes(b4, (int)paraEndFc); tblMs.Write(b4);
    BitConverter.TryWriteBytes(b4, (int)pnChpx);    tblMs.Write(b4);
    int chpxBteLen = (int)tblMs.Position - chpxBteStart;
    var tblBytes = tblMs.ToArray();

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
            using (var s = root.CreateStream("0Table"))       s.Write(tblBytes);
        }
        using var fs = File.OpenRead(tmp);
        var doc = new PolyDonky.Convert.Doc.DocBinaryReader().Read(fs);
        var para = (Paragraph)doc.Sections[0].Blocks[0];

        SmokeHarness.Equal("1234", para.GetPlainText(), "plain text = 1234");
        SmokeHarness.Equal(4, para.Runs.Count, $"Runs.Count = 4 (got {para.Runs.Count})");

        SmokeHarness.Equal(FieldType.Seq, para.Runs[0].Field, "Run[0].Field = Seq");
        SmokeHarness.Equal("Figure", para.Runs[0].FieldArg, "Run[0].FieldArg = Figure");

        SmokeHarness.Equal(FieldType.Ref, para.Runs[1].Field, "Run[1].Field = Ref");
        SmokeHarness.Equal("MyBookmark", para.Runs[1].FieldArg, "Run[1].FieldArg = MyBookmark");

        SmokeHarness.Equal(FieldType.StyleRef, para.Runs[2].Field, "Run[2].Field = StyleRef");
        SmokeHarness.Equal("Heading 1", para.Runs[2].FieldArg, "Run[2].FieldArg = 'Heading 1'");

        SmokeHarness.Equal(FieldType.IncludeText, para.Runs[3].Field, "Run[3].Field = IncludeText");
        SmokeHarness.Equal("C:\\file.doc", para.Runs[3].FieldArg, "Run[3].FieldArg = C:\\file.doc");
    }
    finally { try { File.Delete(tmp); } catch { } }
}

// Phase 3f-3 — FBSE 가 임베드 BLIP 없이 foDelay 만 가지고 Data stream 안의 BLIP 을 가리키는 케이스.
//   Table 안 FBSE: recLen = 36 (body 만), foDelay = Data stream 안 BLIP_PNG 위치.
//   Data stream: 일부 padding 뒤에 BLIP_PNG 레코드 — image bytes = 99 88 77 66.
//   결과: reader.BStoreImages[0] = (image/png, 99 88 77 66).
static void DocBinaryPhase3f3FbseFoDelay()
{
    const string text = "Body\r";
    int ccp = text.Length;
    var textBytes = Encoding.Unicode.GetBytes(text);
    int fcText = 0x200;
    int pnPapx = 4, pnChpx = 5;
    int fcPapxFkp = pnPapx * 512;
    int fcChpxFkp = pnChpx * 512;
    int wdSize = fcChpxFkp + 512;
    var wd = new byte[wdSize];
    Buffer.BlockCopy(textBytes, 0, wd, fcText, textBytes.Length);

    // PAPX FKP — 1 단락.
    int cpara = 1;
    BitConverter.TryWriteBytes(wd.AsSpan(fcPapxFkp + 0), (int)0x200);
    BitConverter.TryWriteBytes(wd.AsSpan(fcPapxFkp + 4), (int)(0x200 + ccp * 2));
    int papx0Off = 64;
    wd[fcPapxFkp + 8 + 0 * 13] = (byte)(papx0Off / 2);
    int p = fcPapxFkp + papx0Off;
    wd[p++] = 0; wd[p++] = 1; BitConverter.TryWriteBytes(wd.AsSpan(p), (ushort)0);
    wd[fcPapxFkp + 511] = (byte)cpara;
    // CHPX FKP — 빈.
    int crun = 1;
    BitConverter.TryWriteBytes(wd.AsSpan(fcChpxFkp + 0), (int)0x200);
    BitConverter.TryWriteBytes(wd.AsSpan(fcChpxFkp + 4), (int)(0x200 + ccp * 2));
    int chpx0Off = 64;
    wd[fcChpxFkp + 4 * (crun + 1) + 0] = (byte)(chpx0Off / 2);
    wd[fcChpxFkp + chpx0Off] = 0;
    wd[fcChpxFkp + 511] = (byte)crun;

    // Data stream — [0..3] padding, [4..32] BLIP_PNG 레코드.
    //   BLIP_PNG (29 byte) = header(8) + UID(16) + tag(1) + image(4).
    var dataStream = new byte[33];
    int blipOffset = 4;
    BitConverter.TryWriteBytes(dataStream.AsSpan(blipOffset + 0), (ushort)0x6E00);  // recInst=0x6E0, recVer=0
    BitConverter.TryWriteBytes(dataStream.AsSpan(blipOffset + 2), (ushort)0xF01E);
    BitConverter.TryWriteBytes(dataStream.AsSpan(blipOffset + 4), (uint)21);
    // UID 16 zeros @ [12..27]
    dataStream[blipOffset + 8 + 16] = 0xFF;  // tag (@28)
    dataStream[blipOffset + 8 + 17] = 0x99;
    dataStream[blipOffset + 8 + 18] = 0x88;
    dataStream[blipOffset + 8 + 19] = 0x77;
    dataStream[blipOffset + 8 + 20] = 0x66;

    // Table stream — CLX + BTE + DggContainer/BStoreContainer/FBSE × 1 (foDelay → blipOffset).
    var tblMs = new MemoryStream();
    Span<byte> b4 = stackalloc byte[4];
    Span<byte> b2 = stackalloc byte[2];
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

    // DggContainer (8+52) > BStoreContainer (8+44) > FBSE (8+36, foDelay→blipOffset).
    int dggStart = (int)tblMs.Position;
    // DggContainer header
    BitConverter.TryWriteBytes(b2, (ushort)0x000F); tblMs.Write(b2);
    BitConverter.TryWriteBytes(b2, (ushort)0xF000); tblMs.Write(b2);
    BitConverter.TryWriteBytes(b4, (uint)52);       tblMs.Write(b4);
    // BStoreContainer header (cBlips=1)
    BitConverter.TryWriteBytes(b2, (ushort)0x001F); tblMs.Write(b2);
    BitConverter.TryWriteBytes(b2, (ushort)0xF001); tblMs.Write(b2);
    BitConverter.TryWriteBytes(b4, (uint)44);       tblMs.Write(b4);
    // FBSE header (recVer=2, recInst=msoblipPNG=6, recLen=36)
    BitConverter.TryWriteBytes(b2, (ushort)0x0062); tblMs.Write(b2);
    BitConverter.TryWriteBytes(b2, (ushort)0xF007); tblMs.Write(b2);
    BitConverter.TryWriteBytes(b4, (uint)36);       tblMs.Write(b4);
    // FBSE body (36 byte):
    //   bytes [0..1] btWin32/btMacOS, [2..17] rgbUid, [18..19] tag, [20..23] size,
    //   [24..27] cRef, [28..31] foDelay, [32..35] unused1/cbName/unused2/unused3.
    var fbseBody = new byte[36];
    BitConverter.TryWriteBytes(fbseBody.AsSpan(28), (uint)blipOffset);  // foDelay = 4
    tblMs.Write(fbseBody, 0, 36);

    int dggLen = (int)tblMs.Position - dggStart;
    var tblBytes = tblMs.ToArray();

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
    BitConverter.TryWriteBytes(wd.AsSpan(0x0312), (uint)dggStart);
    BitConverter.TryWriteBytes(wd.AsSpan(0x0316), (uint)dggLen);

    var tmp = Path.Combine(Path.GetTempPath(), $"polydonky-smoke-{Guid.NewGuid():N}.doc");
    try
    {
        using (var root = OpenMcdf.RootStorage.Create(tmp))
        {
            using (var s = root.CreateStream("WordDocument")) s.Write(wd);
            using (var s = root.CreateStream("0Table"))       s.Write(tblBytes);
            using (var s = root.CreateStream("Data"))         s.Write(dataStream);
        }
        using var fs = File.OpenRead(tmp);
        var reader = new PolyDonky.Convert.Doc.DocBinaryReader();
        reader.Read(fs);

        SmokeHarness.Equal(1, reader.BStoreImages.Count,
            $"BStoreImages.Count = 1 (got {reader.BStoreImages.Count})");
        SmokeHarness.Equal("image/png", reader.BStoreImages[0].MediaType, "MediaType = image/png");
        SmokeHarness.Equal(4, reader.BStoreImages[0].Data.Length, "Data.Length = 4");
        SmokeHarness.True(
            reader.BStoreImages[0].Data[0] == 0x99 && reader.BStoreImages[0].Data[1] == 0x88 &&
            reader.BStoreImages[0].Data[2] == 0x77 && reader.BStoreImages[0].Data[3] == 0x66,
            "Data bytes = 99 88 77 66 (Data stream foDelay path)");
    }
    finally { try { File.Delete(tmp); } catch { } }
}

// Phase 3f-4 — PlcSpaMom 의 FSPA entries 파싱 검증.
//   Table stream 에 PlcSpaMom 박는다 — aCP[3] + aSpa[2], 2 개 floating shape anchor.
static void DocBinaryPhase3f4PlcSpaMom()
{
    const string text = "Body\r";
    int ccp = text.Length;
    var textBytes = Encoding.Unicode.GetBytes(text);
    int fcText = 0x200;
    int pnPapx = 4, pnChpx = 5;
    int fcPapxFkp = pnPapx * 512;
    int fcChpxFkp = pnChpx * 512;
    int wdSize = fcChpxFkp + 512;
    var wd = new byte[wdSize];
    Buffer.BlockCopy(textBytes, 0, wd, fcText, textBytes.Length);

    int cpara = 1;
    BitConverter.TryWriteBytes(wd.AsSpan(fcPapxFkp + 0), (int)0x200);
    BitConverter.TryWriteBytes(wd.AsSpan(fcPapxFkp + 4), (int)(0x200 + ccp * 2));
    int papx0Off = 64;
    wd[fcPapxFkp + 8 + 0 * 13] = (byte)(papx0Off / 2);
    int p = fcPapxFkp + papx0Off;
    wd[p++] = 0; wd[p++] = 1; BitConverter.TryWriteBytes(wd.AsSpan(p), (ushort)0);
    wd[fcPapxFkp + 511] = (byte)cpara;
    int crun = 1;
    BitConverter.TryWriteBytes(wd.AsSpan(fcChpxFkp + 0), (int)0x200);
    BitConverter.TryWriteBytes(wd.AsSpan(fcChpxFkp + 4), (int)(0x200 + ccp * 2));
    int chpx0Off = 64;
    wd[fcChpxFkp + 4 * (crun + 1) + 0] = (byte)(chpx0Off / 2);
    wd[fcChpxFkp + chpx0Off] = 0;
    wd[fcChpxFkp + 511] = (byte)crun;

    var tblMs = new MemoryStream();
    Span<byte> b4 = stackalloc byte[4];
    Span<byte> b2 = stackalloc byte[2];
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

    // PlcSpaMom — 2 shapes:
    //   shape 0: cp=2, spid=1024, rect=(100, 200, 1540, 1640)
    //   shape 1: cp=3, spid=2048, rect=(500, 600, 1940, 2040)
    //   plus aCP[2] = 4 (end CP, conventionally).
    int spaStart = (int)tblMs.Position;
    BitConverter.TryWriteBytes(b4, (int)2); tblMs.Write(b4);     // aCP[0] = 2
    BitConverter.TryWriteBytes(b4, (int)3); tblMs.Write(b4);     // aCP[1] = 3
    BitConverter.TryWriteBytes(b4, (int)4); tblMs.Write(b4);     // aCP[2] = 4 (end)
    // FSPA #1
    BitConverter.TryWriteBytes(b4, (int)1024); tblMs.Write(b4);  // spid
    BitConverter.TryWriteBytes(b4, (int)100);  tblMs.Write(b4);  // xaLeft
    BitConverter.TryWriteBytes(b4, (int)200);  tblMs.Write(b4);  // yaTop
    BitConverter.TryWriteBytes(b4, (int)1540); tblMs.Write(b4);  // xaRight
    BitConverter.TryWriteBytes(b4, (int)1640); tblMs.Write(b4);  // yaBottom
    tblMs.Write(new byte[6]);                                     // flags
    // FSPA #2
    BitConverter.TryWriteBytes(b4, (int)2048); tblMs.Write(b4);
    BitConverter.TryWriteBytes(b4, (int)500);  tblMs.Write(b4);
    BitConverter.TryWriteBytes(b4, (int)600);  tblMs.Write(b4);
    BitConverter.TryWriteBytes(b4, (int)1940); tblMs.Write(b4);
    BitConverter.TryWriteBytes(b4, (int)2040); tblMs.Write(b4);
    tblMs.Write(new byte[6]);
    int spaLen = (int)tblMs.Position - spaStart;

    var tblBytes = tblMs.ToArray();

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
    BitConverter.TryWriteBytes(wd.AsSpan(0x011A), (uint)spaStart);  // fcPlcSpaMom
    BitConverter.TryWriteBytes(wd.AsSpan(0x011E), (uint)spaLen);    // lcbPlcSpaMom

    var tmp = Path.Combine(Path.GetTempPath(), $"polydonky-smoke-{Guid.NewGuid():N}.doc");
    try
    {
        using (var root = OpenMcdf.RootStorage.Create(tmp))
        {
            using (var s = root.CreateStream("WordDocument")) s.Write(wd);
            using (var s = root.CreateStream("0Table"))       s.Write(tblBytes);
        }
        using var fs = File.OpenRead(tmp);
        var reader = new PolyDonky.Convert.Doc.DocBinaryReader();
        reader.Read(fs);

        SmokeHarness.Equal(2, reader.FspaEntries.Count, $"FspaEntries.Count = 2 (got {reader.FspaEntries.Count})");

        SmokeHarness.Equal(2,    reader.FspaEntries[0].Cp,              "FSPA[0].Cp = 2");
        SmokeHarness.Equal(1024, reader.FspaEntries[0].Spid,            "FSPA[0].Spid = 1024");
        SmokeHarness.Equal(100,  reader.FspaEntries[0].XaLeftTwips,     "FSPA[0].XaLeft = 100");
        SmokeHarness.Equal(200,  reader.FspaEntries[0].YaTopTwips,      "FSPA[0].YaTop = 200");
        SmokeHarness.Equal(1540, reader.FspaEntries[0].XaRightTwips,    "FSPA[0].XaRight = 1540");
        SmokeHarness.Equal(1640, reader.FspaEntries[0].YaBottomTwips,   "FSPA[0].YaBottom = 1640");

        SmokeHarness.Equal(3,    reader.FspaEntries[1].Cp,              "FSPA[1].Cp = 3");
        SmokeHarness.Equal(2048, reader.FspaEntries[1].Spid,            "FSPA[1].Spid = 2048");
        SmokeHarness.Equal(500,  reader.FspaEntries[1].XaLeftTwips,     "FSPA[1].XaLeft = 500");
        SmokeHarness.Equal(600,  reader.FspaEntries[1].YaTopTwips,      "FSPA[1].YaTop = 600");
        SmokeHarness.Equal(1940, reader.FspaEntries[1].XaRightTwips,    "FSPA[1].XaRight = 1940");
        SmokeHarness.Equal(2040, reader.FspaEntries[1].YaBottomTwips,   "FSPA[1].YaBottom = 2040");
    }
    finally { try { File.Delete(tmp); } catch { } }
}

// Phase 3f-5 — DggContainer 의 OfficeArtSpContainer 들에서 spid → pib (BStore index) 맵 추출.
//   2 개 SpContainer, 각각 Sp atom (spid) + OPT atom (pib property).
//   shape A: spid=100 → pib=1, shape B: spid=200 → pib=2.
static void DocBinaryPhase3f5SpidPibMap()
{
    const string text = "Body\r";
    int ccp = text.Length;
    var textBytes = Encoding.Unicode.GetBytes(text);
    int fcText = 0x200;
    int pnPapx = 4, pnChpx = 5;
    int fcPapxFkp = pnPapx * 512;
    int fcChpxFkp = pnChpx * 512;
    int wdSize = fcChpxFkp + 512;
    var wd = new byte[wdSize];
    Buffer.BlockCopy(textBytes, 0, wd, fcText, textBytes.Length);

    int cpara = 1;
    BitConverter.TryWriteBytes(wd.AsSpan(fcPapxFkp + 0), (int)0x200);
    BitConverter.TryWriteBytes(wd.AsSpan(fcPapxFkp + 4), (int)(0x200 + ccp * 2));
    int papx0Off = 64;
    wd[fcPapxFkp + 8 + 0 * 13] = (byte)(papx0Off / 2);
    int p = fcPapxFkp + papx0Off;
    wd[p++] = 0; wd[p++] = 1; BitConverter.TryWriteBytes(wd.AsSpan(p), (ushort)0);
    wd[fcPapxFkp + 511] = (byte)cpara;
    int crun = 1;
    BitConverter.TryWriteBytes(wd.AsSpan(fcChpxFkp + 0), (int)0x200);
    BitConverter.TryWriteBytes(wd.AsSpan(fcChpxFkp + 4), (int)(0x200 + ccp * 2));
    int chpx0Off = 64;
    wd[fcChpxFkp + 4 * (crun + 1) + 0] = (byte)(chpx0Off / 2);
    wd[fcChpxFkp + chpx0Off] = 0;
    wd[fcChpxFkp + 511] = (byte)crun;

    var tblMs = new MemoryStream();
    Span<byte> b4 = stackalloc byte[4];
    Span<byte> b2 = stackalloc byte[2];
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

    // DggContainer 합성 — SpContainer 2 개 (BStoreContainer 는 생략, 맵 추출만 검증).
    //   각 SpContainer (38 byte) = header(8) + Sp(16) + OPT(14)
    //     Sp atom (0xF009): header(8) + body(8: spid + flags)
    //     OPT atom (0xF00B, recInst=1, recVer=3): header(8) + property(6: opid + op)
    //   DggContainer body = 2 * 38 = 76 bytes. total = 84.
    int dggStart = (int)tblMs.Position;
    // DggContainer header
    BitConverter.TryWriteBytes(b2, (ushort)0x000F); tblMs.Write(b2);
    BitConverter.TryWriteBytes(b2, (ushort)0xF000); tblMs.Write(b2);
    BitConverter.TryWriteBytes(b4, (uint)76);       tblMs.Write(b4);

    static byte[] BuildSpContainer(int spid, int pib)
    {
        // 38 byte: SpContainer header(8) + Sp(16) + OPT(14)
        var buf = new byte[38];
        // SpContainer header (recVer=0xF, recInst=0, recLen=30)
        BitConverter.TryWriteBytes(buf.AsSpan(0),  (ushort)0x000F);
        BitConverter.TryWriteBytes(buf.AsSpan(2),  (ushort)0xF004);
        BitConverter.TryWriteBytes(buf.AsSpan(4),  (uint)30);
        // Sp atom (recVer=2, recInst=0, recType=0xF009, recLen=8) + body (spid + flags)
        BitConverter.TryWriteBytes(buf.AsSpan(8),  (ushort)0x0002);
        BitConverter.TryWriteBytes(buf.AsSpan(10), (ushort)0xF009);
        BitConverter.TryWriteBytes(buf.AsSpan(12), (uint)8);
        BitConverter.TryWriteBytes(buf.AsSpan(16), (int)spid);
        BitConverter.TryWriteBytes(buf.AsSpan(20), (int)0);
        // OPT atom (recVer=3, recInst=1, recType=0xF00B, recLen=6) + 1 property
        BitConverter.TryWriteBytes(buf.AsSpan(24), (ushort)0x0013);
        BitConverter.TryWriteBytes(buf.AsSpan(26), (ushort)0xF00B);
        BitConverter.TryWriteBytes(buf.AsSpan(28), (uint)6);
        BitConverter.TryWriteBytes(buf.AsSpan(32), (ushort)0x4104);  // propId=260, fBid=1
        BitConverter.TryWriteBytes(buf.AsSpan(34), (uint)pib);
        return buf;
    }

    var spA = BuildSpContainer(100, 1);
    var spB = BuildSpContainer(200, 2);
    tblMs.Write(spA, 0, spA.Length);
    tblMs.Write(spB, 0, spB.Length);

    int dggLen = (int)tblMs.Position - dggStart;
    var tblBytes = tblMs.ToArray();

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
    BitConverter.TryWriteBytes(wd.AsSpan(0x0312), (uint)dggStart);
    BitConverter.TryWriteBytes(wd.AsSpan(0x0316), (uint)dggLen);

    var tmp = Path.Combine(Path.GetTempPath(), $"polydonky-smoke-{Guid.NewGuid():N}.doc");
    try
    {
        using (var root = OpenMcdf.RootStorage.Create(tmp))
        {
            using (var s = root.CreateStream("WordDocument")) s.Write(wd);
            using (var s = root.CreateStream("0Table"))       s.Write(tblBytes);
        }
        using var fs = File.OpenRead(tmp);
        var reader = new PolyDonky.Convert.Doc.DocBinaryReader();
        reader.Read(fs);

        SmokeHarness.Equal(2, reader.ShapeImageIndex.Count,
            $"ShapeImageIndex.Count = 2 (got {reader.ShapeImageIndex.Count})");
        SmokeHarness.True(reader.ShapeImageIndex.TryGetValue(100, out int pibA),
            "spid 100 in map");
        SmokeHarness.Equal(1, pibA, "spid=100 → pib=1");
        SmokeHarness.True(reader.ShapeImageIndex.TryGetValue(200, out int pibB),
            "spid 200 in map");
        SmokeHarness.Equal(2, pibB, "spid=200 → pib=2");
    }
    finally { try { File.Delete(tmp); } catch { } }
}

// Phase 3f-6 — FspaEntries + ShapeImageIndex + BStoreImages 결합 → floating ImageBlock 생성.
//   shape A (spid=100, pib=1): anchor (1440, 720, 2880, 2160) twips → (25.4, 12.7) mm @ 25.4×25.4
//     BLIP image = AA BB CC DD
//   shape B (spid=200, pib=2): anchor (2880, 2880, 4320, 4320) twips → (50.8, 50.8) mm @ 25.4×25.4
//     BLIP image = 11 22 33 44
//   결과: doc.Sections[0].Blocks 의 마지막 두 항목이 floating ImageBlock.
static void DocBinaryPhase3f6FloatingShapeImages()
{
    const string text = "Body\r";
    int ccp = text.Length;
    var textBytes = Encoding.Unicode.GetBytes(text);
    int fcText = 0x200;
    int pnPapx = 4, pnChpx = 5;
    int fcPapxFkp = pnPapx * 512;
    int fcChpxFkp = pnChpx * 512;
    int wdSize = fcChpxFkp + 512;
    var wd = new byte[wdSize];
    Buffer.BlockCopy(textBytes, 0, wd, fcText, textBytes.Length);

    int cpara = 1;
    BitConverter.TryWriteBytes(wd.AsSpan(fcPapxFkp + 0), (int)0x200);
    BitConverter.TryWriteBytes(wd.AsSpan(fcPapxFkp + 4), (int)(0x200 + ccp * 2));
    int papx0Off = 64;
    wd[fcPapxFkp + 8 + 0 * 13] = (byte)(papx0Off / 2);
    int p = fcPapxFkp + papx0Off;
    wd[p++] = 0; wd[p++] = 1; BitConverter.TryWriteBytes(wd.AsSpan(p), (ushort)0);
    wd[fcPapxFkp + 511] = (byte)cpara;
    int crun = 1;
    BitConverter.TryWriteBytes(wd.AsSpan(fcChpxFkp + 0), (int)0x200);
    BitConverter.TryWriteBytes(wd.AsSpan(fcChpxFkp + 4), (int)(0x200 + ccp * 2));
    int chpx0Off = 64;
    wd[fcChpxFkp + 4 * (crun + 1) + 0] = (byte)(chpx0Off / 2);
    wd[fcChpxFkp + chpx0Off] = 0;
    wd[fcChpxFkp + 511] = (byte)crun;

    var tblMs = new MemoryStream();
    Span<byte> b4 = stackalloc byte[4];
    Span<byte> b2 = stackalloc byte[2];
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

    // PlcSpaMom — 2 shapes (CP 1 / CP 2, spid 100/200, 사각형 좌표).
    int spaStart = (int)tblMs.Position;
    BitConverter.TryWriteBytes(b4, (int)1); tblMs.Write(b4);
    BitConverter.TryWriteBytes(b4, (int)2); tblMs.Write(b4);
    BitConverter.TryWriteBytes(b4, (int)3); tblMs.Write(b4);
    // FSPA #1: spid=100, anchor (1440, 720, 2880, 2160) — 25.4 mm @ 12.7 mm, 25.4 × 25.4 mm
    BitConverter.TryWriteBytes(b4, (int)100);  tblMs.Write(b4);
    BitConverter.TryWriteBytes(b4, (int)1440); tblMs.Write(b4);
    BitConverter.TryWriteBytes(b4, (int)720);  tblMs.Write(b4);
    BitConverter.TryWriteBytes(b4, (int)2880); tblMs.Write(b4);
    BitConverter.TryWriteBytes(b4, (int)2160); tblMs.Write(b4);
    tblMs.Write(new byte[6]);
    // FSPA #2: spid=200, anchor (2880, 2880, 4320, 4320)
    BitConverter.TryWriteBytes(b4, (int)200);  tblMs.Write(b4);
    BitConverter.TryWriteBytes(b4, (int)2880); tblMs.Write(b4);
    BitConverter.TryWriteBytes(b4, (int)2880); tblMs.Write(b4);
    BitConverter.TryWriteBytes(b4, (int)4320); tblMs.Write(b4);
    BitConverter.TryWriteBytes(b4, (int)4320); tblMs.Write(b4);
    tblMs.Write(new byte[6]);
    int spaLen = (int)tblMs.Position - spaStart;

    // DggContainer — BStoreContainer (2 BLIP_PNG) + SpContainer × 2 (spid 100/200 → pib 1/2).
    // BStoreContainer 부분 = header(8) + 2 FBSE(각 73) = 154
    // SpContainer 부분    = 2 × 38 = 76
    // DggContainer body   = 8(BStoreContainer hdr) + 146(FBSE) + 76 = 230. total = 238.
    int dggStart = (int)tblMs.Position;
    BitConverter.TryWriteBytes(b2, (ushort)0x000F); tblMs.Write(b2);
    BitConverter.TryWriteBytes(b2, (ushort)0xF000); tblMs.Write(b2);
    BitConverter.TryWriteBytes(b4, (uint)230);      tblMs.Write(b4);

    // BStoreContainer (cBlips=2, body = 2 × 73 = 146)
    BitConverter.TryWriteBytes(b2, (ushort)0x002F); tblMs.Write(b2);
    BitConverter.TryWriteBytes(b2, (ushort)0xF001); tblMs.Write(b2);
    BitConverter.TryWriteBytes(b4, (uint)146);      tblMs.Write(b4);

    static byte[] BuildFbse(byte[] img4)
    {
        var f = new byte[73];
        BitConverter.TryWriteBytes(f.AsSpan(0),  (ushort)0x0062);
        BitConverter.TryWriteBytes(f.AsSpan(2),  (ushort)0xF007);
        BitConverter.TryWriteBytes(f.AsSpan(4),  (uint)65);
        // body[44..51] embedded BLIP_PNG header
        BitConverter.TryWriteBytes(f.AsSpan(44), (ushort)0x6E00);
        BitConverter.TryWriteBytes(f.AsSpan(46), (ushort)0xF01E);
        BitConverter.TryWriteBytes(f.AsSpan(48), (uint)21);
        f[68] = 0xFF;
        Buffer.BlockCopy(img4, 0, f, 69, 4);
        return f;
    }
    var fbse1 = BuildFbse(new byte[] { 0xAA, 0xBB, 0xCC, 0xDD });
    var fbse2 = BuildFbse(new byte[] { 0x11, 0x22, 0x33, 0x44 });
    tblMs.Write(fbse1, 0, fbse1.Length);
    tblMs.Write(fbse2, 0, fbse2.Length);

    // SpContainer × 2.
    static byte[] BuildSp(int spid, int pib)
    {
        var buf = new byte[38];
        BitConverter.TryWriteBytes(buf.AsSpan(0),  (ushort)0x000F);
        BitConverter.TryWriteBytes(buf.AsSpan(2),  (ushort)0xF004);
        BitConverter.TryWriteBytes(buf.AsSpan(4),  (uint)30);
        BitConverter.TryWriteBytes(buf.AsSpan(8),  (ushort)0x0002);
        BitConverter.TryWriteBytes(buf.AsSpan(10), (ushort)0xF009);
        BitConverter.TryWriteBytes(buf.AsSpan(12), (uint)8);
        BitConverter.TryWriteBytes(buf.AsSpan(16), (int)spid);
        BitConverter.TryWriteBytes(buf.AsSpan(20), (int)0);
        BitConverter.TryWriteBytes(buf.AsSpan(24), (ushort)0x0013);
        BitConverter.TryWriteBytes(buf.AsSpan(26), (ushort)0xF00B);
        BitConverter.TryWriteBytes(buf.AsSpan(28), (uint)6);
        BitConverter.TryWriteBytes(buf.AsSpan(32), (ushort)0x4104);
        BitConverter.TryWriteBytes(buf.AsSpan(34), (uint)pib);
        return buf;
    }
    var spA = BuildSp(100, 1);
    var spB = BuildSp(200, 2);
    tblMs.Write(spA, 0, spA.Length);
    tblMs.Write(spB, 0, spB.Length);
    int dggLen = (int)tblMs.Position - dggStart;
    var tblBytes = tblMs.ToArray();

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
    BitConverter.TryWriteBytes(wd.AsSpan(0x011A), (uint)spaStart);
    BitConverter.TryWriteBytes(wd.AsSpan(0x011E), (uint)spaLen);
    BitConverter.TryWriteBytes(wd.AsSpan(0x0312), (uint)dggStart);
    BitConverter.TryWriteBytes(wd.AsSpan(0x0316), (uint)dggLen);

    var tmp = Path.Combine(Path.GetTempPath(), $"polydonky-smoke-{Guid.NewGuid():N}.doc");
    try
    {
        using (var root = OpenMcdf.RootStorage.Create(tmp))
        {
            using (var s = root.CreateStream("WordDocument")) s.Write(wd);
            using (var s = root.CreateStream("0Table"))       s.Write(tblBytes);
        }
        using var fs = File.OpenRead(tmp);
        var doc = new PolyDonky.Convert.Doc.DocBinaryReader().Read(fs);
        var blocks = doc.Sections[0].Blocks;

        // 본문 "Body" 단락 + floating ImageBlock 2 개.
        var imgs = blocks.OfType<ImageBlock>().ToList();
        SmokeHarness.Equal(2, imgs.Count, $"floating ImageBlock 수 = 2 (got {imgs.Count})");

        var iA = imgs[0];
        SmokeHarness.Equal("image/png", iA.MediaType, "ImageBlock #1 MediaType");
        SmokeHarness.True(iA.Data.Length == 4 &&
            iA.Data[0] == 0xAA && iA.Data[1] == 0xBB && iA.Data[2] == 0xCC && iA.Data[3] == 0xDD,
            "ImageBlock #1 bytes = AA BB CC DD");
        SmokeHarness.Equal(ImageWrapMode.InFrontOfText, iA.WrapMode, "ImageBlock #1 WrapMode");
        // 1440 twips / 56.692 ≈ 25.40 mm
        SmokeHarness.True(Math.Abs(iA.OverlayXMm - 25.4) < 0.05,
            $"ImageBlock #1 OverlayXMm ≈ 25.4 (got {iA.OverlayXMm:F2})");
        SmokeHarness.True(Math.Abs(iA.OverlayYMm - 12.7) < 0.05,
            $"ImageBlock #1 OverlayYMm ≈ 12.7 (got {iA.OverlayYMm:F2})");
        SmokeHarness.True(Math.Abs(iA.WidthMm - 25.4) < 0.05,
            $"ImageBlock #1 WidthMm ≈ 25.4 (got {iA.WidthMm:F2})");
        SmokeHarness.True(Math.Abs(iA.HeightMm - 25.4) < 0.05,
            $"ImageBlock #1 HeightMm ≈ 25.4 (got {iA.HeightMm:F2})");

        var iB = imgs[1];
        SmokeHarness.True(iB.Data.Length == 4 &&
            iB.Data[0] == 0x11 && iB.Data[1] == 0x22 && iB.Data[2] == 0x33 && iB.Data[3] == 0x44,
            "ImageBlock #2 bytes = 11 22 33 44");
        SmokeHarness.True(Math.Abs(iB.OverlayXMm - 50.8) < 0.05,
            $"ImageBlock #2 OverlayXMm ≈ 50.8 (got {iB.OverlayXMm:F2})");
    }
    finally { try { File.Delete(tmp); } catch { } }
}

// Phase 3g-2 — 본문 0x02 (footnote/endnote ref) → Run.FootnoteId / EndnoteId 자동 연결.
//   main "A<0x02>B<0x02>C\r" — CP 1 의 0x02 는 footnote ref, CP 3 의 0x02 는 endnote ref.
//   PlcfFndRef aCP = [1, 6] (sentinel=ccpText). PlcfendRef aCP = [3, 6].
//   결과: Paragraph 의 Runs = [A, "" FootnoteId=fn1, B, "" EndnoteId=en1, C].
static void DocBinaryPhase3g2FootnoteRefAutoLink()
{
    // main(6) + footnote(15) + header(0) + comment(0) + endnote(14) = 35 chars
    const string text = "ABC\r \rFootnoteText\r \rEndnoteText\r";
    int ccpMain = 6;
    int ccpFtn  = 15;
    int ccpEdn  = 14;
    int totalCcp = ccpMain + ccpFtn + ccpEdn;  // = 35
    SmokeHarness.Equal(totalCcp, text.Length, "text 길이 검증");

    var textBytes = Encoding.Unicode.GetBytes(text);
    int fcText = 0x200;
    int pnPapx = 4, pnChpx = 5;
    int fcPapxFkp = pnPapx * 512;
    int fcChpxFkp = pnChpx * 512;
    int wdSize = fcChpxFkp + 512;
    var wd = new byte[wdSize];
    Buffer.BlockCopy(textBytes, 0, wd, fcText, textBytes.Length);

    // PAPX FKP — 1 단락 (main 영역만).
    int cpara = 1;
    int paraEndFc = fcText + ccpMain * 2;
    BitConverter.TryWriteBytes(wd.AsSpan(fcPapxFkp + 0), (int)fcText);
    BitConverter.TryWriteBytes(wd.AsSpan(fcPapxFkp + 4), (int)paraEndFc);
    int papx0Off = 64;
    wd[fcPapxFkp + 8 + 0 * 13] = (byte)(papx0Off / 2);
    int p = fcPapxFkp + papx0Off;
    wd[p++] = 0; wd[p++] = 1; BitConverter.TryWriteBytes(wd.AsSpan(p), (ushort)0);
    wd[fcPapxFkp + 511] = (byte)cpara;
    int crun = 1;
    BitConverter.TryWriteBytes(wd.AsSpan(fcChpxFkp + 0), (int)fcText);
    BitConverter.TryWriteBytes(wd.AsSpan(fcChpxFkp + 4), (int)paraEndFc);
    int chpx0Off = 64;
    wd[fcChpxFkp + 4 * (crun + 1) + 0] = (byte)(chpx0Off / 2);
    wd[fcChpxFkp + chpx0Off] = 0;
    wd[fcChpxFkp + 511] = (byte)crun;

    var tblMs = new MemoryStream();
    Span<byte> b4 = stackalloc byte[4];
    Span<byte> b2 = stackalloc byte[2];
    tblMs.WriteByte(0x02);
    BitConverter.TryWriteBytes(b4, (uint)16); tblMs.Write(b4);
    BitConverter.TryWriteBytes(b4, (uint)0);        tblMs.Write(b4);
    BitConverter.TryWriteBytes(b4, (uint)totalCcp); tblMs.Write(b4);
    tblMs.WriteByte(0); tblMs.WriteByte(0);
    BitConverter.TryWriteBytes(b4, (uint)fcText); tblMs.Write(b4);
    tblMs.WriteByte(0); tblMs.WriteByte(0);
    int clxEnd = (int)tblMs.Position;
    int papxBteStart = clxEnd;
    BitConverter.TryWriteBytes(b4, (int)fcText);      tblMs.Write(b4);
    BitConverter.TryWriteBytes(b4, (int)(fcText + totalCcp * 2 + 0x100)); tblMs.Write(b4);
    BitConverter.TryWriteBytes(b4, (int)pnPapx);      tblMs.Write(b4);
    int papxBteLen = (int)tblMs.Position - papxBteStart;
    int chpxBteStart = (int)tblMs.Position;
    BitConverter.TryWriteBytes(b4, (int)fcText);      tblMs.Write(b4);
    BitConverter.TryWriteBytes(b4, (int)(fcText + totalCcp * 2 + 0x100)); tblMs.Write(b4);
    BitConverter.TryWriteBytes(b4, (int)pnChpx);      tblMs.Write(b4);
    int chpxBteLen = (int)tblMs.Position - chpxBteStart;

    // PlcfFndTxt: aCP = [0, 2, 15] (separator + 1 footnote). lcb = 12.
    int fndTxtStart = (int)tblMs.Position;
    BitConverter.TryWriteBytes(b4, (int)0);  tblMs.Write(b4);
    BitConverter.TryWriteBytes(b4, (int)2);  tblMs.Write(b4);
    BitConverter.TryWriteBytes(b4, (int)15); tblMs.Write(b4);
    int fndTxtLen = (int)tblMs.Position - fndTxtStart;

    // PlcfendTxt: aCP = [0, 2, 14] (separator + 1 endnote). lcb = 12.
    int ednTxtStart = (int)tblMs.Position;
    BitConverter.TryWriteBytes(b4, (int)0);  tblMs.Write(b4);
    BitConverter.TryWriteBytes(b4, (int)2);  tblMs.Write(b4);
    BitConverter.TryWriteBytes(b4, (int)14); tblMs.Write(b4);
    int ednTxtLen = (int)tblMs.Position - ednTxtStart;

    // PlcfFndRef: aCP = [1 (ref CP in main), 6 (sentinel = ccpText)].
    //   FRD = 2 byte placeholder. element = 6 byte. lcb = 4*2 + 2*1 = 10.
    int fndRefStart = (int)tblMs.Position;
    BitConverter.TryWriteBytes(b4, (int)1);  tblMs.Write(b4);   // aCP[0]
    BitConverter.TryWriteBytes(b4, (int)6);  tblMs.Write(b4);   // aCP[1]
    BitConverter.TryWriteBytes(b2, (ushort)0); tblMs.Write(b2); // FRD[0]
    int fndRefLen = (int)tblMs.Position - fndRefStart;

    // PlcfendRef: aCP = [3, 6].
    int ednRefStart = (int)tblMs.Position;
    BitConverter.TryWriteBytes(b4, (int)3);  tblMs.Write(b4);
    BitConverter.TryWriteBytes(b4, (int)6);  tblMs.Write(b4);
    BitConverter.TryWriteBytes(b2, (ushort)0); tblMs.Write(b2);
    int ednRefLen = (int)tblMs.Position - ednRefStart;

    var tblBytes = tblMs.ToArray();

    BitConverter.TryWriteBytes(wd.AsSpan(0x00),   (ushort)0xA5EC);
    BitConverter.TryWriteBytes(wd.AsSpan(0x02),   (ushort)193);
    BitConverter.TryWriteBytes(wd.AsSpan(0x0A),   (ushort)0x0000);
    BitConverter.TryWriteBytes(wd.AsSpan(0x18),   (uint)fcText);
    BitConverter.TryWriteBytes(wd.AsSpan(0x4C),   (uint)ccpMain);
    BitConverter.TryWriteBytes(wd.AsSpan(0x50),   (uint)ccpFtn);
    BitConverter.TryWriteBytes(wd.AsSpan(0x54),   (uint)0);
    BitConverter.TryWriteBytes(wd.AsSpan(0x5C),   (uint)0);
    BitConverter.TryWriteBytes(wd.AsSpan(0x60),   (uint)ccpEdn);
    BitConverter.TryWriteBytes(wd.AsSpan(0xAA),   (uint)fndRefStart);  // fcPlcffndRef
    BitConverter.TryWriteBytes(wd.AsSpan(0xAE),   (uint)fndRefLen);    // lcbPlcffndRef
    BitConverter.TryWriteBytes(wd.AsSpan(0xB2),   (uint)fndTxtStart);
    BitConverter.TryWriteBytes(wd.AsSpan(0xB6),   (uint)fndTxtLen);
    BitConverter.TryWriteBytes(wd.AsSpan(0x14A),  (uint)ednRefStart);  // fcPlcfendRef
    BitConverter.TryWriteBytes(wd.AsSpan(0x14E),  (uint)ednRefLen);    // lcbPlcfendRef
    BitConverter.TryWriteBytes(wd.AsSpan(0x152),  (uint)ednTxtStart);
    BitConverter.TryWriteBytes(wd.AsSpan(0x156),  (uint)ednTxtLen);
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
            using (var s = root.CreateStream("0Table"))       s.Write(tblBytes);
        }
        using var fs = File.OpenRead(tmp);
        var doc = new PolyDonky.Convert.Doc.DocBinaryReader().Read(fs);

        // Footnote / Endnote 영역도 확인.
        SmokeHarness.Equal(1, doc.Footnotes.Count, "Footnotes.Count");
        SmokeHarness.Equal("fn1", doc.Footnotes[0].Id, "Footnote Id");
        SmokeHarness.Equal(1, doc.Endnotes.Count, "Endnotes.Count");
        SmokeHarness.Equal("en1", doc.Endnotes[0].Id, "Endnote Id");

        var para = (Paragraph)doc.Sections[0].Blocks[0];
        SmokeHarness.Equal("ABC", para.GetPlainText(), "plain text = ABC (ref marker 는 Text 없음)");
        SmokeHarness.Equal(5, para.Runs.Count, $"Runs.Count = 5 (got {para.Runs.Count})");

        SmokeHarness.Equal("A", para.Runs[0].Text, "Run[0].Text = A");
        SmokeHarness.True(para.Runs[0].FootnoteId is null && para.Runs[0].EndnoteId is null,
            "Run[0] 일반 텍스트");

        SmokeHarness.Equal("",      para.Runs[1].Text, "Run[1].Text 빈문자");
        SmokeHarness.Equal("fn1",   para.Runs[1].FootnoteId, "Run[1].FootnoteId = fn1");

        SmokeHarness.Equal("B", para.Runs[2].Text, "Run[2].Text = B");

        SmokeHarness.Equal("",      para.Runs[3].Text, "Run[3].Text 빈문자");
        SmokeHarness.Equal("en1",   para.Runs[3].EndnoteId, "Run[3].EndnoteId = en1");

        SmokeHarness.Equal("C", para.Runs[4].Text, "Run[4].Text = C");
    }
    finally { try { File.Delete(tmp); } catch { } }
}

// Phase 3h — 변경추적 (revision marks). sprmCFRMarkIns (0x0801) / sprmCFRMarkDel (0x0807) 가
// CHPX 에 있으면 해당 char 범위가 insert / delete revision mark.
//   본문 "ABCDEF\r" — 3 CHPX run:
//     run 0 (AB): 일반
//     run 1 (CD): sprmCFRMarkIns
//     run 2 (EF\r): sprmCFRMarkDel
//   결과 paragraph runs = [AB, CD (Inserted), EF (Deleted)].
static void DocBinaryPhase3hRevisionMarks()
{
    const string text = "ABCDEF\r";
    int ccp = text.Length;
    var textBytes = Encoding.Unicode.GetBytes(text);
    int fcText = 0x200;
    int pnPapx = 4, pnChpx = 5;
    int fcPapxFkp = pnPapx * 512;
    int fcChpxFkp = pnChpx * 512;
    int wdSize = fcChpxFkp + 512;
    var wd = new byte[wdSize];
    Buffer.BlockCopy(textBytes, 0, wd, fcText, textBytes.Length);

    // PAPX FKP — 1 단락.
    int cpara = 1;
    int paraEndFc = fcText + ccp * 2;
    BitConverter.TryWriteBytes(wd.AsSpan(fcPapxFkp + 0), (int)fcText);
    BitConverter.TryWriteBytes(wd.AsSpan(fcPapxFkp + 4), (int)paraEndFc);
    int papx0Off = 64;
    wd[fcPapxFkp + 8 + 0 * 13] = (byte)(papx0Off / 2);
    int p = fcPapxFkp + papx0Off;
    wd[p++] = 0; wd[p++] = 1; BitConverter.TryWriteBytes(wd.AsSpan(p), (ushort)0);
    wd[fcPapxFkp + 511] = (byte)cpara;

    // CHPX FKP — 3 runs:
    //   run 0: [0x200, 0x204) "AB"   — empty CHPX
    //   run 1: [0x204, 0x208) "CD"   — sprmCFRMarkIns (0x0801) + op=1
    //   run 2: [0x208, 0x20E) "EF\r" — sprmCFRMarkDel (0x0807) + op=1
    int crun = 3;
    BitConverter.TryWriteBytes(wd.AsSpan(fcChpxFkp + 0),  (int)0x200);
    BitConverter.TryWriteBytes(wd.AsSpan(fcChpxFkp + 4),  (int)0x204);
    BitConverter.TryWriteBytes(wd.AsSpan(fcChpxFkp + 8),  (int)0x208);
    BitConverter.TryWriteBytes(wd.AsSpan(fcChpxFkp + 12), (int)0x20E);
    int chpx0Off = 64, chpx1Off = 80, chpx2Off = 96;
    int rgbBase  = 4 * (crun + 1);
    wd[fcChpxFkp + rgbBase + 0] = (byte)(chpx0Off / 2);
    wd[fcChpxFkp + rgbBase + 1] = (byte)(chpx1Off / 2);
    wd[fcChpxFkp + rgbBase + 2] = (byte)(chpx2Off / 2);
    // ChpxInFkp[0]: cb=0
    wd[fcChpxFkp + chpx0Off + 0] = 0;
    // ChpxInFkp[1]: cb=3, sprm=0x0801, op=1
    int c1 = fcChpxFkp + chpx1Off;
    wd[c1++] = 3;
    BitConverter.TryWriteBytes(wd.AsSpan(c1), (ushort)0x0801); c1 += 2;
    wd[c1] = 1;
    // ChpxInFkp[2]: cb=3, sprm=0x0807, op=1
    int c2 = fcChpxFkp + chpx2Off;
    wd[c2++] = 3;
    BitConverter.TryWriteBytes(wd.AsSpan(c2), (ushort)0x0807); c2 += 2;
    wd[c2] = 1;
    wd[fcChpxFkp + 511] = (byte)crun;

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
    BitConverter.TryWriteBytes(b4, (int)0x200);     tblMs.Write(b4);
    BitConverter.TryWriteBytes(b4, (int)paraEndFc); tblMs.Write(b4);
    BitConverter.TryWriteBytes(b4, (int)pnPapx);    tblMs.Write(b4);
    int papxBteLen = (int)tblMs.Position - papxBteStart;
    int chpxBteStart = (int)tblMs.Position;
    BitConverter.TryWriteBytes(b4, (int)0x200);     tblMs.Write(b4);
    BitConverter.TryWriteBytes(b4, (int)paraEndFc); tblMs.Write(b4);
    BitConverter.TryWriteBytes(b4, (int)pnChpx);    tblMs.Write(b4);
    int chpxBteLen = (int)tblMs.Position - chpxBteStart;
    var tblBytes = tblMs.ToArray();

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
            using (var s = root.CreateStream("0Table"))       s.Write(tblBytes);
        }
        using var fs = File.OpenRead(tmp);
        var doc = new PolyDonky.Convert.Doc.DocBinaryReader().Read(fs);
        var para = (Paragraph)doc.Sections[0].Blocks[0];

        SmokeHarness.Equal("ABCDEF", para.GetPlainText(), "plain text = ABCDEF");
        SmokeHarness.Equal(3, para.Runs.Count, $"Runs.Count = 3 (got {para.Runs.Count})");

        SmokeHarness.Equal("AB", para.Runs[0].Text, "Run[0].Text = AB");
        SmokeHarness.True(!para.Runs[0].IsInsertedRevision && !para.Runs[0].IsDeletedRevision,
            "Run[0] 일반");

        SmokeHarness.Equal("CD", para.Runs[1].Text, "Run[1].Text = CD");
        SmokeHarness.True(para.Runs[1].IsInsertedRevision, "Run[1].IsInsertedRevision = true");
        SmokeHarness.True(!para.Runs[1].IsDeletedRevision, "Run[1].IsDeletedRevision = false");

        SmokeHarness.Equal("EF", para.Runs[2].Text, "Run[2].Text = EF");
        SmokeHarness.True(para.Runs[2].IsDeletedRevision, "Run[2].IsDeletedRevision = true");
        SmokeHarness.True(!para.Runs[2].IsInsertedRevision, "Run[2].IsInsertedRevision = false");
    }
    finally { try { File.Delete(tmp); } catch { } }
}

// Phase 3h-2 — revision 작성자 + 타임스탬프. SttbfRMark (FIB 0x142/0x146) 에 작성자 배열,
// sprmCIbstRMark (0x4804) 가 인덱스, sprmCDttmRMark (0x6805) 가 packed DTTM.
//   본문 "INS\r" — 전체 한 run, Ins + author=Alice + DTTM=2024-01-15 10:30.
//   결과: Run.IsInsertedRevision=true, RevisionAuthor="Alice", RevisionDate=2024-01-15T10:30Z.
static void DocBinaryPhase3h2RevisionMeta()
{
    const string text = "INS\r";
    int ccp = text.Length;
    var textBytes = Encoding.Unicode.GetBytes(text);
    int fcText = 0x200;
    int pnPapx = 4, pnChpx = 5;
    int fcPapxFkp = pnPapx * 512;
    int fcChpxFkp = pnChpx * 512;
    int wdSize = fcChpxFkp + 512;
    var wd = new byte[wdSize];
    Buffer.BlockCopy(textBytes, 0, wd, fcText, textBytes.Length);

    int cpara = 1;
    int paraEndFc = fcText + ccp * 2;
    BitConverter.TryWriteBytes(wd.AsSpan(fcPapxFkp + 0), (int)fcText);
    BitConverter.TryWriteBytes(wd.AsSpan(fcPapxFkp + 4), (int)paraEndFc);
    int papx0Off = 64;
    wd[fcPapxFkp + 8 + 0 * 13] = (byte)(papx0Off / 2);
    int p = fcPapxFkp + papx0Off;
    wd[p++] = 0; wd[p++] = 1; BitConverter.TryWriteBytes(wd.AsSpan(p), (ushort)0);
    wd[fcPapxFkp + 511] = (byte)cpara;

    // CHPX FKP — 1 run with sprmCFRMarkIns + sprmCIbstRMark(0) + sprmCDttmRMark(DTTM).
    // 2024-01-15 10:30 → DTTM = 0x07C17A9E.
    uint dttm = ((uint)(2024 - 1900) << 20) | (1u << 16) | (15u << 11) | (10u << 6) | 30u;
    SmokeHarness.Equal(0x07C17A9Eu, dttm, "DTTM 계산 검증");

    int crun = 1;
    BitConverter.TryWriteBytes(wd.AsSpan(fcChpxFkp + 0), (int)fcText);
    BitConverter.TryWriteBytes(wd.AsSpan(fcChpxFkp + 4), (int)paraEndFc);
    int chpx0Off = 64;
    wd[fcChpxFkp + 4 * (crun + 1) + 0] = (byte)(chpx0Off / 2);
    int c = fcChpxFkp + chpx0Off;
    // cb = (2+1) + (2+2) + (2+4) = 13
    wd[c++] = 13;
    BitConverter.TryWriteBytes(wd.AsSpan(c), (ushort)0x0801); c += 2;
    wd[c++] = 1;                                                       // Ins=true
    BitConverter.TryWriteBytes(wd.AsSpan(c), (ushort)0x4804); c += 2;
    BitConverter.TryWriteBytes(wd.AsSpan(c), (ushort)0);  c += 2;       // author idx = 0
    BitConverter.TryWriteBytes(wd.AsSpan(c), (ushort)0x6805); c += 2;
    BitConverter.TryWriteBytes(wd.AsSpan(c), dttm); c += 4;
    wd[fcChpxFkp + 511] = (byte)crun;

    // Table stream: CLX + BTE + SttbfRMark ["Alice"].
    var tblMs = new MemoryStream();
    Span<byte> b4 = stackalloc byte[4];
    Span<byte> b2 = stackalloc byte[2];
    tblMs.WriteByte(0x02);
    BitConverter.TryWriteBytes(b4, (uint)16); tblMs.Write(b4);
    BitConverter.TryWriteBytes(b4, (uint)0);   tblMs.Write(b4);
    BitConverter.TryWriteBytes(b4, (uint)ccp); tblMs.Write(b4);
    tblMs.WriteByte(0); tblMs.WriteByte(0);
    BitConverter.TryWriteBytes(b4, (uint)fcText); tblMs.Write(b4);
    tblMs.WriteByte(0); tblMs.WriteByte(0);
    int clxEnd = (int)tblMs.Position;
    int papxBteStart = clxEnd;
    BitConverter.TryWriteBytes(b4, (int)0x200);     tblMs.Write(b4);
    BitConverter.TryWriteBytes(b4, (int)paraEndFc); tblMs.Write(b4);
    BitConverter.TryWriteBytes(b4, (int)pnPapx);    tblMs.Write(b4);
    int papxBteLen = (int)tblMs.Position - papxBteStart;
    int chpxBteStart = (int)tblMs.Position;
    BitConverter.TryWriteBytes(b4, (int)0x200);     tblMs.Write(b4);
    BitConverter.TryWriteBytes(b4, (int)paraEndFc); tblMs.Write(b4);
    BitConverter.TryWriteBytes(b4, (int)pnChpx);    tblMs.Write(b4);
    int chpxBteLen = (int)tblMs.Position - chpxBteStart;

    // SttbfRMark extended Sttb: fExtend(0xFFFF) + cData(1) + cbExtra(0) + [cch(5) + "Alice"(10)]
    int rmarkStart = (int)tblMs.Position;
    BitConverter.TryWriteBytes(b2, (ushort)0xFFFF); tblMs.Write(b2);
    BitConverter.TryWriteBytes(b2, (ushort)1);      tblMs.Write(b2);
    BitConverter.TryWriteBytes(b2, (ushort)0);      tblMs.Write(b2);
    BitConverter.TryWriteBytes(b2, (ushort)5);      tblMs.Write(b2);
    var aliceBytes = Encoding.Unicode.GetBytes("Alice");
    tblMs.Write(aliceBytes, 0, aliceBytes.Length);
    int rmarkLen = (int)tblMs.Position - rmarkStart;
    var tblBytes = tblMs.ToArray();

    BitConverter.TryWriteBytes(wd.AsSpan(0x00),   (ushort)0xA5EC);
    BitConverter.TryWriteBytes(wd.AsSpan(0x02),   (ushort)193);
    BitConverter.TryWriteBytes(wd.AsSpan(0x0A),   (ushort)0x0000);
    BitConverter.TryWriteBytes(wd.AsSpan(0x18),   (uint)fcText);
    BitConverter.TryWriteBytes(wd.AsSpan(0x4C),   (uint)ccp);
    BitConverter.TryWriteBytes(wd.AsSpan(0x142),  (uint)rmarkStart);  // fcSttbfRMark
    BitConverter.TryWriteBytes(wd.AsSpan(0x146),  (uint)rmarkLen);    // lcbSttbfRMark
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
            using (var s = root.CreateStream("0Table"))       s.Write(tblBytes);
        }
        using var fs = File.OpenRead(tmp);
        var doc = new PolyDonky.Convert.Doc.DocBinaryReader().Read(fs);
        var para = (Paragraph)doc.Sections[0].Blocks[0];

        SmokeHarness.Equal("INS", para.GetPlainText(), "plain text = INS");
        SmokeHarness.Equal(1, para.Runs.Count, "Runs.Count = 1");
        var run = para.Runs[0];
        SmokeHarness.True(run.IsInsertedRevision, "Run.IsInsertedRevision");
        SmokeHarness.Equal("Alice", run.RevisionAuthor, "Run.RevisionAuthor");
        SmokeHarness.True(run.RevisionDate.HasValue, "Run.RevisionDate.HasValue");
        var d = run.RevisionDate!.Value;
        SmokeHarness.Equal(2024, d.Year,   "year");
        SmokeHarness.Equal(1,    d.Month,  "month");
        SmokeHarness.Equal(15,   d.Day,    "day");
        SmokeHarness.Equal(10,   d.Hour,   "hour");
        SmokeHarness.Equal(30,   d.Minute, "minute");
    }
    finally { try { File.Delete(tmp); } catch { } }
}

// Phase 3h-3 — Paragraph-level revision marks. 단락 마크(\r) 의 CHPX 에 sprmCFRMarkIns/Del 가
// 있으면 Paragraph.IsInsertedRevision/IsDeletedRevision 으로 표시.
//   text "A\rB\r" — 첫 \r 만 sprmCFRMarkIns. 결과:
//     Paragraph 0 ("A") — IsInsertedRevision=true (Run 자체에는 rev 없음 — \r 만 marked)
//     Paragraph 1 ("B") — no rev
static void DocBinaryPhase3h3ParagraphRevision()
{
    const string text = "A\rB\r";
    int ccp = text.Length;
    var textBytes = Encoding.Unicode.GetBytes(text);
    int fcText = 0x200;
    int pnPapx = 4, pnChpx = 5;
    int fcPapxFkp = pnPapx * 512;
    int fcChpxFkp = pnChpx * 512;
    int wdSize = fcChpxFkp + 512;
    var wd = new byte[wdSize];
    Buffer.BlockCopy(textBytes, 0, wd, fcText, textBytes.Length);

    // PAPX FKP — 2 paragraphs.
    int cpara = 2;
    BitConverter.TryWriteBytes(wd.AsSpan(fcPapxFkp + 0), (int)0x200);
    BitConverter.TryWriteBytes(wd.AsSpan(fcPapxFkp + 4), (int)0x204);
    BitConverter.TryWriteBytes(wd.AsSpan(fcPapxFkp + 8), (int)0x208);
    int papx0Off = 200, papx1Off = 220;
    wd[fcPapxFkp + 12 + 0 * 13] = (byte)(papx0Off / 2);
    wd[fcPapxFkp + 12 + 1 * 13] = (byte)(papx1Off / 2);
    int p = fcPapxFkp + papx0Off;
    wd[p++] = 0; wd[p++] = 1; BitConverter.TryWriteBytes(wd.AsSpan(p), (ushort)0);
    p = fcPapxFkp + papx1Off;
    wd[p++] = 0; wd[p++] = 1; BitConverter.TryWriteBytes(wd.AsSpan(p), (ushort)0);
    wd[fcPapxFkp + 511] = (byte)cpara;

    // CHPX FKP — 4 runs:
    //   Run 0: fc [0x200, 0x202) "A" — empty
    //   Run 1: fc [0x202, 0x204) first \r — sprmCFRMarkIns(=1)
    //   Run 2: fc [0x204, 0x206) "B" — empty
    //   Run 3: fc [0x206, 0x208) second \r — empty
    int crun = 4;
    BitConverter.TryWriteBytes(wd.AsSpan(fcChpxFkp + 0),  (int)0x200);
    BitConverter.TryWriteBytes(wd.AsSpan(fcChpxFkp + 4),  (int)0x202);
    BitConverter.TryWriteBytes(wd.AsSpan(fcChpxFkp + 8),  (int)0x204);
    BitConverter.TryWriteBytes(wd.AsSpan(fcChpxFkp + 12), (int)0x206);
    BitConverter.TryWriteBytes(wd.AsSpan(fcChpxFkp + 16), (int)0x208);
    int chpx0Off = 64, chpx1Off = 80, chpx2Off = 96, chpx3Off = 112;
    int rgbBase  = 4 * (crun + 1);
    wd[fcChpxFkp + rgbBase + 0] = (byte)(chpx0Off / 2);
    wd[fcChpxFkp + rgbBase + 1] = (byte)(chpx1Off / 2);
    wd[fcChpxFkp + rgbBase + 2] = (byte)(chpx2Off / 2);
    wd[fcChpxFkp + rgbBase + 3] = (byte)(chpx3Off / 2);
    wd[fcChpxFkp + chpx0Off] = 0;
    // Run 1: cb=3, sprmCFRMarkIns
    int c1 = fcChpxFkp + chpx1Off;
    wd[c1++] = 3;
    BitConverter.TryWriteBytes(wd.AsSpan(c1), (ushort)0x0801); c1 += 2;
    wd[c1] = 1;
    wd[fcChpxFkp + chpx2Off] = 0;
    wd[fcChpxFkp + chpx3Off] = 0;
    wd[fcChpxFkp + 511] = (byte)crun;

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
    var tblBytes = tblMs.ToArray();

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
            using (var s = root.CreateStream("0Table"))       s.Write(tblBytes);
        }
        using var fs = File.OpenRead(tmp);
        var doc = new PolyDonky.Convert.Doc.DocBinaryReader().Read(fs);
        var blocks = doc.Sections[0].Blocks;
        SmokeHarness.True(blocks.Count >= 2, $"blocks.Count >= 2 (got {blocks.Count})");

        var para0 = (Paragraph)blocks[0];
        SmokeHarness.Equal("A", para0.GetPlainText(), "para0 text");
        SmokeHarness.True(para0.IsInsertedRevision, "para0.IsInsertedRevision = true (\\r 마킹)");
        SmokeHarness.True(!para0.IsDeletedRevision, "para0.IsDeletedRevision = false");
        // Run 자체에는 rev 없음 — Run 0 의 CHPX 는 empty.
        SmokeHarness.True(para0.Runs.Count >= 1, "para0.Runs >= 1");
        SmokeHarness.True(!para0.Runs[0].IsInsertedRevision,
            "para0.Runs[0].IsInsertedRevision = false (Run 자체에는 rev 없음)");

        var para1 = (Paragraph)blocks[1];
        SmokeHarness.Equal("B", para1.GetPlainText(), "para1 text");
        SmokeHarness.True(!para1.IsInsertedRevision, "para1.IsInsertedRevision = false");
    }
    finally { try { File.Delete(tmp); } catch { } }
}

// Phase 3i — 책갈피 (Bookmarks). SttbfBkmk (이름) + PlcfBkf (시작 CP+BKF) + PlcfBkl (끝 CP).
//   2 bookmarks: intro (CP 0..2), body (CP 3..5).
static void DocBinaryPhase3iBookmarks()
{
    const string text = "Hello\r";
    int ccp = text.Length;
    var textBytes = Encoding.Unicode.GetBytes(text);
    int fcText = 0x200;
    int pnPapx = 4, pnChpx = 5;
    int fcPapxFkp = pnPapx * 512;
    int fcChpxFkp = pnChpx * 512;
    int wdSize = fcChpxFkp + 512;
    var wd = new byte[wdSize];
    Buffer.BlockCopy(textBytes, 0, wd, fcText, textBytes.Length);

    int cpara = 1;
    BitConverter.TryWriteBytes(wd.AsSpan(fcPapxFkp + 0), (int)0x200);
    BitConverter.TryWriteBytes(wd.AsSpan(fcPapxFkp + 4), (int)(0x200 + ccp * 2));
    int papx0Off = 64;
    wd[fcPapxFkp + 8 + 0 * 13] = (byte)(papx0Off / 2);
    int p = fcPapxFkp + papx0Off;
    wd[p++] = 0; wd[p++] = 1; BitConverter.TryWriteBytes(wd.AsSpan(p), (ushort)0);
    wd[fcPapxFkp + 511] = (byte)cpara;
    int crun = 1;
    BitConverter.TryWriteBytes(wd.AsSpan(fcChpxFkp + 0), (int)0x200);
    BitConverter.TryWriteBytes(wd.AsSpan(fcChpxFkp + 4), (int)(0x200 + ccp * 2));
    int chpx0Off = 64;
    wd[fcChpxFkp + 4 * (crun + 1) + 0] = (byte)(chpx0Off / 2);
    wd[fcChpxFkp + chpx0Off] = 0;
    wd[fcChpxFkp + 511] = (byte)crun;

    // Table stream: CLX + BTE + SttbfBkmk + PlcfBkf + PlcfBkl.
    var tblMs = new MemoryStream();
    Span<byte> b4 = stackalloc byte[4];
    Span<byte> b2 = stackalloc byte[2];
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

    // SttbfBkmk extended Sttb: fExtend + cData(2) + cbExtra(0) + ["intro"(5) + "body"(4)]
    int sttbStart = (int)tblMs.Position;
    BitConverter.TryWriteBytes(b2, (ushort)0xFFFF); tblMs.Write(b2);
    BitConverter.TryWriteBytes(b2, (ushort)2);      tblMs.Write(b2);
    BitConverter.TryWriteBytes(b2, (ushort)0);      tblMs.Write(b2);
    BitConverter.TryWriteBytes(b2, (ushort)5);      tblMs.Write(b2);
    var introBytes = Encoding.Unicode.GetBytes("intro");
    tblMs.Write(introBytes, 0, introBytes.Length);
    BitConverter.TryWriteBytes(b2, (ushort)4);      tblMs.Write(b2);
    var bodyBytes  = Encoding.Unicode.GetBytes("body");
    tblMs.Write(bodyBytes, 0, bodyBytes.Length);
    int sttbLen = (int)tblMs.Position - sttbStart;

    // PlcfBkf: aCP[3] = {0, 3, 6} + aBKF[2] each 4 byte (ibkl=0/1, flags=0).
    int bkfStart = (int)tblMs.Position;
    BitConverter.TryWriteBytes(b4, (int)0); tblMs.Write(b4);
    BitConverter.TryWriteBytes(b4, (int)3); tblMs.Write(b4);
    BitConverter.TryWriteBytes(b4, (int)6); tblMs.Write(b4);
    // BKF #0
    BitConverter.TryWriteBytes(b2, (ushort)0); tblMs.Write(b2);  // ibkl = 0
    BitConverter.TryWriteBytes(b2, (ushort)0); tblMs.Write(b2);  // flags
    // BKF #1
    BitConverter.TryWriteBytes(b2, (ushort)1); tblMs.Write(b2);  // ibkl = 1
    BitConverter.TryWriteBytes(b2, (ushort)0); tblMs.Write(b2);
    int bkfLen = (int)tblMs.Position - bkfStart;

    // PlcfBkl: aCP[3] = {2, 5, 6} only (data element size = 0).
    int bklStart = (int)tblMs.Position;
    BitConverter.TryWriteBytes(b4, (int)2); tblMs.Write(b4);
    BitConverter.TryWriteBytes(b4, (int)5); tblMs.Write(b4);
    BitConverter.TryWriteBytes(b4, (int)6); tblMs.Write(b4);
    int bklLen = (int)tblMs.Position - bklStart;

    var tblBytes = tblMs.ToArray();

    BitConverter.TryWriteBytes(wd.AsSpan(0x00),   (ushort)0xA5EC);
    BitConverter.TryWriteBytes(wd.AsSpan(0x02),   (ushort)193);
    BitConverter.TryWriteBytes(wd.AsSpan(0x0A),   (ushort)0x0000);
    BitConverter.TryWriteBytes(wd.AsSpan(0x18),   (uint)fcText);
    BitConverter.TryWriteBytes(wd.AsSpan(0x4C),   (uint)ccp);
    BitConverter.TryWriteBytes(wd.AsSpan(0x182),  (uint)sttbStart);   // fcSttbfBkmk
    BitConverter.TryWriteBytes(wd.AsSpan(0x186),  (uint)sttbLen);     // lcbSttbfBkmk
    BitConverter.TryWriteBytes(wd.AsSpan(0x18A),  (uint)bkfStart);    // fcPlcfBkf
    BitConverter.TryWriteBytes(wd.AsSpan(0x18E),  (uint)bkfLen);      // lcbPlcfBkf
    BitConverter.TryWriteBytes(wd.AsSpan(0x192),  (uint)bklStart);    // fcPlcfBkl
    BitConverter.TryWriteBytes(wd.AsSpan(0x196),  (uint)bklLen);      // lcbPlcfBkl
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
            using (var s = root.CreateStream("0Table"))       s.Write(tblBytes);
        }
        using var fs = File.OpenRead(tmp);
        var reader = new PolyDonky.Convert.Doc.DocBinaryReader();
        reader.Read(fs);

        SmokeHarness.Equal(2, reader.Bookmarks.Count, $"Bookmarks.Count = 2 (got {reader.Bookmarks.Count})");

        SmokeHarness.Equal("intro", reader.Bookmarks[0].Name,    "Bookmark[0].Name = intro");
        SmokeHarness.Equal(0,       reader.Bookmarks[0].StartCp, "Bookmark[0].StartCp = 0");
        SmokeHarness.Equal(2,       reader.Bookmarks[0].EndCp,   "Bookmark[0].EndCp = 2");

        SmokeHarness.Equal("body",  reader.Bookmarks[1].Name,    "Bookmark[1].Name = body");
        SmokeHarness.Equal(3,       reader.Bookmarks[1].StartCp, "Bookmark[1].StartCp = 3");
        SmokeHarness.Equal(5,       reader.Bookmarks[1].EndCp,   "Bookmark[1].EndCp = 5");
    }
    finally { try { File.Delete(tmp); } catch { } }
}

// Phase 3i-2 — 책갈피 marker Run 자동 삽입. 본문 "Hello" 안에 bookmark "intro" 가 CP 1..4
// (= "ell") 를 감싸면 paragraph runs = ["H", BookmarkStart, "ell", BookmarkEnd, "o"].
static void DocBinaryPhase3i2BookmarkRunMarkers()
{
    const string text = "Hello\r";
    int ccp = text.Length;
    var textBytes = Encoding.Unicode.GetBytes(text);
    int fcText = 0x200;
    int pnPapx = 4, pnChpx = 5;
    int fcPapxFkp = pnPapx * 512;
    int fcChpxFkp = pnChpx * 512;
    int wdSize = fcChpxFkp + 512;
    var wd = new byte[wdSize];
    Buffer.BlockCopy(textBytes, 0, wd, fcText, textBytes.Length);

    int cpara = 1;
    BitConverter.TryWriteBytes(wd.AsSpan(fcPapxFkp + 0), (int)0x200);
    BitConverter.TryWriteBytes(wd.AsSpan(fcPapxFkp + 4), (int)(0x200 + ccp * 2));
    int papx0Off = 64;
    wd[fcPapxFkp + 8 + 0 * 13] = (byte)(papx0Off / 2);
    int p = fcPapxFkp + papx0Off;
    wd[p++] = 0; wd[p++] = 1; BitConverter.TryWriteBytes(wd.AsSpan(p), (ushort)0);
    wd[fcPapxFkp + 511] = (byte)cpara;
    int crun = 1;
    BitConverter.TryWriteBytes(wd.AsSpan(fcChpxFkp + 0), (int)0x200);
    BitConverter.TryWriteBytes(wd.AsSpan(fcChpxFkp + 4), (int)(0x200 + ccp * 2));
    int chpx0Off = 64;
    wd[fcChpxFkp + 4 * (crun + 1) + 0] = (byte)(chpx0Off / 2);
    wd[fcChpxFkp + chpx0Off] = 0;
    wd[fcChpxFkp + 511] = (byte)crun;

    var tblMs = new MemoryStream();
    Span<byte> b4 = stackalloc byte[4];
    Span<byte> b2 = stackalloc byte[2];
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

    // SttbfBkmk: "intro" 한 개
    int sttbStart = (int)tblMs.Position;
    BitConverter.TryWriteBytes(b2, (ushort)0xFFFF); tblMs.Write(b2);
    BitConverter.TryWriteBytes(b2, (ushort)1);      tblMs.Write(b2);
    BitConverter.TryWriteBytes(b2, (ushort)0);      tblMs.Write(b2);
    BitConverter.TryWriteBytes(b2, (ushort)5);      tblMs.Write(b2);
    var introBytes = Encoding.Unicode.GetBytes("intro");
    tblMs.Write(introBytes, 0, introBytes.Length);
    int sttbLen = (int)tblMs.Position - sttbStart;

    // PlcfBkf: aCP[2] = {1, 6} + BKF[1] (ibkl=0)
    int bkfStart = (int)tblMs.Position;
    BitConverter.TryWriteBytes(b4, (int)1); tblMs.Write(b4);
    BitConverter.TryWriteBytes(b4, (int)6); tblMs.Write(b4);
    BitConverter.TryWriteBytes(b2, (ushort)0); tblMs.Write(b2);   // ibkl
    BitConverter.TryWriteBytes(b2, (ushort)0); tblMs.Write(b2);
    int bkfLen = (int)tblMs.Position - bkfStart;

    // PlcfBkl: aCP[2] = {4, 6}
    int bklStart = (int)tblMs.Position;
    BitConverter.TryWriteBytes(b4, (int)4); tblMs.Write(b4);
    BitConverter.TryWriteBytes(b4, (int)6); tblMs.Write(b4);
    int bklLen = (int)tblMs.Position - bklStart;

    var tblBytes = tblMs.ToArray();

    BitConverter.TryWriteBytes(wd.AsSpan(0x00),   (ushort)0xA5EC);
    BitConverter.TryWriteBytes(wd.AsSpan(0x02),   (ushort)193);
    BitConverter.TryWriteBytes(wd.AsSpan(0x0A),   (ushort)0x0000);
    BitConverter.TryWriteBytes(wd.AsSpan(0x18),   (uint)fcText);
    BitConverter.TryWriteBytes(wd.AsSpan(0x4C),   (uint)ccp);
    BitConverter.TryWriteBytes(wd.AsSpan(0x182),  (uint)sttbStart);
    BitConverter.TryWriteBytes(wd.AsSpan(0x186),  (uint)sttbLen);
    BitConverter.TryWriteBytes(wd.AsSpan(0x18A),  (uint)bkfStart);
    BitConverter.TryWriteBytes(wd.AsSpan(0x18E),  (uint)bkfLen);
    BitConverter.TryWriteBytes(wd.AsSpan(0x192),  (uint)bklStart);
    BitConverter.TryWriteBytes(wd.AsSpan(0x196),  (uint)bklLen);
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
            using (var s = root.CreateStream("0Table"))       s.Write(tblBytes);
        }
        using var fs = File.OpenRead(tmp);
        var doc = new PolyDonky.Convert.Doc.DocBinaryReader().Read(fs);
        var para = (Paragraph)doc.Sections[0].Blocks[0];

        SmokeHarness.Equal("Hello", para.GetPlainText(), "plain text = Hello (marker Run 은 Text 없음)");
        SmokeHarness.Equal(5, para.Runs.Count, $"Runs.Count = 5 (got {para.Runs.Count})");
        SmokeHarness.Equal("H",   para.Runs[0].Text, "Run[0] = H");
        SmokeHarness.Equal("",    para.Runs[1].Text, "Run[1].Text = ''");
        SmokeHarness.Equal("intro", para.Runs[1].BookmarkStart, "Run[1].BookmarkStart = intro");
        SmokeHarness.Equal("ell", para.Runs[2].Text, "Run[2] = ell");
        SmokeHarness.Equal("",    para.Runs[3].Text, "Run[3].Text = ''");
        SmokeHarness.Equal("intro", para.Runs[3].BookmarkEnd, "Run[3].BookmarkEnd = intro");
        SmokeHarness.Equal("o",   para.Runs[4].Text, "Run[4] = o");
    }
    finally { try { File.Delete(tmp); } catch { } }
}

// Phase 3j — 주석 (annotation/comment) ingest. ccpAtn 영역의 텍스트를 PlcfAtnTxt 로 분할해
// doc.Comments 에, 본문 0x05 char 를 PlcfAtnRef 와 매칭해 Run.CommentId 부여.
//   main "A<0x05>B\r"  (ccpText=4)
//   comment 영역 " \rComment text\r"  (ccpAtn=15, separator + 1 comment)
//   결과: doc.Comments = [{Id=cmt1, "Comment text"}],
//         paragraph runs = [A, "" CommentId=cmt1, B].
static void DocBinaryPhase3jComments()
{
    const string text = "AB\r \rComment text\r";  // 0x05 박은 후 19 chars 가 되도록 placeholder
    int ccpMain = 4;
    int ccpAtn  = 15;
    int totalCcp = ccpMain + ccpAtn;
    SmokeHarness.Equal(totalCcp, text.Length, "text 길이 검증");

    var textBytes = Encoding.Unicode.GetBytes(text);
    int fcText = 0x200;
    int pnPapx = 4, pnChpx = 5;
    int fcPapxFkp = pnPapx * 512;
    int fcChpxFkp = pnChpx * 512;
    int wdSize = fcChpxFkp + 512;
    var wd = new byte[wdSize];
    Buffer.BlockCopy(textBytes, 0, wd, fcText, textBytes.Length);

    int cpara = 1;
    int paraEndFc = fcText + ccpMain * 2;
    BitConverter.TryWriteBytes(wd.AsSpan(fcPapxFkp + 0), (int)fcText);
    BitConverter.TryWriteBytes(wd.AsSpan(fcPapxFkp + 4), (int)paraEndFc);
    int papx0Off = 64;
    wd[fcPapxFkp + 8 + 0 * 13] = (byte)(papx0Off / 2);
    int p = fcPapxFkp + papx0Off;
    wd[p++] = 0; wd[p++] = 1; BitConverter.TryWriteBytes(wd.AsSpan(p), (ushort)0);
    wd[fcPapxFkp + 511] = (byte)cpara;
    int crun = 1;
    BitConverter.TryWriteBytes(wd.AsSpan(fcChpxFkp + 0), (int)fcText);
    BitConverter.TryWriteBytes(wd.AsSpan(fcChpxFkp + 4), (int)paraEndFc);
    int chpx0Off = 64;
    wd[fcChpxFkp + 4 * (crun + 1) + 0] = (byte)(chpx0Off / 2);
    wd[fcChpxFkp + chpx0Off] = 0;
    wd[fcChpxFkp + 511] = (byte)crun;

    var tblMs = new MemoryStream();
    Span<byte> b4 = stackalloc byte[4];
    Span<byte> b2 = stackalloc byte[2];
    tblMs.WriteByte(0x02);
    BitConverter.TryWriteBytes(b4, (uint)16); tblMs.Write(b4);
    BitConverter.TryWriteBytes(b4, (uint)0);          tblMs.Write(b4);
    BitConverter.TryWriteBytes(b4, (uint)totalCcp);   tblMs.Write(b4);
    tblMs.WriteByte(0); tblMs.WriteByte(0);
    BitConverter.TryWriteBytes(b4, (uint)fcText);     tblMs.Write(b4);
    tblMs.WriteByte(0); tblMs.WriteByte(0);
    int clxEnd = (int)tblMs.Position;
    int papxBteStart = clxEnd;
    BitConverter.TryWriteBytes(b4, (int)fcText);      tblMs.Write(b4);
    BitConverter.TryWriteBytes(b4, (int)(fcText + totalCcp * 2 + 0x100)); tblMs.Write(b4);
    BitConverter.TryWriteBytes(b4, (int)pnPapx);      tblMs.Write(b4);
    int papxBteLen = (int)tblMs.Position - papxBteStart;
    int chpxBteStart = (int)tblMs.Position;
    BitConverter.TryWriteBytes(b4, (int)fcText);      tblMs.Write(b4);
    BitConverter.TryWriteBytes(b4, (int)(fcText + totalCcp * 2 + 0x100)); tblMs.Write(b4);
    BitConverter.TryWriteBytes(b4, (int)pnChpx);      tblMs.Write(b4);
    int chpxBteLen = (int)tblMs.Position - chpxBteStart;

    // PlcfAtnTxt: aCP = [0, 2, 15] (separator + 1 comment). lcb = 12.
    int atnTxtStart = (int)tblMs.Position;
    BitConverter.TryWriteBytes(b4, (int)0);  tblMs.Write(b4);
    BitConverter.TryWriteBytes(b4, (int)2);  tblMs.Write(b4);
    BitConverter.TryWriteBytes(b4, (int)15); tblMs.Write(b4);
    int atnTxtLen = (int)tblMs.Position - atnTxtStart;

    // PlcfAtnRef: aCP = [1 (0x05 위치), 4 (ccpText)] + 1 ATRDPre10 (30 byte). lcb = 4*2 + 30 = 38.
    int atnRefStart = (int)tblMs.Position;
    BitConverter.TryWriteBytes(b4, (int)1); tblMs.Write(b4);
    BitConverter.TryWriteBytes(b4, (int)4); tblMs.Write(b4);
    tblMs.Write(new byte[30]);  // ATRDPre10 (all zeros — author info 등은 비사용)
    int atnRefLen = (int)tblMs.Position - atnRefStart;

    var tblBytes = tblMs.ToArray();

    BitConverter.TryWriteBytes(wd.AsSpan(0x00),   (ushort)0xA5EC);
    BitConverter.TryWriteBytes(wd.AsSpan(0x02),   (ushort)193);
    BitConverter.TryWriteBytes(wd.AsSpan(0x0A),   (ushort)0x0000);
    BitConverter.TryWriteBytes(wd.AsSpan(0x18),   (uint)fcText);
    BitConverter.TryWriteBytes(wd.AsSpan(0x4C),   (uint)ccpMain);
    BitConverter.TryWriteBytes(wd.AsSpan(0x50),   (uint)0);          // ccpFtn
    BitConverter.TryWriteBytes(wd.AsSpan(0x54),   (uint)0);          // ccpHdd
    BitConverter.TryWriteBytes(wd.AsSpan(0x5C),   (uint)ccpAtn);     // ccpAtn
    BitConverter.TryWriteBytes(wd.AsSpan(0x60),   (uint)0);          // ccpEdn
    BitConverter.TryWriteBytes(wd.AsSpan(0xBA),   (uint)atnRefStart); // fcPlcfAtnRef
    BitConverter.TryWriteBytes(wd.AsSpan(0xBE),   (uint)atnRefLen);   // lcbPlcfAtnRef
    BitConverter.TryWriteBytes(wd.AsSpan(0xC2),   (uint)atnTxtStart); // fcPlcfAtnTxt
    BitConverter.TryWriteBytes(wd.AsSpan(0xC6),   (uint)atnTxtLen);   // lcbPlcfAtnTxt
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
            using (var s = root.CreateStream("0Table"))       s.Write(tblBytes);
        }
        using var fs = File.OpenRead(tmp);
        var doc = new PolyDonky.Convert.Doc.DocBinaryReader().Read(fs);

        // 코멘트 본문.
        SmokeHarness.Equal(1, doc.Comments.Count, "Comments.Count = 1");
        SmokeHarness.Equal("cmt1", doc.Comments[0].Id, "Comment Id = cmt1");
        var cmtPara = (Paragraph)doc.Comments[0].Blocks[0];
        SmokeHarness.Equal("Comment text", cmtPara.GetPlainText(), "Comment 본문");

        // 본문 ref Run.
        var para = (Paragraph)doc.Sections[0].Blocks[0];
        SmokeHarness.Equal("AB", para.GetPlainText(), "본문 plain text = AB");
        // Run 분할: A | "" (CommentId=cmt1) | B
        SmokeHarness.True(para.Runs.Count >= 3, $"Runs.Count >= 3 (got {para.Runs.Count})");
        var refRun = para.Runs.FirstOrDefault(r => r.CommentId is not null);
        SmokeHarness.True(refRun is not null, "CommentId Run 존재");
        SmokeHarness.Equal("cmt1", refRun!.CommentId, "Run.CommentId = cmt1");
        SmokeHarness.Equal("", refRun.Text, "ref Run Text 빈문자");
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
