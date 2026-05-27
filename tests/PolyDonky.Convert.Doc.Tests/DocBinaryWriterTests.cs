using System.Text;
using OpenMcdf;
using PolyDonky.Convert.Doc;
using PolyDonky.Core;

namespace PolyDonky.Convert.Doc.Tests;

/// <summary>
/// DocBinaryWriter (IWPF → Word 97-2003 OLE2 .doc) round-trip 및 컨테이너 구조 검증.
/// Phase F1-W2a 는 본문 평문만 보존하므로, 서식·이미지·표 등은 후속 단계에서 검증.
/// </summary>
public class DocBinaryWriterTests
{
    private static byte[] WriteDoc(PolyDonkyument doc)
    {
        using var ms = new MemoryStream();
        new DocBinaryWriter().Write(doc, ms);
        return ms.ToArray();
    }

    private static PolyDonkyument RoundTrip(PolyDonkyument doc)
    {
        var bytes = WriteDoc(doc);
        using var ms = new MemoryStream(bytes);
        return new DocBinaryReader().Read(ms);
    }

    private static PolyDonkyument DocWith(params Paragraph[] paragraphs)
    {
        var doc = new PolyDonkyument();
        var sec = new Section();
        foreach (var p in paragraphs) sec.Blocks.Add(p);
        doc.Sections.Add(sec);
        return doc;
    }

    private static IEnumerable<string> NonEmptyParagraphTexts(PolyDonkyument doc) =>
        doc.Sections.SelectMany(s => s.Blocks).OfType<Paragraph>()
                    .Select(p => p.GetPlainText())
                    .Where(t => !string.IsNullOrEmpty(t));

    // ── CFB 컨테이너 구조 ────────────────────────────────────────────────────

    [Fact]
    public void Output_Is_Valid_OLE2_CFB_With_Both_Word97_Streams()
    {
        var bytes = WriteDoc(DocWith(Paragraph.Of("hello")));
        using var ms = new MemoryStream(bytes);
        using var root = RootStorage.Open(ms, StorageModeFlags.LeaveOpen);

        var names = root.EnumerateEntries().Select(e => e.Name).ToHashSet();
        Assert.Contains("WordDocument", names);
        Assert.Contains("1Table",       names);
    }

    [Fact]
    public void WordDocument_Stream_Starts_With_FIB_Magic()
    {
        var bytes = WriteDoc(DocWith(Paragraph.Of("x")));
        using var ms = new MemoryStream(bytes);
        using var root = RootStorage.Open(ms, StorageModeFlags.LeaveOpen);
        using var wd  = root.OpenStream("WordDocument");

        var head = new byte[4];
        wd.ReadExactly(head, 0, 4);
        // [MS-DOC] §2.5.2 wIdent = 0xA5EC
        Assert.Equal(0xEC, head[0]);
        Assert.Equal(0xA5, head[1]);
        // nFib = 0x00C1 (Word 97-2003 표준)
        Assert.Equal(0xC1, head[2]);
        Assert.Equal(0x00, head[3]);
    }

    [Fact]
    public void WordDocument_Stream_Length_Is_At_Least_Fib_Pad_Plus_Text()
    {
        var bytes = WriteDoc(DocWith(Paragraph.Of("hello")));
        using var ms = new MemoryStream(bytes);
        using var root = RootStorage.Open(ms, StorageModeFlags.LeaveOpen);
        using var wd  = root.OpenStream("WordDocument");

        // FIB 0x400 + Unicode "hello\r" (6 chars × 2 byte) = 0x40C
        Assert.True(wd.Length >= 0x400 + 12);
    }

    [Fact]
    public void Table_Stream_Is_Named_1Table_Per_Flag_Bit_9()
    {
        var bytes = WriteDoc(DocWith(Paragraph.Of("x")));
        using var ms = new MemoryStream(bytes);
        using var root = RootStorage.Open(ms, StorageModeFlags.LeaveOpen);

        Assert.True(root.ContainsEntry("1Table"));
        Assert.False(root.ContainsEntry("0Table"));
    }

    // ── 라운드트립 ───────────────────────────────────────────────────────────

    [Fact]
    public void Empty_Document_RoundTrips_Without_Exception()
    {
        var doc = new PolyDonkyument();
        doc.Sections.Add(new Section());

        var doc2 = RoundTrip(doc);

        Assert.NotNull(doc2);
        Assert.Single(doc2.Sections);
        // 최소 한 개의 단락 마크는 항상 emit 됨 (Word 호환)
        Assert.NotEmpty(doc2.Sections[0].Blocks);
    }

    [Fact]
    public void Single_Paragraph_Text_RoundTrips()
    {
        var doc = DocWith(Paragraph.Of("Hello, world!"));
        var doc2 = RoundTrip(doc);

        Assert.Contains("Hello, world!", NonEmptyParagraphTexts(doc2));
    }

    [Fact]
    public void Multiple_Paragraphs_RoundTrip_In_Order()
    {
        var doc = DocWith(
            Paragraph.Of("첫 단락"),
            Paragraph.Of("두 번째"),
            Paragraph.Of("세 번째 단락"));
        var doc2 = RoundTrip(doc);

        var texts = NonEmptyParagraphTexts(doc2).ToArray();
        Assert.Equal(new[] { "첫 단락", "두 번째", "세 번째 단락" }, texts);
    }

    [Fact]
    public void Korean_Unicode_Text_RoundTrips_Without_Corruption()
    {
        var doc = DocWith(Paragraph.Of("한글 텍스트 — emoji 🦒 와 함께"));
        var doc2 = RoundTrip(doc);

        // BMP 안 한글은 그대로 보존. 보충면(🦒 = U+1F992) 은 surrogate pair 로 UTF-16 에 인코딩되어 보존.
        var text = string.Concat(NonEmptyParagraphTexts(doc2));
        Assert.Contains("한글 텍스트", text);
        Assert.Contains("🦒", text);
    }

    [Fact]
    public void Multi_Run_Paragraph_Concatenates_Text()
    {
        var p = new Paragraph();
        p.AddText("Bold ", new RunStyle { Bold = true });
        p.AddText("italic ", new RunStyle { Italic = true });
        p.AddText("plain.");
        var doc = DocWith(p);

        var doc2 = RoundTrip(doc);

        // Phase F1-W2a 는 서식 무시. 텍스트만 합쳐져 보존.
        Assert.Contains("Bold italic plain.", NonEmptyParagraphTexts(doc2));
    }

    [Fact]
    public void Embedded_Newline_In_Run_Is_Sanitized_To_Space()
    {
        var doc = DocWith(Paragraph.Of("line1\nline2"));
        var doc2 = RoundTrip(doc);

        // Run 안 \n 은 Word 본문 형식에서 의미가 다르므로 공백으로 치환됨
        var text = string.Concat(NonEmptyParagraphTexts(doc2));
        Assert.Contains("line1 line2", text);
        Assert.DoesNotContain("\n", text);
    }

    [Fact]
    public void Tab_Character_Is_Preserved()
    {
        var doc = DocWith(Paragraph.Of("col1\tcol2"));
        var doc2 = RoundTrip(doc);

        Assert.Contains("col1\tcol2", NonEmptyParagraphTexts(doc2));
    }

    [Fact]
    public void Multiple_Sections_Flatten_Into_Body_Paragraphs()
    {
        var doc = new PolyDonkyument();
        var s1 = new Section();
        s1.Blocks.Add(Paragraph.Of("section 1"));
        var s2 = new Section();
        s2.Blocks.Add(Paragraph.Of("section 2"));
        doc.Sections.Add(s1);
        doc.Sections.Add(s2);

        var doc2 = RoundTrip(doc);

        var texts = NonEmptyParagraphTexts(doc2).ToArray();
        Assert.Contains("section 1", texts);
        Assert.Contains("section 2", texts);
    }

    [Fact]
    public void Document_Without_Sections_Still_Produces_Valid_Output()
    {
        var doc = new PolyDonkyument();   // 섹션 0개

        var bytes = WriteDoc(doc);

        // 빈 입력이라도 CFB + WordDocument + 1Table + 단락 마크 \r 는 emit
        using var ms = new MemoryStream(bytes);
        using var root = RootStorage.Open(ms, StorageModeFlags.LeaveOpen);
        Assert.True(root.ContainsEntry("WordDocument"));
        Assert.True(root.ContainsEntry("1Table"));
    }

    [Fact]
    public void Long_Paragraph_Multi_Sector_RoundTrips()
    {
        // CFB sector 크기(V3 = 512 byte) 를 넘는 텍스트로 multi-sector chain 검증
        var longText = new string('A', 3000);
        var doc = DocWith(Paragraph.Of(longText));

        var doc2 = RoundTrip(doc);

        Assert.Contains(longText, NonEmptyParagraphTexts(doc2));
    }

    [Fact]
    public void Successive_Round_Trips_Are_Stable()
    {
        // CFB sector 배치는 OpenMcdf 의 내부 결정에 따라 byte 가 일치하지 않을 수 있지만,
        // 의미적 round-trip (text 보존) 은 안정적이어야 한다.
        var doc1 = RoundTrip(DocWith(Paragraph.Of("stable")));
        var doc2 = RoundTrip(doc1);
        var doc3 = RoundTrip(doc2);

        Assert.Contains("stable", NonEmptyParagraphTexts(doc3));
    }
}
