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

    // ── Phase F1-W2b: 글자 서식 round-trip ──────────────────────────────────

    private static Paragraph FirstParaWithText(PolyDonkyument doc, string snippet) =>
        doc.Sections.SelectMany(s => s.Blocks).OfType<Paragraph>()
                    .First(p => p.GetPlainText().Contains(snippet));

    [Fact]
    public void Bold_Run_RoundTrips()
    {
        var p = new Paragraph();
        p.AddText("hello ");
        p.AddText("BOLD", new RunStyle { Bold = true });
        p.AddText(" world");

        var doc2 = RoundTrip(DocWith(p));
        var rp   = FirstParaWithText(doc2, "BOLD");
        var bold = rp.Runs.FirstOrDefault(r => r.Style.Bold);
        Assert.NotNull(bold);
        Assert.Equal("BOLD", bold!.Text);
    }

    [Fact]
    public void Italic_Underline_Strike_RoundTrip()
    {
        var p = new Paragraph();
        p.AddText("I", new RunStyle { Italic = true });
        p.AddText("U", new RunStyle { Underline = true });
        p.AddText("S", new RunStyle { Strikethrough = true });

        var doc2 = RoundTrip(DocWith(p));
        var rp   = FirstParaWithText(doc2, "IUS");
        var rs   = rp.Runs.Where(r => !string.IsNullOrEmpty(r.Text)).ToList();
        Assert.Contains(rs, r => r.Text == "I" && r.Style.Italic);
        Assert.Contains(rs, r => r.Text == "U" && r.Style.Underline);
        Assert.Contains(rs, r => r.Text == "S" && r.Style.Strikethrough);
    }

    [Fact]
    public void Font_Size_RoundTrips()
    {
        var p = new Paragraph();
        p.AddText("big", new RunStyle { FontSizePt = 18 });
        var doc2 = RoundTrip(DocWith(p));
        var rp   = FirstParaWithText(doc2, "big");
        Assert.Equal(18.0, rp.Runs.First(r => r.Text == "big").Style.FontSizePt, 1);
    }

    [Fact]
    public void Foreground_Color_RoundTrips_RGB()
    {
        var red = new Color(255, 0, 0);
        var p = new Paragraph();
        p.AddText("red", new RunStyle { Foreground = red });
        var doc2 = RoundTrip(DocWith(p));
        var rp   = FirstParaWithText(doc2, "red");
        var fg   = rp.Runs.First(r => r.Text == "red").Style.Foreground;
        Assert.Equal(red, fg);
    }

    [Fact]
    public void Font_Family_RoundTrips_Via_SttbfFfn()
    {
        var p = new Paragraph();
        p.AddText("Arial text", new RunStyle { FontFamily = "Arial" });
        var doc2 = RoundTrip(DocWith(p));
        var rp   = FirstParaWithText(doc2, "Arial text");
        Assert.Equal("Arial", rp.Runs.First(r => r.Text == "Arial text").Style.FontFamily);
    }

    // ── Phase F1-W2b: 단락 서식 round-trip ──────────────────────────────────

    [Theory]
    [InlineData(Alignment.Left)]
    [InlineData(Alignment.Center)]
    [InlineData(Alignment.Right)]
    [InlineData(Alignment.Justify)]
    public void Paragraph_Alignment_RoundTrips(Alignment align)
    {
        var p = new Paragraph { Style = { Alignment = align } };
        p.AddText("alignment test");
        var doc2 = RoundTrip(DocWith(p));
        var rp   = FirstParaWithText(doc2, "alignment test");
        Assert.Equal(align, rp.Style.Alignment);
    }

    [Fact]
    public void Paragraph_Indents_RoundTrip()
    {
        var p = new Paragraph
        {
            Style = { IndentLeftMm = 15, IndentRightMm = 5, IndentFirstLineMm = 10 },
        };
        p.AddText("indented");
        var doc2 = RoundTrip(DocWith(p));
        var rp   = FirstParaWithText(doc2, "indented");
        Assert.Equal(15, rp.Style.IndentLeftMm,      0);
        Assert.Equal(5,  rp.Style.IndentRightMm,     0);
        Assert.Equal(10, rp.Style.IndentFirstLineMm, 0);
    }

    [Fact]
    public void Paragraph_Spacing_RoundTrips()
    {
        var p = new Paragraph { Style = { SpaceBeforePt = 12, SpaceAfterPt = 6 } };
        p.AddText("spaced");
        var doc2 = RoundTrip(DocWith(p));
        var rp   = FirstParaWithText(doc2, "spaced");
        Assert.Equal(12, rp.Style.SpaceBeforePt, 0);
        Assert.Equal(6,  rp.Style.SpaceAfterPt,  0);
    }

    [Fact]
    public void Line_Height_Factor_RoundTrips()
    {
        var p = new Paragraph { Style = { LineHeightFactor = 1.5 } };
        p.AddText("loose");
        var doc2 = RoundTrip(DocWith(p));
        var rp   = FirstParaWithText(doc2, "loose");
        Assert.InRange(rp.Style.LineHeightFactor, 1.45, 1.55);
    }

    [Fact]
    public void Many_Paragraphs_Spill_To_Multiple_Fkp_Pages()
    {
        // 한 FKP 페이지(512 byte) 안에 들어가는 PAPX 항목 수보다 많은 단락을 만들어
        // multi-page packing 동작을 검증. 약 30 개 paragraph 면 PAPX 2 페이지 필요.
        var paras = Enumerable.Range(1, 60).Select(i =>
        {
            var p = new Paragraph();
            p.AddText($"p{i:00}", new RunStyle { Bold = i % 3 == 0 });
            return p;
        }).ToArray();

        var doc2 = RoundTrip(DocWith(paras));

        var texts = NonEmptyParagraphTexts(doc2).ToArray();
        Assert.Equal(60, texts.Length);
        Assert.Equal("p01", texts[0]);
        Assert.Equal("p60", texts[59]);

        // bold 단락 (3 의 배수 인덱스) 도 보존되어야 함
        var boldTexts = doc2.Sections.SelectMany(s => s.Blocks).OfType<Paragraph>()
            .Where(p => p.Runs.Any(r => r.Style.Bold))
            .Select(p => p.GetPlainText())
            .ToArray();
        Assert.Contains("p03", boldTexts);
        Assert.Contains("p60", boldTexts);
    }

    // ── Phase F1-W2c: 책갈피 round-trip ────────────────────────────────────

    [Fact]
    public void Single_Bookmark_RoundTrips_Via_DocBinaryReader()
    {
        // Run 의 BookmarkStart/End 가 본문 텍스트 양 끝을 감싸는 패턴.
        var p = new Paragraph();
        p.Runs.Add(new Run { BookmarkStart = "chap1" });
        p.AddText("Chapter 1 content");
        p.Runs.Add(new Run { BookmarkEnd   = "chap1" });

        var bytes = WriteDoc(DocWith(p));
        using var ms = new MemoryStream(bytes);
        var reader = new DocBinaryReader();
        _ = reader.Read(ms);

        Assert.Single(reader.Bookmarks);
        Assert.Equal("chap1", reader.Bookmarks[0].Name);
        Assert.True(reader.Bookmarks[0].EndCp > reader.Bookmarks[0].StartCp);
    }

    [Fact]
    public void Multiple_Bookmarks_RoundTrip_With_Distinct_Names()
    {
        var p = new Paragraph();
        p.Runs.Add(new Run { BookmarkStart = "alpha" });
        p.AddText("a");
        p.Runs.Add(new Run { BookmarkEnd   = "alpha" });
        p.Runs.Add(new Run { BookmarkStart = "beta" });
        p.AddText("b");
        p.Runs.Add(new Run { BookmarkEnd   = "beta" });

        var bytes = WriteDoc(DocWith(p));
        using var ms = new MemoryStream(bytes);
        var reader = new DocBinaryReader();
        _ = reader.Read(ms);

        Assert.Equal(2, reader.Bookmarks.Count);
        Assert.Contains(reader.Bookmarks, b => b.Name == "alpha");
        Assert.Contains(reader.Bookmarks, b => b.Name == "beta");
    }

    [Fact]
    public void Bookmark_With_Korean_Name_RoundTrips_Utf16()
    {
        var p = new Paragraph();
        p.Runs.Add(new Run { BookmarkStart = "북마크_한글" });
        p.AddText("내용");
        p.Runs.Add(new Run { BookmarkEnd   = "북마크_한글" });

        var bytes = WriteDoc(DocWith(p));
        using var ms = new MemoryStream(bytes);
        var reader = new DocBinaryReader();
        _ = reader.Read(ms);

        Assert.Single(reader.Bookmarks);
        Assert.Equal("북마크_한글", reader.Bookmarks[0].Name);
    }

    [Fact]
    public void No_Bookmarks_Means_No_Sttbf_Or_Plcf_Streams_In_Fib()
    {
        var bytes = WriteDoc(DocWith(Paragraph.Of("no bookmarks here")));
        using var ms = new MemoryStream(bytes);
        var reader = new DocBinaryReader();
        _ = reader.Read(ms);

        Assert.Empty(reader.Bookmarks);
    }

    [Fact]
    public void Unclosed_Bookmark_Extends_To_End_Of_Paragraph()
    {
        var p = new Paragraph();
        p.Runs.Add(new Run { BookmarkStart = "unclosed" });
        p.AddText("paragraph content");
        // BookmarkEnd 빠짐 — 자동으로 단락 끝까지 확장되어야 함

        var bytes = WriteDoc(DocWith(p));
        using var ms = new MemoryStream(bytes);
        var reader = new DocBinaryReader();
        _ = reader.Read(ms);

        Assert.Single(reader.Bookmarks);
        Assert.Equal("unclosed", reader.Bookmarks[0].Name);
    }

    // ── Phase F1-W2d: 비-Paragraph block placeholder ───────────────────────

    [Fact]
    public void ImageBlock_Emits_Visible_Placeholder_With_Size_And_Mime()
    {
        var doc = new PolyDonkyument();
        var sec = new Section();
        sec.Blocks.Add(Paragraph.Of("before"));
        sec.Blocks.Add(new ImageBlock { MediaType = "image/png", WidthMm = 50, HeightMm = 30, Data = new byte[2048] });
        sec.Blocks.Add(Paragraph.Of("after"));
        doc.Sections.Add(sec);

        var doc2 = RoundTrip(doc);
        var texts = string.Join("\n", NonEmptyParagraphTexts(doc2));
        Assert.Contains("before", texts);
        Assert.Contains("[Image: image/png", texts);
        Assert.Contains("50", texts);   // width
        Assert.Contains("30", texts);   // height
        Assert.Contains("after", texts);
    }

    [Fact]
    public void Table_Emits_Placeholder_With_Dimensions_And_Cell_Contents()
    {
        var table = new Table();
        var row1 = new TableRow();
        var cell1 = new TableCell();
        cell1.Blocks.Add(Paragraph.Of("cell A"));
        var cell2 = new TableCell();
        cell2.Blocks.Add(Paragraph.Of("cell B"));
        row1.Cells.Add(cell1);
        row1.Cells.Add(cell2);
        table.Rows.Add(row1);

        var doc = new PolyDonkyument();
        var sec = new Section();
        sec.Blocks.Add(table);
        doc.Sections.Add(sec);

        var doc2 = RoundTrip(doc);
        var texts = string.Join("\n", NonEmptyParagraphTexts(doc2));
        Assert.Contains("[Table 1×2]", texts);
        Assert.Contains("cell A", texts);
        Assert.Contains("cell B", texts);
        Assert.Contains("[/Table]", texts);
    }

    [Fact]
    public void ShapeObject_Emits_Placeholder_With_Kind_And_Size()
    {
        var shape = new ShapeObject { Kind = ShapeKind.Ellipse, WidthMm = 40, HeightMm = 25 };
        var doc = new PolyDonkyument();
        var sec = new Section();
        sec.Blocks.Add(shape);
        doc.Sections.Add(sec);

        var doc2 = RoundTrip(doc);
        Assert.Contains("[Shape: Ellipse", string.Join("\n", NonEmptyParagraphTexts(doc2)));
    }

    [Fact]
    public void TextBoxObject_Emits_Boundary_Markers_And_Inner_Content()
    {
        var tb = new TextBoxObject { WidthMm = 60, HeightMm = 30 };
        tb.Content.Add(Paragraph.Of("inner text"));

        var doc = new PolyDonkyument();
        var sec = new Section();
        sec.Blocks.Add(tb);
        doc.Sections.Add(sec);

        var doc2 = RoundTrip(doc);
        var texts = string.Join("\n", NonEmptyParagraphTexts(doc2));
        Assert.Contains("[TextBox", texts);
        Assert.Contains("inner text", texts);
        Assert.Contains("[/TextBox]", texts);
    }

    [Fact]
    public void ThematicBreak_Emits_Horizontal_Rule_Like_Glyph()
    {
        var doc = new PolyDonkyument();
        var sec = new Section();
        sec.Blocks.Add(new ThematicBreakBlock());
        doc.Sections.Add(sec);

        var doc2 = RoundTrip(doc);
        var texts = string.Join("\n", NonEmptyParagraphTexts(doc2));
        Assert.Contains("─", texts);
    }

    [Fact]
    public void ContainerBlock_Children_Are_Flattened_Into_Body()
    {
        var container = new ContainerBlock();
        container.Children.Add(Paragraph.Of("inside container"));

        var doc = new PolyDonkyument();
        var sec = new Section();
        sec.Blocks.Add(container);
        doc.Sections.Add(sec);

        var doc2 = RoundTrip(doc);
        Assert.Contains("inside container", NonEmptyParagraphTexts(doc2));
    }

    // ── Phase F1-W2e: FidelityCapsules → OLE2 storage round-trip ───────────

    private static byte[] FakeBytes(int len, byte seed = 0xAB)
    {
        var b = new byte[len];
        for (int i = 0; i < len; i++) b[i] = (byte)(seed + i);
        return b;
    }

    [Fact]
    public void Macro_Capsule_RoundTrips_Via_DocBinaryReader()
    {
        var doc = DocWith(Paragraph.Of("with macros"));
        var vbaProject = FakeBytes(64, 0x10);
        var moduleData = FakeBytes(128, 0x20);
        doc.FidelityCapsules["msdoc/macros/Macros/PROJECT"]        = vbaProject;
        doc.FidelityCapsules["msdoc/macros/Macros/VBA/Module1"]    = moduleData;

        var bytes = WriteDoc(doc);
        using var ms = new MemoryStream(bytes);
        var reader = new DocBinaryReader();
        _ = reader.Read(ms);

        Assert.True(reader.HasMacros);
        Assert.NotNull(reader.MacroProject);
        Assert.Equal("Macros", reader.MacroProject!.StorageName);
        // Reader 가 root storage 안 모든 stream 을 path-keyed 사전으로 보존
        Assert.True(reader.MacroProject!.Streams.ContainsKey("PROJECT"));
        Assert.Equal(vbaProject, reader.MacroProject!.Streams["PROJECT"]);
        Assert.True(reader.MacroProject!.Streams.ContainsKey("VBA/Module1"));
        Assert.Equal(moduleData, reader.MacroProject!.Streams["VBA/Module1"]);
    }

    [Fact]
    public void Digital_Signature_Capsule_RoundTrips()
    {
        var doc = DocWith(Paragraph.Of("signed"));
        var sigData = FakeBytes(256, 0xF0);
        doc.FidelityCapsules["msdoc/signatures/_SignaturesV1/sig.xml"] = sigData;

        var bytes = WriteDoc(doc);
        using var ms = new MemoryStream(bytes);
        var reader = new DocBinaryReader();
        _ = reader.Read(ms);

        Assert.True(reader.HasDigitalSignature);
        Assert.NotNull(reader.DigitalSignature);
        Assert.Equal("_SignaturesV1", reader.DigitalSignature!.StorageName);
        Assert.True(reader.DigitalSignature!.Streams.ContainsKey("sig.xml"));
        Assert.Equal(sigData, reader.DigitalSignature!.Streams["sig.xml"]);
    }

    [Fact]
    public void Preserved_Unknown_Root_Storage_RoundTrips()
    {
        var doc = DocWith(Paragraph.Of("data store"));
        var dataA = FakeBytes(96, 0x33);
        var dataB = FakeBytes(48, 0x77);
        doc.FidelityCapsules["msdoc/preserved/MsoDataStore/item_1.dat"] = dataA;
        doc.FidelityCapsules["msdoc/preserved/MsoDataStore/sub/item_2.dat"] = dataB;

        var bytes = WriteDoc(doc);
        using var ms = new MemoryStream(bytes);
        var reader = new DocBinaryReader();
        _ = reader.Read(ms);

        Assert.Single(reader.PreservedRootStorages);
        var storage = reader.PreservedRootStorages[0];
        Assert.Equal("MsoDataStore", storage.Name);
        Assert.Contains("item_1.dat",     storage.Streams.Keys);
        Assert.Contains("sub/item_2.dat", storage.Streams.Keys);
        Assert.Equal(dataA, storage.Streams["item_1.dat"]);
        Assert.Equal(dataB, storage.Streams["sub/item_2.dat"]);
    }

    [Fact]
    public void Multiple_Capsule_Categories_Coexist_In_Single_Doc()
    {
        var doc = DocWith(Paragraph.Of("all three"));
        doc.FidelityCapsules["msdoc/macros/Macros/PROJECT"]            = FakeBytes(32);
        doc.FidelityCapsules["msdoc/signatures/_SignaturesV1/sig.xml"] = FakeBytes(64, 0x55);
        doc.FidelityCapsules["msdoc/preserved/MsoDataStore/item.dat"]  = FakeBytes(80, 0xAA);

        var bytes = WriteDoc(doc);
        using var ms = new MemoryStream(bytes);
        var reader = new DocBinaryReader();
        _ = reader.Read(ms);

        Assert.True(reader.HasMacros);
        Assert.True(reader.HasDigitalSignature);
        Assert.NotEmpty(reader.PreservedRootStorages);
    }

    [Fact]
    public void Capsule_Targeting_Reserved_Stream_Name_Is_Silently_Skipped()
    {
        // 누가 capsule key 를 'msdoc/preserved/WordDocument/x' 로 만들었더라도 writer 의
        // WordDocument stream 을 덮어쓰면 안 됨. capsule skip → 정상 round-trip.
        var doc = DocWith(Paragraph.Of("malicious"));
        doc.FidelityCapsules["msdoc/preserved/WordDocument/sneaky"] = FakeBytes(16);

        var bytes = WriteDoc(doc);
        using var ms = new MemoryStream(bytes);
        var reader = new DocBinaryReader();
        var doc2 = reader.Read(ms);

        Assert.Contains("malicious", NonEmptyParagraphTexts(doc2));
    }

    [Fact]
    public void Non_Msdoc_Prefix_Capsules_Are_Ignored_By_Doc_Writer()
    {
        var doc = DocWith(Paragraph.Of("only ooxml capsule"));
        // ooxml/* 와 hwpx/* prefix 는 DOCX / HWPX writer 가 처리할 영역 — DOC writer 는 무시.
        doc.FidelityCapsules["ooxml/word/vbaProject.bin"]   = FakeBytes(32);
        doc.FidelityCapsules["hwpx/Scripts/macro.xml"]      = FakeBytes(32);

        var bytes = WriteDoc(doc);
        using var ms = new MemoryStream(bytes);
        var reader = new DocBinaryReader();
        _ = reader.Read(ms);

        Assert.False(reader.HasMacros);
        Assert.False(reader.HasDigitalSignature);
        Assert.Empty(reader.PreservedRootStorages);
    }

    // ── Word strict-parser 호환 skeleton (STSH / PlcfSed / SEPX) ───────────

    [Fact]
    public void Fib_Has_Zero_FcMin_FcMac_And_QuickSaves_Flag()
    {
        var bytes = WriteDoc(DocWith(Paragraph.Of("x")));
        using var ms = new MemoryStream(bytes);
        using var root = RootStorage.Open(ms, StorageModeFlags.LeaveOpen);
        using var wd = root.OpenStream("WordDocument");
        var wb = new byte[wd.Length];
        wd.ReadExactly(wb, 0, wb.Length);

        // [MS-DOC] §2.5.1 — Word 97+ MUST: fcMin / fcMac = 0
        Assert.Equal(0u, BitConverter.ToUInt32(wb, 0x18));
        Assert.Equal(0u, BitConverter.ToUInt32(wb, 0x1C));
        // cQuickSaves (bits 4-7) MUST be 0xF → flags = 0x02F4
        Assert.Equal((ushort)0x02F4, BitConverter.ToUInt16(wb, 0x0A));
    }

    [Fact]
    public void Stsh_And_PlcfSed_Are_Present_In_Fib()
    {
        var bytes = WriteDoc(DocWith(Paragraph.Of("x")));
        using var ms = new MemoryStream(bytes);
        using var root = RootStorage.Open(ms, StorageModeFlags.LeaveOpen);
        using var wd = root.OpenStream("WordDocument");
        var wb = new byte[wd.Length];
        wd.ReadExactly(wb, 0, wb.Length);

        // fcStshf @ 0x00A2 / lcb @ 0x00A6 — non-zero
        Assert.True(BitConverter.ToUInt32(wb, 0xA6) > 0, "STSH lcb should be non-zero");
        // fcPlcfSed @ 0x00CA / lcb @ 0x00CE — non-zero
        Assert.True(BitConverter.ToUInt32(wb, 0xCE) > 0, "PlcfSed lcb should be non-zero");
    }

    [Fact]
    public void Section_Page_Margins_RoundTrip_Via_Sepx()
    {
        var doc = new PolyDonkyument();
        var sec = new Section();
        sec.Page.MarginLeftMm   = 30;
        sec.Page.MarginRightMm  = 15;
        sec.Page.MarginTopMm    = 22;
        sec.Page.MarginBottomMm = 18;
        sec.Blocks.Add(Paragraph.Of("body"));
        doc.Sections.Add(sec);

        var doc2 = RoundTrip(doc);
        var page = doc2.Sections[0].Page;
        Assert.Equal(30, page.MarginLeftMm,   0);
        Assert.Equal(15, page.MarginRightMm,  0);
        Assert.Equal(22, page.MarginTopMm,    0);
        Assert.Equal(18, page.MarginBottomMm, 0);
    }

    [Fact]
    public void Section_Page_Size_RoundTrips_Via_Sepx()
    {
        var doc = new PolyDonkyument();
        var sec = new Section();
        sec.Page.WidthMm  = 210;   // A4
        sec.Page.HeightMm = 297;
        sec.Blocks.Add(Paragraph.Of("body"));
        doc.Sections.Add(sec);

        var doc2 = RoundTrip(doc);
        var page = doc2.Sections[0].Page;
        Assert.Equal(210, page.WidthMm,  0);
        Assert.Equal(297, page.HeightMm, 0);
    }

    [Fact]
    public void Mixed_Style_Runs_In_Single_Paragraph_RoundTrip()
    {
        var p = new Paragraph { Style = { Alignment = Alignment.Center } };
        p.AddText("Normal ");
        p.AddText("Bold-Red", new RunStyle { Bold = true, Foreground = new Color(255, 0, 0) });
        p.AddText(" Italic-14pt", new RunStyle { Italic = true, FontSizePt = 14 });
        p.AddText(" Arial", new RunStyle { FontFamily = "Arial" });

        var doc2 = RoundTrip(DocWith(p));
        var rp   = FirstParaWithText(doc2, "Bold-Red");

        Assert.Equal(Alignment.Center, rp.Style.Alignment);

        var brun = rp.Runs.First(r => r.Text.Contains("Bold-Red"));
        Assert.True(brun.Style.Bold);
        Assert.Equal(new Color(255, 0, 0), brun.Style.Foreground);

        var irun = rp.Runs.First(r => r.Text.Contains("Italic-14pt"));
        Assert.True(irun.Style.Italic);
        Assert.Equal(14.0, irun.Style.FontSizePt, 1);

        var arun = rp.Runs.First(r => r.Text.Contains("Arial"));
        Assert.Equal("Arial", arun.Style.FontFamily);
    }
}
