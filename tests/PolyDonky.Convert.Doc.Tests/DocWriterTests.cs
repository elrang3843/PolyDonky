using System.Text;
using PolyDonky.Convert.Doc;
using PolyDonky.Core;

namespace PolyDonky.Convert.Doc.Tests;

/// <summary>
/// DocWriter (IWPF → RTF) 출력 검증.
/// RTF 는 텍스트 기반이라 출력 문자열 내 control word 존재 여부로 검증한다.
/// </summary>
public class DocWriterTests
{
    private static string WriteRtf(PolyDonkyument doc)
    {
        using var ms = new MemoryStream();
        new DocWriter().Write(doc, ms);
        // RTF 출력은 모든 비-ASCII 가 \uNNNN? 로 이스케이프되어 순수 ASCII 이므로 Latin1 로 안전하게 디코딩.
        return Encoding.Latin1.GetString(ms.ToArray());
    }

    private static PolyDonkyument DocWith(params Block[] blocks)
    {
        var doc = new PolyDonkyument();
        var sec = new Section();
        foreach (var b in blocks) sec.Blocks.Add(b);
        doc.Sections.Add(sec);
        return doc;
    }

    // ── 기본 RTF 구조 ──────────────────────────────────────────────────────────

    [Fact]
    public void Header_Has_Rtf_Signature_And_Font_And_Color_Tables()
    {
        var doc = DocWith(Paragraph.Of("hello"));
        string rtf = WriteRtf(doc);

        Assert.StartsWith(@"{\rtf1\ansi", rtf);
        Assert.Contains(@"\fonttbl", rtf);
        Assert.Contains(@"\colortbl", rtf);
        Assert.EndsWith("}", rtf.TrimEnd());
    }

    [Fact]
    public void Plain_Paragraph_Emits_Text_With_Pard_And_Par()
    {
        var rtf = WriteRtf(DocWith(Paragraph.Of("hello world")));
        Assert.Contains(@"\pard", rtf);
        Assert.Contains("hello world", rtf);
        Assert.Contains(@"\par", rtf);
    }

    [Fact]
    public void Run_Styles_Bold_Italic_Underline_Emit_Toggles()
    {
        var p = new Paragraph();
        p.AddText("B", new RunStyle { Bold = true });
        p.AddText("I", new RunStyle { Italic = true });
        p.AddText("U", new RunStyle { Underline = true });
        p.AddText("S", new RunStyle { Strikethrough = true });

        var rtf = WriteRtf(DocWith(p));
        Assert.Contains(@"\b ", rtf);
        Assert.Contains(@"\b0", rtf);
        Assert.Contains(@"\i ", rtf);
        Assert.Contains(@"\i0", rtf);
        Assert.Contains(@"\ul", rtf);
        Assert.Contains(@"\ulnone", rtf);
        Assert.Contains(@"\strike", rtf);
    }

    [Fact]
    public void Empty_Document_Produces_Valid_Skeleton()
    {
        var rtf = WriteRtf(new PolyDonkyument());
        Assert.StartsWith(@"{\rtf1\ansi", rtf);
        Assert.EndsWith("}", rtf.TrimEnd());
    }

    // ── 하이퍼링크 ─────────────────────────────────────────────────────────────

    [Fact]
    public void Hyperlink_Run_Emits_Field_With_HYPERLINK_Instruction()
    {
        var p = new Paragraph();
        p.Runs.Add(new Run { Text = "click here", Url = "https://example.com/page?x=1" });
        var rtf = WriteRtf(DocWith(p));

        Assert.Contains(@"{\field{\*\fldinst HYPERLINK ""https://example.com/page?x=1""}", rtf);
        Assert.Contains(@"{\fldrslt", rtf);
        Assert.Contains("click here", rtf);
    }

    [Fact]
    public void Hyperlink_With_Bold_Run_Wraps_Inside_Fldrslt()
    {
        var p = new Paragraph();
        p.Runs.Add(new Run
        {
            Text = "PolyDonky",
            Url = "https://github.com/elrang3843/polydonky",
            Style = { Bold = true }
        });
        var rtf = WriteRtf(DocWith(p));

        // 하이퍼링크 텍스트 안에 굵게 토글이 적용되어야 함
        int instStart = rtf.IndexOf(@"\fldrslt", StringComparison.Ordinal);
        Assert.True(instStart > 0);
        Assert.Contains(@"\b ", rtf[instStart..]);
        Assert.Contains("PolyDonky", rtf[instStart..]);
    }

    // ── 자동 필드 ─────────────────────────────────────────────────────────────

    [Theory]
    [InlineData(FieldType.Page,      "PAGE")]
    [InlineData(FieldType.NumPages,  "NUMPAGES")]
    [InlineData(FieldType.Date,      "DATE")]
    [InlineData(FieldType.Time,      "TIME")]
    [InlineData(FieldType.Author,    "AUTHOR")]
    [InlineData(FieldType.Title,     "TITLE")]
    [InlineData(FieldType.NumChars,  "NUMCHARS")]
    [InlineData(FieldType.FileName,  "FILENAME")]
    [InlineData(FieldType.Subject,   "SUBJECT")]
    [InlineData(FieldType.Keywords,  "KEYWORDS")]
    [InlineData(FieldType.Comments,  "COMMENTS")]
    [InlineData(FieldType.If,        "IF")]
    public void Field_Run_Emits_Expected_Instruction(FieldType type, string keyword)
    {
        var p = new Paragraph();
        p.Runs.Add(new Run { Text = "1", Field = type });
        var rtf = WriteRtf(DocWith(p));

        Assert.Contains($@"\fldinst {keyword}", rtf);
    }

    [Fact]
    public void Field_Seq_Run_Preserves_Category_Arg()
    {
        var p = new Paragraph();
        p.Runs.Add(new Run { Text = "3", Field = FieldType.Seq, FieldArg = "Figure" });
        var rtf = WriteRtf(DocWith(p));

        Assert.Contains(@"\fldinst SEQ Figure", rtf);
    }

    [Fact]
    public void Field_StyleRef_Run_Quotes_Argument()
    {
        var p = new Paragraph();
        p.Runs.Add(new Run { Text = "Chapter Title", Field = FieldType.StyleRef, FieldArg = "Heading 1" });
        var rtf = WriteRtf(DocWith(p));

        Assert.Contains(@"\fldinst STYLEREF ""Heading 1""", rtf);
    }

    [Fact]
    public void Empty_Field_Run_Still_Emits_Field_Group()
    {
        var p = new Paragraph();
        p.Runs.Add(new Run { Text = string.Empty, Field = FieldType.Page });
        var rtf = WriteRtf(DocWith(p));

        Assert.Contains(@"\fldinst PAGE", rtf);
    }

    // ── 책갈피 ────────────────────────────────────────────────────────────────

    [Fact]
    public void Bookmark_Start_End_Runs_Emit_Markers()
    {
        var p = new Paragraph();
        p.Runs.Add(new Run { Text = "",      BookmarkStart = "chapter_1" });
        p.Runs.Add(new Run { Text = "Chapter 1" });
        p.Runs.Add(new Run { Text = "",      BookmarkEnd   = "chapter_1" });
        var rtf = WriteRtf(DocWith(p));

        Assert.Contains(@"{\*\bkmkstart chapter_1}", rtf);
        Assert.Contains(@"{\*\bkmkend chapter_1}",   rtf);

        // 마커 위치 — start 가 본문보다 앞, end 가 본문보다 뒤에 와야 함
        int s = rtf.IndexOf(@"\bkmkstart", StringComparison.Ordinal);
        int t = rtf.IndexOf("Chapter 1", StringComparison.Ordinal);
        int e = rtf.IndexOf(@"\bkmkend",   StringComparison.Ordinal);
        Assert.True(s < t && t < e);
    }

    [Fact]
    public void Bookmark_Name_Strips_Non_Ascii_And_Whitespace()
    {
        var p = new Paragraph();
        p.Runs.Add(new Run { Text = "", BookmarkStart = "북마크 1" });
        p.Runs.Add(new Run { Text = "x" });
        p.Runs.Add(new Run { Text = "", BookmarkEnd   = "북마크 1" });
        var rtf = WriteRtf(DocWith(p));

        // 한글·공백은 제거 → "1" 만 남음
        Assert.Contains(@"\bkmkstart 1}", rtf);
        Assert.Contains(@"\bkmkend 1}",   rtf);
    }

    [Fact]
    public void Bookmark_Empty_Name_Falls_Back_To_Default()
    {
        var p = new Paragraph();
        p.Runs.Add(new Run { Text = "", BookmarkStart = "!!!" });
        p.Runs.Add(new Run { Text = "x" });
        p.Runs.Add(new Run { Text = "", BookmarkEnd   = "!!!" });
        var rtf = WriteRtf(DocWith(p));

        Assert.Contains(@"\bkmkstart bm}", rtf);
        Assert.Contains(@"\bkmkend bm}",   rtf);
    }

    [Fact]
    public void Run_With_Bookmark_And_Hyperlink_Emits_Both_Markers()
    {
        var p = new Paragraph();
        p.Runs.Add(new Run
        {
            Text          = "anchor",
            Url           = "https://example.com",
            BookmarkStart = "ref1",
            BookmarkEnd   = "ref1",
        });
        var rtf = WriteRtf(DocWith(p));

        int s = rtf.IndexOf(@"\bkmkstart ref1", StringComparison.Ordinal);
        int h = rtf.IndexOf("HYPERLINK",         StringComparison.Ordinal);
        int e = rtf.IndexOf(@"\bkmkend ref1",   StringComparison.Ordinal);
        Assert.True(s > 0);
        Assert.True(h > 0);
        Assert.True(e > 0);
        Assert.True(s < h && h < e);
    }

    // ── 변경추적 ─────────────────────────────────────────────────────────────

    [Fact]
    public void Inserted_Run_Wraps_In_Revised_RevAuth_Group()
    {
        var p = new Paragraph();
        p.Runs.Add(new Run
        {
            Text                = "added",
            IsInsertedRevision  = true,
            RevisionAuthor      = "Alice",
        });
        var rtf = WriteRtf(DocWith(p));

        Assert.Contains(@"{\*\revtbl", rtf);
        Assert.Contains("{Alice;}",     rtf);
        Assert.Contains(@"\revised\revauth1", rtf);
        Assert.Contains("added",        rtf);
    }

    [Fact]
    public void Deleted_Run_Wraps_In_Deleted_RevAuthDel_Group()
    {
        var p = new Paragraph();
        p.Runs.Add(new Run
        {
            Text               = "gone",
            IsDeletedRevision  = true,
            RevisionAuthor     = "Bob",
        });
        var rtf = WriteRtf(DocWith(p));

        Assert.Contains(@"\deleted\revauthdel1", rtf);
        Assert.Contains("{Bob;}", rtf);
    }

    [Fact]
    public void Revision_Without_Author_Uses_Index_0_Unknown()
    {
        var p = new Paragraph();
        p.Runs.Add(new Run { Text = "x", IsInsertedRevision = true });
        var rtf = WriteRtf(DocWith(p));

        Assert.Contains(@"\revised\revauth0", rtf);
        // 작성자가 Unknown 한 명뿐이면 revtbl 자체를 생략
        Assert.DoesNotContain(@"\*\revtbl", rtf);
    }

    [Fact]
    public void Revision_With_Date_Emits_RevDttm()
    {
        var p = new Paragraph();
        p.Runs.Add(new Run
        {
            Text               = "x",
            IsInsertedRevision = true,
            RevisionAuthor     = "Alice",
            RevisionDate       = new DateTimeOffset(2026, 5, 27, 10, 30, 0, TimeSpan.Zero),
        });
        var rtf = WriteRtf(DocWith(p));

        Assert.Matches(@"\\revised\\revauth1\\revdttm-?\d+", rtf);
    }

    [Fact]
    public void Revision_Author_Table_Indexes_Are_Stable()
    {
        var p = new Paragraph();
        p.Runs.Add(new Run { Text = "a", IsInsertedRevision = true, RevisionAuthor = "Alice" });
        p.Runs.Add(new Run { Text = "b", IsInsertedRevision = true, RevisionAuthor = "Bob" });
        p.Runs.Add(new Run { Text = "c", IsInsertedRevision = true, RevisionAuthor = "Alice" });
        var rtf = WriteRtf(DocWith(p));

        // Alice = 1, Bob = 2 (Unknown = 0)
        Assert.Contains("{Unknown;}{Alice;}{Bob;}", rtf);
        Assert.Equal(2, System.Text.RegularExpressions.Regex.Matches(rtf, @"\\revauth1\b").Count);
        Assert.Single(System.Text.RegularExpressions.Regex.Matches(rtf, @"\\revauth2\b"));
    }

    [Fact]
    public void Paragraph_Inserted_Revision_Emits_Revised_Before_Par()
    {
        var p = new Paragraph { IsInsertedRevision = true };
        p.AddText("new line");
        var rtf = WriteRtf(DocWith(p));

        int revIdx = rtf.LastIndexOf(@"\revised\revauth0", StringComparison.Ordinal);
        int parIdx = rtf.IndexOf(@"\par", revIdx, StringComparison.Ordinal);
        Assert.True(revIdx > 0);
        Assert.True(parIdx > revIdx);
    }

    [Fact]
    public void Paragraph_Deleted_Revision_Emits_Deleted_Before_Par()
    {
        var p = new Paragraph { IsDeletedRevision = true };
        p.AddText("gone");
        var rtf = WriteRtf(DocWith(p));

        Assert.Contains(@"\deleted\revauthdel0", rtf);
    }

    // ── 각주 / 미주 ───────────────────────────────────────────────────────────

    [Fact]
    public void Footnote_Ref_Emits_Inline_Footnote_With_Body()
    {
        var doc = new PolyDonkyument();
        doc.Sections.Add(new Section());
        var p = new Paragraph();
        p.AddText("Body text");
        p.Runs.Add(new Run { FootnoteId = "fn1" });
        doc.Sections[0].Blocks.Add(p);

        var note = new FootnoteEntry { Id = "fn1" };
        note.Blocks.Add(Paragraph.Of("This is the footnote."));
        doc.Footnotes.Add(note);

        var rtf = WriteRtf(doc);

        Assert.Contains(@"{\super\chftn", rtf);
        Assert.Contains(@"{\*\footnote", rtf);
        Assert.Contains("This is the footnote.", rtf);
        Assert.DoesNotContain(@"\ftnalt", rtf);  // 각주이므로 미주 marker 없어야 함
    }

    [Fact]
    public void Endnote_Ref_Emits_Footnote_With_FtnAlt_Marker()
    {
        var doc = new PolyDonkyument();
        doc.Sections.Add(new Section());
        var p = new Paragraph();
        p.AddText("Body");
        p.Runs.Add(new Run { EndnoteId = "en1" });
        doc.Sections[0].Blocks.Add(p);

        var note = new FootnoteEntry { Id = "en1" };
        note.Blocks.Add(Paragraph.Of("This is an endnote."));
        doc.Endnotes.Add(note);

        var rtf = WriteRtf(doc);

        Assert.Contains(@"\ftnalt", rtf);
        Assert.Contains("This is an endnote.", rtf);
    }

    [Fact]
    public void Dangling_Footnote_Ref_Still_Emits_Marker()
    {
        var p = new Paragraph();
        p.AddText("x");
        p.Runs.Add(new Run { FootnoteId = "fnX" });  // 본문 없음
        var rtf = WriteRtf(DocWith(p));

        Assert.Contains(@"\chftn", rtf);
    }

    [Fact]
    public void Footnote_Body_Does_Not_Recurse_Into_Its_Own_Footnote_Ref()
    {
        var doc = new PolyDonkyument();
        doc.Sections.Add(new Section());
        var p = new Paragraph();
        p.AddText("x");
        p.Runs.Add(new Run { FootnoteId = "fn1" });
        doc.Sections[0].Blocks.Add(p);

        // 각주 본문 안에 또 다른 각주 ref — 무한 재귀 방지 확인
        var note = new FootnoteEntry { Id = "fn1" };
        var noteBody = new Paragraph();
        noteBody.AddText("inner ");
        noteBody.Runs.Add(new Run { FootnoteId = "fn2" });
        note.Blocks.Add(noteBody);
        doc.Footnotes.Add(note);

        var rtf = WriteRtf(doc);  // 무한 재귀가 있으면 StackOverflow 로 떨어짐
        Assert.Contains("inner", rtf);
    }

    // ── 주석 (annotations) ───────────────────────────────────────────────────

    [Fact]
    public void Comment_Ref_Emits_Atnauthor_And_Annotation()
    {
        var doc = new PolyDonkyument();
        doc.Sections.Add(new Section());
        var p = new Paragraph();
        p.AddText("Body ");
        p.Runs.Add(new Run { CommentId = "cmt1" });
        doc.Sections[0].Blocks.Add(p);

        var cm = new CommentEntry
        {
            Id     = "cmt1",
            Author = "Charlie",
            Date   = new DateTimeOffset(2026, 1, 1, 9, 0, 0, TimeSpan.Zero),
        };
        cm.Blocks.Add(Paragraph.Of("Please review."));
        doc.Comments.Add(cm);

        var rtf = WriteRtf(doc);

        Assert.Contains(@"{\*\atnauthor Charlie}", rtf);
        Assert.Contains(@"{\*\atnid Charlie}",     rtf);
        Assert.Matches(@"\\\*\\atndate -?\d+",     rtf);
        Assert.Contains(@"\chatn",                 rtf);
        Assert.Contains(@"{\*\annotation",         rtf);
        Assert.Contains("Please review.",          rtf);
    }

    [Fact]
    public void Comment_Without_Author_Skips_Atnauthor_Headers()
    {
        var doc = new PolyDonkyument();
        doc.Sections.Add(new Section());
        var p = new Paragraph();
        p.Runs.Add(new Run { Text = "x", CommentId = "c1" });
        doc.Sections[0].Blocks.Add(p);

        var cm = new CommentEntry { Id = "c1" };
        cm.Blocks.Add(Paragraph.Of("anon"));
        doc.Comments.Add(cm);

        var rtf = WriteRtf(doc);

        Assert.DoesNotContain(@"\atnauthor", rtf);
        Assert.Contains(@"\chatn", rtf);
        Assert.Contains("anon",    rtf);
    }

    // ── 페이지 설정 ───────────────────────────────────────────────────────────

    [Fact]
    public void Default_A4_Portrait_Page_Settings_Are_Emitted()
    {
        var doc = DocWith(Paragraph.Of("body"));
        var rtf = WriteRtf(doc);

        // A4 portrait: 210×297mm → 11905×16838 twips (constant 56.692)
        Assert.Contains(@"\paperw11905",  rtf);
        Assert.Contains(@"\paperh16838",  rtf);
        Assert.Contains(@"\margl",        rtf);
        Assert.Contains(@"\margr",        rtf);
        Assert.Contains(@"\margt",        rtf);
        Assert.Contains(@"\margb",        rtf);
        Assert.Contains(@"\sectd",        rtf);
        Assert.DoesNotContain(@"\landscape",  rtf);
        Assert.DoesNotContain(@"\lndscpsxn",  rtf);
    }

    [Fact]
    public void Landscape_Orientation_Emits_Landscape_Keyword()
    {
        var doc = new PolyDonkyument();
        var sec = new Section();
        sec.Page.Orientation = PageOrientation.Landscape;
        sec.Blocks.Add(Paragraph.Of("x"));
        doc.Sections.Add(sec);

        var rtf = WriteRtf(doc);

        Assert.Contains(@"\landscape", rtf);
        Assert.Contains(@"\lndscpsxn", rtf);
        // 가로 모드에서 EffectiveWidth = HeightMm, EffectiveHeight = WidthMm
        Assert.Contains(@"\paperw16838", rtf);  // 297mm
        Assert.Contains(@"\paperh11905", rtf);  // 210mm
    }

    [Fact]
    public void Custom_Margins_Emit_Twip_Values()
    {
        var doc = new PolyDonkyument();
        var sec = new Section();
        sec.Page.MarginLeftMm   = 30;
        sec.Page.MarginRightMm  = 15;
        sec.Page.MarginTopMm    = 25;
        sec.Page.MarginBottomMm = 18;
        sec.Blocks.Add(Paragraph.Of("x"));
        doc.Sections.Add(sec);

        var rtf = WriteRtf(doc);

        Assert.Contains(@"\margl1701", rtf);   // 30mm
        Assert.Contains(@"\margr850",  rtf);   // 15mm
        Assert.Contains(@"\margt1417", rtf);   // 25mm
        Assert.Contains(@"\margb1020", rtf);   // 18mm
    }

    [Fact]
    public void Different_First_Page_Emits_Titlepg_In_Section()
    {
        var doc = new PolyDonkyument();
        var sec = new Section();
        sec.Page.DifferentFirstPage = true;
        sec.Blocks.Add(Paragraph.Of("x"));
        doc.Sections.Add(sec);

        Assert.Contains(@"\titlepg", WriteRtf(doc));
    }

    [Fact]
    public void Different_Odd_Even_Emits_Facingp_At_Document_Level()
    {
        var doc = new PolyDonkyument();
        var sec = new Section();
        sec.Page.DifferentOddEven = true;
        sec.Blocks.Add(Paragraph.Of("x"));
        doc.Sections.Add(sec);

        Assert.Contains(@"\facingp", WriteRtf(doc));
    }

    [Fact]
    public void Multi_Column_Section_Emits_Cols_And_Colsx()
    {
        var doc = new PolyDonkyument();
        var sec = new Section();
        sec.Page.ColumnCount = 2;
        sec.Page.ColumnGapMm = 10;
        sec.Blocks.Add(Paragraph.Of("x"));
        doc.Sections.Add(sec);

        var rtf = WriteRtf(doc);
        Assert.Contains(@"\cols2",  rtf);
        Assert.Contains(@"\colsx",  rtf);
    }

    [Fact]
    public void Multiple_Sections_Emit_Sect_Separator()
    {
        var doc = new PolyDonkyument();
        var s1 = new Section();
        s1.Blocks.Add(Paragraph.Of("first"));
        var s2 = new Section();
        s2.Blocks.Add(Paragraph.Of("second"));
        doc.Sections.Add(s1);
        doc.Sections.Add(s2);

        var rtf = WriteRtf(doc);

        int first = rtf.IndexOf("first",  StringComparison.Ordinal);
        int sect  = rtf.IndexOf(@"\sect\sectd", StringComparison.Ordinal);
        int second = rtf.IndexOf("second", StringComparison.Ordinal);
        Assert.True(first  > 0);
        Assert.True(sect   > first);
        Assert.True(second > sect);
    }

    // ── 머리말 / 꼬리말 ──────────────────────────────────────────────────────

    [Fact]
    public void Header_With_Center_Slot_Emits_Header_Group()
    {
        var doc = new PolyDonkyument();
        var sec = new Section();
        sec.Page.Header.Center.Paragraphs.Add(Paragraph.Of("페이지 머리말"));
        sec.Blocks.Add(Paragraph.Of("본문"));
        doc.Sections.Add(sec);

        var rtf = WriteRtf(doc);

        Assert.Contains(@"{\header", rtf);
        Assert.Contains(@"\qc",      rtf);  // Center 슬롯은 가운데 정렬
    }

    [Fact]
    public void Footer_With_Three_Slots_Emits_All_Alignments()
    {
        var doc = new PolyDonkyument();
        var sec = new Section();
        sec.Page.Footer.Left.Paragraphs.Add(Paragraph.Of("LEFT"));
        sec.Page.Footer.Center.Paragraphs.Add(Paragraph.Of("MID"));
        sec.Page.Footer.Right.Paragraphs.Add(Paragraph.Of("RIGHT"));
        sec.Blocks.Add(Paragraph.Of("body"));
        doc.Sections.Add(sec);

        var rtf = WriteRtf(doc);
        int footer = rtf.IndexOf(@"{\footer", StringComparison.Ordinal);
        Assert.True(footer > 0);
        string ft = rtf[footer..];
        Assert.Contains("LEFT",  ft);
        Assert.Contains("MID",   ft);
        Assert.Contains("RIGHT", ft);
        // 슬롯별 정렬 control word 확인
        Assert.Contains(@"\ql",  ft);
        Assert.Contains(@"\qc",  ft);
        Assert.Contains(@"\qr",  ft);
    }

    [Fact]
    public void Header_Empty_Does_Not_Emit_Header_Group()
    {
        var doc = DocWith(Paragraph.Of("x"));
        var rtf = WriteRtf(doc);
        Assert.DoesNotContain(@"{\header", rtf);
        Assert.DoesNotContain(@"{\footer", rtf);
    }

    // ── 글상자 / 부유 이미지 / 머리말 토큰 ─────────────────────────────────────

    [Fact]
    public void TextBox_Emits_Shp_TextBox_With_Content()
    {
        var tb = new TextBoxObject { OverlayXMm = 20, OverlayYMm = 180, WidthMm = 50, HeightMm = 25 };
        tb.Content.Clear();
        tb.Content.Add(Paragraph.Of("box text"));
        var rtf = WriteRtf(DocWith(tb));

        Assert.Contains(@"{\sv 202}", rtf);     // shapeType 202 = text box
        Assert.Contains(@"\shptxt",   rtf);
        Assert.Contains("box text",   rtf);
    }

    [Fact]
    public void Floating_Image_Emits_Page_Positioned_Frame()
    {
        var doc = new PolyDonkyument();
        var sec = new Section();
        sec.Blocks.Add(new ImageBlock
        {
            MediaType = "image/png",
            Data      = new byte[] { 1, 2, 3, 4 },
            WidthMm   = 30, HeightMm = 20,
            WrapMode  = ImageWrapMode.InFrontOfText,
            OverlayXMm = 120, OverlayYMm = 90,
        });
        doc.Sections.Add(sec);
        var rtf = WriteRtf(doc);

        Assert.Contains(@"\phpg\pvpg", rtf);   // page-relative absolute frame
        Assert.Contains(@"\posx",      rtf);
        Assert.Contains(@"\posy",      rtf);
        Assert.Contains(@"\pict",      rtf);
    }

    [Fact]
    public void Inline_Image_Stays_Inline_Pict()
    {
        var doc = new PolyDonkyument();
        var sec = new Section();
        sec.Blocks.Add(new ImageBlock
        {
            MediaType = "image/png", Data = new byte[] { 1, 2, 3, 4 },
            WidthMm = 30, HeightMm = 20, WrapMode = ImageWrapMode.Inline,
        });
        doc.Sections.Add(sec);
        var rtf = WriteRtf(doc);

        Assert.Contains(@"\pict", rtf);
        Assert.DoesNotContain(@"\phpg\pvpg", rtf);   // 인라인은 frame 안 씀
    }

    [Theory]
    [InlineData("{PAGE}",     "PAGE")]
    [InlineData("{페이지}",    "PAGE")]
    [InlineData("{NUMPAGES}", "NUMPAGES")]
    [InlineData("{전체페이지}", "NUMPAGES")]
    [InlineData("{DATE}",     "DATE")]
    [InlineData("{TITLE}",    "TITLE")]
    public void Footer_Token_Becomes_Rtf_Field(string token, string keyword)
    {
        var doc = new PolyDonkyument();
        var sec = new Section();
        sec.Page.Footer.Center.Paragraphs.Add(Paragraph.Of(token));
        sec.Blocks.Add(Paragraph.Of("body"));
        doc.Sections.Add(sec);
        var rtf = WriteRtf(doc);

        Assert.Contains($@"\fldinst {keyword}", rtf);
        Assert.DoesNotContain(token, rtf);   // 리터럴 토큰이 남으면 안 됨
    }

    [Fact]
    public void Header_Token_Mixed_With_Literal_Text_Splits_Correctly()
    {
        var doc = new PolyDonkyument();
        var sec = new Section();
        sec.Page.Header.Center.Paragraphs.Add(Paragraph.Of("Page {PAGE} of {NUMPAGES}"));
        sec.Blocks.Add(Paragraph.Of("body"));
        doc.Sections.Add(sec);
        var rtf = WriteRtf(doc);

        Assert.Contains("Page",            rtf);   // 리터럴 보존
        Assert.Contains(" of ",            rtf);
        Assert.Contains(@"\fldinst PAGE",     rtf);
        Assert.Contains(@"\fldinst NUMPAGES", rtf);
    }
}
