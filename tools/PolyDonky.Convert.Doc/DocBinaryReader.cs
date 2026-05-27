using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using OpenMcdf;
using PolyDonky.Core;

namespace PolyDonky.Convert.Doc;

/// <summary>
/// Word 97-2003 binary (.doc, OLE2 Compound File) → IWPF 변환기 (Phase 1a~1g — 텍스트·단락·서식·폰트·스타일·체인 상속).
///
/// 참고 공식 공개 명세:
///   - [MS-DOC]      Word (.doc) Binary File Format — FIB·CLX·piece table·PAPX/CHPX·SEPX 등
///   - [MS-CFB]      Compound File Binary File Format — OLE2 컨테이너 (OpenMcdf 가 처리)
///   - [MS-OLEPS]    Object Linking and Embedding (OLE) Property Set Format — SummaryInformation
///                   / DocumentSummaryInformation 의 PID·Variant 인코딩
///   - [MS-OFFCRYPTO] Office Document Cryptography Structure — fEncrypted 비트 / EncryptionInfo
///                   stream / RC4·AES 키 유도 (Phase 1a 는 인식 후 거부만)
///
/// 본 단계 범위:
///   - FIB(File Information Block) 파싱 + 암호화/obfuscation 사전 거부 (MS-OFFCRYPTO)
///   - CLX / Piece Table 따라 WordDocument stream 에서 텍스트 추출 (MS-DOC §2.3, §2.8)
///   - 단락 마커 0x0D 기준 단락 분리; 표 셀 마커 0x07 / 줄바꿈 0x0B / 페이지 break 0x0C /
///     필드 제어 0x13..0x15 등은 공백·줄바꿈으로 정규화하거나 폐기
///   - PAPX 운영 → 정렬(sprmPJc80 0x2461) / 들여쓰기(sprmPDxaLeft 0x845D, sprmPDxaRight 0x845E,
///     sprmPDxaLeft1 0x8460) / 단락 간격(sprmPDyaBefore 0xA413, sprmPDyaAfter 0xA415) /
///     줄 간격(sprmPDyaLine 0x6412 LSPD) / istd → STSH lookup → OutlineLevel(H1..H6)  (Phase 1b/1c/1e)
///   - CHPX 운영 → 굵게(0x0835)·이탤릭(0x0836)·취소선(0x0837)·밑줄(0x2A3E)·
///     글자 크기(sprmCHps 0x4A43) / 전경색(sprmCIco 0x2A42 팔레트, sprmCCv 0x6870 RGB) /
///     하이라이트(sprmCHighlight 0x2A0C) / 폰트 패밀리(sprmCRgFtc0/1/2 0x4A4F/0x4A50/0x4A51
///     + STTB FFN 운영) 로 Run 분할                                                  (Phase 1b/1c/1d)
///   - SummaryInformation PropertySet (MS-OLEPS §2.18) 에서 Title/Subject/Author/Keywords/
///     LastSavedBy/Created/Modified/AppName 추출 (VT_LPSTR/VT_LPWSTR/VT_FILETIME)
///   - 비-OLE2 입력·손상 CFB 컨테이너는 한국어 진단으로 명확히 거부
///
/// 의도적으로 다음 작업의 범위 밖 (다음 PR 들에서 다룸):
///   - 표 / 이미지 / 필드 / 헤더·푸터 / 섹션                                          (Phase 2)
///   - MS-OFFCRYPTO 의 실제 해독 (RC4/AES 키 유도·복호화)
///   - 리스트·번호 (LST/LFO) 와 STSH 의 stk=4 numbering 스타일                       (Phase 3)
///
/// 자체 구현인 이유는 CLAUDE.md: HWP 코덱과 마찬가지로 비-OSS 상용 라이브러리(Aspose 등) 의존을 피한다.
/// </summary>
public class DocBinaryReader
{
    /// <summary>
    /// Phase 3f — 마지막 Read 에서 추출한 OfficeArtBStoreContainer 의 BLIP 목록.
    /// FBSE 가 가리키는 공유 이미지들. 인라인 PICF 와 무관하게 문서 전역에서 참조 가능한 자원.
    /// 향후 단계에서 PICF 의 BSE-index 참조 해석에 사용할 인덱스 — 1-based.
    /// </summary>
    public IReadOnlyList<(string MediaType, byte[] Data)> BStoreImages { get; private set; }
        = Array.Empty<(string, byte[])>();

    /// <summary>
    /// Phase 3f-4 — PlcSpaMom (Plex of FSPA in Main document) 에서 추출한 floating shape anchor 목록.
    /// 각 entry 는 본문 CP (0x08 drawing char 위치) + shape ID + 앵커 사각형 (twips).
    /// </summary>
    public IReadOnlyList<FspaEntry> FspaEntries { get; private set; } = Array.Empty<FspaEntry>();

    /// <summary>
    /// Phase 3f-5 — DggContainer 의 OfficeArtSpContainer 들을 walk 해서 만든 spid → BStore 1-based index 맵.
    /// FspaEntry.Spid 와 결합하면 floating shape 의 위치 + 이미지를 결정 가능:
    ///   image = BStoreImages[ShapeImageIndex[fspa.Spid] - 1] at (fspa.XaLeft, fspa.YaTop).
    /// </summary>
    public IReadOnlyDictionary<int, int> ShapeImageIndex { get; private set; }
        = new Dictionary<int, int>();

    /// <summary>
    /// Phase 3i — 본문 책갈피 목록. 각 entry 는 (Name, StartCp, EndCp). SttbfBkmk + PlcfBkf + PlcfBkl 결합.
    /// </summary>
    public IReadOnlyList<BookmarkEntry> Bookmarks { get; private set; } = Array.Empty<BookmarkEntry>();

    /// <summary>
    /// Phase 3l — ObjectPool sub-storage 들에서 추출한 임베드 OLE 객체 목록.
    /// 각 entry 는 storage 이름과 안에 든 stream 들의 raw bytes 사전.
    /// EquationNative / Ole10Native / Workbook 등의 stream 이 실제 임베드 객체의 데이터를 담는다.
    /// </summary>
    public IReadOnlyList<OleEmbedEntry> OleEmbeds { get; private set; } = Array.Empty<OleEmbedEntry>();

    /// <summary>
    /// Phase 3n — VBA 매크로 프로젝트 (Macros / _VBA_PROJECT_CUR storage) 의 격리된 raw bytes.
    /// 콘텐츠는 절대 실행하지 않으며 fidelity 보존용으로만 유지. UI 는 <see cref="HasMacros"/> 로 사용자에게 경고.
    /// CLAUDE.md §"활성 콘텐츠는 격리 저장" 원칙.
    /// </summary>
    public MacroProjectInfo? MacroProject { get; private set; }

    /// <summary>Phase 3n — 매크로 프로젝트 존재 여부. UI 에서 "이 문서에는 매크로가 있습니다" 경고에 사용.</summary>
    public bool HasMacros => MacroProject is not null;

    public PolyDonkyument Read(Stream input)
    {
        // OpenMcdf 의 RootStorage 는 파일 경로 또는 Seekable Stream 을 받는데, 안전성을 위해
        // 임시 파일에 복사 후 OpenRead 로 다룬다 (HwpReader 와 같은 패턴).
        var tmpPath = Path.GetTempFileName();
        try
        {
            using (var fs = File.Create(tmpPath))
                input.CopyTo(fs);

            RootStorage root;
            try
            {
                // [MS-CFB] 컨테이너 열기. OpenMcdf 가 sector/FAT 무결성을 검사하므로 비-OLE2 / 손상 파일은 여기서 거부됨.
                root = RootStorage.OpenRead(tmpPath);
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException(
                    "이 파일은 Word 97-2003 (.doc) 형식이 아니거나 OLE2 컨테이너가 손상되었습니다. " +
                    "RTF 는 .rtf 로, DOCX 는 .docx 확장자로 시도해 보세요.",
                    ex);
            }
            using var _root = root;  // dispose 보장

            // WordDocument stream 전체를 메모리에 적재 — Word97-2003 문서는 보통 수 MB 이하.
            byte[] wd = ReadAll(root, "WordDocument")
                ?? throw new InvalidOperationException("이 파일은 Word 97-2003 (.doc) 형식이 아닙니다 — WordDocument stream 이 없습니다.");

            var fib = ParseFib(wd);

            // [MS-OFFCRYPTO] 암호화·obfuscation 사전 거부.
            // fEncrypted 비트가 켜져 있으면 본문 byte 가 RC4/AES 로 변조되어 piece table 도 의미 없음.
            // 일부 변종은 EncryptionInfo / EncryptedSummary stream 으로도 신호.
            if (fib.Encrypted || HasEncryptionStream(root))
            {
                throw new InvalidOperationException(
                    "암호화된 Word 문서는 현재 지원하지 않습니다 (MS-OFFCRYPTO). " +
                    "Word 에서 암호를 풀어 다시 저장한 뒤 시도해 주세요.");
            }

            byte[] table = ReadAll(root, fib.TableStreamName)
                ?? throw new InvalidOperationException($"{fib.TableStreamName} stream 이 없습니다.");

            var (text, fcs) = ExtractTextWithFcs(wd, table, fib);
            // Phase 1b — PAPX/CHPX 바인 테이블을 한 번 로드해서 단락·문자 서식 조회에 재사용.
            var fmt = FormatStyles.Build(wd, table, fib);
            // Phase 3e-2 — Data stream (선택적, 이미지 PICF 가 여기에). 없으면 null.
            fmt.DataStream = ReadAll(root, "Data");
            // Phase 3f / 3f-3 — OfficeArtBStoreContainer 의 FBSE 들에서 공유 BLIP 추출
            // (Data stream 도 함께 넘겨 FBSE.foDelay 가 Data stream 안의 BLIP 을 가리키는 케이스 지원).
            fmt.BStoreImages = ParseBStoreImages(table, fib, fmt.DataStream);
            BStoreImages = fmt.BStoreImages;
            // Phase 3f-4 — PlcSpaMom 에서 floating shape anchor 목록 추출.
            FspaEntries  = ParseFspaEntries(table, fib);
            // Phase 3f-5 — DggContainer 의 SpContainer 들에서 spid → pib (BStore index) 맵 추출.
            ShapeImageIndex = ParseShapeImageIndex(table, fib);
            // Phase 3i — Bookmarks 데이터 추출 (BuildDocument 보다 먼저 — char-walk 이 CP-event 를 본다).
            Bookmarks = ParseBookmarks(table, fib);
            fmt.SetBookmarks(Bookmarks);
            // Phase 3l — ObjectPool sub-storage 들에서 임베드 OLE 객체 추출.
            OleEmbeds = ParseOleEmbeds(root);
            // Phase 3n — VBA 매크로 프로젝트 격리 저장 (절대 실행 X, fidelity 보존만).
            MacroProject = ParseMacroProject(root);
            var doc = BuildDocument(text, fcs, fmt, OleEmbeds);

            // Phase 3f-6 — FspaEntries + ShapeImageIndex + BStoreImages 결합 → floating ImageBlock 생성.
            ApplyFloatingShapeImages(doc);

            // Phase 3d — 헤더/푸터 영역 (subdocument) 텍스트 추출 후 doc.Sections[0] 에 매핑.
            ApplyHeaderFooter(wd, table, fib, doc, fmt);

            // Phase 3g — 각주/미주 sub-document 텍스트 추출 후 doc.Footnotes / doc.Endnotes 에 매핑.
            ApplyFootnotesAndEndnotes(wd, table, fib, doc, fmt);
            // Phase 3j-2 — comment author/date 메타 (ATRDPre10 + SttbfRMark) 를 CommentEntry 에 매핑.
            ApplyCommentsMetadata(table, fib, doc.Comments, fmt.RMarkAuthors);

            // 메타데이터 (best-effort)
            var summary = ReadAll(root, "SummaryInformation");
            if (summary is { Length: > 0 })
                ApplySummaryInformation(summary, doc);

            return doc;
        }
        finally
        {
            try { File.Delete(tmpPath); } catch { }
        }
    }

    // ─────────────────────────────── FIB ────────────────────────────────────────

    private sealed record Fib(
        string TableStreamName,
        uint   FcMin,
        uint   CcpText,
        uint   FcClx,
        uint   LcbClx,
        ushort NFib,
        bool   Encrypted,
        bool   Obfuscated,
        // Phase 1b — PAPX / CHPX bin tables (FIB §2.5.5 FibRgFcLcb97)
        uint   FcPlcfBteChpx,
        uint   LcbPlcfBteChpx,
        uint   FcPlcfBtePapx,
        uint   LcbPlcfBtePapx,
        // Phase 1d — STTB FFN
        uint   FcSttbfFfn,
        uint   LcbSttbfFfn,
        // Phase 1e — Style Sheet (STSH)
        uint   FcStshf,
        uint   LcbStshf,
        // Phase 3c — Section descriptor plex (FibRgFcLcb97 fcPlcfSed @ 0x00CA / lcbPlcfSed @ 0x00CE)
        uint   FcPlcfSed,
        uint   LcbPlcfSed,
        // Phase 3d — Header/footer subdocument: ccpFtn (footnote 길이) + ccpHdd (header/footer 길이)
        //          + PlcfHdd (sub-story 경계 CP 들). 텍스트 영역은 main text 다음에 위치.
        uint   CcpFtn,
        uint   CcpHdd,
        uint   FcPlcfHdd,
        uint   LcbPlcfHdd,
        // Phase 3f — OfficeArtDggContainer (fcDggInfo @ 0x0312, lcbDggInfo @ 0x0316).
        //            Table stream 의 이 영역이 OfficeArtBStoreContainer 를 포함해 문서 전역 BLIP store 를 담는다.
        uint   FcDggInfo,
        uint   LcbDggInfo,
        // Phase 3f-4 — PlcSpaMom (main doc floating shape anchors). FibRgFcLcb97 pair 16:
        //   fcPlcSpaMom  @ 0x011A, lcbPlcSpaMom @ 0x011E.
        uint   FcPlcSpaMom,
        uint   LcbPlcSpaMom,
        // Phase 3g — Footnote / Endnote sub-document.
        //   FibRgLw97: ccpAtn @ 0x005C, ccpEdn @ 0x0060.
        //   FibRgFcLcb97 pair 3: fcPlcffndTxt @ 0x00B2, lcbPlcffndTxt @ 0x00B6.
        //   FibRgFcLcb97 pair 23: fcPlcfendTxt @ 0x0152, lcbPlcfendTxt @ 0x0156.
        uint   CcpAtn,
        uint   CcpEdn,
        uint   FcPlcffndTxt,
        uint   LcbPlcffndTxt,
        uint   FcPlcfendTxt,
        uint   LcbPlcfendTxt,
        // Phase 3g-2 — Footnote / Endnote 본문 참조 plex.
        //   FibRgFcLcb97 pair 2:  fcPlcffndRef @ 0x00AA, lcbPlcffndRef @ 0x00AE.
        //   FibRgFcLcb97 pair 22: fcPlcfendRef @ 0x014A, lcbPlcfendRef @ 0x014E.
        uint   FcPlcffndRef,
        uint   LcbPlcffndRef,
        uint   FcPlcfendRef,
        uint   LcbPlcfendRef,
        // Phase 3h-2 — SttbfRMark (revision mark 작성자 이름 SttbExtend).
        //   FibRgFcLcb97 pair 25: fcSttbfRMark @ 0x0142, lcbSttbfRMark @ 0x0146.
        uint   FcSttbfRMark,
        uint   LcbSttbfRMark,
        // Phase 3i — Bookmarks.
        //   FibRgFcLcb97 pair 29: fcSttbfBkmk @ 0x0182, lcbSttbfBkmk @ 0x0186.
        //   FibRgFcLcb97 pair 30: fcPlcfBkf   @ 0x018A, lcbPlcfBkf   @ 0x018E.
        //   FibRgFcLcb97 pair 31: fcPlcfBkl   @ 0x0192, lcbPlcfBkl   @ 0x0196.
        uint   FcSttbfBkmk,
        uint   LcbSttbfBkmk,
        uint   FcPlcfBkf,
        uint   LcbPlcfBkf,
        uint   FcPlcfBkl,
        uint   LcbPlcfBkl,
        // Phase 3j — Comments (annotations).
        //   FibRgFcLcb97 pair 4: fcPlcfAtnRef @ 0x00BA, lcbPlcfAtnRef @ 0x00BE.
        //   FibRgFcLcb97 pair 5: fcPlcfAtnTxt @ 0x00C2, lcbPlcfAtnTxt @ 0x00C6.
        uint   FcPlcfAtnRef,
        uint   LcbPlcfAtnRef,
        uint   FcPlcfAtnTxt,
        uint   LcbPlcfAtnTxt);

    // FIB (File Information Block) — WordDocument stream 의 첫 부분. 크기는 nFib 에 따라 다르지만
    // 우리가 필요한 모든 필드는 첫 0x200 byte 안에 있다.
    private static Fib ParseFib(byte[] wd)
    {
        if (wd.Length < 0x200)
            throw new InvalidOperationException("WordDocument stream 이 너무 짧음 — FIB 헤더 부족.");

        ushort magic = BitConverter.ToUInt16(wd, 0x0000);
        if (magic != 0xA5EC)
            throw new InvalidOperationException($"Word 97-2003 시그니처 불일치 (0x{magic:X4}). 다른 형식(RTF/DOCX/HWP)이 .doc 로 잘못 명명되었을 수 있습니다.");

        ushort nFib  = BitConverter.ToUInt16(wd, 0x0002);
        ushort flags = BitConverter.ToUInt16(wd, 0x000A);
        // [MS-DOC] §2.5.1 FibBase flag bits @ 0x000A:
        //   bit 8  (0x0100) = fEncrypted        — RC4/AES 암호화
        //   bit 9  (0x0200) = fWhichTblStm      — 1 → "1Table", 0 → "0Table"
        //   bit 15 (0x8000) = fObfuscated       — XOR obfuscation (legacy)
        bool   encrypted   = (flags & 0x0100) != 0;
        bool   obfuscated  = (flags & 0x8000) != 0;
        string tableName   = (flags & 0x0200) != 0 ? "1Table" : "0Table";

        // fcMin (보통 0) — 본문 텍스트 시작 file character position. Complex 문서에선 piece table 가 우선.
        uint fcMin   = BitConverter.ToUInt32(wd, 0x0018);
        // ccpText — main text 길이 (CP 단위)
        uint ccpText = BitConverter.ToUInt32(wd, 0x004C);

        // fcClx / lcbClx — Complex File Information (piece table 포함)
        uint fcClx   = BitConverter.ToUInt32(wd, 0x01A2);
        uint lcbClx  = BitConverter.ToUInt32(wd, 0x01A6);

        // PAPX / CHPX bin tables — Phase 1b
        // [MS-DOC] §2.5.5 FibRgFcLcb97:
        //   fcPlcfBteChpx @ 0x00FA, lcbPlcfBteChpx @ 0x00FE
        //   fcPlcfBtePapx @ 0x0102, lcbPlcfBtePapx @ 0x0106
        uint fcPlcfBteChpx  = BitConverter.ToUInt32(wd, 0x00FA);
        uint lcbPlcfBteChpx = BitConverter.ToUInt32(wd, 0x00FE);
        uint fcPlcfBtePapx  = BitConverter.ToUInt32(wd, 0x0102);
        uint lcbPlcfBtePapx = BitConverter.ToUInt32(wd, 0x0106);

        // Phase 1d — STTB FFN: [MS-DOC] FibRgFcLcb97 fcSttbfFfn @ 0x0112, lcbSttbfFfn @ 0x0116
        uint fcSttbfFfn  = BitConverter.ToUInt32(wd, 0x0112);
        uint lcbSttbfFfn = BitConverter.ToUInt32(wd, 0x0116);

        // Phase 1e — STSH (Style Sheet): [MS-DOC] FibRgFcLcb97 fcStshf @ 0x00A2, lcbStshf @ 0x00A6
        uint fcStshf  = BitConverter.ToUInt32(wd, 0x00A2);
        uint lcbStshf = BitConverter.ToUInt32(wd, 0x00A6);

        // Phase 3c — PlcfSed (Section descriptor plex): fcPlcfSed @ 0x00CA, lcbPlcfSed @ 0x00CE
        uint fcPlcfSed  = BitConverter.ToUInt32(wd, 0x00CA);
        uint lcbPlcfSed = BitConverter.ToUInt32(wd, 0x00CE);

        // Phase 3d — Subdocument lengths + PlcfHdd:
        //   ccpFtn @ 0x0050, ccpHdd @ 0x0054
        //   fcPlcfHdd @ 0x00F2, lcbPlcfHdd @ 0x00F6
        uint ccpFtn     = BitConverter.ToUInt32(wd, 0x0050);
        uint ccpHdd     = BitConverter.ToUInt32(wd, 0x0054);
        uint fcPlcfHdd  = BitConverter.ToUInt32(wd, 0x00F2);
        uint lcbPlcfHdd = BitConverter.ToUInt32(wd, 0x00F6);

        // Phase 3f — OfficeArtDggContainer offsets (FibRgFcLcb97 §2.5.5 pair 79):
        //   fcDggInfo @ 0x0312, lcbDggInfo @ 0x0316
        //   FIB 가 너무 작아 이 offset 까지 없으면 0/0 (BStore 없음) 으로 간주.
        uint fcDggInfo  = wd.Length >= 0x0316 ? BitConverter.ToUInt32(wd, 0x0312) : 0u;
        uint lcbDggInfo = wd.Length >= 0x031A ? BitConverter.ToUInt32(wd, 0x0316) : 0u;

        // Phase 3f-4 — PlcSpaMom (FibRgFcLcb97 pair 16). fcPlcSpaMom @ 0x011A, lcbPlcSpaMom @ 0x011E.
        uint fcPlcSpaMom  = wd.Length >= 0x011E ? BitConverter.ToUInt32(wd, 0x011A) : 0u;
        uint lcbPlcSpaMom = wd.Length >= 0x0122 ? BitConverter.ToUInt32(wd, 0x011E) : 0u;

        // Phase 3g — Footnote / Endnote sub-document fields.
        //   FibRgLw97: ccpAtn @ 0x005C, ccpEdn @ 0x0060.
        //   FibRgFcLcb97 pair 3: fcPlcffndTxt @ 0x00B2, lcbPlcffndTxt @ 0x00B6.
        //   FibRgFcLcb97 pair 23: fcPlcfendTxt @ 0x0152, lcbPlcfendTxt @ 0x0156.
        uint ccpAtn        = wd.Length >= 0x0060 ? BitConverter.ToUInt32(wd, 0x005C) : 0u;
        uint ccpEdn        = wd.Length >= 0x0064 ? BitConverter.ToUInt32(wd, 0x0060) : 0u;
        uint fcPlcffndTxt  = wd.Length >= 0x00B6 ? BitConverter.ToUInt32(wd, 0x00B2) : 0u;
        uint lcbPlcffndTxt = wd.Length >= 0x00BA ? BitConverter.ToUInt32(wd, 0x00B6) : 0u;
        uint fcPlcfendTxt  = wd.Length >= 0x0156 ? BitConverter.ToUInt32(wd, 0x0152) : 0u;
        uint lcbPlcfendTxt = wd.Length >= 0x015A ? BitConverter.ToUInt32(wd, 0x0156) : 0u;

        // Phase 3g-2 — Footnote / Endnote ref PLC offsets.
        //   FibRgFcLcb97 pair 2:  fcPlcffndRef @ 0x00AA, lcbPlcffndRef @ 0x00AE.
        //   FibRgFcLcb97 pair 22: fcPlcfendRef @ 0x014A, lcbPlcfendRef @ 0x014E.
        uint fcPlcffndRef  = wd.Length >= 0x00AE ? BitConverter.ToUInt32(wd, 0x00AA) : 0u;
        uint lcbPlcffndRef = wd.Length >= 0x00B2 ? BitConverter.ToUInt32(wd, 0x00AE) : 0u;
        uint fcPlcfendRef  = wd.Length >= 0x014E ? BitConverter.ToUInt32(wd, 0x014A) : 0u;
        uint lcbPlcfendRef = wd.Length >= 0x0152 ? BitConverter.ToUInt32(wd, 0x014E) : 0u;

        // Phase 3h-2 — SttbfRMark (revision authors).
        //   FibRgFcLcb97 pair 25: fcSttbfRMark @ 0x0142, lcbSttbfRMark @ 0x0146.
        uint fcSttbfRMark  = wd.Length >= 0x0146 ? BitConverter.ToUInt32(wd, 0x0142) : 0u;
        uint lcbSttbfRMark = wd.Length >= 0x014A ? BitConverter.ToUInt32(wd, 0x0146) : 0u;

        // Phase 3i — Bookmarks (FibRgFcLcb97 pair 29 / 30 / 31).
        uint fcSttbfBkmk  = wd.Length >= 0x0186 ? BitConverter.ToUInt32(wd, 0x0182) : 0u;
        uint lcbSttbfBkmk = wd.Length >= 0x018A ? BitConverter.ToUInt32(wd, 0x0186) : 0u;
        uint fcPlcfBkf    = wd.Length >= 0x018E ? BitConverter.ToUInt32(wd, 0x018A) : 0u;
        uint lcbPlcfBkf   = wd.Length >= 0x0192 ? BitConverter.ToUInt32(wd, 0x018E) : 0u;
        uint fcPlcfBkl    = wd.Length >= 0x0196 ? BitConverter.ToUInt32(wd, 0x0192) : 0u;
        uint lcbPlcfBkl   = wd.Length >= 0x019A ? BitConverter.ToUInt32(wd, 0x0196) : 0u;

        // Phase 3j — Comments (FibRgFcLcb97 pair 4 / 5).
        uint fcPlcfAtnRef  = wd.Length >= 0x00BE ? BitConverter.ToUInt32(wd, 0x00BA) : 0u;
        uint lcbPlcfAtnRef = wd.Length >= 0x00C2 ? BitConverter.ToUInt32(wd, 0x00BE) : 0u;
        uint fcPlcfAtnTxt  = wd.Length >= 0x00C6 ? BitConverter.ToUInt32(wd, 0x00C2) : 0u;
        uint lcbPlcfAtnTxt = wd.Length >= 0x00CA ? BitConverter.ToUInt32(wd, 0x00C6) : 0u;

        return new Fib(tableName, fcMin, ccpText, fcClx, lcbClx, nFib, encrypted, obfuscated,
                       fcPlcfBteChpx, lcbPlcfBteChpx, fcPlcfBtePapx, lcbPlcfBtePapx,
                       fcSttbfFfn, lcbSttbfFfn, fcStshf, lcbStshf, fcPlcfSed, lcbPlcfSed,
                       ccpFtn, ccpHdd, fcPlcfHdd, lcbPlcfHdd,
                       fcDggInfo, lcbDggInfo,
                       fcPlcSpaMom, lcbPlcSpaMom,
                       ccpAtn, ccpEdn, fcPlcffndTxt, lcbPlcffndTxt, fcPlcfendTxt, lcbPlcfendTxt,
                       fcPlcffndRef, lcbPlcffndRef, fcPlcfendRef, lcbPlcfendRef,
                       fcSttbfRMark, lcbSttbfRMark,
                       fcSttbfBkmk, lcbSttbfBkmk, fcPlcfBkf, lcbPlcfBkf, fcPlcfBkl, lcbPlcfBkl,
                       fcPlcfAtnRef, lcbPlcfAtnRef, fcPlcfAtnTxt, lcbPlcfAtnTxt);
    }

    // [MS-OFFCRYPTO] EncryptionInfo / EncryptedSummary stream 존재 검사 — fEncrypted 비트가
    // 없더라도 일부 변종이 stream 만으로 암호화를 표시할 수 있어 보조 신호로 본다.
    private static bool HasEncryptionStream(RootStorage root)
    {
        foreach (var entry in root.EnumerateEntries())
        {
            if (entry.Name.Equals("EncryptionInfo",   StringComparison.OrdinalIgnoreCase) ||
                entry.Name.Equals("EncryptedSummary", StringComparison.OrdinalIgnoreCase))
                return true;
        }
        return false;
    }

    // ─────────────────────────────── 텍스트 추출 ────────────────────────────────

    // 본문 main text 영역(CP 0..ccpText)에 대해 piece table 따라가며 raw 텍스트를 구성.
    // Phase 1b 부터 각 char 에 대응하는 file character position(FC byte offset) 도 함께 반환 —
    // PAPX/CHPX 바인 테이블이 FC 기준이므로 단락/문자 서식 조회에 필수.
    private static (string Text, int[] Fcs) ExtractTextWithFcs(byte[] wd, byte[] table, Fib fib)
    {
        if (fib.LcbClx == 0)
        {
            // 비-복합 문서 — 매우 드물지만 fall-back. fcMin..fcMin+ccpText 를 1252 로 본다.
            var s = DecodeAnsi(wd, (int)fib.FcMin, (int)fib.CcpText);
            var fc = new int[s.Length];
            for (int i = 0; i < s.Length; i++) fc[i] = (int)fib.FcMin + i;
            return (s, fc);
        }

        var pcds = ParsePieceTable(table, (int)fib.FcClx, (int)fib.LcbClx, out int[] cps);

        var sb  = new StringBuilder((int)fib.CcpText);
        var fcs = new List<int>((int)fib.CcpText);
        for (int i = 0; i < pcds.Count; i++)
        {
            int cpStart = cps[i];
            int cpEnd   = cps[i + 1];
            // 본문 main text 범위(0..ccpText) 와 교집합만 채택 — 그 너머는 헤더/푸터/각주/주석 등.
            if (cpStart >= fib.CcpText) break;
            int effEnd = Math.Min(cpEnd, (int)fib.CcpText);
            int len    = effEnd - cpStart;
            if (len <= 0) continue;

            uint fcRaw = pcds[i].Fc;
            bool compressed = (fcRaw & 0x40000000u) != 0;
            int  fc   = (int)(fcRaw & 0x3FFFFFFFu);
            if (compressed)
            {
                // ANSI (CP1252) 1-byte chars, fc /= 2.
                fc /= 2;
                if (fc < 0 || fc + len > wd.Length) continue;
                var piece = DecodeAnsi(wd, fc, len);
                for (int j = 0; j < piece.Length; j++)
                {
                    sb.Append(piece[j]);
                    fcs.Add(fc + j);
                }
            }
            else
            {
                // UTF-16LE 2-byte chars.
                int byteLen = len * 2;
                if (fc < 0 || fc + byteLen > wd.Length) continue;
                var piece = Encoding.Unicode.GetString(wd, fc, byteLen);
                for (int j = 0; j < piece.Length; j++)
                {
                    sb.Append(piece[j]);
                    fcs.Add(fc + j * 2);
                }
            }
        }

        return (sb.ToString(), fcs.ToArray());
    }

    private sealed record Pcd(uint Fc);

    // CLX 영역에서 PCDT(0x02) 를 찾아 piece table 을 파싱.
    private static List<Pcd> ParsePieceTable(byte[] table, int fcClx, int lcbClx, out int[] cps)
    {
        if (fcClx < 0 || lcbClx < 5 || fcClx + lcbClx > table.Length)
            throw new InvalidOperationException("CLX 영역이 Table stream 범위를 벗어남.");

        int pos = fcClx;
        int end = fcClx + lcbClx;
        while (pos < end)
        {
            byte clxt = table[pos];
            if (clxt == 0x01)
            {
                // PRC — skip (2-byte length + content)
                if (pos + 3 > end) break;
                int cbGrpprl = BitConverter.ToUInt16(table, pos + 1);
                pos += 3 + cbGrpprl;
                continue;
            }
            if (clxt == 0x02)
            {
                if (pos + 5 > end) break;
                uint lcb = BitConverter.ToUInt32(table, pos + 1);
                int plcStart = pos + 5;
                if (plcStart + lcb > end || lcb < 4) break;

                // PlcPcd: aCP[n+1] + aPcd[n] 으로 구성, lcb = 4(n+1) + 8n = 12n + 4
                int n = (int)((lcb - 4) / 12);
                if (n < 0) break;

                cps = new int[n + 1];
                for (int i = 0; i <= n; i++)
                    cps[i] = BitConverter.ToInt32(table, plcStart + i * 4);

                int pcdBase = plcStart + (n + 1) * 4;
                var list = new List<Pcd>(n);
                for (int i = 0; i < n; i++)
                {
                    // PCD: 2 byte flags, 4 byte fc, 2 byte prm. fc 는 LE u32.
                    uint fc = BitConverter.ToUInt32(table, pcdBase + i * 8 + 2);
                    list.Add(new Pcd(fc));
                }
                return list;
            }
            // 알 수 없는 clxt — 안전하게 중단
            break;
        }
        throw new InvalidOperationException("Piece Table (PCDT 0x02) 를 CLX 영역에서 찾을 수 없음.");
    }

    private static string DecodeAnsi(byte[] data, int offset, int length)
    {
        if (offset < 0 || length <= 0 || offset + length > data.Length) return string.Empty;
        // CP1252 는 .NET 10 CodePagesEncodingProvider 가 BCL 에 포함되어 별도 등록 불필요.
        try
        {
            return Encoding.GetEncoding(1252).GetString(data, offset, length);
        }
        catch
        {
            return Encoding.Latin1.GetString(data, offset, length);
        }
    }

    // ─────────────────────────────── 단락 분리 ──────────────────────────────────

    // Word 의 내부 텍스트 제어 문자 처리:
    //   0x0D '\r'  → 단락 끝
    //   0x0B       → soft line break (단락 내 줄바꿈) → '\n' 으로 보존
    //   0x07       → 표 셀 끝 / 표 행 끝 → 일단 탭으로 평탄화 (표 처리는 Phase 2)
    //   0x0C       → 페이지 break → 단락 분리로 처리
    //   0x13/0x14/0x15 → 필드 시작/구분/끝 → Phase 2 까지는 제거
    //   0x02       → 각주/주석 참조 → 제거
    //   0x05       → 주석 참조 → 제거
    //   0x08       → drawing anchor → 제거
    //   0x01       → picture anchor → 제거 (Phase 2 에서 이미지 처리)
    //   기타 0x00..0x1F 중 \t(0x09) 외 → 제거
    // Phase 1b — 단락 경계(\r 또는 \f) 도달 시 그 FC 로 PAPX 를 조회해 단락 정렬 적용.
    //          단락 내 각 char 의 FC 로 CHPX 를 조회해 같은 RunStyle 끼리 묶어 Run 분할.
    // Phase 2h — 중첩 표 지원을 위해 Stack<TableState> 기반으로 흐름 일반화.
    // 각 단락의 itap level 에 따라 stack push/pop, level=0 면 본문, level>=1 면 해당 깊이의 표.
    private static PolyDonkyument BuildDocument(string raw, int[] fcs, FormatStyles fmt,
        IReadOnlyList<OleEmbedEntry> OleEmbeds)
    {
        var doc     = new PolyDonkyument();
        var section = new Section();
        doc.Sections.Add(section);
        // Phase 3c-2 — 첫 section 에 첫 SED 의 SEPX 적용.
        if (fmt.SectionFcSepx.Count > 0)
            ApplySepx(section, fmt.TableBytes, fmt.SectionFcSepx[0]);

        var paraChars = new List<char>();
        var paraFcs   = new List<int>();
        int lastFc    = 0;
        // Phase 3l-2 — sprmCFOle2 가 set 된 0x01 만날 때마다 OleEmbeds 와 1:1 매칭하기 위한 counter.
        int oleSeenCount = 0;
        // Phase 2h — itap level 별 누적 상태. itap=0 면 stack 비어 있음, itap=1 면 외곽 표 한 개,
        //          itap=2 면 외곽+내부 두 개.
        var tableStack = new Stack<TableState>();
        // Phase 3a — 필드 처리 상태. 0=일반, 1=field code (폐기), 2=field result (포함).
        //   [MS-DOC] §2.8.25 — 0x13 field begin, 0x14 separator, 0x15 end. 중첩 필드는 1-level
        //   단순화 (대부분의 실문서에서 충분).
        int fieldMode = 0;
        // Phase 3a-2 / 3a-5 — 활성 필드 instr / result 범위 추적.
        //   fieldInstr: 0x13 ~ 0x14 사이 누적된 instr 문자 (HYPERLINK "url", PAGE, DATE 등).
        //   resultStartFc: 0x14 직후 fc — result 영역의 시작.
        //   activeUrl / activeFieldType / activeFieldArg: instr 파싱 결과 — result 영역의 각 fc 에 매핑됨.
        StringBuilder? fieldInstr = null;
        int        resultStartFc  = -1;
        string?    activeUrl      = null;
        FieldType? activeFieldType = null;
        string?    activeFieldArg  = null;
        // Phase 3c — 다음 처리할 섹션 boundary 의 인덱스. SectionBoundaryCps[1..] 가 본문 내 break.
        //          [0]=0 은 시작, [last]=ccpText 는 본문 끝. 단락이 boundary 를 넘으면 새 Section.
        int nextSecIdx = 1;

        for (int i = 0; i < raw.Length; i++)
        {
            char c  = raw[i];
            int  fc = i < fcs.Length ? fcs[i] : lastFc;
            lastFc = fc;

            // Phase 3i-2 — bookmark events at this CP. End 먼저(닫고) 그 다음 Start(열기) — 표현상 자연스럽다.
            //   각 marker 는 char-walk 시점에 EnqueueBookmarkEvent 로 등록, BuildParaFromChars 에서 dequeue.
            var bkEnds = fmt.GetBookmarkEndsAtCp(i);
            if (bkEnds is not null) foreach (var name in bkEnds)
            {
                fmt.EnqueueBookmarkEvent(fc, isStart: false, name);
                paraChars.Add('￼'); paraFcs.Add(fc);
            }
            var bkStarts = fmt.GetBookmarkStartsAtCp(i);
            if (bkStarts is not null) foreach (var name in bkStarts)
            {
                fmt.EnqueueBookmarkEvent(fc, isStart: true, name);
                paraChars.Add('￼'); paraFcs.Add(fc);
            }

            switch (c)
            {
                case '\r':
                case '\f':
                    // 단락 경계에서 fieldMode 강제 reset — 정상 Word 는 단락 안에 0x15 가 있어
                    // 단락 경계에서 이미 0 이지만, 0x15 누락된 손상 파일/합성 입력 안전성.
                    fieldMode = 0;
                    FlushParagraph(section, paraChars, paraFcs, fc, fmt, tableStack);
                    // Phase 3c — 이 단락의 CP (i) 가 다음 섹션 boundary 를 넘었으면 새 Section.
                    //   마지막 boundary 는 본문 끝 (cps[^1] = ccpText) 이므로 건드리지 않는다.
                    while (nextSecIdx < fmt.SectionBoundaryCps.Count - 1 &&
                           i + 1 >= fmt.SectionBoundaryCps[nextSecIdx])
                    {
                        // 진행 중인 표 마감 후 새 Section 으로 전환.
                        FinalizeStack(section, tableStack, targetDepth: 0);
                        section = new Section();
                        doc.Sections.Add(section);
                        // Phase 3c-2 — 새 section 의 SEPX 적용. nextSecIdx 가 새 section 의 index.
                        if (nextSecIdx < fmt.SectionFcSepx.Count)
                            ApplySepx(section, fmt.TableBytes, fmt.SectionFcSepx[nextSecIdx]);
                        nextSecIdx++;
                    }
                    break;
                case '\v':
                    paraChars.Add('\n'); paraFcs.Add(fc);
                    break;
                case '\u0007':  // cell mark
                    paraChars.Add('\u0007'); paraFcs.Add(fc);
                    break;
                case '\u0013':  // field begin → field code 모드 (폐기)
                    fieldMode = 1;
                    fieldInstr = new StringBuilder();
                    resultStartFc  = -1;
                    activeUrl      = null;
                    activeFieldType = null;
                    activeFieldArg  = null;
                    break;
                case '\u0014':  // field separator → field result 모드 (포함)
                    fieldMode = 2;
                    if (fieldInstr is not null)
                    {
                        var (t, u, a) = ParseFieldInstr(fieldInstr.ToString());
                        activeFieldType = t;
                        activeUrl       = u;
                        activeFieldArg  = a;
                    }
                    resultStartFc = fc;
                    break;
                case '\u0015':  // field end → 일반 모드 복귀
                    if (resultStartFc >= 0)
                        fmt.AddFieldRange(resultStartFc, fc, activeUrl, activeFieldType, activeFieldArg);
                    fieldMode = 0;
                    fieldInstr = null;
                    resultStartFc  = -1;
                    activeUrl      = null;
                    activeFieldType = null;
                    activeFieldArg  = null;
                    break;
                case '\u0001':  // picture marker (inline image OR OLE embed if sprmCFOle2 set)
                    // Phase 3e — char-walk 중 picture marker 만나면 현재 단락 flush 후 ImageBlock 삽입.
                    // Phase 3l-2 — sprmCFOle2 가 set 이면 임베드 OLE 객체 placeholder 로 변환.
                    //   OleEmbeds 와 char-walk 순서로 1:1 매칭 (counter oleSeenCount).
                    if (fieldMode != 1)
                    {
                        FlushParagraph(section, paraChars, paraFcs, fc, fmt, tableStack);
                        bool isOle = fmt.GetCharIsOle(fc);
                        ImageBlock img;
                        if (isOle)
                        {
                            var match = oleSeenCount < OleEmbeds.Count ? OleEmbeds[oleSeenCount] : null;
                            string cls = match?.ClassName ?? "OLE";
                            img = new ImageBlock
                            {
                                Description = $"[OLE {cls}]",
                                MediaType   = "application/x-ole-embed",
                                Data        = match?.PrimaryContent ?? Array.Empty<byte>(),
                            };
                            oleSeenCount++;
                        }
                        else
                        {
                            img = new ImageBlock
                            {
                                Description = "[image]",
                                MediaType   = "application/octet-stream",
                            };
                            int? picFc = fmt.GetPictureFc(fc);
                            if (picFc.HasValue && fmt.DataStream is not null)
                            {
                                var extracted = TryExtractImage(fmt.DataStream, picFc.Value, fmt.BStoreImages);
                                if (extracted.HasValue)
                                {
                                    img.MediaType = extracted.Value.MediaType;
                                    img.Data      = extracted.Value.Data;
                                }
                            }
                        }
                        if (tableStack.Count > 0)
                            tableStack.Peek().CellBlocks.Add(img);
                        else
                            section.Blocks.Add(img);
                    }
                    break;
                case '\u0002':  // Phase 3g-2 — footnote / endnote 참조 char.
                    if (fieldMode != 1)
                    {
                        string? fnId = null, enId = null;
                        int fnIdx = fmt.FindFootnoteRefIndex(i);
                        if (fnIdx >= 0) fnId = $"fn{fnIdx + 1}";
                        else
                        {
                            int enIdx = fmt.FindEndnoteRefIndex(i);
                            if (enIdx >= 0) enId = $"en{enIdx + 1}";
                        }
                        if (fnId is not null || enId is not null)
                        {
                            fmt.RegisterRefFc(fc, fnId, enId);
                            paraChars.Add('\uFFFC');
                            paraFcs.Add(fc);
                        }
                    }
                    break;
                case '\u0005':  // Phase 3j — comment 참조 char.
                    if (fieldMode != 1)
                    {
                        int cIdx = fmt.FindCommentRefIndex(i);
                        if (cIdx >= 0)
                        {
                            string cmtId = $"cmt{cIdx + 1}";
                            fmt.RegisterRefFc(fc, fnId: null, enId: null, cmtId: cmtId);
                            paraChars.Add('\uFFFC');
                            paraFcs.Add(fc);
                        }
                    }
                    break;
                case '\u0008':  // drawing
                    break;
                case '\t':
                case '\n':
                    if (fieldMode == 1) fieldInstr?.Append(c);
                    else { paraChars.Add(c); paraFcs.Add(fc); }
                    break;
                default:
                    if (c < 0x20) break;
                    if (fieldMode == 1) fieldInstr?.Append(c);
                    else { paraChars.Add(c); paraFcs.Add(fc); }
                    break;
            }
        }
        FlushParagraph(section, paraChars, paraFcs, lastFc, fmt, tableStack);

        // 본문 끝에서 stack 에 남은 표들을 모두 마감.
        FinalizeStack(section, tableStack, targetDepth: 0);

        if (section.Blocks.Count == 0) section.Blocks.Add(new Paragraph());
        return doc;
    }

    // Phase 2h — itap level 별 표 누적 상태.
    private sealed class TableState
    {
        public Table Table { get; } = new();
        public TableRow? Row { get; set; }
        public List<(TableRow Row, TableCellProps[]? Cp)> RowsRaw { get; } = new();
        // 셀의 누적 Block 들. Paragraph + 중첩 Table 둘 다 가능.
        public List<Block> CellBlocks { get; } = new();
    }

    // Phase 2h — stack 을 targetDepth 까지 pop 하면서 각 표를 마감. 부모가 있으면 그 셀의 CellBlocks
    // 에 Table 추가; 없으면 section.Blocks 에 직접 추가.
    private static void FinalizeStack(Section section, Stack<TableState> stack, int targetDepth)
    {
        while (stack.Count > targetDepth)
        {
            var top = stack.Pop();
            if (top.Row is { Cells.Count: > 0 }) top.RowsRaw.Add((top.Row, null));
            FinalizeTable(top.Table, top.RowsRaw);
            if (top.Table.Rows.Count == 0) continue;  // 빈 표는 버림
            if (stack.Count > 0) stack.Peek().CellBlocks.Add(top.Table);
            else                  section.Blocks.Add(top.Table);
        }
    }

    // Phase 1b — 누적된 (char, fc) 쌍을 한 단락으로 묶어 만든다.
    // Phase 2a/2h — itap level 에 따라 stack 조정 후 단락을 본문 또는 표 셀 안에 배치.
    private static void FlushParagraph(
        Section section, List<char> paraChars, List<int> paraFcs, int paraEndFc, FormatStyles fmt,
        Stack<TableState> tableStack)
    {
        var (paraIstd, ps, inTable, isTtp, rgdxa, cellProps, itap) = fmt.GetParagraphInfo(paraEndFc);

        // 1) Stack depth 를 itap 에 맞춤. depth > itap 면 표 마감, depth < itap 면 새 표 push.
        FinalizeStack(section, tableStack, targetDepth: itap);
        while (tableStack.Count < itap)
            tableStack.Push(new TableState());

        if (itap == 0)
        {
            // 본문 단락.
            var para = BuildParaFromChars(paraChars, paraFcs, paraIstd, ps, fmt);
            // Phase 3h-3 — 단락 마크(\r) 의 CHPX 에서 revision flag 추출 → Paragraph 자체에 적용.
            //   sprmCFRMarkIns / Del 가 paragraph mark 에 붙으면 단락 자체가 inserted/deleted.
            var (_, paraRev) = fmt.GetRunStyle(paraEndFc, paraIstd);
            if (paraRev.Inserted) para.IsInsertedRevision = true;
            if (paraRev.Deleted)  para.IsDeletedRevision  = true;
            section.Blocks.Add(para);
            paraChars.Clear();
            paraFcs.Clear();
            return;
        }

        var top = tableStack.Peek();

        // TTP — 행 종료. Phase 2b: rgdxa → 컬럼 너비.
        if (isTtp)
        {
            if (top.Row is { Cells.Count: > 0 }) top.RowsRaw.Add((top.Row, cellProps));
            top.Row = null;
            if (rgdxa is { Length: > 1 } && top.Table.Columns.Count == 0)
            {
                for (int j = 0; j < rgdxa.Length - 1; j++)
                {
                    double widthMm = (rgdxa[j + 1] - rgdxa[j]) / 56.692;
                    if (widthMm < 0) widthMm = 0;
                    top.Table.Columns.Add(new TableColumn { WidthMm = widthMm });
                }
            }
            paraChars.Clear();
            paraFcs.Clear();
            return;
        }

        // 표 안 셀 단락 — 0x07 으로 셀 분리, 없으면 셀 중간 단락으로 누적.
        top.Row ??= new TableRow();
        SplitIntoCells(paraChars, paraFcs, paraIstd, fmt, top.Row, top.CellBlocks);
        paraChars.Clear();
        paraFcs.Clear();
    }

    // Phase 3a-2 / 3a-5 — 필드 instr 파싱.
    //   "HYPERLINK \"https://example.com\""  → (null, Url="https://example.com", Arg=null)
    //   "PAGE \\* MERGEFORMAT"               → (FieldType.Page, null, null)
    //   "SEQ Figure \\* ARABIC"              → (FieldType.Seq, null, Arg="Figure")
    //   "REF MyBookmark"                     → (FieldType.Ref, null, Arg="MyBookmark")
    //   "STYLEREF \"Heading 1\""             → (FieldType.StyleRef, null, Arg="Heading 1")
    //   "INCLUDETEXT \"file.doc\""           → (FieldType.IncludeText, null, Arg="file.doc")
    //   기타 미지원 instr → (null, null, null)
    private static (FieldType? Type, string? Url, string? Arg) ParseFieldInstr(string instr)
    {
        var s = instr.TrimStart();
        if (s.Length == 0) return (null, null, null);

        if (s.StartsWith("HYPERLINK", StringComparison.OrdinalIgnoreCase))
        {
            int q1 = s.IndexOf('"');
            if (q1 < 0) return (null, null, null);
            int q2 = s.IndexOf('"', q1 + 1);
            if (q2 <= q1) return (null, null, null);
            return (null, s.Substring(q1 + 1, q2 - q1 - 1), null);
        }

        // 첫 토큰 = 필드 종류. " " / "\t" / "\\" 로 끝.
        int wordEnd = 0;
        while (wordEnd < s.Length && !char.IsWhiteSpace(s[wordEnd]) && s[wordEnd] != '\\')
            wordEnd++;
        var head = s[..wordEnd].ToUpperInvariant();
        FieldType? type = head switch
        {
            "PAGE"        => FieldType.Page,
            "NUMPAGES"    => FieldType.NumPages,
            "DATE"        => FieldType.Date,
            "TIME"        => FieldType.Time,
            "AUTHOR"      => FieldType.Author,
            "TITLE"       => FieldType.Title,
            // Phase 3a-3
            "NUMCHARS"    => FieldType.NumChars,
            "FILENAME"    => FieldType.FileName,
            "SUBJECT"     => FieldType.Subject,
            "KEYWORDS"    => FieldType.Keywords,
            "COMMENTS"    => FieldType.Comments,
            // Phase 3a-4
            "SEQ"         => FieldType.Seq,
            "REF"         => FieldType.Ref,
            "STYLEREF"    => FieldType.StyleRef,
            "INCLUDETEXT" => FieldType.IncludeText,
            "IF"          => FieldType.If,
            _             => (FieldType?)null,
        };
        if (type is null) return (null, null, null);

        // Phase 3a-5 — head 다음 첫 의미 인자 추출. 스위치 (\) 만나면 인자 없음.
        //   "ABC" 형식이면 따옴표 안쪽, 아니면 다음 공백/스위치 까지의 토큰.
        string? arg = null;
        int p = wordEnd;
        while (p < s.Length && char.IsWhiteSpace(s[p])) p++;
        if (p < s.Length && s[p] != '\\')
        {
            if (s[p] == '"')
            {
                int q = s.IndexOf('"', p + 1);
                if (q > p) arg = s.Substring(p + 1, q - p - 1);
            }
            else
            {
                int end = p;
                while (end < s.Length && !char.IsWhiteSpace(s[end]) && s[end] != '\\')
                    end++;
                if (end > p) arg = s.Substring(p, end - p);
            }
        }
        return (type, null, arg);
    }

    // 단일 셀 단락 만들기 — Phase 1 의 Run 분할 알고리즘 재사용. Phase 3a-2 — URL/FieldType 도
    // run break 기준에 포함시켜 fmt.GetFieldAtFc 의 결과를 Run.Url / Run.Field 에 매핑.
    private static Paragraph BuildParaFromChars(
        List<char> chars, List<int> fcs, int paraIstd, ParagraphStyle? ps, FormatStyles fmt)
    {
        var para = new Paragraph();
        if (ps is not null) para.Style = ps;
        if (chars.Count == 0) return para;

        RunStyle?  curStyle = null;
        string?    curUrl   = null;
        FieldType? curField = null;
        string?    curArg   = null;
        bool       curIns   = false;
        bool       curDel   = false;
        string?    curRevAuthor = null;
        System.DateTimeOffset? curRevDate = null;
        var curText = new StringBuilder();

        void Flush()
        {
            if (curText.Length == 0 || curStyle is null) return;
            var run = new Run { Text = curText.ToString(), Style = curStyle };
            if (curUrl is { Length: > 0 }) run.Url = curUrl;
            if (curField is not null)      run.Field = curField;
            if (curArg is { Length: > 0 }) run.FieldArg = curArg;
            if (curIns) run.IsInsertedRevision = true;
            if (curDel) run.IsDeletedRevision  = true;
            if (curRevAuthor is { Length: > 0 }) run.RevisionAuthor = curRevAuthor;
            if (curRevDate is not null)          run.RevisionDate   = curRevDate;
            para.Runs.Add(run);
            curText.Clear();
        }

        for (int i = 0; i < chars.Count; i++)
        {
            char c = chars[i];
            int fc = fcs[i];

            // Phase 3g-2 / 3i-2 / 3j — OBJECT REPLACEMENT CHARACTER (U+FFFC) marker:
            //   footnote/endnote/comment ref (3g-2 / 3j) → bookmark start/end (3i-2) 순으로 매칭.
            if (c == '￼')
            {
                var (fnId, enId, cmtId) = fmt.GetRefAtFc(fc);
                if (fnId is not null || enId is not null || cmtId is not null)
                {
                    Flush();
                    para.Runs.Add(new Run
                    {
                        Text       = "",
                        Style      = curStyle ?? new RunStyle(),
                        FootnoteId = fnId,
                        EndnoteId  = enId,
                        CommentId  = cmtId,
                    });
                    continue;
                }
                var bk = fmt.DequeueBookmarkEvent(fc);
                if (bk is not null)
                {
                    Flush();
                    var run = new Run { Text = "", Style = curStyle ?? new RunStyle() };
                    if (bk.Value.IsStart) run.BookmarkStart = bk.Value.Name;
                    else                  run.BookmarkEnd   = bk.Value.Name;
                    para.Runs.Add(run);
                    continue;
                }
            }

            var (rsOrNull, rev) = fmt.GetRunStyle(fc, paraIstd);
            var rs = rsOrNull ?? new RunStyle();
            var (url, field, arg) = fmt.GetFieldAtFc(fc);
            var revDate = rev.HasDttm ? FormatStyles.UnpackDttm(rev.RMarkDttm) : null;
            bool styleBreak = curStyle is null || !RunStyleEquals(curStyle, rs);
            bool fieldBreak = !string.Equals(curUrl, url, StringComparison.Ordinal)
                            || curField != field
                            || !string.Equals(curArg, arg, StringComparison.Ordinal);
            bool revBreak   = curIns != rev.Inserted || curDel != rev.Deleted
                            || !string.Equals(curRevAuthor, rev.Author, StringComparison.Ordinal)
                            || !Nullable.Equals(curRevDate, revDate);
            if (styleBreak || fieldBreak || revBreak)
            {
                Flush();
                curStyle     = rs;
                curUrl       = url;
                curField     = field;
                curArg       = arg;
                curIns       = rev.Inserted;
                curDel       = rev.Deleted;
                curRevAuthor = rev.Author;
                curRevDate   = revDate;
            }
            curText.Append(c);
        }
        Flush();
        return para;
    }

    // Phase 2f — 표 마감 시 raw rows + cellProps 를 walk 해 (세로 → 가로 병합) 순으로 적용.
    private static void FinalizeTable(
        Table table, List<(TableRow Row, TableCellProps[]? Cp)> rows)
    {
        if (rows.Count == 0) return;

        // 1. 세로 병합 — column 별 chain.
        int maxCols = rows.Max(r => r.Cp?.Length ?? 0);
        var toRemove = new bool[rows.Count][];
        for (int r = 0; r < rows.Count; r++)
            toRemove[r] = new bool[rows[r].Cp?.Length ?? 0];

        for (int col = 0; col < maxCols; col++)
        {
            int startRow = -1;
            for (int r = 0; r < rows.Count; r++)
            {
                var cp = rows[r].Cp;
                if (cp is null || col >= cp.Length) { startRow = -1; continue; }
                var p = cp[col];
                if (p.IsVertMerge && p.IsVertRestart) startRow = r;
                else if (p.IsVertMerge && !p.IsVertRestart && startRow >= 0)
                {
                    if (col < rows[startRow].Row.Cells.Count)
                        rows[startRow].Row.Cells[col].RowSpan++;
                    toRemove[r][col] = true;
                }
                else startRow = -1;
            }
        }

        // 2. 각 row 에 (세로 흡수 제거 + 가로 병합) 적용.
        for (int r = 0; r < rows.Count; r++)
        {
            var (row, cp) = rows[r];
            row.Cells = MergeRowCells(row.Cells, cp, toRemove[r]);
            if (row.Cells.Count > 0) table.Rows.Add(row);
        }
    }

    // 한 행의 row.Cells 에 (테두리·배경 적용 → 세로 흡수 제거 + 가로 병합) 처리.
    private static IList<TableCell> MergeRowCells(
        IList<TableCell> cells, TableCellProps[]? cp, bool[]? toRemove)
    {
        if (cp is null) return cells;

        int n = Math.Min(cells.Count, cp.Length);
        for (int i = 0; i < n; i++)
        {
            var cell = cells[i];
            var p    = cp[i];
            if (p.Top    is not null) cell.BorderTop    = p.Top;
            if (p.Left   is not null) cell.BorderLeft   = p.Left;
            if (p.Bottom is not null) cell.BorderBottom = p.Bottom;
            if (p.Right  is not null) cell.BorderRight  = p.Right;
            if (p.BackgroundHex is not null) cell.BackgroundColor = p.BackgroundHex;
        }

        var merged = new List<TableCell>();
        for (int col = 0; col < n; col++)
        {
            if (toRemove is not null && col < toRemove.Length && toRemove[col]) continue;
            bool absorb = cp[col].IsMerged && !cp[col].IsFirstMerged && merged.Count > 0;
            if (absorb) merged[^1].ColumnSpan++;
            else        merged.Add(cells[col]);
        }
        for (int k = n; k < cells.Count; k++) merged.Add(cells[k]);
        return merged;
    }

    // Phase 2a/2g/2h — paraChars 를 0x07 (cell mark) 기준으로 셀 종료 시점에 분리.
    //   0x07 만남: curChars 의 텍스트 + 누적 cellBlocks 를 cell.Blocks 로 묶어 pendingRow 에 추가.
    //   0x07 없이 단락 끝남: cellBlocks 에 단락 누적만, 셀 종료 X. 다음 단락에서 이어진다.
    // cellBlocks 는 List<Block> 이라 중첩 표(Phase 2h)도 셀 안에 자연스럽게 들어간다 — 중첩 표는
    // 상위 FlushParagraph 가 stack pop 시 부모의 CellBlocks 에 Table 을 추가한 결과.
    private static void SplitIntoCells(
        List<char> paraChars, List<int> paraFcs, int paraIstd, FormatStyles fmt,
        TableRow pendingRow, List<Block> cellBlocks)
    {
        var curChars = new List<char>();
        var curFcs   = new List<int>();
        for (int i = 0; i < paraChars.Count; i++)
        {
            if (paraChars[i] == '\u0007')
            {
                // 셀 종료. curChars 가 비어 있고 cellBlocks 가 채워져 있는 케이스 (예: 외곽 셀의
                // 마지막 0x07 이 단락 안에 단독 — 그 직전에 inner 표가 누적됨) 면 빈 단락 추가 X.
                if (curChars.Count > 0)
                    cellBlocks.Add(BuildParaFromChars(curChars, curFcs, paraIstd, null, fmt));
                var cell = new TableCell();
                if (cellBlocks.Count == 0) cell.Blocks.Add(new Paragraph());
                else foreach (var b in cellBlocks) cell.Blocks.Add(b);
                pendingRow.Cells.Add(cell);
                cellBlocks.Clear();
                curChars.Clear();
                curFcs.Clear();
            }
            else
            {
                curChars.Add(paraChars[i]);
                curFcs.Add(paraFcs[i]);
            }
        }
        if (curChars.Count > 0)
            cellBlocks.Add(BuildParaFromChars(curChars, curFcs, paraIstd, null, fmt));
    }
    private static bool RunStyleEquals(RunStyle a, RunStyle b)
        => a.Bold == b.Bold
        && a.Italic == b.Italic
        && a.Underline == b.Underline
        && a.Strikethrough == b.Strikethrough
        && a.FontSizePt == b.FontSizePt
        && string.Equals(a.FontFamily, b.FontFamily, StringComparison.Ordinal)
        && Nullable.Equals(a.Foreground, b.Foreground)
        && Nullable.Equals(a.Background, b.Background);

    // ─────────────────────────────── 메타데이터 ─────────────────────────────────

    // [MS-OLEPS] §2.18 SummaryInformation PropertySet → DocumentMetadata 매핑.
    // PIDSI 코드: 0x02 Title · 0x03 Subject · 0x04 Author · 0x05 Keywords · 0x06 Comments ·
    //             0x08 LastSavedBy · 0x09 RevisionNumber · 0x0C CreateTime · 0x0D LastSavedTime ·
    //             0x12 AppName.
    // Phase 3e-2 — PICF 영역 안에서 PNG/JPEG/GIF signature 검색 후 raw byte 추출.
    // Phase 3e-3 — WMF placeable header + DIB BITMAPINFOHEADER 추가.
    // Phase 3e-5 — EMF EMR_HEADER (iType=1 + " EMF" dSignature@40) 추가.
    // Phase 3e-6 — OfficeArt 컨테이너 정식 파싱. recVer=0xF 컨테이너 재귀, BLIP_PNG(0xF01E)/
    //              BLIP_JPEG(0xF01D, 0xF02A) atom 검출 시 UID + tag 건너뛰고 정확한 image data 추출.
    //              올바른 OfficeArt 구조면 signature heuristic 보다 정확하며 image 경계도 정확.
    //              실패 시 (raw inline embed) 기존 signature 스캔으로 fallback.
    // [MS-DOC] §2.9.197 PICF 의 lcb 가 전체 영역 크기. PICF 내부에 OfficeArt blob 가 들어가는데
    // 가장 흔한 modern Word 이미지는 그 blob 끝부분에 raw byte 가 인라인. signature 위치부터
    // PICF 끝까지를 image data 로 본다.
    private static (string MediaType, byte[] Data)? TryExtractImage(
        byte[] data, int fcPic,
        IReadOnlyList<(string MediaType, byte[] Data)>? bstore)
    {
        if (fcPic < 0 || fcPic + 4 > data.Length) return null;
        int lcb = BitConverter.ToInt32(data, fcPic);
        if (lcb <= 8 || fcPic + lcb > data.Length) return null;

        int start = fcPic + 4;                 // lcb 이후
        int picEnd = Math.Min(fcPic + lcb, data.Length);

        // Phase 3e-6 — OfficeArt 컨테이너 우선 시도. Phase 3f-2 — bstore 도 함께 넘겨 pib 참조 해석.
        var fromOa = TryExtractFromOfficeArt(data, start, picEnd, bstore, depth: 0);
        if (fromOa.HasValue) return fromOa;

        int end = picEnd - 8;  // signature 최소 8 byte 여유

        for (int i = start; i < end; i++)
        {
            // PNG: 89 50 4E 47 0D 0A 1A 0A
            if (data[i] == 0x89 && data[i + 1] == 0x50 && data[i + 2] == 0x4E && data[i + 3] == 0x47 &&
                data[i + 4] == 0x0D && data[i + 5] == 0x0A && data[i + 6] == 0x1A && data[i + 7] == 0x0A)
                return ("image/png", SliceTo(data, i, fcPic + lcb));
            // JPEG: FF D8 FF
            if (data[i] == 0xFF && data[i + 1] == 0xD8 && data[i + 2] == 0xFF)
                return ("image/jpeg", SliceTo(data, i, fcPic + lcb));
            // GIF: 47 49 46 38 (GIF87a/89a)
            if (data[i] == 0x47 && data[i + 1] == 0x49 && data[i + 2] == 0x46 && data[i + 3] == 0x38)
                return ("image/gif", SliceTo(data, i, fcPic + lcb));
            // Phase 3e-3 — Placeable WMF header: D7 CD C6 9A 00 00 (§Placeable Metafile)
            if (data[i] == 0xD7 && data[i + 1] == 0xCD && data[i + 2] == 0xC6 && data[i + 3] == 0x9A)
                return ("image/wmf", SliceTo(data, i, fcPic + lcb));
            // Phase 3e-5 — EMF EMR_HEADER record:
            //   bytes 0..3  : iType (= 1, EMR_HEADER)         = 01 00 00 00
            //   bytes 40..43: dSignature " EMF" (0x464D4520)  = 20 45 4D 46
            if (i + 44 <= data.Length &&
                data[i] == 0x01 && data[i + 1] == 0 && data[i + 2] == 0 && data[i + 3] == 0 &&
                data[i + 40] == 0x20 && data[i + 41] == 0x45 && data[i + 42] == 0x4D && data[i + 43] == 0x46)
                return ("image/emf", SliceTo(data, i, fcPic + lcb));
            // Phase 3e-3 — DIB BITMAPINFOHEADER heuristic:
            //   byte 0..3 = 0x28 0x00 0x00 0x00 (biSize=40)
            //   byte 12..13 = biPlanes (LE) = 1
            //   byte 14..15 = biBitCount in {1,4,8,16,24,32}
            //   biSize=40 만으로는 false-positive 가 흔해 추가 검사가 필요.
            if (i + 16 < data.Length &&
                data[i] == 0x28 && data[i + 1] == 0 && data[i + 2] == 0 && data[i + 3] == 0 &&
                data[i + 12] == 0x01 && data[i + 13] == 0x00)
            {
                ushort bitCount = BitConverter.ToUInt16(data, i + 14);
                if (bitCount is 1 or 4 or 8 or 16 or 24 or 32)
                    return ("image/x-dib", SliceTo(data, i, fcPic + lcb));
            }
        }
        return null;

        static byte[] SliceTo(byte[] src, int from, int toExclusive)
        {
            int len = toExclusive - from;
            var dst = new byte[len];
            Buffer.BlockCopy(src, from, dst, 0, len);
            return dst;
        }
    }

    // Phase 3e-6 — OfficeArt 레코드 walker. PICF 안의 OfficeArt blob 을 진짜 record tree 로 해석.
    //   각 레코드 헤더 8 byte:
    //     recVerAndInstance (ushort LE): low 4 bits = recVer, high 12 bits = recInstance.
    //     recType           (ushort LE): 0xF000 대역.
    //     recLen            (uint   LE): 헤더 다음 데이터 길이.
    //   recVer == 0xF → container, 자식 레코드 재귀.
    //   BLIP atom:
    //     - Bitmap BLIP (0xF01D BLIP_JPEG, 0xF01E BLIP_PNG, 0xF01F BLIP_DIB, 0xF02A BLIP_JPEGCMYK):
    //         body = UID(16) [+ UID2(16) if (recInstance & 1) == 1] + tag(1) + image bytes.
    //     - Metafile BLIP (0xF01A BLIP_EMF, 0xF01B BLIP_WMF) — Phase 3e-7:
    //         body = UID(16) [+ UID2(16)] + OfficeArtMetafileHeader(34) + image bytes (raw 또는 zlib).
    //         MetafileHeader: cbSize(4) + rcBounds(16) + ptSize(8) + cbSave(4) + compression(1) + filter(1).
    //         compression == 0xFE: raw. compression == 0x00: zlib 압축 — ZLibStream 으로 풀어 원본 복원.
    private static (string MediaType, byte[] Data)? TryExtractFromOfficeArt(
        byte[] data, int start, int end,
        IReadOnlyList<(string MediaType, byte[] Data)>? bstore,
        int depth)
    {
        if (depth > 8) return null;  // 손상된 입력에서 무한 재귀 방지
        int pos = start;
        while (pos + 8 <= end)
        {
            ushort verInst = BitConverter.ToUInt16(data, pos);
            ushort recType = BitConverter.ToUInt16(data, pos + 2);
            uint   recLen  = BitConverter.ToUInt32(data, pos + 4);
            int dataStart  = pos + 8;
            long dataEnd64 = (long)dataStart + recLen;
            if (dataEnd64 > end || dataEnd64 < dataStart) return null;
            int dataEnd = (int)dataEnd64;

            int recVer  = verInst & 0x000F;
            int recInst = (verInst >> 4) & 0x0FFF;

            if (recVer == 0xF)
            {
                var child = TryExtractFromOfficeArt(data, dataStart, dataEnd, bstore, depth + 1);
                if (child.HasValue) return child;
            }
            else
            {
                // Bitmap BLIP — UID(16) [+ UID2(16)] + tag(1) + image data.
                string? bitmapMime = recType switch
                {
                    0xF01D => "image/jpeg",   // BLIP_JPEG (RGB)
                    0xF01E => "image/png",    // BLIP_PNG
                    0xF01F => "image/x-dib",  // BLIP_DIB
                    0xF02A => "image/jpeg",   // BLIP_JPEGCMYK
                    _      => null,
                };
                if (bitmapMime is not null)
                {
                    int hdrExtra = 16 + 1;
                    if ((recInst & 1) == 1) hdrExtra += 16;
                    int imgStart = dataStart + hdrExtra;
                    if (imgStart >= dataEnd) return null;
                    int imgLen = dataEnd - imgStart;
                    var img = new byte[imgLen];
                    Buffer.BlockCopy(data, imgStart, img, 0, imgLen);
                    return (bitmapMime, img);
                }

                // Metafile BLIP — UID(16) [+ UID2(16)] + OfficeArtMetafileHeader(34) + image bytes.
                string? mfMime = recType switch
                {
                    0xF01A => "image/emf",   // BLIP_EMF
                    0xF01B => "image/wmf",   // BLIP_WMF
                    _      => null,
                };
                if (mfMime is not null)
                {
                    int hdrUids = 16;
                    if ((recInst & 1) == 1) hdrUids += 16;
                    int mfHdrStart = dataStart + hdrUids;
                    if (mfHdrStart + 34 > dataEnd) return null;
                    uint cbSize = BitConverter.ToUInt32(data, mfHdrStart + 0);
                    byte compression = data[mfHdrStart + 32];
                    int imgStart = mfHdrStart + 34;
                    if (imgStart > dataEnd) return null;
                    int storedLen = dataEnd - imgStart;

                    if (compression == 0xFE)
                    {
                        // 비압축 — cbSize 만큼 또는 남은 만큼.
                        int len = Math.Min(storedLen, (int)Math.Min(cbSize, int.MaxValue));
                        var raw = new byte[len];
                        Buffer.BlockCopy(data, imgStart, raw, 0, len);
                        return (mfMime, raw);
                    }
                    if (compression == 0x00)
                    {
                        try
                        {
                            using var src = new MemoryStream(data, imgStart, storedLen, writable: false);
                            using var zs  = new System.IO.Compression.ZLibStream(
                                src, System.IO.Compression.CompressionMode.Decompress);
                            using var dst = new MemoryStream();
                            zs.CopyTo(dst);
                            return (mfMime, dst.ToArray());
                        }
                        catch
                        {
                            return null;  // 손상된 zlib — fallback 양보.
                        }
                    }
                }

                // Phase 3f-2 — OfficeArtFOPT (0xF00B) property block 안의 pib (property ID 260)
                //   를 1-based BLIP index 로 해석해 bstore 에서 lookup.
                //   각 property = opid(2) + op(4):
                //     opid bits 0..13 = propId, bit 14 = fBid, bit 15 = fComplex
                if (recType == 0xF00B && bstore is { Count: > 0 })
                {
                    int propCount = recInst;
                    int avail = (dataEnd - dataStart) / 6;
                    if (propCount > avail) propCount = avail;
                    for (int i = 0; i < propCount; i++)
                    {
                        int off = dataStart + i * 6;
                        ushort opid = BitConverter.ToUInt16(data, off);
                        uint   op   = BitConverter.ToUInt32(data, off + 2);
                        int  propId = opid & 0x3FFF;
                        bool fBid   = (opid & 0x4000) != 0;
                        if (propId == 260 && fBid)
                        {
                            int idx = (int)op;
                            if (idx >= 1 && idx <= bstore.Count) return bstore[idx - 1];
                        }
                    }
                }
            }
            pos = dataEnd;
        }
        return null;
    }

    // Phase 3f — OfficeArtDggContainer (Table stream) → OfficeArtBStoreContainer (0xF001) →
    //            OfficeArtFBSE atoms (0xF007) → 임베드된 BLIP 추출 후 인덱스 작성.
    //   FBSE body (36 byte fixed): btWin32 / btMacOS / rgbUid(16) / tag(2) / size(4) / cRef(4) /
    //                              foDelay(4) / unused1(1) / cbName(1) / unused2(1) / unused3(1)
    //   이후 nameData(cbName byte) 다음에 임베드 BLIP 가 옵션으로 위치.
    // Phase 3f-3 — 임베드 BLIP 가 없으면 foDelay 가 가리키는 Data stream 의 BLIP 을 탐색.
    private static IReadOnlyList<(string MediaType, byte[] Data)> ParseBStoreImages(
        byte[] table, Fib fib, byte[]? dataStream)
    {
        if (fib.LcbDggInfo == 0) return Array.Empty<(string, byte[])>();
        int start = (int)fib.FcDggInfo;
        long endL  = (long)start + fib.LcbDggInfo;
        if (start < 0 || endL > table.Length) return Array.Empty<(string, byte[])>();
        int end = (int)endL;

        var list = new List<(string, byte[])>();
        WalkForBStore(table, start, end, dataStream, list, depth: 0);
        return list;
    }

    // DggContainer (0xF000) 안에서 BStoreContainer (0xF001) 를 찾아 FBSE 들을 처리.
    private static void WalkForBStore(byte[] data, int start, int end,
        byte[]? dataStream, List<(string, byte[])> sink, int depth)
    {
        if (depth > 8) return;
        int pos = start;
        while (pos + 8 <= end)
        {
            ushort verInst = BitConverter.ToUInt16(data, pos);
            ushort recType = BitConverter.ToUInt16(data, pos + 2);
            uint   recLen  = BitConverter.ToUInt32(data, pos + 4);
            int dataStart  = pos + 8;
            long dataEnd64 = (long)dataStart + recLen;
            if (dataEnd64 > end || dataEnd64 < dataStart) return;
            int dataEnd = (int)dataEnd64;
            int recVer = verInst & 0x000F;

            if (recType == 0xF001)
            {
                // BStoreContainer (container 0xF001). 자식 FBSE 처리.
                ExtractFbses(data, dataStart, dataEnd, dataStream, sink);
            }
            else if (recVer == 0xF)
            {
                // 다른 컨테이너 — DggContainer / OptContainer 등 — 재귀로 BStore 탐색.
                WalkForBStore(data, dataStart, dataEnd, dataStream, sink, depth + 1);
            }
            pos = dataEnd;
        }
    }

    // BStoreContainer 안의 FBSE atom 들을 walk 하며 임베드 BLIP 또는 foDelay 가 가리키는 Data stream BLIP 추출.
    private static void ExtractFbses(byte[] data, int start, int end,
        byte[]? dataStream, List<(string, byte[])> sink)
    {
        int pos = start;
        while (pos + 8 <= end)
        {
            ushort recType = BitConverter.ToUInt16(data, pos + 2);
            uint   recLen  = BitConverter.ToUInt32(data, pos + 4);
            int dataStart  = pos + 8;
            long dataEnd64 = (long)dataStart + recLen;
            if (dataEnd64 > end || dataEnd64 < dataStart) return;
            int dataEnd = (int)dataEnd64;

            if (recType == 0xF007 && dataStart + 36 <= dataEnd)
            {
                uint foDelay = BitConverter.ToUInt32(data, dataStart + 28);
                byte cbName  = data[dataStart + 33];
                int blipStart = dataStart + 36 + cbName;

                // (a) 임베드 BLIP — FBSE 본체 안에 BLIP 레코드가 함께 있는 경우 (Phase 3f).
                (string MediaType, byte[] Data)? blip = null;
                if (blipStart < dataEnd)
                    blip = TryExtractFromOfficeArt(data, blipStart, dataEnd, bstore: null, depth: 0);

                // (b) Phase 3f-3 — 임베드가 없으면 foDelay 의 Data stream 안에서 BLIP 찾기.
                //     foDelay == 0xFFFFFFFF 는 "BLIP 없음" sentinel.
                if (!blip.HasValue && foDelay != 0xFFFFFFFF && dataStream is not null)
                {
                    int doff = (int)foDelay;
                    if (doff >= 0 && doff < dataStream.Length)
                        blip = TryExtractFromOfficeArt(
                            dataStream, doff, dataStream.Length, bstore: null, depth: 0);
                }

                if (blip.HasValue) sink.Add(blip.Value);
            }
            pos = dataEnd;
        }
    }

    // Phase 3f-4 — PlcSpaMom (Plex of FSPA) 파싱. [MS-DOC] §2.8.32.
    //   aCP[N+1] (각 4 byte) + aSpa[N] (각 26 byte). lcb = 4*(N+1) + 26*N = 4 + 30N.
    //   각 FSPA: spid(4) + xaLeft(4) + yaTop(4) + xaRight(4) + yaBottom(4) + flags(6 byte).
    //   aCP[i] = 본문 안 0x08 (drawing char) 의 CP 위치 — 도형이 anchor 되는 지점.
    private static IReadOnlyList<FspaEntry> ParseFspaEntries(byte[] table, Fib fib)
    {
        if (fib.LcbPlcSpaMom < 4 + 30) return Array.Empty<FspaEntry>();
        int start = (int)fib.FcPlcSpaMom;
        long endL  = (long)start + fib.LcbPlcSpaMom;
        if (start < 0 || endL > table.Length) return Array.Empty<FspaEntry>();

        int lcb = (int)fib.LcbPlcSpaMom;
        int n = (lcb - 4) / 30;
        if (n <= 0) return Array.Empty<FspaEntry>();

        var list = new List<FspaEntry>(n);
        int cpsBase = start;
        int spaBase = start + 4 * (n + 1);
        for (int i = 0; i < n; i++)
        {
            int cp        = BitConverter.ToInt32(table, cpsBase + i * 4);
            int spaOff    = spaBase + i * 26;
            int spid      = BitConverter.ToInt32(table, spaOff + 0);
            int xaLeft    = BitConverter.ToInt32(table, spaOff + 4);
            int yaTop     = BitConverter.ToInt32(table, spaOff + 8);
            int xaRight   = BitConverter.ToInt32(table, spaOff + 12);
            int yaBottom  = BitConverter.ToInt32(table, spaOff + 16);
            list.Add(new FspaEntry(cp, spid, xaLeft, yaTop, xaRight, yaBottom));
        }
        return list;
    }

    // Phase 3f-6 — FspaEntries 의 각 항목을 ShapeImageIndex → BStoreImages 로 resolve 해
    //   floating ImageBlock 을 doc.Sections[0].Blocks 에 추가.
    //   좌표는 twips → mm 변환 (1 inch = 1440 twips = 25.4 mm → / 56.692).
    //   AnchorPageIndex 는 0 (페이지 단위 분배는 메인 앱의 페이지네이션 단계에서 처리).
    private void ApplyFloatingShapeImages(PolyDonkyument doc)
    {
        if (FspaEntries.Count == 0 || ShapeImageIndex.Count == 0
            || BStoreImages.Count == 0 || doc.Sections.Count == 0)
            return;

        var section = doc.Sections[0];
        foreach (var fspa in FspaEntries)
        {
            if (!ShapeImageIndex.TryGetValue(fspa.Spid, out int pib)) continue;
            if (pib < 1 || pib > BStoreImages.Count) continue;
            var (mime, data) = BStoreImages[pib - 1];

            double widthMm  = (fspa.XaRightTwips  - fspa.XaLeftTwips) / 56.692;
            double heightMm = (fspa.YaBottomTwips - fspa.YaTopTwips)  / 56.692;
            if (widthMm  < 0) widthMm  = 0;
            if (heightMm < 0) heightMm = 0;

            section.Blocks.Add(new ImageBlock
            {
                MediaType       = mime,
                Data            = data,
                WrapMode        = ImageWrapMode.InFrontOfText,
                AnchorPageIndex = 0,
                OverlayXMm      = fspa.XaLeftTwips / 56.692,
                OverlayYMm      = fspa.YaTopTwips  / 56.692,
                WidthMm         = widthMm,
                HeightMm        = heightMm,
                Description     = $"[floating shape spid={fspa.Spid}]",
            });
        }
    }

    // Phase 3i — 본문 책갈피 추출. SttbfBkmk (이름 배열) + PlcfBkf (시작 CP) + PlcfBkl (끝 CP).
    //   PlcfBkf 형식: aCP[N+1] + aBKF[N], BKF = 4 byte (ibkl 2 + flags 2). lcb = 4*(N+1) + 4*N = 4 + 8*N.
    //   PlcfBkl 형식: aCP[M+1] only. lcb = 4*(M+1).
    //   각 BKF.ibkl 가 PlcfBkl 의 인덱스 (해당 책갈피의 끝 CP 위치).
    //   책갈피 이름은 SttbfBkmk[i] — 인덱스 i 는 PlcfBkf 의 순서와 일치.
    private static IReadOnlyList<BookmarkEntry> ParseBookmarks(byte[] table, Fib fib)
    {
        // 이름.
        var names = FormatStyles.ReadSttbExtend(table, (int)fib.FcSttbfBkmk, (int)fib.LcbSttbfBkmk);

        // PlcfBkf — 시작 CP + ibkl.
        var bkfPlc = ReadPlcCpsAndU16Pair(table, (int)fib.FcPlcfBkf, (int)fib.LcbPlcfBkf, dataSize: 4);
        if (bkfPlc.Cps.Length == 0) return Array.Empty<BookmarkEntry>();

        // PlcfBkl — 끝 CP. data 없음 (element = 4 byte CP).
        var bklCps = FormatStyles.ReadPlcCps(table, (int)fib.FcPlcfBkl, (int)fib.LcbPlcfBkl, frdSize: 0);

        int n = bkfPlc.Cps.Length;
        var list = new List<BookmarkEntry>(n);
        for (int i = 0; i < n; i++)
        {
            int startCp = bkfPlc.Cps[i];
            int ibkl    = bkfPlc.DataU16s[i];  // BKF body[0..1] = ibkl
            int endCp   = (ibkl >= 0 && ibkl < bklCps.Length) ? bklCps[ibkl] : startCp;
            string name = i < names.Length ? names[i] : $"_Bookmark{i}";
            list.Add(new BookmarkEntry(name, startCp, endCp));
        }
        return list;
    }

    // PlcfBkf 같이 element 가 CP + 가변 크기 data 인 plex 에서 CP 와 data 의 첫 2 byte (U16) 를 함께 추출.
    //   lcb = 4*(N+1) + dataSize*N.
    private static (int[] Cps, int[] DataU16s) ReadPlcCpsAndU16Pair(byte[] table, int fc, int lcb, int dataSize)
    {
        if (lcb < 4 + dataSize || fc < 0 || fc + lcb > table.Length)
            return (Array.Empty<int>(), Array.Empty<int>());
        int element = 4 + dataSize;
        int n = (lcb - 4) / element;
        if (n <= 0) return (Array.Empty<int>(), Array.Empty<int>());
        var cps = new int[n];
        var u16 = new int[n];
        int dataBase = fc + 4 * (n + 1);
        for (int i = 0; i < n; i++)
        {
            cps[i] = BitConverter.ToInt32(table, fc + i * 4);
            u16[i] = dataSize >= 2 ? BitConverter.ToUInt16(table, dataBase + i * dataSize) : 0;
        }
        return (cps, u16);
    }

    // Phase 3f-5 — DggContainer 안의 OfficeArtSpContainer (0xF004) 들을 walk 해 spid → pib 맵 작성.
    //   각 SpContainer:
    //     - Sp atom (0xF009, body 8 byte: spid(4) + flags(4)) — shape ID
    //     - OPT atom (0xF00B, body = N 개 property × 6 byte) — pib (id=260, fBid=1) 가 BLIP 인덱스
    //   spid != 0 이고 pib > 0 일 때만 맵에 등록.
    private static IReadOnlyDictionary<int, int> ParseShapeImageIndex(byte[] table, Fib fib)
    {
        var map = new Dictionary<int, int>();
        if (fib.LcbDggInfo == 0) return map;
        int start = (int)fib.FcDggInfo;
        long endL  = (long)start + fib.LcbDggInfo;
        if (start < 0 || endL > table.Length) return map;
        WalkSpContainers(table, start, (int)endL, map, depth: 0);
        return map;
    }

    private static void WalkSpContainers(byte[] data, int start, int end, Dictionary<int, int> map, int depth)
    {
        if (depth > 12) return;  // 손상된 입력 안전
        int pos = start;
        while (pos + 8 <= end)
        {
            ushort verInst = BitConverter.ToUInt16(data, pos);
            ushort recType = BitConverter.ToUInt16(data, pos + 2);
            uint   recLen  = BitConverter.ToUInt32(data, pos + 4);
            int dataStart  = pos + 8;
            long dataEnd64 = (long)dataStart + recLen;
            if (dataEnd64 > end || dataEnd64 < dataStart) return;
            int dataEnd = (int)dataEnd64;
            int recVer = verInst & 0x000F;

            if (recType == 0xF004)
            {
                // SpContainer — 자식들에서 spid + pib 추출.
                int spid = ExtractSpid(data, dataStart, dataEnd);
                int pib  = ExtractPib(data, dataStart, dataEnd);
                if (spid != 0 && pib > 0) map[spid] = pib;
            }
            if (recVer == 0xF) WalkSpContainers(data, dataStart, dataEnd, map, depth + 1);
            pos = dataEnd;
        }
    }

    // SpContainer 의 자식 중 Sp atom (0xF009) 을 찾아 spid (body[0..3]) 반환. 없으면 0.
    private static int ExtractSpid(byte[] data, int start, int end)
    {
        int pos = start;
        while (pos + 8 <= end)
        {
            ushort recType = BitConverter.ToUInt16(data, pos + 2);
            uint   recLen  = BitConverter.ToUInt32(data, pos + 4);
            int dataStart  = pos + 8;
            long dataEnd64 = (long)dataStart + recLen;
            if (dataEnd64 > end) return 0;
            if (recType == 0xF009 && dataStart + 4 <= dataEnd64)
                return BitConverter.ToInt32(data, dataStart);
            pos = (int)dataEnd64;
        }
        return 0;
    }

    // SpContainer 의 자식 중 OPT atom (0xF00B) 의 properties 에서 pib (propId=260, fBid=1) 값 추출.
    private static int ExtractPib(byte[] data, int start, int end)
    {
        int pos = start;
        while (pos + 8 <= end)
        {
            ushort verInst = BitConverter.ToUInt16(data, pos);
            ushort recType = BitConverter.ToUInt16(data, pos + 2);
            uint   recLen  = BitConverter.ToUInt32(data, pos + 4);
            int dataStart  = pos + 8;
            long dataEnd64 = (long)dataStart + recLen;
            if (dataEnd64 > end) return 0;
            int dataEnd = (int)dataEnd64;

            if (recType == 0xF00B)
            {
                int propCount = (verInst >> 4) & 0x0FFF;
                int avail = (dataEnd - dataStart) / 6;
                if (propCount > avail) propCount = avail;
                for (int i = 0; i < propCount; i++)
                {
                    int off = dataStart + i * 6;
                    ushort opid = BitConverter.ToUInt16(data, off);
                    uint   op   = BitConverter.ToUInt32(data, off + 2);
                    int propId = opid & 0x3FFF;
                    bool fBid  = (opid & 0x4000) != 0;
                    if (propId == 260 && fBid) return (int)op;
                }
            }
            pos = dataEnd;
        }
        return 0;
    }

    // Phase 3c-2 — SEPX (Section Properties Exception) 의 sprm 들을 Section.Page 에 매핑.
    // SED 의 fcSepx 가 Table stream 내 SEPX 위치. 0xFFFFFFFF/음수 → 적용 안 함 (default).
    // SEPX layout (§2.9.249): cb(2) + grpprl(cb byte).
    private static void ApplySepx(Section section, byte[] table, int fcSepx)
    {
        if (fcSepx < 0 || fcSepx == unchecked((int)0xFFFFFFFF)) return;
        if (fcSepx + 2 > table.Length) return;
        int cb = BitConverter.ToUInt16(table, fcSepx);
        if (cb <= 0 || fcSepx + 2 + cb > table.Length) return;
        var grpprl = new byte[cb];
        Buffer.BlockCopy(table, fcSepx + 2, grpprl, 0, cb);

        WalkSprmsTopLevel(grpprl, (sprm, operand) =>
        {
            // [MS-DOC] section sprm — sgc=4. 흔한 페이지 속성:
            //   sprmSDxaPage   0xB016 (spra=5, 2-byte unsigned) — page width (twips)
            //   sprmSDyaPage   0xB017                            — page height
            //   sprmSDxaLeft   0xB02C                            — margin left
            //   sprmSDxaRight  0xB02D                            — margin right
            //   sprmSDyaTop    0xB02E (signed)                   — margin top
            //   sprmSDyaBottom 0xB02F (signed)                   — margin bottom
            const double TwipsToMm = 1.0 / 56.692;
            switch (sprm)
            {
                case 0xB016 when operand.Length >= 2:
                    section.Page.WidthMm  = BitConverter.ToUInt16(operand, 0) * TwipsToMm; break;
                case 0xB017 when operand.Length >= 2:
                    section.Page.HeightMm = BitConverter.ToUInt16(operand, 0) * TwipsToMm; break;
                case 0xB02C when operand.Length >= 2:
                    section.Page.MarginLeftMm   = BitConverter.ToUInt16(operand, 0) * TwipsToMm; break;
                case 0xB02D when operand.Length >= 2:
                    section.Page.MarginRightMm  = BitConverter.ToUInt16(operand, 0) * TwipsToMm; break;
                case 0xB02E when operand.Length >= 2:
                    section.Page.MarginTopMm    = BitConverter.ToInt16(operand, 0)  * TwipsToMm; break;
                case 0xB02F when operand.Length >= 2:
                    section.Page.MarginBottomMm = BitConverter.ToInt16(operand, 0)  * TwipsToMm; break;
            }
        });
    }

    // FormatStyles 의 nested WalkSprms 는 private. 동일 알고리즘을 본 클래스 scope 에서 재사용.
    // (spra 비트로 operand 크기 결정; sprmTDefTable 류만 2-byte cb 변형.)
    private static void WalkSprmsTopLevel(byte[] sprms, Action<ushort, byte[]> onSprm)
    {
        int i = 0;
        while (i + 2 <= sprms.Length)
        {
            ushort sprm = BitConverter.ToUInt16(sprms, i);
            i += 2;
            int spra = (sprm >> 13) & 0x07;
            int operandSize;
            switch (spra)
            {
                case 0: case 1: operandSize = 1; break;
                case 2: case 4: case 5: operandSize = 2; break;
                case 3: operandSize = 4; break;
                case 7: operandSize = 3; break;
                case 6:
                    if (i >= sprms.Length) return;
                    if (sprm == 0xD608 || sprm == 0xD605 || sprm == 0xC615)
                    {
                        if (i + 2 > sprms.Length) return;
                        operandSize = BitConverter.ToUInt16(sprms, i);
                        i += 2;
                    }
                    else { operandSize = sprms[i]; i += 1; }
                    break;
                default: return;
            }
            if (i + operandSize > sprms.Length) return;
            var operand = new byte[operandSize];
            Buffer.BlockCopy(sprms, i, operand, 0, operandSize);
            i += operandSize;
            onSprm(sprm, operand);
        }
    }

    // Phase 3d — 헤더/푸터 sub-document 영역의 텍스트 추출 후 doc.Sections[0] 에 매핑.
    // Word 의 본문 다음에 footnote, 그 다음 header/footer subdocument 가 위치.
    // PlcfHdd 의 aCP[n+1] 이 영역 내 sub-story 들의 경계. 본 단계는 첫 두 sub-story 만
    // Header.Center / Footer.Center 로 매핑 (홀/짝/첫페이지 구분은 후속 SEPX 단계).
    // Phase 3d / 3m — 헤더/푸터 sub-document 텍스트 추출 후 doc.Sections[0] 의 Page.Header/Footer 에 매핑.
    //   Phase 3m: CleanSubdocText (plain text) 대신 BuildSubdocParagraphs (Run-rich) 로 처리해
    //   하이퍼링크 / 책갈피 / 굵게·이탤릭 등 서식 보존.
    private static void ApplyHeaderFooter(byte[] wd, byte[] table, Fib fib, PolyDonkyument doc, FormatStyles fmt)
    {
        if (fib.CcpHdd == 0 || fib.LcbPlcfHdd < 4 || doc.Sections.Count == 0) return;
        int hddBase = (int)(fib.CcpText + fib.CcpFtn);
        int hddEnd  = hddBase + (int)fib.CcpHdd;

        int plcStart = (int)fib.FcPlcfHdd;
        int plcLen   = (int)fib.LcbPlcfHdd;
        if (plcStart < 0 || plcStart + plcLen > table.Length) return;
        int n = plcLen / 4 - 1;
        if (n <= 0) return;
        var subCps = new int[n + 1];
        for (int i = 0; i <= n; i++)
            subCps[i] = BitConverter.ToInt32(table, plcStart + i * 4);

        // Phase 3d-2 — [MS-DOC] §2.8.7 PlcfHdd 의 sub-story 순서 (section 별 6 stories block):
        //   index%6 == 0 : even page footer
        //   index%6 == 1 : odd page footer  ← 일반 푸터
        //   index%6 == 2 : even page header
        //   index%6 == 3 : odd page header  ← 일반 헤더
        //   index%6 == 4 : first page footer
        //   index%6 == 5 : first page header
        // IWPF 의 Page.Header/Footer 는 단일 슬롯이라 odd 만 매핑 (fallback: even → first).
        // 첫 섹션의 6 stories 만 처리.
        for (int i = 0; i < Math.Min(n, 6); i++)
        {
            int cpStart = hddBase + subCps[i];
            int cpEnd   = hddBase + subCps[i + 1];
            if (cpEnd <= cpStart || cpEnd > hddEnd) continue;
            // Phase 3m — Run-rich 단락 추출.
            var paras = BuildSubdocParagraphs(wd, table, fib, cpStart, cpEnd, fmt);
            if (paras.Count == 0) continue;

            var page = doc.Sections[0].Page;
            switch (i)
            {
                case 1: case 0: case 4:  // odd footer (1) / even (0) / first (4) — odd 우선, 비어 있으면 fallback
                    if (page.Footer.Center.IsEmpty)
                        foreach (var pa in paras) page.Footer.Center.Paragraphs.Add(pa);
                    break;
                case 3: case 2: case 5:  // odd header (3) / even (2) / first (5)
                    if (page.Header.Center.IsEmpty)
                        foreach (var pa in paras) page.Header.Center.Paragraphs.Add(pa);
                    break;
            }
        }
    }

    // Phase 3g — 각주/미주 sub-document 추출. WordDocument 의 sub-document 순서:
    //   [0, ccpText)                                                  : 본문
    //   [ccpText, +ccpFtn)                                            : 각주
    //   [+ccpHdd)                                                     : 헤더/푸터
    //   [+ccpAtn)                                                     : 주석 (comment)
    //   [+ccpEdn)                                                     : 미주
    // PlcfFndTxt / PlcfEdnTxt 의 aCP 배열이 각 영역 안의 sub-story 경계 (separator + 각 항목).
    // 빈/공백 sub-story (separator) 는 자동 skip.
    private static void ApplyFootnotesAndEndnotes(byte[] wd, byte[] table, Fib fib, PolyDonkyument doc, FormatStyles fmt)
    {
        // 각주 영역.
        if (fib.CcpFtn > 0 && fib.LcbPlcffndTxt >= 8)
        {
            int ftnBase = (int)fib.CcpText;
            int ftnEnd  = ftnBase + (int)fib.CcpFtn;
            ExtractStoriesInto(wd, table, fib,
                (int)fib.FcPlcffndTxt, (int)fib.LcbPlcffndTxt,
                ftnBase, ftnEnd,
                doc.Footnotes, idPrefix: "fn", fmt);
        }

        // 미주 영역 (본문 + 각주 + 헤더/푸터 + 주석 뒤).
        if (fib.CcpEdn > 0 && fib.LcbPlcfendTxt >= 8)
        {
            int ednBase = (int)(fib.CcpText + fib.CcpFtn + fib.CcpHdd + fib.CcpAtn);
            int ednEnd  = ednBase + (int)fib.CcpEdn;
            ExtractStoriesInto(wd, table, fib,
                (int)fib.FcPlcfendTxt, (int)fib.LcbPlcfendTxt,
                ednBase, ednEnd,
                doc.Endnotes, idPrefix: "en", fmt);
        }

        // Phase 3j — 주석 영역 (본문 + 각주 + 헤더/푸터 뒤).
        if (fib.CcpAtn > 0 && fib.LcbPlcfAtnTxt >= 8)
        {
            int atnBase = (int)(fib.CcpText + fib.CcpFtn + fib.CcpHdd);
            int atnEnd  = atnBase + (int)fib.CcpAtn;
            ExtractCommentsInto(wd, table, fib,
                (int)fib.FcPlcfAtnTxt, (int)fib.LcbPlcfAtnTxt,
                atnBase, atnEnd, doc.Comments, fmt);
        }
    }

    // Phase 3j-2 — PlcfAtnRef 의 ATRDPre10 (30 byte) 들에서 author 인덱스 (ibstName) + DTTM 추출,
    //   이미 만들어진 CommentEntry 들에 Author + Date 적용. 1:1 매핑 (PlcfAtnRef 의 i 번째 entry 가
    //   i 번째 comment 와 대응). ATRDPre10 layout:
    //     [0..1]  xstUsrInitl.cch (의 wide-char 개수)
    //     [2..21] xstUsrInitl chars (10 wide chars)
    //     [22..23] ibst       (SttbfAtnBkmk index — skip)
    //     [24..25] ibstName   (SttbfRMark index — author full name)
    //     [26..29] dttm       (packed DateTime)
    private static void ApplyCommentsMetadata(
        byte[] table, Fib fib, IList<CommentEntry> comments, IReadOnlyList<string> authors)
    {
        if (comments.Count == 0 || fib.LcbPlcfAtnRef < 4 + 30) return;
        int fc = (int)fib.FcPlcfAtnRef;
        int lcb = (int)fib.LcbPlcfAtnRef;
        if (fc < 0 || fc + lcb > table.Length) return;

        const int atrdSize = 30;
        int element = 4 + atrdSize;
        int n = (lcb - 4) / element;
        if (n <= 0) return;

        int dataBase = fc + 4 * (n + 1);
        int count = Math.Min(comments.Count, n);
        for (int i = 0; i < count; i++)
        {
            int off = dataBase + i * atrdSize;
            if (off + atrdSize > fc + lcb) break;
            int ibstName = BitConverter.ToUInt16(table, off + 24);
            uint dttm    = BitConverter.ToUInt32(table, off + 26);

            if (ibstName > 0 && ibstName <= authors.Count)
                comments[i].Author = authors[ibstName - 1];   // ibstName is 1-based per spec
            else if (ibstName >= 0 && ibstName < authors.Count)
                comments[i].Author = authors[ibstName];        // 0-based fallback
            if (dttm != 0)
                comments[i].Date = FormatStyles.UnpackDttm(dttm);
        }
    }

    // Phase 3j — PlcfAtnTxt 의 sub-story 들에서 코멘트 텍스트 추출 후 CommentEntry 로 sink 에 추가.
    //   ExtractStoriesInto 의 CommentEntry 변종 (FootnoteEntry vs CommentEntry 가 서로 호환 안 되어 별도 helper).
    // Phase 3m-2 — plain text 변환 대신 BuildSubdocParagraphs 로 Run-rich 단락 추출.
    private static void ExtractCommentsInto(
        byte[] wd, byte[] table, Fib fib,
        int plcStart, int plcLen,
        int subBase, int subEnd,
        IList<CommentEntry> sink, FormatStyles fmt)
    {
        if (plcStart < 0 || plcStart + plcLen > table.Length) return;
        int n = plcLen / 4 - 1;
        if (n <= 0) return;
        var cps = new int[n + 1];
        for (int i = 0; i <= n; i++)
            cps[i] = BitConverter.ToInt32(table, plcStart + i * 4);

        for (int i = 0; i < n; i++)
        {
            int cpStart = subBase + cps[i];
            int cpEnd   = subBase + cps[i + 1];
            if (cpEnd <= cpStart || cpEnd > subEnd) continue;
            var paras = BuildSubdocParagraphs(wd, table, fib, cpStart, cpEnd, fmt);
            // 빈 sub-story (separator) skip.
            if (paras.Count == 0 || paras.All(p => string.IsNullOrWhiteSpace(p.GetPlainText()))) continue;
            var entry = new CommentEntry { Id = $"cmt{sink.Count + 1}" };
            foreach (var pa in paras) entry.Blocks.Add(pa);
            sink.Add(entry);
        }
    }

    // PlcfTxt 형식의 aCP 배열을 walk 해서 각 sub-story 의 텍스트를 추출, FootnoteEntry 로 sink 에 추가.
    //   aCP[0] 는 separator (보통 빈 텍스트), 1.. 부터 실제 항목.
    //   빈 sub-story 는 skip. Phase 3m-2 — Run-rich BuildSubdocParagraphs 사용.
    private static void ExtractStoriesInto(
        byte[] wd, byte[] table, Fib fib,
        int plcStart, int plcLen,
        int subBase, int subEnd,
        IList<FootnoteEntry> sink, string idPrefix, FormatStyles fmt)
    {
        if (plcStart < 0 || plcStart + plcLen > table.Length) return;
        int n = plcLen / 4 - 1;
        if (n <= 0) return;
        var cps = new int[n + 1];
        for (int i = 0; i <= n; i++)
            cps[i] = BitConverter.ToInt32(table, plcStart + i * 4);

        for (int i = 0; i < n; i++)
        {
            int cpStart = subBase + cps[i];
            int cpEnd   = subBase + cps[i + 1];
            if (cpEnd <= cpStart || cpEnd > subEnd) continue;
            var paras = BuildSubdocParagraphs(wd, table, fib, cpStart, cpEnd, fmt);
            if (paras.Count == 0 || paras.All(p => string.IsNullOrWhiteSpace(p.GetPlainText()))) continue;
            var entry = new FootnoteEntry { Id = $"{idPrefix}{sink.Count + 1}" };
            foreach (var pa in paras) entry.Blocks.Add(pa);
            sink.Add(entry);
        }
    }

    // 임의 CP 범위 [cpStart, cpEnd) 의 텍스트를 piece table 따라 추출. ExtractTextWithFcs 의 일반화.
    private static string ExtractSubdocText(byte[] wd, byte[] table, Fib fib, int cpStart, int cpEnd)
    {
        if (fib.LcbClx == 0 || cpEnd <= cpStart) return string.Empty;
        var pcds = ParsePieceTable(table, (int)fib.FcClx, (int)fib.LcbClx, out int[] cps);
        var sb = new StringBuilder();
        for (int i = 0; i < pcds.Count; i++)
        {
            int pcStart = cps[i];
            int pcEnd   = cps[i + 1];
            int effStart = Math.Max(pcStart, cpStart);
            int effEnd   = Math.Min(pcEnd, cpEnd);
            if (effEnd <= effStart) continue;
            int len = effEnd - effStart;
            int offsetInPiece = effStart - pcStart;

            uint fcRaw = pcds[i].Fc;
            bool compressed = (fcRaw & 0x40000000u) != 0;
            int  fc   = (int)(fcRaw & 0x3FFFFFFFu);
            if (compressed)
            {
                fc /= 2;
                int absFc = fc + offsetInPiece;
                if (absFc < 0 || absFc + len > wd.Length) continue;
                sb.Append(DecodeAnsi(wd, absFc, len));
            }
            else
            {
                int absFc = fc + offsetInPiece * 2;
                int byteLen = len * 2;
                if (absFc < 0 || absFc + byteLen > wd.Length) continue;
                sb.Append(Encoding.Unicode.GetString(wd, absFc, byteLen));
            }
        }
        return sb.ToString();
    }

    // Phase 3m — ExtractSubdocText 의 fcs 동행 버전. 각 char 의 file character position 도 함께 반환,
    //   FormatStyles.GetRunStyle 등 fc-기반 lookup 이 가능해진다.
    private static (string Text, int[] Fcs) ExtractSubdocTextWithFcs(
        byte[] wd, byte[] table, Fib fib, int cpStart, int cpEnd)
    {
        if (fib.LcbClx == 0 || cpEnd <= cpStart) return (string.Empty, Array.Empty<int>());
        var pcds = ParsePieceTable(table, (int)fib.FcClx, (int)fib.LcbClx, out int[] cps);
        var sb  = new StringBuilder();
        var fcs = new List<int>();
        for (int i = 0; i < pcds.Count; i++)
        {
            int pcStart = cps[i];
            int pcEnd   = cps[i + 1];
            int effStart = Math.Max(pcStart, cpStart);
            int effEnd   = Math.Min(pcEnd, cpEnd);
            if (effEnd <= effStart) continue;
            int len = effEnd - effStart;
            int offsetInPiece = effStart - pcStart;

            uint fcRaw = pcds[i].Fc;
            bool compressed = (fcRaw & 0x40000000u) != 0;
            int  fc   = (int)(fcRaw & 0x3FFFFFFFu);
            if (compressed)
            {
                fc /= 2;
                int absFc = fc + offsetInPiece;
                if (absFc < 0 || absFc + len > wd.Length) continue;
                var piece = DecodeAnsi(wd, absFc, len);
                for (int j = 0; j < piece.Length; j++) { sb.Append(piece[j]); fcs.Add(absFc + j); }
            }
            else
            {
                int absFc = fc + offsetInPiece * 2;
                int byteLen = len * 2;
                if (absFc < 0 || absFc + byteLen > wd.Length) continue;
                var piece = Encoding.Unicode.GetString(wd, absFc, byteLen);
                for (int j = 0; j < piece.Length; j++) { sb.Append(piece[j]); fcs.Add(absFc + j * 2); }
            }
        }
        return (sb.ToString(), fcs.ToArray());
    }

    // Phase 3m — 머리말/꼬리말 같은 sub-document 영역 [cpStart, cpEnd) 을 BuildParaFromChars 기반 풀
    //   Run 빌더로 처리. \r → 단락 경계, 0x13/0x14/0x15 → 필드, bookmark event → marker.
    //   각주/미주/픽처 (0x01/0x02/0x05/0x08) 는 sub-doc 에서 보통 안 쓰여 skip.
    //   각 Paragraph 는 fc 기반 CHPX/STSH 적용 → 헤더/푸터의 굵게·이탤릭·하이퍼링크 등 보존.
    private static IList<Paragraph> BuildSubdocParagraphs(
        byte[] wd, byte[] table, Fib fib, int cpStart, int cpEnd, FormatStyles fmt)
    {
        var (text, fcs) = ExtractSubdocTextWithFcs(wd, table, fib, cpStart, cpEnd);
        var result = new List<Paragraph>();
        if (text.Length == 0) return result;

        var paraChars = new List<char>();
        var paraFcs   = new List<int>();
        int lastFc    = 0;

        // 필드 추적 (Phase 3a-2/3a-5 와 같은 흐름).
        StringBuilder? fieldInstr = null;
        int fieldMode = 0;
        int resultStartFc  = -1;
        string?    activeUrl       = null;
        FieldType? activeFieldType = null;
        string?    activeFieldArg  = null;

        void FlushPara(int paraEndFc)
        {
            var (paraIstd, ps, _, _, _, _, _) = fmt.GetParagraphInfo(paraEndFc);
            var para = BuildParaFromChars(paraChars, paraFcs, paraIstd, ps, fmt);
            // 빈 단락이라도 \r 만으로 emit (헤더에 연속 \r 있는 케이스 보존).
            // Phase 3h-3 — paragraph mark 의 rev flag 도 적용.
            var (_, paraRev) = fmt.GetRunStyle(paraEndFc, paraIstd);
            if (paraRev.Inserted) para.IsInsertedRevision = true;
            if (paraRev.Deleted)  para.IsDeletedRevision  = true;
            result.Add(para);
            paraChars.Clear();
            paraFcs.Clear();
        }

        for (int i = 0; i < text.Length; i++)
        {
            char c  = text[i];
            int  fc = fcs[i];
            lastFc = fc;
            int absCp = cpStart + i;

            // Phase 3i-2 — bookmark events at this CP (whole-doc CP space).
            var bkEnds = fmt.GetBookmarkEndsAtCp(absCp);
            if (bkEnds is not null) foreach (var name in bkEnds)
            {
                fmt.EnqueueBookmarkEvent(fc, isStart: false, name);
                paraChars.Add('￼'); paraFcs.Add(fc);
            }
            var bkStarts = fmt.GetBookmarkStartsAtCp(absCp);
            if (bkStarts is not null) foreach (var name in bkStarts)
            {
                fmt.EnqueueBookmarkEvent(fc, isStart: true, name);
                paraChars.Add('￼'); paraFcs.Add(fc);
            }

            switch (c)
            {
                case '\r':
                case '\f':
                    fieldMode = 0;
                    FlushPara(fc);
                    break;
                case '':
                    fieldMode = 1; fieldInstr = new StringBuilder();
                    resultStartFc = -1; activeUrl = null; activeFieldType = null; activeFieldArg = null;
                    break;
                case '':
                    fieldMode = 2;
                    if (fieldInstr is not null)
                    {
                        var (t, u, a) = ParseFieldInstr(fieldInstr.ToString());
                        activeFieldType = t; activeUrl = u; activeFieldArg = a;
                    }
                    resultStartFc = fc;
                    break;
                case '':
                    if (resultStartFc >= 0)
                        fmt.AddFieldRange(resultStartFc, fc, activeUrl, activeFieldType, activeFieldArg);
                    fieldMode = 0; fieldInstr = null;
                    resultStartFc = -1; activeUrl = null; activeFieldType = null; activeFieldArg = null;
                    break;
                case '': case '': case '': case '': case '':
                    // 픽처/footnote/comment/cell-mark/drawing — 헤더에선 보통 무시.
                    break;
                case '\t':
                case '\n':
                    if (fieldMode == 1) fieldInstr?.Append(c);
                    else { paraChars.Add(c); paraFcs.Add(fc); }
                    break;
                case '\v':
                    paraChars.Add('\n'); paraFcs.Add(fc);
                    break;
                default:
                    if (c < 0x20) break;
                    if (fieldMode == 1) fieldInstr?.Append(c);
                    else { paraChars.Add(c); paraFcs.Add(fc); }
                    break;
            }
        }
        // 남은 chars 가 있으면 한 단락으로 flush — 종료 \r 없는 sub-doc 안전 처리.
        if (paraChars.Count > 0) FlushPara(lastFc);
        // 빈 trailing 단락 (전부 separator) 제거.
        while (result.Count > 0 && result[^1].Runs.Count == 0)
            result.RemoveAt(result.Count - 1);
        return result;
    }

    // 헤더/푸터 sub-story 의 raw 텍스트에서 \r → \n, 그 외 제어 문자 (필드/cell mark 등) 폐기.
    private static string CleanSubdocText(string raw)
    {
        var sb = new StringBuilder(raw.Length);
        foreach (char c in raw)
        {
            if (c == '\r' || c == '\f') sb.Append('\n');
            else if (c >= 0x20 || c == '\t' || c == '\n') sb.Append(c);
        }
        return sb.ToString().TrimEnd('\n', '\r', '\t', ' ');
    }

    // Variant 타입: VT_LPSTR (0x001E, CP1252) · VT_LPWSTR (0x001F, UTF-16LE) · VT_FILETIME (0x0040)
    // 만 지원 — 그 외 PID/Variant 는 무시 (메타데이터는 본문 무효화 사유가 되지 않음).
    private static void ApplySummaryInformation(byte[] data, PolyDonkyument doc)
    {
        try
        {
            // PropertySetStream header: byteOrder(2)·version(2)·sysID(4)·CLSID(16)·numPropertySets(4)
            //                          + FormatID(16) + sectionOffset(4) = 48 byte until first section.
            if (data.Length < 48) return;
            int sectionOffset = BitConverter.ToInt32(data, 44);
            if (sectionOffset < 0 || sectionOffset + 8 > data.Length) return;

            // PropertySection header: cb(4), cProperties(4)
            int cProps = BitConverter.ToInt32(data, sectionOffset + 4);
            if (cProps <= 0 || cProps > 1024) return;

            int idxBase = sectionOffset + 8;
            for (int i = 0; i < cProps; i++)
            {
                int entryOff = idxBase + i * 8;
                if (entryOff + 8 > data.Length) break;
                int pid    = BitConverter.ToInt32(data, entryOff);
                int offset = BitConverter.ToInt32(data, entryOff + 4);
                int valOff = sectionOffset + offset;
                if (valOff < 0 || valOff + 4 > data.Length) continue;

                int type = BitConverter.ToInt32(data, valOff);
                switch (type)
                {
                    case 0x001E:  // VT_LPSTR (CP1252)
                    case 0x001F:  // VT_LPWSTR (UTF-16LE)
                    {
                        var str = ReadLpStr(data, valOff + 4, ansi: type == 0x001E);
                        if (string.IsNullOrEmpty(str)) break;
                        ApplyStringPid(pid, str, doc);
                        break;
                    }
                    case 0x0040:  // VT_FILETIME (Win32 FILETIME, 8 bytes LE)
                    {
                        if (valOff + 12 > data.Length) break;
                        long ft = BitConverter.ToInt64(data, valOff + 4);
                        if (ft <= 0) break;
                        DateTimeOffset when;
                        try { when = new DateTimeOffset(DateTime.FromFileTimeUtc(ft), TimeSpan.Zero); }
                        catch { break; }
                        switch (pid)
                        {
                            case 0x0C: doc.Metadata.Created  = when; break;
                            case 0x0D: doc.Metadata.Modified = when; break;
                        }
                        break;
                    }
                }
            }
        }
        catch
        {
            // 메타데이터는 best-effort. 어떤 예외도 본문 텍스트 추출을 무효화하지 않는다.
        }
    }

    private static void ApplyStringPid(int pid, string value, PolyDonkyument doc)
    {
        switch (pid)
        {
            case 0x02: doc.Metadata.Title       = value; break;
            case 0x04: doc.Metadata.Author      = value; break;
            case 0x08: doc.Metadata.Editor      = value; break;
            case 0x12: doc.Metadata.Application = value; break;
            case 0x03: doc.Metadata.Custom["Subject"]        = value; break;
            case 0x05: doc.Metadata.Custom["Keywords"]       = value; break;
            case 0x06: doc.Metadata.Custom["Comments"]       = value; break;
            case 0x09: doc.Metadata.Custom["RevisionNumber"] = value; break;
        }
    }

    private static string? ReadLpStr(byte[] data, int offset, bool ansi)
    {
        if (offset + 4 > data.Length) return null;
        int len = BitConverter.ToInt32(data, offset);
        if (len <= 0) return null;
        int byteLen = ansi ? len : len * 2;
        if (offset + 4 + byteLen > data.Length) return null;
        string s;
        if (ansi)
        {
            try { s = Encoding.GetEncoding(1252).GetString(data, offset + 4, byteLen); }
            catch { s = Encoding.Latin1.GetString(data, offset + 4, byteLen); }
        }
        else
        {
            s = Encoding.Unicode.GetString(data, offset + 4, byteLen);
        }
        return s.TrimEnd('\0');
    }

    // ─────────────────────────────── 표 셀 메타 (Phase 2c~2f) ────────────────────

    // 한 행의 셀별 테두리/배경/병합 묶음. ScanTableProps (FormatStyles) 가 sprmTDefTable/sprmTSetShd
    // 에서 채우고 FlushParagraph/FinalizeTable 가 pendingRow.Cells 에 적용.
    //   Phase 2c — 4면 BRC (Top/Left/Bottom/Right)
    //   Phase 2d — sprmTSetShd 의 cvBack → BackgroundHex
    //   Phase 2e — TC97 bf bit 0 (fFirstMerged) / bit 1 (fMerged) — 가로 병합
    //   Phase 2f — TC97 bf bit 5 (fVertMerge)   / bit 6 (fVertRestart) — 세로 병합
    internal sealed record TableCellProps(
        CellBorderSide? Top, CellBorderSide? Left,
        CellBorderSide? Bottom, CellBorderSide? Right,
        string? BackgroundHex,
        bool IsFirstMerged, bool IsMerged,
        bool IsVertMerge,   bool IsVertRestart);

    // ─────────────────────────────── PAPX / CHPX 운영 (Phase 1b) ────────────────
    //
    // [MS-DOC] §2.4.6 BTE (Bin Table Entry) plex: aFC[n+1] + aPnFkp[n], 각 PnFkp 는 4-byte 페이지 번호.
    // [MS-DOC] §2.8.1 PapxFkp (512 byte): rgfc[cpara+1] + rgbx[cpara] (각 13 byte BXPap) + cpara(1 byte at off 511).
    //                                     BXPap = bOffset(1) + PHE(12); bOffset 단위는 2-byte.
    //                                     bOffset 위치의 PapxInFkp = cb(1) + grpprlInPapx(...).
    //                                       cb!=0 → 크기 = cb*2 - 1, grpprl 은 +1 에서 시작
    //                                       cb==0 → 다음 1 byte cb'; 크기 = cb'*2, grpprl 은 +2 에서 시작
    //                                     grpprlInPapx 의 첫 2 byte 는 istd; 그 뒤가 sprm 배열.
    // [MS-DOC] §2.8.2 ChpxFkp (512 byte): rgfc[crun+1] + rgb[crun] (각 1 byte bOffset) + crun(1 byte at off 511).
    //                                     ChpxInFkp = cb(1) + grpprlInChpx(cb byte). istd 없음.
    //
    // 본 단계에서 인식하는 sprm (Word 97+):
    //   PAPX:  sprmPJc80     (0x2461)  — 단락 정렬 1 byte (0=L, 1=C, 2=R, 3=J)
    //          sprmPDxaLeft  (0x845D)  — 왼쪽 들여쓰기 (signed twips, 2-byte)
    //          sprmPDxaRight (0x845E)  — 오른쪽 들여쓰기 (signed twips, 2-byte)
    //          sprmPDxaLeft1 (0x8460)  — 첫 줄 들여쓰기 (signed twips, 2-byte)
    //          sprmPDyaBefore(0xA413)  — 단락 앞 여백 (unsigned twips, 2-byte)
    //          sprmPDyaAfter (0xA415)  — 단락 뒤 여백 (unsigned twips, 2-byte)
    //          sprmPDyaLine  (0x6412)  — 줄 간격 LSPD (4-byte: dyaLine + fMultLinespace)
    //   CHPX:  sprmCFBold (0x0835), sprmCFItalic (0x0836)        — 1 byte toggle/on/off
    //          sprmCFStrike    (0x0837)                           — 1 byte
    //          sprmCKul        (0x2A3E)                           — 1 byte (0=none, 그 외=ON)
    //          sprmCHps        (0x4A43)                           — 2 byte 글자크기 (half-points)
    //          sprmCIco        (0x2A42)                           — 1 byte Word16 팔레트 전경색
    //          sprmCCv         (0x6870)                           — 4 byte RGB+1 (Word 2002+) 전경색
    //          sprmCHighlight  (0x2A0C)                           — 1 byte Word16 팔레트 하이라이트
    // 그 외 sprm 은 본 단계에서 무시 — Phase 1d/2 에서 점진적으로 추가한다 (폰트 패밀리는 STSH/STTB FFN 필요).

    private sealed class FormatStyles
    {
        private readonly byte[] _wd;
        private readonly List<BteEntry> _papxBte;
        private readonly List<BteEntry> _chpxBte;
        // Phase 1d — STTB FFN 에서 추출한 폰트명. sprmCRgFtc{0,1,2} 의 operand 인 ftc 가 이 인덱스.
        private readonly IReadOnlyList<string> _fonts;
        // Phase 1e — STSH (Style Sheet) 의 STD 배열. istd → sti/stk/name. null = 빈 슬롯.
        private readonly IReadOnlyList<StyleDef?> _styles;
        // Phase 3c — PlcfSed 에서 추출한 section 경계 CP. 0 부터 시작, 마지막 = ccpText.
        public IReadOnlyList<int> SectionBoundaryCps { get; }
        // Phase 3c-2 — 각 section 의 SED (fcSepx). null = 빈 SED (default 속성 사용).
        public IReadOnlyList<int> SectionFcSepx { get; }
        // Phase 3c-2 — SEPX 적용에 필요하므로 보존.
        public byte[] TableBytes { get; }
        // Phase 3e-2 — Data stream (이미지 PICF 가 들어 있는 OLE2 stream). null = 없음.
        public byte[]? DataStream { get; set; }
        // Phase 3f-2 — OfficeArtBStoreContainer 에서 추출한 공유 BLIP 목록. 1-based 인덱스.
        // PICF body 안의 OfficeArtFOPT pib (property ID 260) 가 이 인덱스를 참조.
        public IReadOnlyList<(string MediaType, byte[] Data)>? BStoreImages { get; set; }

        // Phase 3a-2 / 3a-5 — 필드 범위 (fcStart ≤ fc < fcEnd 의 result text 에 Url/FieldType/FieldArg 적용).
        //   char-walk 가 0x13(field begin) / 0x14(separator) / 0x15(field end) 를 처리하면서
        //   instr (0x13 ~ 0x14 사이) 를 파싱해 HYPERLINK URL / FieldType (PAGE/DATE/...) / FieldArg
        //   (SEQ 카테고리명, REF 책갈피명, STYLEREF 스타일명, INCLUDETEXT 경로) 를 결정,
        //   result 영역 (0x14 ~ 0x15 사이) 의 fc 범위에 매핑한다.
        private readonly List<(int FcStart, int FcEnd, string? Url, FieldType? Type, string? Arg)> _fieldRanges = new();

        public void AddFieldRange(int fcStart, int fcEnd, string? url, FieldType? type, string? arg)
        {
            if (fcEnd > fcStart && (url is not null || type is not null || arg is not null))
                _fieldRanges.Add((fcStart, fcEnd, url, type, arg));
        }

        public (string? Url, FieldType? Type, string? Arg) GetFieldAtFc(int fc)
        {
            // 범위는 작은 수 — 평탄한 linear scan 으로 충분.
            for (int i = _fieldRanges.Count - 1; i >= 0; i--)
            {
                var r = _fieldRanges[i];
                if (fc >= r.FcStart && fc < r.FcEnd) return (r.Url, r.Type, r.Arg);
            }
            return (null, null, null);
        }

        // FKP 페이지(512 byte) 파싱은 매번 동일 데이터를 다시 만지지 않도록 page → grpprl 캐시.
        // PAPX 는 (istd, sprms) 가 페어로 필요하므로 별도 캐시.
        private readonly Dictionary<(int Pn, int RgIdx), (int Istd, byte[] Sprms)?> _papxCache = new();
        private readonly Dictionary<(int Pn, int RgIdx), byte[]?> _chpxCache = new();

        // Phase 3g-2 — Footnote/Endnote ref CP 배열 (sorted, ascending).
        //   PlcfFndRef/PlcfendRef 의 aCP[0..N-1] — aCP[N] 은 sentinel 이라 제외.
        //   char-walk 에서 0x02 만났을 때 CP lookup 으로 footnote/endnote 인덱스 결정.
        private int[] _footnoteRefCps = Array.Empty<int>();
        private int[] _endnoteRefCps  = Array.Empty<int>();
        // Phase 3j — Comment ref CP 배열 (0x05 char 의 위치).
        private int[] _commentRefCps  = Array.Empty<int>();
        // 본문 char-walk 가 0x02 / 0x05 를 발견하면 fc → (fnId, enId, cmtId) 를 등록.
        // BuildParaFromChars 가 '￼' 마커 만났을 때 lookup.
        private readonly Dictionary<int, (string? FnId, string? EnId, string? CmtId)> _refsByFc = new();

        // Phase 3h-2 — SttbfRMark 의 author 이름 배열. sprmCIbstRMark / sprmCIbstRMarkDel 의 인덱스 참조.
        private string[] _rmarkAuthors = Array.Empty<string>();
        public string? GetRMarkAuthor(int index)
            => index >= 0 && index < _rmarkAuthors.Length ? _rmarkAuthors[index] : null;
        // Phase 3j-2 — outer 코드가 SttbfRMark 배열을 직접 사용해야 할 때 (e.g., ATRDPre10 의 ibstName 해소).
        public IReadOnlyList<string> RMarkAuthors => _rmarkAuthors;

        // Phase 3i-2 — Bookmark events. CP 단위로 시작/끝 이벤트를 저장; char-walk 이 매 CP 마다 확인.
        //   같은 CP 에 여러 책갈피 시작/끝 이 가능하므로 list 형식.
        //   marker fc → events queue: BuildParaFromChars 가 '￼' marker 만나면 하나씩 dequeue.
        private readonly Dictionary<int, List<string>> _bookmarkStartsByCp = new();
        private readonly Dictionary<int, List<string>> _bookmarkEndsByCp   = new();
        private readonly Dictionary<int, List<(bool IsStart, string Name)>> _bookmarkEventsByFc = new();

        public void SetBookmarks(IReadOnlyList<BookmarkEntry> bks)
        {
            _bookmarkStartsByCp.Clear();
            _bookmarkEndsByCp.Clear();
            foreach (var bk in bks)
            {
                if (!_bookmarkStartsByCp.TryGetValue(bk.StartCp, out var s)) _bookmarkStartsByCp[bk.StartCp] = s = new();
                s.Add(bk.Name);
                if (!_bookmarkEndsByCp.TryGetValue(bk.EndCp, out var e)) _bookmarkEndsByCp[bk.EndCp] = e = new();
                e.Add(bk.Name);
            }
        }

        public IReadOnlyList<string>? GetBookmarkStartsAtCp(int cp)
            => _bookmarkStartsByCp.TryGetValue(cp, out var lst) ? lst : null;
        public IReadOnlyList<string>? GetBookmarkEndsAtCp(int cp)
            => _bookmarkEndsByCp.TryGetValue(cp, out var lst) ? lst : null;

        public void EnqueueBookmarkEvent(int fc, bool isStart, string name)
        {
            if (!_bookmarkEventsByFc.TryGetValue(fc, out var lst)) _bookmarkEventsByFc[fc] = lst = new();
            lst.Add((isStart, name));
        }
        public (bool IsStart, string Name)? DequeueBookmarkEvent(int fc)
        {
            if (_bookmarkEventsByFc.TryGetValue(fc, out var lst) && lst.Count > 0)
            {
                var ev = lst[0];
                lst.RemoveAt(0);
                return ev;
            }
            return null;
        }

        public int FindFootnoteRefIndex(int cp) => Array.IndexOf(_footnoteRefCps, cp);
        public int FindEndnoteRefIndex(int cp)  => Array.IndexOf(_endnoteRefCps,  cp);
        public int FindCommentRefIndex(int cp)  => Array.IndexOf(_commentRefCps,  cp);

        public void RegisterRefFc(int fc, string? fnId, string? enId, string? cmtId = null)
        {
            if (fnId is not null || enId is not null || cmtId is not null)
                _refsByFc[fc] = (fnId, enId, cmtId);
        }

        public (string? FnId, string? EnId, string? CmtId) GetRefAtFc(int fc)
            => _refsByFc.TryGetValue(fc, out var v) ? v : (null, null, null);

        private FormatStyles(byte[] wd, byte[] table, List<BteEntry> papx, List<BteEntry> chpx,
                             IReadOnlyList<string> fonts, IReadOnlyList<StyleDef?> styles,
                             IReadOnlyList<int> sectionCps, IReadOnlyList<int> sectionFcSepx,
                             int[] footnoteRefCps, int[] endnoteRefCps, int[] commentRefCps,
                             string[] rmarkAuthors)
        {
            _wd = wd; _papxBte = papx; _chpxBte = chpx; _fonts = fonts; _styles = styles;
            SectionBoundaryCps = sectionCps;
            SectionFcSepx = sectionFcSepx;
            TableBytes = table;
            _footnoteRefCps = footnoteRefCps;
            _endnoteRefCps  = endnoteRefCps;
            _commentRefCps  = commentRefCps;
            _rmarkAuthors   = rmarkAuthors;
        }

        public static FormatStyles Build(byte[] wd, byte[] table, Fib fib)
        {
            var papx   = ReadBte(table, (int)fib.FcPlcfBtePapx, (int)fib.LcbPlcfBtePapx);
            var chpx   = ReadBte(table, (int)fib.FcPlcfBteChpx, (int)fib.LcbPlcfBteChpx);
            var fonts  = ReadSttbfFfn(table, (int)fib.FcSttbfFfn, (int)fib.LcbSttbfFfn);
            var styles = ReadStsh(table, (int)fib.FcStshf, (int)fib.LcbStshf);
            var (sectionCps, sectionFcSepx) = ReadPlcfSed(table, (int)fib.FcPlcfSed, (int)fib.LcbPlcfSed);
            // Phase 3g-2 — PlcfFndRef / PlcfendRef. Plc element = aCP(4) + FRD(2) = 6 byte.
            //   lcb = 4*(N+1) + 2*N = 4 + 6*N. N = (lcb - 4) / 6.
            var fnRefCps = ReadPlcCps(table, (int)fib.FcPlcffndRef, (int)fib.LcbPlcffndRef, frdSize: 2);
            var enRefCps = ReadPlcCps(table, (int)fib.FcPlcfendRef, (int)fib.LcbPlcfendRef, frdSize: 2);
            // Phase 3j — PlcfAtnRef. element = aCP(4) + ATRDPre10(30) = 34 byte. lcb = 4*(N+1) + 30*N = 4 + 34*N.
            var cmtRefCps = ReadPlcCps(table, (int)fib.FcPlcfAtnRef, (int)fib.LcbPlcfAtnRef, frdSize: 30);
            // Phase 3h-2 — SttbfRMark: extended SttbExtend (fExtend=0xFFFF + cData + cbExtra + entries).
            var rmarkAuthors = ReadSttbExtend(table, (int)fib.FcSttbfRMark, (int)fib.LcbSttbfRMark);
            return new FormatStyles(wd, table, papx, chpx, fonts, styles, sectionCps, sectionFcSepx,
                                    fnRefCps, enRefCps, cmtRefCps, rmarkAuthors);
        }

        // Phase 3h-2 — SttbExtend (extended Sttb) 파싱. Format:
        //   [0..1]   fExtend = 0xFFFF (extended marker)
        //   [2..3]   cData  (number of strings)
        //   [4..5]   cbExtra (typically 0)
        //   For each entry: [cchData (2 byte)] + [data (cchData * 2 byte UTF-16)] + [cbExtra bytes extra]
        // SttbfRMark 는 항상 extended.
        internal static string[] ReadSttbExtend(byte[] table, int fc, int lcb)
        {
            if (lcb < 6 || fc < 0 || fc + lcb > table.Length) return Array.Empty<string>();
            int pos = fc;
            ushort fExtend = BitConverter.ToUInt16(table, pos); pos += 2;
            if (fExtend != 0xFFFF) return Array.Empty<string>();
            int cData = BitConverter.ToUInt16(table, pos); pos += 2;
            int cbExtra = BitConverter.ToUInt16(table, pos); pos += 2;
            if (cData <= 0) return Array.Empty<string>();
            var arr = new string[cData];
            int end = fc + lcb;
            for (int i = 0; i < cData; i++)
            {
                if (pos + 2 > end) return arr.Take(i).ToArray();
                int cch = BitConverter.ToUInt16(table, pos); pos += 2;
                int byteLen = cch * 2;
                if (pos + byteLen + cbExtra > end) return arr.Take(i).ToArray();
                arr[i] = cch > 0 ? System.Text.Encoding.Unicode.GetString(table, pos, byteLen) : string.Empty;
                pos += byteLen + cbExtra;
            }
            return arr;
        }

        internal static int[] ReadPlcCps(byte[] table, int fc, int lcb, int frdSize)
        {
            if (lcb < 4 || fc < 0 || fc + lcb > table.Length) return Array.Empty<int>();
            int element = 4 + frdSize;
            int n = (lcb - 4) / element;
            if (n <= 0) return Array.Empty<int>();
            var arr = new int[n];
            for (int i = 0; i < n; i++) arr[i] = BitConverter.ToInt32(table, fc + i * 4);
            return arr;
        }

        // [MS-DOC] §2.8.31 PlcfSed — aCP[n+1] (각 4 byte CP) + aSed[n] (각 12 byte SED).
        // SED layout (§2.9.241): fn(2) + fcSepx(4) + fnMpr(2) + fcMpr(4) = 12 byte.
        // Phase 3c-2 — fcSepx 만 추출 (SEPX 위치). 0xFFFFFFFF (-1) 면 SEPX 없음 (default).
        private static (IReadOnlyList<int>, IReadOnlyList<int>) ReadPlcfSed(byte[] table, int fc, int lcb)
        {
            if (lcb < 4 || fc < 0 || fc + lcb > table.Length)
                return (Array.Empty<int>(), Array.Empty<int>());
            int n = (lcb - 4) / 16;
            if (n <= 0) return (Array.Empty<int>(), Array.Empty<int>());
            var cps   = new int[n + 1];
            var fcSepx = new int[n];
            for (int i = 0; i <= n; i++)
                cps[i] = BitConverter.ToInt32(table, fc + i * 4);
            int sedBase = fc + (n + 1) * 4;
            for (int i = 0; i < n; i++)
                fcSepx[i] = BitConverter.ToInt32(table, sedBase + i * 12 + 2);  // fcSepx @ offset 2
            return (cps, fcSepx);
        }

        // Phase 1f — (istd, ParagraphStyle?, InTable, IsTtp) — Phase 1 기본.
        // Phase 2a — InTable / IsTtp 플래그 추가 (sprmPFInTable / sprmPFTtp).
        // Phase 2b — Rgdxa (sprmTDefTable 의 셀 boundary, twips) 추가.
        // Phase 2c — CellProps (sprmTDefTable 의 rgTc 에서 추출한 셀 테두리) 추가.
        // Phase 2h — Itap (sprmPItap 의 nesting level): 0=일반, 1=표 안, 2=중첩 표 안.
        public (int Istd, ParagraphStyle? Style, bool InTable, bool IsTtp,
                short[]? Rgdxa, TableCellProps[]? CellProps, int Itap) GetParagraphInfo(int paraEndFc)
        {
            var papx = LoadPapx(paraEndFc);
            if (papx is null) return (-1, null, false, false, null, null, 0);

            var (istd, directSprms) = papx.Value;
            var style = new ParagraphStyle();
            bool touched = false;
            bool inTable = false;
            bool isTtp   = false;
            short[]? rgdxa = null;
            TableCellProps[]? cellProps = null;
            int itap = 0;

            // 1. STSH built-in sti → Outline (Heading N → HN).
            if (istd >= 0 && istd < _styles.Count && _styles[istd] is { } sd && sd.Sti >= 1 && sd.Sti <= 9)
            {
                int level = Math.Min(sd.Sti, 6);
                style.Outline = (OutlineLevel)level;
                touched = true;
            }

            // 2. Phase 1g — istdBase 체인을 따라 root → leaf 순으로 STD PAPX sprms + 표 sprm 적용.
            foreach (int chainIstd in ResolveStyleChain(istd))
            {
                if (_styles[chainIstd]?.PapxSprms is { Length: > 0 } chainSprms)
                {
                    touched |= ApplyParagraphSprms(chainSprms, style);
                    ScanTableProps(chainSprms, ref inTable, ref isTtp, ref rgdxa, ref cellProps, ref itap);
                }
            }

            // 3. 직접 PAPX sprms — 스타일 상속값을 덮어쓴다.
            touched |= ApplyParagraphSprms(directSprms, style);
            ScanTableProps(directSprms, ref inTable, ref isTtp, ref rgdxa, ref cellProps, ref itap);

            // sprmPItap 가 명시 안 되어도 InTable=true / IsTtp=true 면 1-level 표로 간주 — Word 95
            // legacy 호환 + 합성 케이스. itap > 0 이 명시된 경우 (중첩) 는 그대로.
            if (itap == 0 && (inTable || isTtp)) itap = 1;
            return (istd, touched ? style : null, inTable, isTtp, rgdxa, cellProps, itap);
        }

        // [MS-DOC] 단락의 표 관련 sprm 일괄 스캔:
        //   sprmPFInTable (0x2416, 1-byte) — 단락이 표 안에 있는지
        //   sprmPFTtp     (0x2417, 1-byte) — 단락이 행 종료 단락(TTP)인지
        //   sprmTDefTable (0xD608, variable, spra=6 2-byte length)
        //                — TTP 단락의 PAPX 에 포함. itcMac(1) + rgdxaCenter[itcMac+1] (2 byte signed each)
        //                  + rgTc[itcMac] (20 byte each, §2.9.301 TC97).
        // ref 매개변수는 람다 캡처 불가라 로컬에 모은 뒤 호출부에서 합친다.
        private static void ScanTableProps(byte[] grpprl,
            ref bool inTable, ref bool isTtp, ref short[]? rgdxa, ref TableCellProps[]? cellProps,
            ref int itap)
        {
            bool localIn  = inTable;
            bool localTtp = isTtp;
            short[]? localDxa = rgdxa;
            TableCellProps[]? localCp = cellProps;
            int localItap = itap;
            WalkSprms(grpprl, (sprm, operand) =>
            {
                if (sprm == 0x2416 && operand.Length >= 1) localIn = operand[0] != 0;
                else if (sprm == 0x2417 && operand.Length >= 1) localTtp = operand[0] != 0;
                else if (sprm == 0x6649 && operand.Length >= 4)
                {
                    // sprmPItap (spra=3, 4-byte signed) — nesting level. itap>=1 면 표 안.
                    int v = BitConverter.ToInt32(operand, 0);
                    if (v >= 0 && v <= 8) localItap = v;
                }
                else if (sprm == 0xD608)
                {
                    if (operand.Length < 1) return;
                    int itcMac = operand[0];
                    int dxaBase  = 1;
                    int tcBase   = 1 + 2 * (itcMac + 1);
                    int needed   = tcBase + 20 * itcMac;
                    if (itcMac <= 0 || itcMac > 63 || operand.Length < needed) return;
                    var dxa = new short[itcMac + 1];
                    for (int j = 0; j <= itcMac; j++)
                        dxa[j] = BitConverter.ToInt16(operand, dxaBase + j * 2);
                    localDxa = dxa;
                    // TC97[itcMac] — 각 20 byte: bf(2)+wUnused(2)+brcTop(4)+brcLeft(4)+brcBottom(4)+brcRight(4)
                    // bf bit 0 = fFirstMerged, bit 1 = fMerged (가로 병합).
                    // bf bit 5 = fVertMerge, bit 6 = fVertRestart (세로 병합).
                    var tcs = new TableCellProps[itcMac];
                    for (int j = 0; j < itcMac; j++)
                    {
                        int tcOff = tcBase + j * 20;
                        ushort bf = BitConverter.ToUInt16(operand, tcOff);
                        tcs[j] = new TableCellProps(
                            ParseBrc80(operand, tcOff + 4),
                            ParseBrc80(operand, tcOff + 8),
                            ParseBrc80(operand, tcOff + 12),
                            ParseBrc80(operand, tcOff + 16),
                            BackgroundHex: null,
                            IsFirstMerged: (bf & 0x0001) != 0,
                            IsMerged:      (bf & 0x0002) != 0,
                            IsVertMerge:   (bf & 0x0020) != 0,
                            IsVertRestart: (bf & 0x0040) != 0);
                    }
                    localCp = tcs;
                }
                else if (sprm == 0xD612)
                {
                    // [MS-DOC] §2.9.297 sprmTSetShd — 셀 배경 음영.
                    // operand: itcFirst(1) + itcLim(1) + Shd[New] (10 byte: cvFore(4)+cvBack(4)+ipat(2))
                    // cvBack 의 byte 0~2 = R/G/B, byte 3 = cvType (0xFF=auto). ipat=1(solid) 일 때 적용.
                    if (operand.Length < 12 || localCp is null) return;
                    int itcFirst = operand[0];
                    int itcLim   = operand[1];
                    if (itcFirst < 0 || itcLim <= itcFirst || itcLim > localCp.Length) return;
                    // cvBack @ operand[6..9]
                    byte cvBackAuto = operand[9];
                    if (cvBackAuto == 0xFF) return;  // auto — 적용 안 함
                    // ipat @ operand[10..11], 0=clear/no fill, 1=solid, 그 외 패턴
                    ushort ipat = BitConverter.ToUInt16(operand, 10);
                    if (ipat == 0) return;  // clear — 배경 없음
                    string hex = $"#{operand[6]:X2}{operand[7]:X2}{operand[8]:X2}";
                    for (int j = itcFirst; j < itcLim; j++)
                        localCp[j] = localCp[j] with { BackgroundHex = hex };
                }
            });
            inTable   = localIn;
            isTtp     = localTtp;
            rgdxa     = localDxa;
            cellProps = localCp;
            itap      = localItap;
        }

        // [MS-DOC] §2.9.16 Brc80 (4 byte border code):
        //   byte 0: dptLineWidth (1/8 pt 단위)
        //   byte 1: brcType (0=none, 1=single, 3=double, 5=dotted, 6=dashed, 7=dotDash, 8=dotDotDash, ...)
        //   byte 2: ico (Word 16-color palette index)
        //   byte 3: dptSpace(5)+fShadow(1)+fFrame(1)+reserved(1)
        private static CellBorderSide? ParseBrc80(byte[] tc, int off)
        {
            if (off + 4 > tc.Length) return null;
            byte dpt  = tc[off];
            byte type = tc[off + 1];
            byte ico  = tc[off + 2];
            if (type == 0 || dpt == 0) return null;  // none
            var line = type switch
            {
                3            => BorderLineStyle.Double,
                5            => BorderLineStyle.Dotted,
                6            => BorderLineStyle.Dashed,
                7 or 8       => BorderLineStyle.DashDot,
                _            => BorderLineStyle.Solid,
            };
            string? colorHex = WordPaletteColor(ico)?.ToHex();
            return new CellBorderSide(dpt / 8.0, colorHex, line);
        }

        // TableCellProps 는 외부(DocBinaryReader)에서 BuildDocument 가 raw row + cp 누적에 사용하므로
        // outer scope 로 이동됨. ScanTableProps 가 채우고 FlushParagraph 가 적용.

        // Phase 1g — istd 부터 istdBase 를 따라가 root 까지 chain 을 모은 뒤 root 부터 순회.
        // 부모가 먼저 적용되고 자식이 덮어쓰는 순서. 순환 참조와 nil(0xFFF) 종료를 모두 처리.
        private IEnumerable<int> ResolveStyleChain(int startIstd)
        {
            if (startIstd < 0 || startIstd >= _styles.Count) yield break;
            var chain  = new List<int>();
            var seen   = new HashSet<int>();
            int cur    = startIstd;
            while (cur >= 0 && cur < _styles.Count && cur != IstdNil && seen.Add(cur))
            {
                var sd = _styles[cur];
                if (sd is null) break;
                chain.Add(cur);
                cur = sd.IstdBase;
            }
            // root → leaf 순서로 yield (부모 먼저 → 자식 덮어쓰기).
            for (int i = chain.Count - 1; i >= 0; i--)
                yield return chain[i];
        }

        private static bool ApplyParagraphSprms(byte[] grpprl, ParagraphStyle style)
        {
            bool touched = false;
            WalkSprms(grpprl, (sprm, operand) =>
            {
                if (ApplyParagraphSprm(sprm, operand, style)) touched = true;
            });
            return touched;
        }

        private static bool ApplyParagraphSprm(ushort sprm, byte[] operand, ParagraphStyle style)
        {
            switch (sprm)
            {
                case 0x2461:  // sprmPJc80 — alignment
                    if (operand.Length >= 1)
                    {
                        style.Alignment = operand[0] switch
                        {
                            1 => Alignment.Center,
                            2 => Alignment.Right,
                            3 => Alignment.Justify,
                            _ => Alignment.Left,
                        };
                        return true;
                    }
                    return false;
                case 0x2407:  // sprmPFPageBreakBefore — 단락 앞 강제 페이지 분할
                    if (operand.Length >= 1)
                    { style.ForcePageBreakBefore = operand[0] != 0; return true; }
                    return false;
                case 0x845D:
                    if (operand.Length >= 2)
                    { style.IndentLeftMm  = BitConverter.ToInt16(operand, 0) * TwipsToMm; return true; }
                    return false;
                case 0x845E:
                    if (operand.Length >= 2)
                    { style.IndentRightMm = BitConverter.ToInt16(operand, 0) * TwipsToMm; return true; }
                    return false;
                case 0x8460:
                    if (operand.Length >= 2)
                    { style.IndentFirstLineMm = BitConverter.ToInt16(operand, 0) * TwipsToMm; return true; }
                    return false;
                case 0xA413:
                    if (operand.Length >= 2)
                    { style.SpaceBeforePt = BitConverter.ToUInt16(operand, 0) / 20.0; return true; }
                    return false;
                case 0xA415:
                    if (operand.Length >= 2)
                    { style.SpaceAfterPt  = BitConverter.ToUInt16(operand, 0) / 20.0; return true; }
                    return false;
                case 0x6412:
                    if (operand.Length >= 4)
                    {
                        short  dyaLine = BitConverter.ToInt16(operand, 0);
                        ushort fMult   = BitConverter.ToUInt16(operand, 2);
                        if (fMult == 1 && dyaLine > 0)
                        { style.LineHeightFactor = dyaLine / 240.0; return true; }
                        if (fMult == 0 && dyaLine != 0)
                        {
                            double abs = Math.Abs(dyaLine) / 240.0;
                            if (abs > 0.5 && abs < 5.0)
                            { style.LineHeightFactor = abs; return true; }
                        }
                    }
                    return false;
            }
            return false;
        }

        // Twips → mm 변환 — 1 mm = 56.692 twips (1440/25.4).
        private const double TwipsToMm = 1.0 / 56.692;

        // Phase 3e-2 — 특정 char(picture marker 0x01) 의 CHPX 에서 sprmCPicLocation 추출.
        // [MS-DOC] §2.8.6 sprmCPicLocation (0x6A03, spra=3 4-byte) — Data stream 내 PICF 위치.
        public int? GetPictureFc(int charFc)
        {
            var chpx = LoadChpx(charFc);
            if (chpx is null) return null;
            int? fc = null;
            WalkSprms(chpx, (sprm, operand) =>
            {
                if (sprm == 0x6A03 && operand.Length >= 4)
                    fc = BitConverter.ToInt32(operand, 0);
            });
            return fc;
        }

        // Phase 3l-2 — sprmCFOle2 (0x080A, 1-byte bool) — 이 char 가 OLE 객체 embed 인지 검사.
        //   true 면 0x01 이 이미지가 아닌 임베드 OLE 객체 (Equation, Excel 등) 의 placeholder.
        public bool GetCharIsOle(int charFc)
        {
            var chpx = LoadChpx(charFc);
            if (chpx is null) return false;
            bool isOle = false;
            WalkSprms(chpx, (sprm, operand) =>
            {
                if (sprm == 0x080A && operand.Length >= 1 && operand[0] != 0) isOle = true;
            });
            return isOle;
        }

        // Phase 1f — paraIstd 가 -1 아니면 단락 스타일의 STD chpxSprms 를 먼저 적용한 뒤
        // 직접 CHPX 로 override. Heading 의 폰트/크기/굵게 등이 자동 상속된다.
        // Phase 3h — sprmCFRMarkIns / sprmCFRMarkDel 도 동시에 추출해 revision flags 반환.
        // Phase 3h-2 — sprmCIbstRMark/Del (author index) + sprmCDttmRMark/Del (DTTM) 도 추출.
        public (RunStyle? Style, RunRevInfo Rev) GetRunStyle(int charFc, int paraIstd)
        {
            var rs = new RunStyle();
            var info = new RunRevInfo();
            bool touched = false;

            // 1. Phase 1g — 단락 스타일의 istdBase 체인을 따라 root → leaf 순으로 STD CHPX sprms 적용.
            foreach (int chainIstd in ResolveStyleChain(paraIstd))
            {
                if (_styles[chainIstd]?.ChpxSprms is { Length: > 0 } chainChpx)
                    touched |= ApplyRunSprms(chainChpx, rs, info);
            }

            // 2. 직접 CHPX FKP sprms (override).
            var direct = LoadChpx(charFc);
            if (direct is { Length: > 0 })
                touched |= ApplyRunSprms(direct, rs, info);

            // SttbfRMark 인덱스를 작성자 이름으로 환산.
            if (info.RMarkAuthorIndex >= 0)
                info.Author = GetRMarkAuthor(info.RMarkAuthorIndex);

            return (touched ? rs : null, info);
        }

        // Revision flags 를 lambda 캡처용으로 box 한 holder.
        public sealed class RunRevInfo
        {
            public bool Inserted;
            public bool Deleted;
            public int  RMarkAuthorIndex = -1;   // sprmCIbstRMark / sprmCIbstRMarkDel
            public uint RMarkDttm;                // sprmCDttmRMark / sprmCDttmRMarkDel (packed DTTM)
            public bool HasDttm;
            public string? Author;                // resolved from index
        }

        private bool ApplyRunSprms(byte[] grpprl, RunStyle rs, RunRevInfo info)
        {
            bool touched = false;
            WalkSprms(grpprl, (sprm, operand) =>
            {
                // Phase 3h — revision sprms (1-byte bool).
                if (sprm == 0x0801 && operand.Length >= 1)      { info.Inserted = operand[0] != 0; touched = true; }
                else if (sprm == 0x0807 && operand.Length >= 1) { info.Deleted  = operand[0] != 0; touched = true; }
                // Phase 3h-2 — author index (2-byte) + DTTM (4-byte). Ins / Del 변종 모두 동일 슬롯에 저장
                //   (Run 단위로는 ins XOR del 가 일반적이라 충돌 없음).
                else if ((sprm == 0x4804 || sprm == 0x4863) && operand.Length >= 2)
                {
                    info.RMarkAuthorIndex = BitConverter.ToUInt16(operand, 0);
                    touched = true;
                }
                else if ((sprm == 0x6805 || sprm == 0x6864) && operand.Length >= 4)
                {
                    info.RMarkDttm = BitConverter.ToUInt32(operand, 0);
                    info.HasDttm   = true;
                    touched = true;
                }
                else if (ApplyRunSprm(sprm, operand, rs)) touched = true;
            });
            return touched;
        }

        // Phase 3h-2 — DTTM (packed 32-bit) → DateTimeOffset.
        //   bits  0..5  : minute (0-59)
        //   bits  6..10 : hour   (0-23)
        //   bits 11..15 : day    (1-31)
        //   bits 16..19 : month  (1-12)
        //   bits 20..28 : year - 1900
        //   bits 29..31 : day-of-week (무시)
        public static System.DateTimeOffset? UnpackDttm(uint dttm)
        {
            if (dttm == 0) return null;
            int min  = (int)(dttm        & 0x3F);
            int hour = (int)((dttm >> 6) & 0x1F);
            int day  = (int)((dttm >> 11) & 0x1F);
            int mon  = (int)((dttm >> 16) & 0x0F);
            int year = (int)((dttm >> 20) & 0x1FF) + 1900;
            if (mon < 1 || mon > 12 || day < 1 || day > 31 || hour > 23 || min > 59) return null;
            try { return new System.DateTimeOffset(year, mon, day, hour, min, 0, System.TimeSpan.Zero); }
            catch { return null; }
        }

        private bool ApplyRunSprm(ushort sprm, byte[] operand, RunStyle rs)
        {
            switch (sprm)
            {
                case 0x0835:  // sprmCFBold
                    if (operand.Length >= 1) { rs.Bold = operand[0] != 0; return true; }
                    return false;
                case 0x0836:  // sprmCFItalic
                    if (operand.Length >= 1) { rs.Italic = operand[0] != 0; return true; }
                    return false;
                case 0x0837:  // sprmCFStrike
                    if (operand.Length >= 1) { rs.Strikethrough = operand[0] != 0; return true; }
                    return false;
                case 0x2A3E:  // sprmCKul
                    if (operand.Length >= 1) { rs.Underline = operand[0] != 0; return true; }
                    return false;
                case 0x4A43:  // sprmCHps (font size in half-points)
                    if (operand.Length >= 2)
                    {
                        ushort halfPt = BitConverter.ToUInt16(operand, 0);
                        if (halfPt > 0 && halfPt < 1000) { rs.FontSizePt = halfPt / 2.0; return true; }
                    }
                    return false;
                case 0x2A42:  // sprmCIco
                    if (operand.Length >= 1)
                    {
                        var c = WordPaletteColor(operand[0]);
                        if (c.HasValue) { rs.Foreground = c.Value; return true; }
                    }
                    return false;
                case 0x6870:  // sprmCCv
                    if (operand.Length >= 4 && operand[3] != 0xFF)
                    {
                        rs.Foreground = new Color(operand[0], operand[1], operand[2]);
                        return true;
                    }
                    return false;
                case 0x2A0C:  // sprmCHighlight
                    if (operand.Length >= 1)
                    {
                        var c = WordPaletteColor(operand[0]);
                        if (c.HasValue) { rs.Background = c.Value; return true; }
                    }
                    return false;
                case 0x4A4F:  // sprmCRgFtc0
                case 0x4A50:  // sprmCRgFtc1
                case 0x4A51:  // sprmCRgFtc2
                    if (operand.Length >= 2)
                    {
                        int ftc = BitConverter.ToUInt16(operand, 0);
                        if (ftc >= 0 && ftc < _fonts.Count)
                        {
                            var name = _fonts[ftc];
                            if (!string.IsNullOrEmpty(name))
                            { rs.FontFamily = name; return true; }
                        }
                    }
                    return false;
            }
            return false;
        }

        // [MS-DOC] Word 16-color palette (sprmCIco, sprmCHighlight 의 1-byte 인덱스).
        // 0 = auto (부모 색 상속 — null 반환).
        private static Color? WordPaletteColor(byte ico) => ico switch
        {
            1  => new Color(0,   0,   0),     //  1 Black
            2  => new Color(0,   0,   255),   //  2 Blue
            3  => new Color(0,   255, 255),   //  3 Cyan
            4  => new Color(0,   255, 0),     //  4 Green
            5  => new Color(255, 0,   255),   //  5 Magenta
            6  => new Color(255, 0,   0),     //  6 Red
            7  => new Color(255, 255, 0),     //  7 Yellow
            8  => new Color(255, 255, 255),   //  8 White
            9  => new Color(0,   0,   128),   //  9 Dark Blue
            10 => new Color(0,   128, 128),   // 10 Dark Cyan
            11 => new Color(0,   128, 0),     // 11 Dark Green
            12 => new Color(128, 0,   128),   // 12 Dark Magenta
            13 => new Color(128, 0,   0),     // 13 Dark Red
            14 => new Color(128, 128, 0),     // 14 Dark Yellow
            15 => new Color(128, 128, 128),   // 15 Dark Gray
            16 => new Color(192, 192, 192),   // 16 Light Gray
            _  => null,                       //  0 auto / 그 외 알 수 없음
        };

        // PAPX 로드 — istd 와 sprm bytes 를 한 쌍으로 반환. STSH lookup 에 istd 가 필요해 분리.
        private (int Istd, byte[] Sprms)? LoadPapx(int fc)
        {
            var loc = LocateFkpEntry(_papxBte, fc);
            if (loc is null) return null;
            var (pn, rgIdx) = loc.Value;
            if (_papxCache.TryGetValue((pn, rgIdx), out var cached)) return cached;

            int fkpOff = pn * 512;
            int cRun   = _wd[fkpOff + 511];
            int rgbxBase = fkpOff + 4 * (cRun + 1);
            // BXPap = 13 byte; first byte = bOffset (2-byte units from FKP start).
            int bOffset = _wd[rgbxBase + rgIdx * 13];
            var result  = bOffset == 0 ? null : ReadPapxInFkp(_wd, fkpOff + bOffset * 2);
            _papxCache[(pn, rgIdx)] = result;
            return result;
        }

        // CHPX 로드 — sprm bytes 만 (istd 없음).
        private byte[]? LoadChpx(int fc)
        {
            var loc = LocateFkpEntry(_chpxBte, fc);
            if (loc is null) return null;
            var (pn, rgIdx) = loc.Value;
            if (_chpxCache.TryGetValue((pn, rgIdx), out var cached)) return cached;

            int fkpOff = pn * 512;
            int cRun   = _wd[fkpOff + 511];
            int rgbxBase = fkpOff + 4 * (cRun + 1);
            // ChpxFkp rgb = 1 byte each.
            int bOffset = _wd[rgbxBase + rgIdx];
            var result  = bOffset == 0 ? null : ReadChpxInFkp(_wd, fkpOff + bOffset * 2);
            _chpxCache[(pn, rgIdx)] = result;
            return result;
        }

        private (int Pn, int RgIdx)? LocateFkpEntry(List<BteEntry> bte, int fc)
        {
            if (bte.Count == 0) return null;
            int pn = -1;
            foreach (var e in bte)
            {
                if (fc >= e.FcStart && fc < e.FcEnd) { pn = e.PnFkp; break; }
            }
            if (pn < 0) return null;
            int fkpOff = pn * 512;
            if (fkpOff < 0 || fkpOff + 512 > _wd.Length) return null;
            int cRun = _wd[fkpOff + 511];
            if (cRun == 0) return null;
            for (int i = 0; i < cRun; i++)
            {
                int fc0 = BitConverter.ToInt32(_wd, fkpOff + i * 4);
                int fc1 = BitConverter.ToInt32(_wd, fkpOff + (i + 1) * 4);
                if (fc >= fc0 && fc < fc1) return (pn, i);
            }
            return null;
        }

        private static (int Istd, byte[] Sprms)? ReadPapxInFkp(byte[] data, int off)
        {
            if (off < 0 || off >= data.Length) return null;
            byte cb = data[off];
            int grpprlStart, grpprlSize;
            if (cb == 0)
            {
                if (off + 1 >= data.Length) return null;
                byte cb2 = data[off + 1];
                grpprlSize = cb2 * 2;
                grpprlStart = off + 2;
            }
            else
            {
                grpprlSize = cb * 2 - 1;
                grpprlStart = off + 1;
            }
            if (grpprlSize < 2 || grpprlStart + grpprlSize > data.Length) return null;
            // GrpPrlAndIstd: 처음 2 byte = istd (style identifier), 그 뒤가 sprm 배열.
            int istd = BitConverter.ToUInt16(data, grpprlStart);
            int sprmsLen = grpprlSize - 2;
            var sprms = new byte[sprmsLen];
            Buffer.BlockCopy(data, grpprlStart + 2, sprms, 0, sprmsLen);
            return (istd, sprms);
        }

        private static byte[]? ReadChpxInFkp(byte[] data, int off)
        {
            if (off < 0 || off >= data.Length) return null;
            byte cb = data[off];
            if (cb == 0) return Array.Empty<byte>();
            if (off + 1 + cb > data.Length) return null;
            var sprms = new byte[cb];
            Buffer.BlockCopy(data, off + 1, sprms, 0, cb);
            return sprms;
        }

        // [MS-DOC] §2.6.2 Sprm: 2-byte sprm 헤더 → ispmd(9) + fSpec(1) + sgc(3) + spra(3).
        // spra 가 operand 크기를 결정: 0/1 → 1 byte, 2/4/5 → 2 byte, 3 → 4 byte, 7 → 3 byte,
        // 6 → variable (다음 1 byte 가 길이).
        private static void WalkSprms(byte[] sprms, Action<ushort, byte[]> onSprm)
        {
            int i = 0;
            while (i + 2 <= sprms.Length)
            {
                ushort sprm = BitConverter.ToUInt16(sprms, i);
                i += 2;
                int spra = (sprm >> 13) & 0x07;
                int operandSize;
                switch (spra)
                {
                    case 0: case 1: operandSize = 1; break;
                    case 2: case 4: case 5: operandSize = 2; break;
                    case 3: operandSize = 4; break;
                    case 7: operandSize = 3; break;
                    case 6:
                        if (i >= sprms.Length) return;
                        // [MS-DOC] §2.6.2 spra=6: 가변 길이 sprm.
                        //   기본: 첫 1 byte = operand 길이.
                        //   예외 (§2.9.293 sprmTDefTable / §2.9.292 sprmTDefTable10): 첫 2 byte = 길이.
                        //   sprmPChgTabs (0xC615) 도 같은 2-byte 변형. 그 외는 1-byte.
                        if (sprm == 0xD608 || sprm == 0xD605 || sprm == 0xC615)
                        {
                            if (i + 2 > sprms.Length) return;
                            operandSize = BitConverter.ToUInt16(sprms, i);
                            i += 2;
                        }
                        else
                        {
                            operandSize = sprms[i];
                            i += 1;
                        }
                        break;
                    default: return;
                }
                if (i + operandSize > sprms.Length) return;
                var operand = new byte[operandSize];
                Buffer.BlockCopy(sprms, i, operand, 0, operandSize);
                i += operandSize;
                onSprm(sprm, operand);
            }
        }

        // [MS-DOC] §2.9.262 SttbfFfn — Word 97+ 의 폰트 이름 STTB. 각 원소는 §2.9.85 FFN.
        // Extended STTB header (6 byte): 0xFFFF marker + cData(2) + cbExtra(2).
        // 각 entry: cchData(2 byte, FFN 의 wide-char 수) + FFN(cchData*2 byte) + cbExtra byte.
        // FFN 내부에서 폰트명(xszFfn) 은 offset 40 부터 null-terminated UTF-16LE.
        private static IReadOnlyList<string> ReadSttbfFfn(byte[] table, int fc, int lcb)
        {
            var fonts = new List<string>();
            if (lcb <= 0 || fc < 0 || fc + lcb > table.Length) return fonts;

            int pos = fc;
            int end = fc + lcb;
            bool extended;
            int  cbExtra;
            if (end - pos >= 6 && BitConverter.ToUInt16(table, pos) == 0xFFFF)
            {
                extended = true;
                cbExtra  = BitConverter.ToUInt16(table, pos + 4);
                pos += 6;
            }
            else if (end - pos >= 4)
            {
                // 비-extended (legacy) — Word 97+ 에서는 거의 없지만 fallback.
                extended = false;
                cbExtra  = BitConverter.ToUInt16(table, pos + 2);
                pos += 4;
            }
            else
            {
                return fonts;
            }

            while (pos < end)
            {
                int hdrLen = extended ? 2 : 1;
                if (pos + hdrLen > end) break;
                int cchData = extended ? BitConverter.ToUInt16(table, pos) : table[pos];
                int ffnByteLen = cchData * 2;
                if (pos + hdrLen + ffnByteLen + cbExtra > end) break;

                int ffnStart = pos + hdrLen;
                string name = string.Empty;
                if (ffnByteLen > 40)
                {
                    int nameStart = ffnStart + 40;
                    int nameMax   = ffnStart + ffnByteLen;
                    int nameEnd   = nameStart;
                    while (nameEnd + 1 < nameMax && BitConverter.ToUInt16(table, nameEnd) != 0)
                        nameEnd += 2;
                    if (nameEnd > nameStart)
                        name = Encoding.Unicode.GetString(table, nameStart, nameEnd - nameStart);
                }
                fonts.Add(name);

                pos += hdrLen + ffnByteLen + cbExtra;
            }
            return fonts;
        }

        // [MS-DOC] §2.9.270 STD 의 핵심 필드 — sti(12 bit): built-in style identifier
        // (0=Normal, 1..9=Heading 1..9), stk(4 bit): 1=paragraph, 2=character, 3=table, 4=numbering.
        // Phase 1f — STD 의 grLPUpxSw 에서 추출한 단락/문자 sprm 기본값.
        // Phase 1g — IstdBase(12 bit): 상속 부모 styleId. 0xFFF = nil (체인 종료).
        private sealed record StyleDef(int Sti, int Stk, int IstdBase, string? Name,
                                       byte[]? PapxSprms, byte[]? ChpxSprms);

        private const int IstdNil = 0xFFF;

        // [MS-DOC] §2.9.271 STSH 파서.
        // STSH = LPStshi (2 byte cbStshi + Stshi) + rgLPStd[cstd] (각 LPStd: 2 byte cbStd + STD).
        // 본 단계에서는 STD 의 sti/stk 만 추출 — STD grpprl 상속(default 서식)은 후속 단계.
        private static IReadOnlyList<StyleDef?> ReadStsh(byte[] table, int fc, int lcb)
        {
            var empty = Array.Empty<StyleDef?>();
            if (lcb < 4 || fc < 0 || fc + lcb > table.Length) return empty;

            int pos = fc;
            int end = fc + lcb;

            // LPStshi: cbStshi(2) + Stshi(cbStshi)
            int cbStshi = BitConverter.ToUInt16(table, pos);
            pos += 2;
            if (cbStshi < 4 || pos + cbStshi > end) return empty;

            // Stshi 의 첫 4 byte: cstd(2) + cbSTDBaseInFile(2)
            int cstd            = BitConverter.ToUInt16(table, pos);
            int cbSTDBaseInFile = BitConverter.ToUInt16(table, pos + 2);
            pos += cbStshi;

            if (cstd <= 0 || cstd > 8192) return empty;
            if (cbSTDBaseInFile is not (10 or 18)) cbSTDBaseInFile = 10;

            var styles = new StyleDef?[cstd];
            for (int i = 0; i < cstd && pos + 2 <= end; i++)
            {
                int cbStd = BitConverter.ToUInt16(table, pos);
                pos += 2;
                if (cbStd == 0) { styles[i] = null; continue; }
                if (pos + cbStd > end) break;

                int stdStart = pos;
                int stdEnd   = stdStart + cbStd;
                if (cbStd >= 4)
                {
                    ushort word0 = BitConverter.ToUInt16(table, stdStart);
                    ushort word1 = BitConverter.ToUInt16(table, stdStart + 2);
                    ushort word2 = cbStd >= 6 ? BitConverter.ToUInt16(table, stdStart + 4) : (ushort)0;
                    int sti       = word0 & 0x0FFF;
                    int stk       = word1 & 0x000F;
                    int istdBase  = (word1 >> 4) & 0x0FFF;  // 부모 styleId — 0xFFF 면 nil.
                    int cupx      = word2 & 0x000F;

                    // xstzName — stdfBase(+stdfPost2000) 뒤. Xstz: cchData(2) + chars(2*cchData) + null(2).
                    string? name = null;
                    int nameOff  = stdStart + cbSTDBaseInFile;
                    int afterName = nameOff;
                    if (nameOff + 2 <= stdEnd)
                    {
                        int cch = BitConverter.ToUInt16(table, nameOff);
                        int nameStart = nameOff + 2;
                        int nameByteLen = cch * 2;
                        if (nameStart + nameByteLen + 2 <= stdEnd)
                        {
                            if (cch > 0)
                                name = Encoding.Unicode.GetString(table, nameStart, nameByteLen);
                            afterName = nameStart + nameByteLen + 2;  // +null terminator
                        }
                        else
                        {
                            afterName = stdEnd;  // 손상된 STD — grLPUpxSw 진입 안 함
                        }
                    }
                    // 2-byte 정렬은 STD 내부 offset 기준이라 stdStart 의 parity 와 비교.
                    if (((afterName - stdStart) & 1) != 0) afterName++;

                    // [MS-DOC] §2.9.135 grLPUpxSw — cupx 개의 LPUpx (각각 cbUpx(2) + UPX(cbUpx)).
                    //   paragraph style(stk=1): LPUpx[0]=PAPX (UPX = istd(2)+grpprl), LPUpx[1]=CHPX (UPX=grpprl)
                    //   character style(stk=2): LPUpx[0]=CHPX
                    var (papx, chpx) = ParseGrLPUpxSw(table, stdStart, afterName, stdEnd, cupx, stk);
                    styles[i] = new StyleDef(sti, stk, istdBase, name, papx, chpx);
                }
                pos += cbStd;
            }
            return styles;
        }

        // [MS-DOC] §2.9.135 grLPUpxSw — 한 STD 안의 cupx 개 LPUpx 를 순서대로 파싱.
        // cupx=2 paragraph style: LPUpx[0]=PAPX (UPX는 istd(2)+grpprl), LPUpx[1]=CHPX (UPX=grpprl)
        // cupx=1 character style: LPUpx[0]=CHPX
        // (table/numbering style 은 본 단계에서 무시.)
        private static (byte[]? Papx, byte[]? Chpx) ParseGrLPUpxSw(
            byte[] table, int stdStart, int start, int end, int cupx, int stk)
        {
            byte[]? papx = null;
            byte[]? chpx = null;
            int pos = start;
            for (int idx = 0; idx < cupx && pos + 2 <= end; idx++)
            {
                int cbUpx = BitConverter.ToUInt16(table, pos);
                pos += 2;
                if (pos + cbUpx > end) break;

                if (stk == 1 && idx == 0)
                {
                    // Paragraph style 의 PAPX UPX: istd(2 byte) + grpprl. istd 는 styleId 라 무시.
                    if (cbUpx > 2)
                    {
                        int sprmLen = cbUpx - 2;
                        papx = new byte[sprmLen];
                        Buffer.BlockCopy(table, pos + 2, papx, 0, sprmLen);
                    }
                }
                else if ((stk == 1 && idx == 1) || (stk == 2 && idx == 0))
                {
                    // CHPX UPX = grpprl 그대로.
                    if (cbUpx > 0)
                    {
                        chpx = new byte[cbUpx];
                        Buffer.BlockCopy(table, pos, chpx, 0, cbUpx);
                    }
                }
                pos += cbUpx;
                // LPUpx 는 2-byte 정렬 — STD 내부 offset 기준.
                if (((pos - stdStart) & 1) != 0) pos++;
            }
            return (papx, chpx);
        }

        private readonly record struct BteEntry(int FcStart, int FcEnd, int PnFkp);

        private static List<BteEntry> ReadBte(byte[] table, int fc, int lcb)
        {
            var list = new List<BteEntry>();
            if (lcb <= 4 || fc < 0 || fc + lcb > table.Length) return list;
            // PlcBteFcN: aFC[n+1] (4 byte each) + aPnFkp[n] (4 byte each) → 8n + 4 = lcb
            int n = (lcb - 4) / 8;
            if (n <= 0) return list;
            for (int i = 0; i < n; i++)
            {
                int fcStart = BitConverter.ToInt32(table, fc + i * 4);
                int fcEnd   = BitConverter.ToInt32(table, fc + (i + 1) * 4);
                // PnFkp 의 하위 22 bit 가 페이지 번호 (전체 4 byte 중 bit 0~21).
                int pn      = BitConverter.ToInt32(table, fc + (n + 1) * 4 + i * 4) & 0x003FFFFF;
                list.Add(new BteEntry(fcStart, fcEnd, pn));
            }
            return list;
        }
    }

    // ─────────────────────────────── 유틸 ───────────────────────────────────────

    private static byte[]? ReadAll(RootStorage root, string streamName)
    {
        try
        {
            using var s = root.OpenStream(streamName);
            var ms = new MemoryStream();
            // OpenMcdf v3 stream 은 Seekable. 한 번에 읽는다.
            s.CopyTo(ms);
            return ms.ToArray();
        }
        catch
        {
            return null;
        }
    }

    // Phase 3l — ObjectPool storage 의 child 들 (각각 하나의 임베드 OLE 객체) 을 enumerate 해
    //   storage 이름과 안의 stream 들을 OleEmbedEntry 로 반환.
    //   CompObj 가 있으면 거기서 ClassName (Pascal-style 4-byte length + ANSI string) 도 추출.
    private static IReadOnlyList<OleEmbedEntry> ParseOleEmbeds(OpenMcdf.RootStorage root)
    {
        var list = new List<OleEmbedEntry>();
        if (!root.TryOpenStorage("ObjectPool", out var pool)) return list;

        foreach (var entry in pool.EnumerateEntries())
        {
            // sub-storage 만 다룬다 — stream 직접 자식은 ObjectPool 에서 보통 없음.
            if (!pool.TryOpenStorage(entry.Name, out var objStorage)) continue;

            var streams = new Dictionary<string, byte[]>(StringComparer.Ordinal);
            foreach (var inner in objStorage.EnumerateEntries())
            {
                if (objStorage.TryOpenStream(inner.Name, out var st))
                {
                    using var s = st;
                    using var ms = new MemoryStream();
                    s.CopyTo(ms);
                    streams[inner.Name] = ms.ToArray();
                }
                else
                {
                }
            }
            string? className = streams.TryGetValue("CompObj", out var compObj)
                ? ParseCompObjClass(compObj) : null;
            // 우선순위: EquationNative > Ole10Native > Workbook > 첫 비-제어 stream.
            byte[]? primary = null;
            string? primaryName = null;
            foreach (var preferred in new[] { "EquationNative", "Ole10Native", "Workbook", "PowerPoint Document" })
            {
                if (streams.TryGetValue(preferred, out var data)) { primary = data; primaryName = preferred; break; }
            }
            if (primary is null)
            {
                foreach (var kv in streams)
                {
                    if (kv.Key.Length > 0 && kv.Key[0] >= 0x20)
                    {
                        primary = kv.Value; primaryName = kv.Key; break;
                    }
                }
            }
            // Phase 3l-3 — Ole10Native wrapper 풀어 inner native data + 원본 파일명 추출.
            string? originalFileName = null;
            if (primary is not null && primaryName == "Ole10Native")
            {
                var unwrap = ParseOle10Native(primary);
                if (unwrap.HasValue)
                {
                    primary          = unwrap.Value.Native;
                    originalFileName = unwrap.Value.FileName;
                }
            }
            list.Add(new OleEmbedEntry(entry.Name, className, primaryName, primary,
                                       originalFileName,
                                       (IReadOnlyDictionary<string, byte[]>)streams));
        }
        return list;
    }

    // [MS-OLEDS] CompObj stream — 28-byte header 후 length-prefixed ANSI 클래스 이름.
    //   bytes [0..3]   : Reserved1 (-2)
    //   bytes [4..7]   : Version (보통 0x0A 03 01 00)
    //   bytes [8..27]  : Reserved2 (CLSID, 16 byte) + extra
    //   bytes [28..31] : cch (length of ANSI class string, INCL null terminator)
    //   bytes [32..32+cch-1]: ANSI string (null-terminated)
    // Phase 3l-3 — Ole10Native stream 의 wrapper 파싱. 형식:
    //   [0..3]   TotalSize (uint32 LE) — 이 4 byte 이후 데이터 크기
    //   [4..5]   flag (보통 0x0002) — 임베드 표시
    //   [6..]    null-terminated ANSI class name
    //   다음     null-terminated ANSI original file name
    //   다음     null-terminated ANSI source path
    //   reserved (8 byte: dwReserved + originalPathLen?)
    //   uint32   tempPathLen + 그만큼 ANSI temp path
    //   uint32   nativeDataLen + 그만큼 native data
    private static (byte[] Native, string? FileName)? ParseOle10Native(byte[] stream)
    {
        if (stream.Length < 8) return null;
        int pos = 4;  // skip TotalSize
        ushort flag = BitConverter.ToUInt16(stream, pos); pos += 2;
        if (flag != 0x0002) return null;

        string ReadCString()
        {
            int start = pos;
            while (pos < stream.Length && stream[pos] != 0) pos++;
            if (pos >= stream.Length) return string.Empty;
            var s = System.Text.Encoding.GetEncoding(1252).GetString(stream, start, pos - start);
            pos++;  // skip null
            return s;
        }
        string _className   = ReadCString();
        string fileName     = ReadCString();
        string _sourcePath  = ReadCString();
        if (pos + 8 > stream.Length) return null;
        pos += 8;  // reserved + originalPathLen

        if (pos + 4 > stream.Length) return null;
        uint tempPathLen = BitConverter.ToUInt32(stream, pos); pos += 4;
        if (tempPathLen > 0 && pos + tempPathLen <= stream.Length) pos += (int)tempPathLen;

        if (pos + 4 > stream.Length) return null;
        uint dataLen = BitConverter.ToUInt32(stream, pos); pos += 4;
        if (dataLen == 0 || pos + dataLen > stream.Length) return null;
        var native = new byte[dataLen];
        Buffer.BlockCopy(stream, pos, native, 0, (int)dataLen);
        return (native, fileName.Length > 0 ? fileName : null);
    }

    // Phase 3n — VBA 매크로 프로젝트 storage 를 격리 저장. 콘텐츠는 절대 파싱·실행하지 않고
    //   storage 안의 모든 stream 을 path → bytes 사전으로 그대로 보존 (round-trip 충실도용).
    //   알려진 storage 이름: "Macros" (Word 97-2003), "_VBA_PROJECT_CUR" (older variant).
    private static MacroProjectInfo? ParseMacroProject(OpenMcdf.RootStorage root)
    {
        foreach (var candidate in new[] { "Macros", "_VBA_PROJECT_CUR" })
        {
            if (!root.TryOpenStorage(candidate, out var macroRoot)) continue;
            var streams = new Dictionary<string, byte[]>(StringComparer.Ordinal);
            ReadStorageRecursive(macroRoot, "", streams);
            if (streams.Count == 0) continue;
            return new MacroProjectInfo(candidate, streams);
        }
        return null;
    }

    // 임의 storage 의 모든 stream 을 path-prefixed key 로 dict 에 저장. sub-storage 는 재귀.
    private static void ReadStorageRecursive(
        OpenMcdf.Storage storage, string prefix, Dictionary<string, byte[]> sink)
    {
        foreach (var entry in storage.EnumerateEntries())
        {
            string path = prefix.Length > 0 ? $"{prefix}/{entry.Name}" : entry.Name;
            if (storage.TryOpenStream(entry.Name, out var stm))
            {
                using var s = stm;
                using var ms = new MemoryStream();
                s.CopyTo(ms);
                sink[path] = ms.ToArray();
            }
            else if (storage.TryOpenStorage(entry.Name, out var child))
            {
                ReadStorageRecursive(child, path, sink);
            }
        }
    }

    private static string? ParseCompObjClass(byte[] compObj)
    {
        if (compObj.Length < 36) return null;
        int cch = BitConverter.ToInt32(compObj, 28);
        if (cch <= 1 || cch > 256 || 32 + cch > compObj.Length) return null;
        int strLen = cch;
        if (compObj[32 + cch - 1] == 0) strLen--;
        if (strLen <= 0) return null;
        var result = System.Text.Encoding.GetEncoding(1252).GetString(compObj, 32, strLen);
        return result;
    }
}

/// <summary>
/// Phase 3f-4 — FSPA (Floating Shape Anchor) entry. PlcSpaMom 의 한 entry 는
/// 본문 CP (0x08 drawing char 위치) + shape ID (spid) + 앵커 사각형 (twips) 으로 구성된다.
/// 후속 단계에서 OfficeArtSpContainer 와 매칭해 floating shape 의 위치·이미지·도형 등을 복원한다.
/// </summary>
/// <param name="Cp">본문 안 0x08 char 의 CP 위치.</param>
/// <param name="Spid">OfficeArtSpContainer 의 sp.spid 와 매칭되는 shape ID.</param>
/// <param name="XaLeftTwips">앵커 좌측 (twips, 페이지 좌측 또는 단락 기준).</param>
/// <param name="YaTopTwips">앵커 상단 (twips).</param>
/// <param name="XaRightTwips">앵커 우측 (twips).</param>
/// <param name="YaBottomTwips">앵커 하단 (twips).</param>
public sealed record FspaEntry(
    int Cp,
    int Spid,
    int XaLeftTwips,
    int YaTopTwips,
    int XaRightTwips,
    int YaBottomTwips);

/// <summary>
/// Phase 3i — DOC 본문 책갈피 entry. SttbfBkmk + PlcfBkf + PlcfBkl 의 결합 결과.
/// </summary>
/// <param name="Name">책갈피 이름 (SttbfBkmk 의 string).</param>
/// <param name="StartCp">시작 CP (PlcfBkf aCP).</param>
/// <param name="EndCp">끝 CP (PlcfBkl aCP — BKF.ibkl 로 매핑).</param>
public sealed record BookmarkEntry(string Name, int StartCp, int EndCp);

/// <summary>
/// Phase 3l — DOC OLE2 컨테이너의 ObjectPool sub-storage 한 개. 각 entry 가 임베드된 OLE 객체.
/// </summary>
/// <param name="Name">ObjectPool 안의 storage 이름 (예: "_1234567890").</param>
/// <param name="ClassName">CompObj stream 에서 추출한 OLE class 이름 (예: "Equation.3", "Excel.Sheet.8"). null 이면 미상.</param>
/// <param name="PrimaryStreamName">실제 콘텐츠를 담는 것으로 추정되는 stream 이름 (EquationNative / Ole10Native / Workbook 등).</param>
/// <param name="PrimaryContent">PrimaryStreamName 의 콘텐츠 bytes. Phase 3l-3 — Ole10Native 이면 wrapper 풀린 inner native data.</param>
/// <param name="OriginalFileName">Phase 3l-3 — Ole10Native wrapper 의 원본 파일 이름 (있는 경우만).</param>
/// <param name="Streams">storage 내 모든 stream 이름 → raw bytes 사전.</param>
public sealed record OleEmbedEntry(
    string Name,
    string? ClassName,
    string? PrimaryStreamName,
    byte[]? PrimaryContent,
    string? OriginalFileName,
    IReadOnlyDictionary<string, byte[]> Streams);

/// <summary>
/// Phase 3n — VBA 매크로 프로젝트의 격리된 raw bytes. 절대 실행하지 않으며 fidelity 보존용으로만 유지.
/// CLAUDE.md §"활성 콘텐츠는 격리 저장" 원칙: 매크로/스크립트 (VBA 등) 는 격리 저장, 실행 정책은 별도.
/// </summary>
/// <param name="StorageName">root 안의 매크로 storage 이름 (보통 "Macros", 드물게 "_VBA_PROJECT_CUR").</param>
/// <param name="Streams">매크로 storage 내 모든 stream path → raw bytes 사전. sub-storage 는 "Sub/Stream" path 로 표현.</param>
public sealed record MacroProjectInfo(
    string StorageName,
    IReadOnlyDictionary<string, byte[]> Streams);
