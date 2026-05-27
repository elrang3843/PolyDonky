using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using OpenMcdf;
using PolyDonky.Core;

namespace PolyDonky.Convert.Doc;

/// <summary>
/// IWPF → Word 97-2003 (.doc) OLE2 바이너리 writer.
///
/// Phase F1-W2a (현재 단계): CFB 컨테이너 골격 + FIB + 최소 CLX (piece table) + 본문 텍스트만.
/// 글자/단락 서식, 표, 이미지, 헤더/푸터, 각주, 변경추적, 책갈피 등은 후속 단계 (F1-W2b ~ F1-W2e)
/// 에서 추가된다. 출력된 파일은 <see cref="DocBinaryReader"/> 로 round-trip 가능하며,
/// Word 가 열면 서식 없는 plain text 로 표시된다.
///
/// 참고 사양:
///   [MS-CFB]  Compound File Binary File Format (OpenMcdf 가 처리)
///   [MS-DOC]  Word (.doc) Binary File Format §2.5 FIB, §2.9.177 Pcd, §2.5.5 FibRgFcLcb97
/// </summary>
public sealed class DocBinaryWriter
{
    // ── FIB 레이아웃 상수 (모든 offset 은 WordDocument stream 기준) ──────────────

    /// <summary>FIB 시그니처 — [MS-DOC] §2.5.2 wIdent.</summary>
    private const ushort Magic = 0xA5EC;

    /// <summary>Word 97-2003 표준 nFib 값 (Word 2000+ 저장).</summary>
    private const ushort NFib = 0x00C1;   // 193

    /// <summary>nFibBack — Word 95 fallback 식별자 (호환성).</summary>
    private const ushort NFibBack = 0x00BF;   // 191

    /// <summary>FibBase flags @ 0x000A — bit 2 fComplex + bit 9 fWhichTblStm(=1Table).</summary>
    private const ushort FibFlags = 0x0204;

    /// <summary>FibRgW97 길이 (csw): [MS-DOC] §2.5.1 — nFib 0xC1 에서 14 개 short.</summary>
    private const ushort Csw  = 0x000E;   // 14

    /// <summary>FibRgLw97 길이 (cslw): nFib 0xC1 에서 22 개 long.</summary>
    private const ushort Cslw = 0x0016;   // 22

    /// <summary>FibRgFcLcbBlob 의 (fc, lcb) pair 개수 (cbRgFcLcb): nFib 0xC1 에서 93.</summary>
    private const ushort CbRgFcLcb = 0x005D;   // 93

    /// <summary>FibRgCswNew 길이: nFib 0xC1 에서 2 개 short.</summary>
    private const ushort CswNew = 0x0002;

    /// <summary>FIB 영역을 0x400 byte 로 패딩 — Reader 의 모든 offset 가드(최대 0x0316)를 충족.</summary>
    private const int    FibPadSize = 0x400;

    // FibRgFcLcbBlob 안의 pair index (각 pair = 8 bytes, [fc:4][lcb:4]).
    // 우리가 쓰는 pair 만 정의 — 나머지는 0 으로 두면 Reader 가 자동 무시.
    private const int PairFcClx = 33;   // [MS-DOC] §2.5.5 FibRgFcLcb97.fcClx @ pair 33 → FIB 0x01A2

    // FibRgFcLcbBlob 시작 offset
    private const int RgFcLcbBase = 0x009A;

    // ── 공개 진입점 ────────────────────────────────────────────────────────────

    public void Write(PolyDonkyument doc, Stream output)
    {
        // 1. 본문 텍스트 추출 — Phase F1-W2a 는 plain text 만, 서식 무시.
        //    Word 의 본문은 항상 단락 마크 '\r' 로 끝나야 하므로, 마지막에 강제 추가.
        string text = BuildPlainText(doc);
        byte[] textBytes = Encoding.Unicode.GetBytes(text);
        uint   ccpText   = (uint)text.Length;   // CP 단위 — UTF-16 code unit 1개 = 1 CP

        // 2. WordDocument stream 조립: [FIB 0x400 byte] [Unicode text]
        var wd = new byte[FibPadSize + textBytes.Length];
        uint fcMin = FibPadSize;
        uint fcMac = fcMin + (uint)textBytes.Length;
        Buffer.BlockCopy(textBytes, 0, wd, (int)fcMin, textBytes.Length);

        // 3. Table stream (1Table) 조립: [CLX with single Pcdt]
        //    CLX = 1 byte clxtPcdt(0x02) + 4 byte lcb + plcfpcd(cps + Pcds).
        //    cps = [0, ccpText],  Pcd[0].fc = fcMin (Unicode).
        byte[] table = BuildTableStreamWithSinglePiece(fcMin, ccpText);

        // 4. FIB 채우기
        WriteFib(wd, fcMin, fcMac, ccpText, fcClx: 0, lcbClx: (uint)table.Length);

        // 5. OLE2 CFB 컨테이너에 두 stream 기록.
        //    OpenMcdf v3 의 RootStorage.Create 는 stream 을 가져가는데, LeaveOpen flag 로 호출 측의
        //    output stream 을 보존. 비-Transacted 모드에선 Write 가 즉시 반영되므로 Commit() 불필요.
        using var root = RootStorage.Create(output, OpenMcdf.Version.V3, StorageModeFlags.LeaveOpen);
        WriteStream(root, "WordDocument", wd);
        WriteStream(root, "1Table",       table);
    }

    // ── helpers ─────────────────────────────────────────────────────────────────

    /// <summary>모든 Section/Paragraph 의 평문을 추출하고 '\r' 로 join. 마지막에 단락 마크 보장.
    /// 이미지/표/도형 등은 Phase F1-W2b 이후 — 현재는 무시.</summary>
    private static string BuildPlainText(PolyDonkyument doc)
    {
        var sb = new StringBuilder();
        foreach (var sec in doc.Sections)
        {
            foreach (var blk in sec.Blocks)
            {
                if (blk is Paragraph p)
                {
                    foreach (var r in p.Runs)
                        if (!string.IsNullOrEmpty(r.Text))
                            sb.Append(SanitizeForWord(r.Text));
                    sb.Append('\r');   // 단락 마크
                }
                // 다른 블록은 향후 단계에서 처리
            }
        }
        // 본문은 반드시 '\r' 로 끝나야 함 (Word 호환). 단락이 하나도 없었으면 강제 추가.
        if (sb.Length == 0 || sb[^1] != '\r')
            sb.Append('\r');
        return sb.ToString();
    }

    /// <summary>Word 본문에서 의미가 있는 control char (0x07 셀 끝, 0x0B 강제 줄바꿈 등) 만 허용하고
    /// 그 외 ASCII 제어 문자(예: '\n', '\t' 도 \r 외)는 공백으로 치환. UTF-16 surrogate pair 는 그대로 유지.</summary>
    private static string SanitizeForWord(string s)
    {
        var sb = new StringBuilder(s.Length);
        foreach (char c in s)
        {
            // \r 은 단락 마크 — 본문 Run 안에서는 별도 단락으로 분할되어야 하므로 허용 안 함.
            if (c == '\r' || c == '\n') sb.Append(' ');
            else if (c == '\t') sb.Append('\t');   // tab 은 그대로
            else if (c < 0x20)  sb.Append(' ');    // 기타 제어 문자 → 공백
            else                sb.Append(c);
        }
        return sb.ToString();
    }

    /// <summary>1Table 의 CLX (Pcdt) 직렬화.
    /// 단일 piece, Unicode (fCompressed=0), 전체 본문을 WordDocument 의 fcText 부터 ccpText CP 만큼 매핑.</summary>
    private static byte[] BuildTableStreamWithSinglePiece(uint fcText, uint ccpText)
    {
        // plcfpcd 구조: cps[n+1] (4 byte each) + Pcds[n] (8 byte each).
        // n=1 → cps = [0, ccpText], pcds = [{flags=0, fc=fcText(unicode), prm=0}].
        const int n = 1;
        int cpsSize = (n + 1) * 4;
        int pcdSize = n * 8;
        int plcSize = cpsSize + pcdSize;

        // CLX = clxtPcdt(1) + lcb(4) + plcfpcd
        var ms = new MemoryStream(1 + 4 + plcSize);
        var bw = new BinaryWriter(ms);

        bw.Write((byte)0x02);            // clxtPcdt marker
        bw.Write((uint)plcSize);         // lcb of plcfpcd

        // cps array
        bw.Write((uint)0);               // cp[0] = 0
        bw.Write(ccpText);               // cp[1] = ccpText

        // single Pcd (8 bytes)
        bw.Write((ushort)0);             // A/B/C/D flags = 0 (fNoParaLast=0 means trailing \r is a real para mark)
        bw.Write(fcText);                // fc — Unicode mode (bit 30 = 0)
        bw.Write((ushort)0);             // prm = 0 (no run-level property modifier)

        return ms.ToArray();
    }

    /// <summary>FIB 의 모든 필수 필드를 wd 의 처음 0x400 byte 영역에 채운다.
    /// 미사용 영역은 0 으로 남겨두면 Reader 가 0/0 (= 없음) 으로 해석한다.</summary>
    private static void WriteFib(byte[] wd, uint fcMin, uint fcMac, uint ccpText, uint fcClx, uint lcbClx)
    {
        // [MS-DOC] §2.5.2 FibBase
        WriteUInt16(wd, 0x0000, Magic);
        WriteUInt16(wd, 0x0002, NFib);
        // 0x0004 unused
        // 0x0006 lid (language id) — 0 OK
        // 0x0008 pnNext = 0
        WriteUInt16(wd, 0x000A, FibFlags);
        WriteUInt16(wd, 0x000C, NFibBack);
        // 0x000E lKey = 0 (no obfuscation key)
        // 0x0012 envr, 0x0013 more flags, 0x0014..0x0017 reserved — 0 OK
        WriteUInt32(wd, 0x0018, fcMin);
        WriteUInt32(wd, 0x001C, fcMac);

        // FibRgW97
        WriteUInt16(wd, 0x0020, Csw);
        // 14 short 값 (0x0022 ~ 0x003D) — 모두 0 으로 두면 default.

        // FibRgLw97
        WriteUInt16(wd, 0x003E, Cslw);
        // FibRgLw97 anchor = 0x0040, 22 × 4 byte 필드. 주요 인덱스:
        //   0x004C ccpText, 0x0050 ccpFtn, 0x0054 ccpHdd, 0x005C ccpAtn, 0x0060 ccpEdn
        WriteUInt32(wd, 0x004C, ccpText);
        // 나머지 ccp* 는 0 (각주/미주/주석/헤더 없음)

        // FibRgFcLcbBlob
        WriteUInt16(wd, 0x0098, CbRgFcLcb);
        WritePair(wd, PairFcClx, fcClx, lcbClx);   // fcClx @ 0x01A2, lcbClx @ 0x01A6

        // cswNew + FibRgCswNew (nFib 0xC1 에서 2 short)
        int cswNewOffset = RgFcLcbBase + CbRgFcLcb * 8;   // = 0x0382
        WriteUInt16(wd, cswNewOffset,     CswNew);
        WriteUInt16(wd, cswNewOffset + 2, NFib);          // FibRgCswNew[0] = nFibNew
        // FibRgCswNew[1] = 0 (cQuickSavesNew)
    }

    private static void WritePair(byte[] wd, int pairIndex, uint fc, uint lcb)
    {
        int fcOff  = RgFcLcbBase + pairIndex * 8;
        int lcbOff = fcOff + 4;
        WriteUInt32(wd, fcOff,  fc);
        WriteUInt32(wd, lcbOff, lcb);
    }

    private static void WriteUInt16(byte[] buf, int offset, ushort value)
    {
        buf[offset    ] = (byte)(value & 0xFF);
        buf[offset + 1] = (byte)(value >> 8);
    }

    private static void WriteUInt32(byte[] buf, int offset, uint value)
    {
        buf[offset    ] = (byte) (value        & 0xFF);
        buf[offset + 1] = (byte)((value >>  8) & 0xFF);
        buf[offset + 2] = (byte)((value >> 16) & 0xFF);
        buf[offset + 3] = (byte)((value >> 24) & 0xFF);
    }

    private static void WriteStream(RootStorage root, string name, byte[] data)
    {
        // CfbStream 은 사용 후 dispose 해야 데이터가 directory entry 에 flush 된다 (OpenMcdf v3).
        using var s = root.CreateStream(name);
        s.Write(data, 0, data.Length);
    }
}
