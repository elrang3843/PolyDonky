# PolyDonky.Convert.Doc — CLI 사용 설명서

RTF (Rich Text Format) 와 Word 97-2003 OLE2 바이너리(`.doc`) 문서를 IWPF 로 가져오거나,
IWPF 를 RTF 로 내보내는 명령줄 변환기입니다.
CLAUDE.md §3 의 **외부 변환 모듈 분리 원칙**에 따라 메인 앱에서 분리된 독립 실행 파일입니다.

> **참고** `.doc` (Word 97-2003 OLE2 바이너리) 는 현재 **import 만** 지원합니다.
> IWPF → `.doc` 바이너리 출력은 v1.0.0 이후 자체 OLE2 writer (`OpenMcdf` 기반) 로 추가될 예정이며,
> 그 전까지 IWPF 의 외부 출력은 RTF 만 제공합니다.

---

## 목차
- [개요](#개요)
- [사용법 (Synopsis)](#사용법-synopsis)
- [변환 쌍](#변환-쌍)
- [옵션](#옵션)
- [종료 코드](#종료-코드)
- [RTF 형식 특징](#rtf-형식-특징)
- [진행 표시](#진행-표시)
- [사용 예](#사용-예)
- [제한 사항 및 향후 계획](#제한-사항-및-향후-계획)
- [관련 문서](#관련-문서)

---

## 개요

| 항목 | 내용 |
|------|------|
| 실행 파일 | `PolyDonky.Convert.Doc.exe` (Windows) |
| 대상 포맷 | RTF (Rich Text Format), Word 97-2003 OLE2 바이너리 (`.doc`) |
| 방향 | `.rtf → .iwpf`, `.doc → .iwpf` (import), `.iwpf → .rtf` (export) |

---

## 사용법 (Synopsis)

```
PolyDonky.Convert.Doc <input> <output> [--debug]
PolyDonky.Convert.Doc --version | -v
PolyDonky.Convert.Doc --help    | -h | /?
```

---

## 변환 쌍

| 입력 | 출력 | 설명 |
|------|------|------|
| `*.rtf` | `*.iwpf` | **import** — RTF 를 IWPF 로 변환 |
| `*.doc` | `*.iwpf` | **import** — Word 97-2003 OLE2 바이너리를 IWPF 로 변환 (VBA 매크로·디지털 서명·OLE 임베드는 fidelity capsule 로 격리 저장) |
| `*.iwpf` | `*.rtf` | **export** — IWPF 를 RTF 로 변환 |
| `*.iwpf` | `*.doc` | **export** — Word 97-2003 OLE2 바이너리. 본문 텍스트·글자/단락 서식·책갈피·FidelityCapsules (VBA/서명/미인식 storage) 보존. 이미지/표/도형은 가시 placeholder. 헤더/푸터·각주·필드는 v1.0.0+. |

---

## 옵션

| 옵션 | 설명 |
|------|------|
| `--debug` \| `-d` \| `DEBUG` | 예외 스택 트레이스 등 상세 진단 정보를 stderr 에 출력 |
| `--version` \| `-v` | 버전 정보 출력 후 종료 |
| `--help` \| `-h` \| `/?` | 이 도움말 출력 후 종료 |

---

## 종료 코드

| 코드 | 상수 | 의미 |
|------|------|------|
| `0` | `Ok` | 성공 |
| `2` | `BadArgs` | 인자 오류 |
| `3` | `UnsupportedOp` | 지원하지 않는 변환 쌍 |
| `4` | `IoError` | I/O 실패 |
| `5` | `ConvertError` | 변환 실패 |

---

## RTF 형식 특징

RTF 는 Word 97 이상 및 대부분의 워드프로세서에서 호환되는 텍스트 기반 포맷입니다.

| 기능 | Import (`.rtf → .iwpf`) | Export (`.iwpf → .rtf`) |
|------|:---:|:---:|
| 텍스트·단락 | ✅ | ✅ |
| 글자 서식 (굵기·이탤릭·밑줄·취소선·크기·색상) | ✅ | ✅ |
| 위첨자·아래첨자 | ✅ | ✅ |
| 배경색 (하이라이트) | ✅ | ✅ |
| 들여쓰기·줄 간격·문단 간격 | ✅ | ✅ |
| 표 (병합·정렬·테두리·패딩) | ✅ | ✅ |
| 이미지 (PNG/JPEG/BMP 인라인) | ✅ | ✅ |
| 도형 (`\shp` outline) | ✅ | ✅ |
| OLE 개체 (OpaqueBlock) | ✅ | ✅ |
| 메타데이터 (제목·저자·생성/수정 시각) | ✅ | ✅ |
| 하이퍼링크 (`\field HYPERLINK`) | ✅ | ✅ |
| 자동 필드 (PAGE / NUMPAGES / DATE / TIME / AUTHOR / TITLE 등 16종) | ✅ | ✅ |
| 책갈피 (`\*\bkmkstart` / `\*\bkmkend`) | ✅ | ✅ |
| 각주·미주 (`\chftn` / `\footnote` / `\ftnalt`) | ✅ | ✅ |
| 주석 (`\chatn` / `\annotation` + `\atnauthor` / `\atndate`) | ✅ | ✅ |
| 변경추적 (`\revised` / `\deleted` + `\revtbl`) | ✅ | ✅ |
| 페이지 설정 (크기·여백·가로/세로·다단·첫페이지/홀짝 다름) | ✅ | ✅ |
| 머리말·꼬리말 (Left/Center/Right 슬롯) | ✅ | ✅ |

---

## 진행 표시

RTF import:
```
PROGRESS:0:RTF 읽는 중
PROGRESS:80:IWPF 로 변환 중
PROGRESS:100:완료
```

DOC (Word 97-2003 OLE2) import:
```
PROGRESS:0:DOC (Word 97-2003 binary) 읽는 중
PROGRESS:80:IWPF 로 변환 중
PROGRESS:100:완료
```

RTF export:
```
PROGRESS:0:IWPF 읽는 중
PROGRESS:50:RTF 로 변환 중
PROGRESS:100:완료
```

---

## 사용 예

### RTF → IWPF

```bash
PolyDonky.Convert.Doc 문서.rtf 문서.iwpf
```

### IWPF → RTF

```bash
PolyDonky.Convert.Doc 문서.iwpf 결과.rtf
```

### 상세 진단

```bash
PolyDonky.Convert.Doc 문서.rtf 문서.iwpf --debug
```

---

## 제한 사항 및 향후 계획

| 항목 | 상태 |
|------|------|
| RTF import / export | ✅ 현재 지원 |
| `.doc` (Word 97-2003 OLE2 바이너리) import | ✅ 현재 지원 — VBA 매크로·디지털 서명·OLE 임베드·미인식 root storage 는 IWPF fidelity capsule 로 격리 저장 |
| `.iwpf → .doc` (OLE2 바이너리 export) | ✅ Phase F1-W2 완료 — `OpenMcdf` 기반 자체 writer. CFB + FIB + CLX + 본문 (F1-W2a), CHPX/PAPX 글자/단락 서식 (F1-W2b), SttbfBkmk/PlcfBkf/PlcfBkl 책갈피 (F1-W2c), 이미지/표/도형 가시 placeholder (F1-W2d), FidelityCapsules → OLE2 storage 복원 (F1-W2e). `.doc → .iwpf → .doc` 라운드트립 시 매크로·서명·미인식 storage 가 0% 손실로 보존 |
| 헤더/푸터, 섹션 SEPX, 각주/미주, 주석, 필드 binary 임베드 | ⏳ v1.0.0+ |
| OfficeArt 이미지/도형 binary 임베드 (PICF / BStore / FSPA) | ⏳ v1.0.0+ |
| 수식 (`\equation` 표기) | ⏳ v1.0.0+ |

---

## 관련 문서

- [IWPF 포맷 사양](../IWPF.md)
- [HISTORY.md](../HISTORY.md)
