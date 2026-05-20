# PolyDonky.Convert.Doc — CLI 사용 설명서

RTF (Rich Text Format) 문서를 IWPF 로 가져오거나,
IWPF 를 RTF 로 내보내는 명령줄 변환기입니다.
CLAUDE.md §3 의 **외부 변환 모듈 분리 원칙**에 따라 메인 앱에서 분리된 독립 실행 파일입니다.

> **참고** 이 변환기의 이름(`Convert.Doc`)은 향후 Word 97-2003 바이너리 `.doc` 파일 지원을
> 염두에 두고 붙여진 이름입니다. 현재(v1.0.0 이전)는 **RTF** 만 처리합니다.
> `.doc` OLE2 바이너리 지원은 v1.0.0 이후에 추가될 예정입니다.

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
| 대상 포맷 | RTF (Rich Text Format), 텍스트·서식·배경색 지원 |
| 방향 | `.rtf → .iwpf` (import), `.iwpf → .rtf` (export) |

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
| `*.iwpf` | `*.rtf` | **export** — IWPF 를 RTF 로 변환 |

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

| 기능 | 지원 |
|------|------|
| 텍스트·단락 | ✅ |
| 글자 서식 (굵기·이탤릭·밑줄·크기·색상) | ✅ |
| 배경색 (하이라이트) | ✅ — RTF 의 강점 |
| 표 | ✅ |
| 이미지 | ✅ |
| 머리말·꼬리말 | ✅ |

---

## 진행 표시

```
PROGRESS:0:RTF 읽는 중
PROGRESS:80:IWPF 로 변환 중
PROGRESS:100:완료
```

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
| `.doc` (Word 97-2003 OLE2 바이너리) | ⏳ v1.0.0 이후 지원 예정 |
| 수식·도형 (RTF 기반) | 현재 구현 중 |

`.doc` OLE2 바이너리 지원이 추가되면 이 문서가 갱신됩니다.

---

## 관련 문서

- [IWPF 포맷 사양](../IWPF.md)
- [HISTORY.md](../HISTORY.md)
