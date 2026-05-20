# PolyDonky.Convert.Docx — CLI 사용 설명서

Microsoft Word DOCX (OOXML, ECMA-376) 문서를 IWPF 로 가져오거나,
IWPF 를 DOCX 로 내보내는 명령줄 변환기입니다.
CLAUDE.md §3 의 **외부 변환 모듈 분리 원칙**에 따라 메인 앱에서 분리된 독립 실행 파일입니다.

---

## 목차
- [개요](#개요)
- [사용법 (Synopsis)](#사용법-synopsis)
- [변환 쌍](#변환-쌍)
- [옵션](#옵션)
- [종료 코드](#종료-코드)
- [버전 정책](#버전-정책)
- [진행 표시](#진행-표시)
- [출력 안전성](#출력-안전성)
- [사용 예](#사용-예)
- [제한 사항](#제한-사항)
- [관련 문서](#관련-문서)

---

## 개요

| 항목 | 내용 |
|------|------|
| 실행 파일 | `PolyDonky.Convert.Docx.exe` (Windows) |
| 의존성 | `DocumentFormat.OpenXml` (MIT) |
| 대상 포맷 | DOCX (OOXML, ECMA-376 / ISO/IEC 29500), Word 2013 이상 |
| 방향 | `.docx → .iwpf` (import), `.iwpf → .docx` (export) |

---

## 사용법 (Synopsis)

```
PolyDonky.Convert.Docx <input> <output> [--debug]
PolyDonky.Convert.Docx --version | -v
PolyDonky.Convert.Docx --help    | -h | /?
```

---

## 변환 쌍

| 입력 | 출력 | 설명 |
|------|------|------|
| `*.docx` | `*.iwpf` | **import** — DOCX 를 IWPF 로 변환 |
| `*.iwpf` | `*.docx` | **export** — IWPF 를 DOCX 로 변환 |

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
| `4` | `IoError` | I/O 실패 (파일 없음·권한·디렉터리 없음·잠금) |
| `5` | `ConvertError` | 변환 실패 (DOCX 구조 손상·내부 예외) |
| `6` | `OldVersion` | 지원하지 않는 옛 버전 (Word 2013 미만) |

---

## 버전 정책

PolyDonky 는 **Word 2013 (AppVersion 15.0) 이상**만 처리합니다.
버전은 `docProps/app.xml` 의 `AppVersion` 요소로 판단합니다.

| AppVersion | Word 버전 | 지원 |
|------------|----------|------|
| 15.0 이상 | Word 2013 이상 | ✅ |
| 15.0 미만 | Word 2010 이하 | ❌ (exit 6) |

`AppVersion` 요소가 없는 파일(일부 서드파티 생성 DOCX)은 버전 검사를 통과합니다.

---

## 지원 기능 (import)

| 기능 | 상태 |
|------|------|
| 본문 텍스트·런 스타일 (굵기·이탤릭·밑줄·색상·크기) | ✅ |
| 단락 스타일 (들여쓰기·정렬·줄 간격) | ✅ |
| 표 (셀 병합 포함) | ✅ |
| 이미지 (인라인·플로팅) | ✅ |
| DrawingML 도형·그룹 | ✅ |
| DrawingML 차트 (`<c:chart>`) | ✅ — 막대 그래프로 렌더링 |
| 머리말·꼬리말 | ✅ |
| 각주·미주 | ✅ |
| 하이퍼링크 | ✅ |
| 번호 매기기·개요 목록 | ✅ |
| 변경추적 | 현재 구현 중 |
| 수식 (OMML) | 현재 구현 중 |

---

## 진행 표시

```
PROGRESS:0:DOCX 읽는 중 (AppVersion 16.0)
PROGRESS:60:IWPF 로 변환 중
PROGRESS:100:완료
```

---

## 출력 안전성

임시 파일에 먼저 쓴 뒤 원자적 이름 변경으로 최종 경로에 배치합니다.
Ctrl+C 시에도 임시 파일이 자동으로 정리됩니다.

---

## 사용 예

### DOCX → IWPF

```bash
PolyDonky.Convert.Docx 보고서.docx 보고서.iwpf
```

### IWPF → DOCX

```bash
PolyDonky.Convert.Docx 문서.iwpf 결과.docx
```

### 상세 진단

```bash
PolyDonky.Convert.Docx 문서.docx 문서.iwpf --debug
```

---

## 제한 사항

| 항목 | 상태 |
|------|------|
| Word 2010 이하 (AppVersion < 15.0) | ❌ exit 6 |
| 암호화된 DOCX | ❌ exit 5 |
| 매크로 포함 DOCM | ⚠️ 매크로 격리 보존 |
| 수식 (OMML) | 현재 구현 중 |
| 변경추적 | 현재 구현 중 |

---

## 관련 문서

- [IWPF 포맷 사양](../IWPF.md)
- [HISTORY.md](../HISTORY.md)
