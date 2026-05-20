# PolyDonky.Convert.Hwpx — CLI 사용 설명서

HWPX(Open Word Processor Markup Language, KS X 6101) 문서를 IWPF 로 가져오거나,
IWPF 를 HWPX 로 내보내는 명령줄 변환기입니다.
CLAUDE.md §3 의 **외부 변환 모듈 분리 원칙**에 따라 메인 앱에서 분리된 독립 실행 파일입니다.

---

## 목차
- [개요](#개요)
- [사용법 (Synopsis)](#사용법-synopsis)
- [변환 쌍](#변환-쌍)
- [옵션](#옵션)
- [종료 코드](#종료-코드)
- [사전 검증 (import)](#사전-검증-import)
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
| 실행 파일 | `PolyDonky.Convert.Hwpx.exe` (Windows) |
| 대상 포맷 | HWPX 1.2 이상 (HWP 2014 이상, KS X 6101 기반 ZIP 패키지) |
| 방향 | `.hwpx → .iwpf` (import), `.iwpf → .hwpx` (export) |

---

## 사용법 (Synopsis)

```
PolyDonky.Convert.Hwpx <input> <output> [--debug]
PolyDonky.Convert.Hwpx --version | -v
PolyDonky.Convert.Hwpx --help    | -h | /?
```

---

## 변환 쌍

| 입력 | 출력 | 설명 |
|------|------|------|
| `*.hwpx` | `*.iwpf` | **import** — HWPX 문서를 IWPF 로 변환 |
| `*.iwpf` | `*.hwpx` | **export** — IWPF 를 HWPX 로 변환 |

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
| `4` | `IoError` | I/O 실패 (파일 없음·권한·디스크) |
| `5` | `ConvertError` | 변환 실패 (HWPX 구조 손상·암호화·내부 예외) |
| `6` | `OldVersion` | 지원하지 않는 옛 버전 (HWPX 1.2 미만) |

---

## 사전 검증 (import)

`.hwpx → .iwpf` 변환 시작 전에 세 가지 사전 검증을 수행합니다.
검증 실패 시 변환을 시작하지 않고 exit 5 로 종료합니다.

| 검증 항목 | 설명 |
|----------|------|
| **mimetype** | ZIP 루트의 `mimetype` 엔트리가 `application/hwp+zip` 인지 확인 |
| **암호화 여부** | `META-INF/manifest.xml` 에 `encryption-data` 가 있으면 거부 |
| **핵심 콘텐츠** | `Contents/header.xml` 과 `Contents/section*.xml` 중 하나 이상 존재 확인 |

암호화된 HWPX 를 변환하려면 한컴오피스에서 암호를 해제하고 다시 저장한 뒤 시도하세요.

---

## 버전 정책

PolyDonky 는 **HWPX 1.2 (HWP 2014) 이상**만 처리합니다.
버전은 `Contents/header.xml` 의 `xmlVersion` 속성으로 판단합니다.

| xmlVersion | HWP 버전 | 지원 |
|------------|----------|------|
| 1.2 이상 | HWP 2014 이상 | ✅ |
| 1.2 미만 | HWP 2010 이하 | ❌ (exit 6) |

---

## 진행 표시

```
PROGRESS:0:HWPX 읽는 중 (xmlVersion 1.3, 한글 2022 11.0)
PROGRESS:60:IWPF 로 변환 중
PROGRESS:100:완료
```

---

## 출력 안전성

임시 파일에 먼저 쓴 뒤 원자적 이름 변경(rename)으로 최종 경로에 배치합니다.
Ctrl+C 등 SIGINT 시에도 임시 파일이 자동으로 정리됩니다.

---

## 사용 예

### HWPX → IWPF

```bash
PolyDonky.Convert.Hwpx 보고서.hwpx 보고서.iwpf
```

### IWPF → HWPX

```bash
PolyDonky.Convert.Hwpx 문서.iwpf 결과.hwpx
```

### 상세 진단 (예외 스택 트레이스)

```bash
PolyDonky.Convert.Hwpx 문서.hwpx 문서.iwpf --debug
```

### 버전 확인

```bash
PolyDonky.Convert.Hwpx --version
# PolyDonky.Convert.Hwpx 1.0
```

---

## 제한 사항

| 항목 | 상태 |
|------|------|
| 암호화된 HWPX | ❌ 지원 안 함 (exit 5) |
| HWPX 1.2 미만 | ❌ 지원 안 함 (exit 6) |
| 매크로 / 스크립트 | 격리 보존(opaque) — 실행 안 됨 |
| 변경추적·주석 | 현재 구현 중 |
| 수식 (MathML / TeX) | 현재 구현 중 |

---

## 관련 문서

- [PolyDonky.Convert.Hwp CLI 설명서](cli-hwp.md)
- [IWPF 포맷 사양](../IWPF.md)
- [HISTORY.md](../HISTORY.md)
