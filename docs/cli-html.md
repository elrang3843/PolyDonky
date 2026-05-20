# PolyDonky.Convert.Html — CLI 사용 설명서

HTML / HTM 문서를 IWPF 로 가져오거나, IWPF 를 HTML5 문서로 내보내는
명령줄 변환기입니다. CSS 캐스케이드 계산, 인코딩 자동 감지, 이미지 임베드를 지원합니다.
CLAUDE.md §3 의 **외부 변환 모듈 분리 원칙**에 따라 메인 앱에서 분리된 독립 실행 파일입니다.

---

## 목차
- [개요](#개요)
- [사용법 (Synopsis)](#사용법-synopsis)
- [변환 쌍](#변환-쌍)
- [옵션](#옵션)
- [종료 코드](#종료-코드)
- [인코딩 감지 (import)](#인코딩-감지-import)
- [CSS 지원](#css-지원)
- [이미지 처리](#이미지-처리)
- [바이너리 거부](#바이너리-거부)
- [진행 표시](#진행-표시)
- [출력 안전성](#출력-안전성)
- [사용 예](#사용-예)
- [제한 사항](#제한-사항)
- [관련 문서](#관련-문서)

---

## 개요

| 항목 | 내용 |
|------|------|
| 실행 파일 | `PolyDonky.Convert.Html.exe` (Windows) |
| 의존성 | `AngleSharp` + `AngleSharp.Css` (MIT / beta) |
| 대상 포맷 | HTML5 / HTML4 / XHTML (`.html`, `.htm`) |
| 방향 | `.html|.htm → .iwpf` (import), `.iwpf → .html|.htm` (export) |

---

## 사용법 (Synopsis)

```
PolyDonky.Convert.Html <input> <output> [옵션...]
PolyDonky.Convert.Html --version | -v
PolyDonky.Convert.Html --help    | -h | /?
```

---

## 변환 쌍

| 입력 | 출력 | 설명 |
|------|------|------|
| `*.html`, `*.htm` | `*.iwpf` | **import** — HTML 을 IWPF 로 변환 |
| `*.iwpf` | `*.html`, `*.htm` | **export** — IWPF 를 HTML5 문서로 변환 |

---

## 옵션

| 옵션 | 적용 방향 | 설명 |
|------|----------|------|
| `--fragment` | export | `<!DOCTYPE>`·`<html>`·`<head>`·`<body>` 래퍼 없이 내용만 출력 |
| `--title <text>` | export | `<title>` 요소 텍스트 지정. 생략 시 첫 H1 텍스트 또는 기본값 |
| `--title=<text>` | export | `--title <text>` 의 등호 형식 |
| `--debug` \| `-d` \| `DEBUG` | 공통 | 예외 스택 트레이스 등 상세 진단 정보를 stderr 에 출력 |
| `--version` \| `-v` | 공통 | 버전 정보 출력 후 종료 |
| `--help` \| `-h` \| `/?` | 공통 | 이 도움말 출력 후 종료 |

`--fragment` 와 `--title` 은 export(`*.iwpf → *.html`) 시에만 유효합니다.
import 시 지정하면 exit 2 로 종료됩니다.

---

## 종료 코드

| 코드 | 상수 | 의미 |
|------|------|------|
| `0` | `Ok` | 성공 |
| `2` | `BadArgs` | 인자 오류 (잘못된 옵션·위치 인자 수 오류) |
| `3` | `UnsupportedOp` | 지원하지 않는 변환 쌍 |
| `4` | `IoError` | I/O 실패 |
| `5` | `ConvertError` | 변환 실패 (HTML 파싱 오류·바이너리 입력·내부 예외) |

---

## 인코딩 감지 (import)

HTML 파일의 인코딩은 다음 순서로 자동 감지됩니다.

1. **BOM** — UTF-8 (EF BB BF), UTF-16 LE/BE, UTF-32 LE/BE
2. **유효한 UTF-8 판정** — 파일 전체가 유효한 UTF-8 시퀀스이면 `<meta charset>` 선언보다 UTF-8 우선
   (EUC-KR 을 선언했어도 실제로 UTF-8 로 저장된 파일이 많기 때문)
3. **`<meta charset="X">`** 또는 **`http-equiv Content-Type charset`** — 첫 4KB 안에서 탐색
4. 레거시 코드페이지 지원: `cp949`, `EUC-KR`, `Shift-JIS` 등 (CodePagesEncodingProvider)
5. 위 모두 실패 → **UTF-8** (HTML5 기본값)

감지된 인코딩은 `PROGRESS:` 메시지에 표시됩니다:
```
PROGRESS:0:HTML 읽는 중 (인코딩: EUC-KR)
```

---

## CSS 지원

import 시 CSS 속성을 파싱해 IWPF 공통 모델로 변환합니다.

| CSS 기능 | 지원 |
|---------|------|
| 인라인 `style=""` | ✅ |
| `<style>` 블록 | ✅ |
| 외부 스타일시트 `<link rel="stylesheet">` | ✅ — 로컬 파일 경로 인라인 |
| 선택자 특이도(Specificity) 계산 | ✅ |
| 캐스케이드 계산 | ✅ |
| `::before` / `::after` 가상 요소 | ✅ |
| `::first-letter` 가상 요소 | ✅ |
| `counter()` / `counters()` | ✅ |
| `color` / `background-color` | ✅ |
| `font-size` / `font-weight` / `font-style` | ✅ |
| `text-decoration` (underline·line-through) | ✅ |
| `text-align` | ✅ |
| `margin` / `padding` | ✅ |
| 미디어 쿼리 | ⚠️ 무시됨 |
| CSS 애니메이션 / 전환 | ⚠️ 무시됨 |

---

## 이미지 처리

- import 시 `<img src="...">` 의 로컬 파일 경로를 읽어 `data:` URI 또는 바이너리로 임베드합니다.
- 원격 URL(`http://`, `https://`) 의 이미지는 임베드하지 않고 `ResourcePath` 만 보존합니다.
- export 시 IWPF 내 이미지 데이터를 `<img src="data:image/...;base64,...">` 로 출력합니다.

---

## 바이너리 거부

파일 앞 1KB 안에 **NUL 바이트가 5% 이상**이면 HTML 이 아닌 바이너리 파일로 간주해
변환을 거부합니다 (exit 5).

---

## 진행 표시

```
PROGRESS:0:HTML 읽는 중 (인코딩: UTF-8)
PROGRESS:60:IWPF 로 변환 중
PROGRESS:100:완료
```

---

## 출력 안전성

임시 파일에 먼저 쓴 뒤 원자적 이름 변경으로 최종 경로에 배치합니다.
Ctrl+C 시에도 임시 파일이 자동으로 정리됩니다.

---

## 사용 예

### HTML → IWPF

```bash
PolyDonky.Convert.Html 문서.html 문서.iwpf
```

### EUC-KR HTML → IWPF

```bash
# 인코딩 자동 감지 — 별도 옵션 불필요
PolyDonky.Convert.Html 한국어문서.htm 한국어문서.iwpf
```

### IWPF → HTML (완전 문서)

```bash
PolyDonky.Convert.Html 문서.iwpf 결과.html
```

### IWPF → HTML fragment (래퍼 없음)

```bash
PolyDonky.Convert.Html 문서.iwpf 결과.html --fragment
```

### IWPF → HTML (제목 지정)

```bash
PolyDonky.Convert.Html 문서.iwpf 결과.html --title "2026 연간 보고서"
# 또는
PolyDonky.Convert.Html 문서.iwpf 결과.html --title="2026 연간 보고서"
```

### 상세 진단

```bash
PolyDonky.Convert.Html 문서.html 문서.iwpf --debug
```

---

## 제한 사항

| 항목 | 상태 |
|------|------|
| 미디어 쿼리 | ⚠️ 무시됨 |
| CSS 애니메이션 / 전환 | ⚠️ 무시됨 |
| `<canvas>` / `<video>` / `<audio>` | ⚠️ 무시됨 |
| JavaScript | ⚠️ 실행 안 됨 (보안 정책) |
| 원격 이미지 | ⚠️ 임베드 안 됨 — URL 만 보존 |

---

## 관련 문서

- [PolyDonky.Convert.Xml CLI 설명서](cli-xml.md)
- [IWPF 포맷 사양](../IWPF.md)
- [HISTORY.md](../HISTORY.md)
