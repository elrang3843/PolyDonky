# PolyDonky.Convert.Xml — CLI 사용 설명서

XML / XHTML 문서를 IWPF 로 가져오거나, IWPF 를 XHTML5 polyglot markup 으로
내보내는 명령줄 변환기입니다.
CLAUDE.md §3 의 **외부 변환 모듈 분리 원칙**에 따라 메인 앱에서 분리된 독립 실행 파일입니다.

---

## 목차
- [개요](#개요)
- [사용법 (Synopsis)](#사용법-synopsis)
- [변환 쌍](#변환-쌍)
- [옵션](#옵션)
- [종료 코드](#종료-코드)
- [인코딩 감지 (import)](#인코딩-감지-import)
- [XHTML 감지](#xhtml-감지)
- [보안 정책 (XXE 차단)](#보안-정책-xxe-차단)
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
| 실행 파일 | `PolyDonky.Convert.Xml.exe` (Windows) |
| 의존성 | `AngleSharp` + `AngleSharp.Css` (Convert.Html 위에 polyglot serializer 추가) |
| 대상 포맷 | XML 1.0/1.1, XHTML1 / XHTML5 (`.xml`, `.xhtml`) |
| 방향 | `.xml|.xhtml → .iwpf` (import), `.iwpf → .xml|.xhtml` (export) |

---

## 사용법 (Synopsis)

```
PolyDonky.Convert.Xml <input> <output> [옵션...]
PolyDonky.Convert.Xml --version | -v
PolyDonky.Convert.Xml --help    | -h | /?
```

---

## 변환 쌍

| 입력 | 출력 | 설명 |
|------|------|------|
| `*.xml`, `*.xhtml` | `*.iwpf` | **import** — XML/XHTML 을 IWPF 로 변환 |
| `*.iwpf` | `*.xml`, `*.xhtml` | **export** — IWPF 를 XHTML5 polyglot markup 으로 변환 |

---

## 옵션

| 옵션 | 적용 방향 | 설명 |
|------|----------|------|
| `--fragment` | export | `<?xml ?>`·`<!DOCTYPE>`·`<html>` 래퍼 없이 내용만 출력 |
| `--title <text>` | export | `<title>` 요소 텍스트 지정. 생략 시 첫 H1 텍스트 또는 기본값 |
| `--title=<text>` | export | `--title <text>` 의 등호 형식 |
| `--debug` \| `-d` \| `DEBUG` | 공통 | 예외 스택 트레이스 등 상세 진단 정보를 stderr 에 출력 |
| `--version` \| `-v` | 공통 | 버전 정보 출력 후 종료 |
| `--help` \| `-h` \| `/?` | 공통 | 이 도움말 출력 후 종료 |

`--fragment` 와 `--title` 은 export(`*.iwpf → *.xml`) 시에만 유효합니다.
import 시 지정하면 exit 2 로 종료됩니다.

---

## 종료 코드

| 코드 | 상수 | 의미 |
|------|------|------|
| `0` | `Ok` | 성공 |
| `2` | `BadArgs` | 인자 오류 |
| `3` | `UnsupportedOp` | 지원하지 않는 변환 쌍 |
| `4` | `IoError` | I/O 실패 |
| `5` | `ConvertError` | 변환 실패 (XML 형식 오류·DTD 거부·바이너리 입력) |

---

## 인코딩 감지 (import)

XML 파일의 인코딩은 다음 순서로 자동 감지됩니다.

1. **BOM** — UTF-8 (EF BB BF), UTF-16 LE/BE, UTF-32 LE/BE
2. **XML 선언** — 첫 256바이트 안의 `<?xml version="1.0" encoding="X"?>` 에서 `encoding` 속성
3. 레거시 코드페이지 지원: `cp949`, `EUC-KR`, `Shift-JIS` 등 (CodePagesEncodingProvider)
4. 위 모두 실패 → **UTF-8** (XML 1.0 기본값)

---

## XHTML 감지

import 시 파일이 XHTML 인지 다음 기준으로 자동 감지합니다.

- `<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML ...">` 선언
- 루트 요소에 `xmlns="http://www.w3.org/1999/xhtml"` 네임스페이스

XHTML 로 감지된 경우 HTML 처리 파이프라인으로 위임합니다.
일반 XML (비 XHTML) 은 텍스트 노드를 추출해 단락으로 변환합니다.

---

## 보안 정책 (XXE 차단)

**DTD 처리가 완전히 비활성화**되어 있습니다.
`DOCTYPE` 선언이 포함된 XML 파일은 변환을 거부합니다 (exit 5).

이는 **XXE (XML External Entity) 공격 및 외부 엔티티 참조**를 차단하기 위한 설계입니다.

```xml
<!-- 이 파일은 exit 5 로 거부됩니다 -->
<?xml version="1.0"?>
<!DOCTYPE foo SYSTEM "file:///etc/passwd">
<root>&xxe;</root>
```

DTD 를 포함하지 않도록 수정하거나, XHTML 입력이라면 DOCTYPE 선언을 제거하세요.

---

## 바이너리 거부

파일 앞 1KB 안에 **NUL 바이트가 5% 이상**이면 XML 이 아닌 바이너리 파일로 간주해
변환을 거부합니다 (exit 5).

---

## 진행 표시

```
PROGRESS:0:XML 읽는 중 (인코딩: UTF-8)
PROGRESS:60:IWPF 로 변환 중
PROGRESS:100:완료
```

---

## 출력 안전성

임시 파일에 먼저 쓴 뒤 원자적 이름 변경으로 최종 경로에 배치합니다.
Ctrl+C 시에도 임시 파일이 자동으로 정리됩니다.

---

## 사용 예

### XML (XHTML) → IWPF

```bash
PolyDonky.Convert.Xml 문서.xhtml 문서.iwpf
```

### 일반 XML → IWPF (텍스트 추출)

```bash
PolyDonky.Convert.Xml data.xml data.iwpf
```

### IWPF → XHTML5 (완전 문서)

```bash
PolyDonky.Convert.Xml 문서.iwpf 결과.xhtml
```

### IWPF → XHTML5 fragment

```bash
PolyDonky.Convert.Xml 문서.iwpf 결과.xml --fragment
```

### IWPF → XHTML5 (제목 지정)

```bash
PolyDonky.Convert.Xml 문서.iwpf 결과.xhtml --title "기술 명세서"
```

### 상세 진단

```bash
PolyDonky.Convert.Xml 문서.xml 문서.iwpf --debug
```

---

## 제한 사항

| 항목 | 상태 |
|------|------|
| DTD 포함 XML | ❌ 보안 정책상 거부 (exit 5) |
| XML Schema (XSD) 검증 | ⚠️ 수행 안 됨 — 형식 오류만 거부 |
| XSLT 변환 | ⚠️ 지원 안 됨 |
| XML 네임스페이스 (비 XHTML) | ⚠️ 텍스트 추출만 수행 |

---

## 관련 문서

- [PolyDonky.Convert.Html CLI 설명서](cli-html.md)
- [IWPF 포맷 사양](../IWPF.md)
- [HISTORY.md](../HISTORY.md)
