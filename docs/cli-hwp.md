# PolyDonky.Convert.Hwp — CLI 사용 설명서

HWP 5.x 바이너리 문서를 IWPF 로 가져오거나, IWPF 를 HWP 호환 형식으로 내보내는
명령줄 변환기입니다. CLAUDE.md §3 의 **외부 변환 모듈 분리 원칙**에 따라 메인 앱에서
분리된 독립 실행 파일입니다.

---

## 목차
- [개요](#개요)
- [사용법 (Synopsis)](#사용법-synopsis)
- [변환 쌍](#변환-쌍)
- [옵션](#옵션)
- [종료 코드](#종료-코드)
- [파일 형식 감지](#파일-형식-감지)
- [진단 로그](#진단-로그)
- [진행 표시](#진행-표시)
- [출력 안전성](#출력-안전성)
- [사용 예](#사용-예)
- [제한 사항](#제한-사항)
- [관련 문서](#관련-문서)

---

## 개요

| 항목 | 내용 |
|------|------|
| 실행 파일 | `PolyDonky.Convert.Hwp.exe` (Windows) |
| 대상 포맷 | HWP 5.x (KS X 5700 · OLE2 CFB), ZIP 기반 HWPX 을 `.hwp` 확장자로 배포한 파일 |
| 방향 | `.hwp → .iwpf` (import), `.iwpf → .hwp` (export) |

---

## 사용법 (Synopsis)

```
PolyDonky.Convert.Hwp <input> <output> [--debug]
PolyDonky.Convert.Hwp --version | -v
PolyDonky.Convert.Hwp --help    | -h | /?
```

`<input>` 과 `<output>` 은 상대 경로·절대 경로 모두 허용됩니다.

---

## 변환 쌍

| 입력 | 출력 | 설명 |
|------|------|------|
| `*.hwp` | `*.iwpf` | **import** — HWP 5.x 문서를 IWPF 로 변환 |
| `*.iwpf` | `*.hwp` | **export** — IWPF 를 HWPX 기반 출력으로 변환 (한글 호환) |

> **참고** `*.iwpf → *.hwp` 출력은 내부적으로 `HwpxWriter` 를 사용하며,
> 생성된 파일은 HWP 5.x 바이너리가 아니라 HWPX(ZIP) 형식입니다.
> 한글 2014 이상에서 `.hwp` 확장자로 열 수 있습니다.

---

## 옵션

| 옵션 | 설명 |
|------|------|
| `--debug` \| `-d` \| `DEBUG` | 상세 진단 로그 활성화 (→ [진단 로그](#진단-로그) 참고) |
| `--version` \| `-v` | 버전 정보 출력 후 종료 |
| `--help` \| `-h` \| `/?` | 이 도움말 출력 후 종료 |

옵션은 위치 인자 앞·뒤·사이 어디에나 올 수 있습니다.

```
PolyDonky.Convert.Hwp --debug input.hwp output.iwpf
PolyDonky.Convert.Hwp input.hwp --debug output.iwpf
PolyDonky.Convert.Hwp input.hwp output.iwpf DEBUG
```

---

## 종료 코드

| 코드 | 상수 | 의미 |
|------|------|------|
| `0` | `Ok` | 성공 |
| `2` | `BadArgs` | 인자 오류 (경로 오류·동일 경로·지원 안 하는 쌍) |
| `3` | `UnsupportedOp` | 지원하지 않는 변환 쌍 |
| `4` | `IoError` | I/O 실패 (파일 없음·권한·디스크) |
| `5` | `ConvertError` | 변환 실패 (파일 손상·내부 예외) |

---

## 파일 형식 감지

`.hwp` 확장자를 가진 파일이라도 실제 내부 형식이 다를 수 있습니다.
프로그램은 **매직 바이트(파일 시그니처)** 로 형식을 구별합니다.

| 시그니처 (첫 4바이트) | 형식 | 처리기 |
|----------------------|------|--------|
| `D0 CF 11 E0` (OLE2) | HWP 5.x 바이너리 | `HwpReader` |
| `50 4B 03 04` (ZIP · PK) | HWPX (ZIP 기반) | `HwpxReader` |

---

## 진단 로그

`--debug` 옵션을 전달하면 상세 파싱 로그가 파일에 기록됩니다.

| OS | 로그 파일 경로 |
|----|---------------|
| Windows | `d:\Temp\PolyDonky-HwpReader.log` |
| Linux / macOS | `/tmp/PolyDonky-HwpReader.log` |

- 세션 시작 시 헤더(`=== HwpReader session YYYY-MM-DD HH:mm:ss ===`)가 추가됩니다.
- 기존 로그에 **덧붙임(append)** 됩니다. 필요하면 수동으로 삭제하세요.
- Debug 빌드에서는 `--debug` 없이도 항상 활성화됩니다.

로그에 기록되는 내용:

- 레코드별 파싱 상태 (Body, DocInfo, BinData 등)
- 이미지·도형·OLE 개체의 위치·크기·배치 모드
- 글상자·표의 BorderFill·채우기 색 파싱 결과
- 단락 앵커링 해소 과정
- 차트 데이터 추출 상태

---

## 진행 표시

변환 중 표준 출력(stdout) 에 `PROGRESS:<percent>:<message>` 형식으로 진행 상황을
보고합니다. 메인 앱(`ExternalConverter`)이 이 줄을 파싱해 진행 대화상자에 표시합니다.

```
PROGRESS:0:HWP 읽는 중
PROGRESS:70:IWPF 저장 중
PROGRESS:100:완료
```

`PROGRESS:` 로 시작하지 않는 stdout 줄은 메인 앱에서 무시됩니다.
오류 메시지는 stderr 에 출력됩니다.

---

## 출력 안전성

- 변환 결과는 임시 파일(`<output>.tmp-XXXXXXXX`)에 먼저 쓴 뒤
  완료 후 최종 경로로 이름을 변경합니다.
- 변환 도중 비정상 종료(예외·프로세스 강제 종료)가 발생해도 반쪽짜리 출력 파일이
  남지 않습니다.
- 임시 파일 경로가 정리되지 않으면 `finally` 블록에서 삭제를 시도합니다.

---

## 사용 예

### HWP → IWPF (기본)

```bash
PolyDonky.Convert.Hwp 보고서.hwp 보고서.iwpf
```

### HWP → IWPF (진단 로그 포함)

```bash
PolyDonky.Convert.Hwp 보고서.hwp 보고서.iwpf --debug
# → d:\Temp\PolyDonky-HwpReader.log (Windows)
# → /tmp/PolyDonky-HwpReader.log  (Linux)
```

### IWPF → HWP (내보내기)

```bash
PolyDonky.Convert.Hwp 문서.iwpf 결과.hwp
```

### 버전 확인

```bash
PolyDonky.Convert.Hwp --version
# PolyDonky.Convert.Hwp 1.0
```

---

## 제한 사항

| 항목 | 상태 |
|------|------|
| 암호화된 HWP 파일 | ❌ 지원 안 함 (exit 5) |
| HWP 5.0 미만 구버전 | ⚠️ 레코드 구조 차이로 파싱 오류 가능 |
| 매크로 / 스크립트 | 격리 보존(opaque) — 실행 안 됨 |
| OLE 차트 데이터 값 | HCH 포맷 비공개로 라벨만 추출; 수치는 placeholder |
| 변경추적·주석 | 현재 구현 중 |
| 수식 (MathML / TeX) | 현재 구현 중 |

---

## 관련 문서

- [PolyDonky.Convert.Hwpx CLI 설명서](cli-hwpx.md)
- [IWPF 포맷 사양](../IWPF.md)
- [HISTORY.md](../HISTORY.md) — 변경 이력
