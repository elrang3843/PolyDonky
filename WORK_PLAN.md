# Work Plan — PolyDoc

이 문서는 **세션을 가로질러 작업을 이어받기 위한 운영 계획서**입니다.
한 세션이 끝나도 다음 세션이 이 문서를 읽고 동일한 맥락에서 이어 갈 수 있도록 갱신합니다.

- 사용자: 노진문 (Noh JinMoon)
- 회사: 핸텍 (HANDTECH)
- 메인 브랜치(작업 대상): `claude/create-claude-guide-VK1Pz`
- 정책: 사용자 명시 지시 전까지 모든 빌드는 **테스트 버전(`1.0.0-test.<n>`)**, 최초 정식 릴리스는 `1.0.0`.
- 운영 가정: 한 세션당 일일 사용량의 ~90% 까지 사용을 허용하되, 위험·비가역 결정은 사용자 게이트에서 멈춘다.

---

## 환경 사실관계

| 항목 | 상태 |
|---|---|
| 호스트 OS | Ubuntu 24.04 (root) |
| .NET SDK | **10.0.107** (apt: `dotnet-sdk-10.0`) |
| GUI/Display | 없음 — WPF UI 빌드·실행·스크린샷 불가 |
| 타겟 OS | Windows 10/11 x64 |
| Git remote | `elrang3843/PolyDoc` |
| **NuGet.org** | **503 으로 차단** — `api.nuget.org` 응답 안 함 (curl/dotnet restore 모두). BCL 만 사용하는 src 프로젝트는 복원/빌드 정상. xUnit/Markdig 같은 외부 패키지는 본 환경에서 복원 불가 |

**핵심 제약**:
- WPF 앱(`Microsoft.NET.Sdk.WindowsDesktop`)은 Windows 에서만 빌드 가능.
- 이 환경에서는 **순수 라이브러리·코덱** 만 검증 가능. UI 검증은 사용자 책임.
- xUnit 테스트는 **코드는 작성하되 본 환경에서 실행 불가**. Windows(또는 NuGet 가능 환경)에서 `dotnet restore && dotnet test` 로 검증한다.
- Markdig 같은 외부 의존이 필요한 코덱은 NuGet 회복 시점에 도입한다 (현재는 BCL 서브셋 구현으로 대체).

---

## 기술 스택 (확정)

| 항목 | 값 |
|---|---|
| .NET | **10.0** (LTS) |
| UI | **WPF** |
| MVVM | CommunityToolkit.Mvvm (MIT) |
| 테스트 | xUnit (Apache 2.0) — Assertion 은 xUnit 내장만 사용 (FluentAssertions v8 라이선스 회피) |
| Markdown | Markdig (BSD-2-Clause) |
| DOCX | DocumentFormat.OpenXml (MIT) |
| HTML | AngleSharp (MIT) |
| HWPX | 자체 구현 (KS X 6101) |
| HWP·DOC | LibreOffice headless 위탁 노선 우선, 어려우면 자체 |
| 직렬화 | System.Text.Json (Phase A), 필요 시 IWPF 일부는 XML 로 전환 |

라이선스는 모두 Apache 2.0 호스트 프로젝트와 호환.

---

## 솔루션 레이아웃

```
PolyDoc/
├── PolyDoc.sln
├── Directory.Build.props
├── Directory.Packages.props      # central package management
├── global.json                   # SDK pin (10.0.x)
├── .gitignore                    # .NET 표준 + IDE
├── src/
│   ├── PolyDoc.Core/             # 공통 문서 모델 (POCO)
│   ├── PolyDoc.Iwpf/             # IWPF 패키지 codec (ZIP+JSON)
│   ├── PolyDoc.Codecs.Text/      # TXT codec
│   ├── PolyDoc.Codecs.Markdown/  # MD codec (Markdig)
│   ├── PolyDoc.Codecs.Docx/      # (Phase C) DOCX codec
│   ├── PolyDoc.Codecs.Hwpx/      # (Phase C) HWPX codec
│   └── PolyDoc.App/              # (Phase B) WPF UI shell — Windows-only
├── tests/
│   ├── PolyDoc.Core.Tests/
│   ├── PolyDoc.Iwpf.Tests/
│   ├── PolyDoc.Codecs.Text.Tests/
│   └── PolyDoc.Codecs.Markdown.Tests/
└── samples/
    └── corpus/                   # 골든 테스트 코퍼스
```

---

## Phase 진행표

체크박스: ☐ 미진행 / ◑ 진행중 / ✅ 완료

### Phase A — Core 라이브러리 (Linux 전수 가능)
- ✅ A1 솔루션 스캐폴딩 (PolyDoc.slnx, .NET 10, CPM, 4 src + 4 tests + tools/SmokeTest)
- ✅ A2 PolyDoc.Core (공통 문서 모델)
- ✅ A3 PolyDoc.Iwpf (reader/writer, ZIP+JSON, SHA-256 검증, 위변조 거부)
- ✅ A4 PolyDoc.Codecs.Text (TXT in/out, BOM 감지)
- ✅ A5 PolyDoc.Codecs.Markdown (Markdig 없이 BCL 서브셋: 헤더·리스트·강조)
- ◑ A6 단위 테스트 — xUnit 코드 작성 완료, **NuGet 차단으로 본 환경 미실행**. Windows 에서 G2 시 `dotnet test` 로 검증.
- ✅ A6b 콘솔 스모크 러너 4/4 통과 (라운드트립 3종 + 위변조 검출)
- ✅ A7 커밋·푸시

### Phase B — WPF UI 셸 (Windows 필수)
- ☐ B1 PolyDoc.App 스캐폴딩 (WPF + MVVM)
- ☐ B2 메뉴/툴바/문서 탭/룰러 골격
- ☐ B3 i18n 리소스 (한/영)
- ☐ B4 테마 시스템 (학생~장년 대상 다중 테마)
- ☐ B5 사용자 게이트 G2: Windows에서 첫 빌드/run 결과 보고

### Phase C — DOCX/HWPX 1급 시민 (M2-M3)
- ☐ C1 DOCX reader (OpenXml)
- ☐ C2 DOCX writer + 라운드트립 테스트
- ☐ C3 HWPX reader (KS X 6101)
- ☐ C4 HWPX writer + 라운드트립 테스트
- ☐ G3: 사용자가 Word/한컴에서 결과 시각 검증

### Phase D — 외부 CLI 컨버터 분리
- ☐ D1 PolyDoc.Cli.Docx 분리
- ☐ D2 PolyDoc.Cli.Hwpx 분리
- ☐ D3 메인 앱 ↔ CLI IPC (인자/표준입출력/exit code)
- ☐ G4: LibreOffice 의존 vs 자체 결정

### Phase E — 편집 기능 / M3-M4
- ☐ 표·이미지·머리말/꼬리말·각주/미주
- ☐ 변경추적·주석
- ☐ 수식·도형/텍스트박스
- ☐ 필드/목차

### Phase F — DOC/HWP ingest / M5
- ☐ DOC import (LibreOffice 또는 자체)
- ☐ HWP import (LibreOffice 또는 자체)
- ☐ Opaque island 정책 적용

### Phase G — 고급 기능 / M6-M7
- ☐ 테마 다수, 사인 만들기 독립 앱, 사전·맞춤법 외부 모듈

### Phase H — 인스톨러 / `1.0.0` 릴리즈
- ☐ MSIX 인스톨러
- ☐ G5: 사용자 명시 지시 후 `1.0.0` 컷

---

## 사용자 디버깅·컨펌 게이트

| 게이트 | 시점 | 사용자 작업 |
|---|---|---|
| G0 | 시작 전 | 기술 스택 승인 ✅ (완료) |
| G1 | Phase A 종료 | PR 리뷰 후 머지 결정 |
| G2 | Phase B 시작 | Windows에서 `dotnet build` / `dotnet run` 후 결과 보고 |
| G3 | Phase C 종료 | HWPX·DOCX 결과를 한컴/Word에서 시각 검증 |
| G4 | Phase D 진입 시 | LibreOffice 의존 노선 확정 vs 자체 구현 |
| G5 | Phase H 진입 시 | "릴리즈하자" 명시 — `1.0.0` 컷 |

---

## 다음 세션 인수인계 체크리스트

세션 종료 시점에 이 문서의 **Phase 진행표** 와 아래 항목을 갱신한다:

- [ ] HISTORY.md `[Unreleased]` 정리
- [ ] 미해결 이슈 / 알려진 버그 목록
- [ ] 다음 세션 첫 작업 후보 (구체적 파일/함수명)
- [ ] 사용자 답을 기다리는 게이트 (있다면 어떤 게이트인지)

---

## 현재 인수인계 (Phase A 종료 시점)

### 완료
- 솔루션 골격 및 4개 src 라이브러리 빌드 그린
- 콘솔 스모크 러너 4/4 통과
- xUnit 테스트 코드 작성 완료 (실행은 Windows 에서)

### 사용자(노진문) 작업이 필요한 항목 — G1 직전
- [ ] Windows 머신에서 저장소 클론, `dotnet --info` 로 SDK 10.0.x 확인
- [ ] `dotnet restore PolyDoc.slnx` — NuGet 정상이면 xUnit 자동 복원
- [ ] `dotnet build PolyDoc.slnx` — 0 warning / 0 error 확인
- [ ] `dotnet test PolyDoc.slnx` — 모든 xUnit 테스트 그린 확인
- [ ] (옵션) `dotnet run --project tools/PolyDoc.SmokeTest` — 콘솔 스모크 4/4 통과 재확인
- [ ] PR 리뷰 후 머지 결정 (G1)

### Phase B 진입 전 정해야 할 것
- 정식 회사 로고 파일이 별도로 있다면 `assets/logo.png` 로 추가하고 README 의 GitHub 아바타 참조를 교체할지 (현재는 `https://github.com/elrang3843.png` 사용)
- WPF 앱의 .NET TFM: `net10.0-windows`(권장) vs `net10.0-windows10.0.19041.0`(WinRT API 사용 시) — 사인 만들기·인쇄 미리보기 같은 기능에서 결정

### 다음 세션 첫 작업 후보
1. `src/PolyDoc.App/` WPF 프로젝트 스캐폴딩 (Phase B1)
2. MVVM 골격 + 메인 윈도우 + 메뉴 리소스 한·영 분리
3. IWPF 파일 드래그&드롭 → Core 모델 → 단순 본문 표시 (TextBlock 수준) 까지

### 알려진 한계
- Markdown 코덱이 Markdig 의 풀 CommonMark 가 아닌 **실용 서브셋**. 코드블록·인용·표·이미지·링크 미지원. Phase C 진입 전 Markdig 로 교체하거나 서브셋 확장.
- Block 다형성을 `JsonDerivedType` 로 처리 — 현재 `Paragraph` 만 등록. `Table`, `Image`, `TextBox` 등 추가 시 같은 위치에 등록해야 한다 (`src/PolyDoc.Core/Block.cs`).
- IWPF document.json 본문은 Phase A 에서 JSON. 후속 단계에서 IWPF 사양에 맞춰 일부를 XML 로 전환할 수 있다 (`document.xml`, `styles.xml`).
