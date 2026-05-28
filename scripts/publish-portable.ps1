<#
.SYNOPSIS
    PolyDonky 이식형 ZIP 빌드 스크립트 (Windows x64)

.DESCRIPTION
    메인 앱과 모든 CLI 변환기를 self-contained win-x64 로 게시하고
    압축 해제 후 바로 사용 가능한 ZIP 파일을 생성한다.

    요구 사항:
      - Windows 10 이상
      - .NET 10 SDK (dotnet --version 으로 확인)
      - PowerShell 7+ (pwsh) 또는 Windows PowerShell 5.1

    생성된 ZIP 은 .NET 10 설치 없이 바로 실행 가능하다.

.PARAMETER Version
    버전 문자열 (미지정 시 Directory.Build.props 에서 자동 읽음).

.PARAMETER Rid
    대상 런타임 식별자. 기본값: win-x64.
    win-x86, win-arm64 도 지원 (빌드 머신과 대상 아키텍처가 같아야 한다).

.PARAMETER Config
    빌드 구성. 기본값: Release.

.EXAMPLE
    # 기본 (win-x64, Release, 버전 자동 감지)
    pwsh scripts\publish-portable.ps1

.EXAMPLE
    # 버전 명시
    pwsh scripts\publish-portable.ps1 -Version 1.1.0

.EXAMPLE
    # ARM64 빌드
    pwsh scripts\publish-portable.ps1 -Rid win-arm64
#>
param(
    [string] $Version = "",
    [string] $Rid     = "win-x64",
    [string] $Config  = "Release"
)

$ErrorActionPreference = "Stop"

# ─── 작업 디렉터리: 저장소 루트 ───────────────────────────────────────────
$root = Split-Path $PSScriptRoot -Parent
Set-Location $root

# ─── 버전 자동 감지 ───────────────────────────────────────────────────────
if (-not $Version) {
    [xml]$props = Get-Content "$root\Directory.Build.props"
    $Version = ($props.Project.PropertyGroup | Select-Object -ExpandProperty Version -ErrorAction SilentlyContinue | Select-Object -First 1)
    if (-not $Version) { $Version = "1.0.0" }
}

# ─── 경로 설정 ────────────────────────────────────────────────────────────
$publishRoot = "$root\publish"
$outDir      = "$publishRoot\PolyDonky-$Version-$Rid"
$zipPath     = "$publishRoot\PolyDonky-$Version-$Rid.zip"
$appCsproj   = "$root\src\PolyDonky.App\PolyDonky.App.csproj"

Write-Host ""
Write-Host "  PolyDonky v$Version  ($Rid, $Config)" -ForegroundColor Cyan
Write-Host "  출력 폴더: $outDir" -ForegroundColor DarkGray
Write-Host "  ZIP:       $zipPath" -ForegroundColor DarkGray
Write-Host ""

# ─── 이전 출력 초기화 ─────────────────────────────────────────────────────
if (Test-Path $outDir) {
    Write-Host "[정리] 이전 출력 삭제..." -ForegroundColor DarkGray
    Remove-Item $outDir -Recurse -Force
}
New-Item -ItemType Directory -Path $publishRoot -Force | Out-Null

# ─── 게시 ─────────────────────────────────────────────────────────────────
#  -p:PublishProfile=Portable 와 동일한 설정.
#  AfterTargets="Publish" 의 PublishExternalConverters 타깃이
#  CLI 변환기 6종을 $outDir 에 자동으로 함께 게시한다.
Write-Host "[빌드] 메인 앱 + CLI 변환기 게시 중 (첫 실행은 5~10분 소요)..." -ForegroundColor Yellow

$publishArgs = @(
    "publish", $appCsproj,
    "-c",  $Config,
    "-r",  $Rid,
    "--self-contained", "true",
    "-p:PublishReadyToRun=true",
    "-p:DebugType=none",
    "-p:DebugSymbols=false",
    "-o",  $outDir,
    "--nologo"
)
& dotnet @publishArgs

if ($LASTEXITCODE -ne 0) {
    Write-Host "`n[오류] 게시 실패 (종료 코드: $LASTEXITCODE)" -ForegroundColor Red
    exit $LASTEXITCODE
}

# ─── 불필요 파일 제거 ─────────────────────────────────────────────────────
Write-Host "`n[정리] PDB 및 XML 문서 파일 제거..." -ForegroundColor DarkGray
Get-ChildItem $outDir -Recurse -Include "*.pdb", "*.xml" | Remove-Item -Force

# ─── ZIP 생성 ─────────────────────────────────────────────────────────────
Write-Host "[압축] $zipPath 생성 중..." -ForegroundColor Yellow
if (Test-Path $zipPath) { Remove-Item $zipPath -Force }
Compress-Archive -Path "$outDir\*" -DestinationPath $zipPath

# ─── 결과 요약 ────────────────────────────────────────────────────────────
$sizeMB   = [math]::Round((Get-Item $zipPath).Length / 1MB, 1)
$fileCount = (Get-ChildItem $outDir -File).Count

Write-Host ""
Write-Host "  ✔ 완료!" -ForegroundColor Green
Write-Host "  ZIP : $zipPath  ($sizeMB MB)" -ForegroundColor Green
Write-Host "  파일: $fileCount 개  →  $outDir" -ForegroundColor DarkGray
Write-Host ""
Write-Host "  배포 방법:" -ForegroundColor White
Write-Host "    1) ZIP 파일을 배포한다." -ForegroundColor DarkGray
Write-Host "    2) 사용자는 압축 해제 후 PolyDonky.exe 를 실행한다." -ForegroundColor DarkGray
Write-Host "    3) .NET 10 별도 설치 불필요." -ForegroundColor DarkGray
Write-Host ""
