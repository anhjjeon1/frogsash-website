# Laptop Setup Script (2026-05-17)
# 노트북에 Claude Code CLI PATH 추가 + 메모리 Junction + git clone 일괄 자동
# 사용: 노트북 관리자 PowerShell에서 한 줄
#   irm https://raw.githubusercontent.com/anhjjeon1/frogsash-website/main/laptop-setup.ps1 | iex

$ErrorActionPreference = 'Continue'
Write-Host ""
Write-Host "=========================================" -ForegroundColor Cyan
Write-Host "  Laptop Setup Script (2026-05-17)" -ForegroundColor Cyan
Write-Host "=========================================" -ForegroundColor Cyan
Write-Host ""

$U = $env:USERNAME
Write-Host "현재 사용자: $U" -ForegroundColor Gray
Write-Host ""

# STEP 1: PATH 영구 추가
Write-Host "[STEP 1] PATH 영구 추가..." -ForegroundColor Yellow
$bin = "C:\Users\$U\.local\bin"
$userPath = [Environment]::GetEnvironmentVariable('Path','User')
if ($userPath -notlike "*$bin*") {
    [Environment]::SetEnvironmentVariable('Path', $userPath + ";$bin", 'User')
    $env:Path += ";$bin"
    Write-Host "  ✓ PATH 추가: $bin" -ForegroundColor Green
} else {
    Write-Host "  - 이미 PATH에 있음" -ForegroundColor DarkGray
}

# STEP 2: Claude 검증
Write-Host ""
Write-Host "[STEP 2] Claude Code CLI 검증..." -ForegroundColor Yellow
$claudeExe = "$bin\claude.exe"
if (Test-Path $claudeExe) {
    $ver = & $claudeExe --version 2>&1
    Write-Host "  ✓ $ver" -ForegroundColor Green
} else {
    Write-Host "  ❌ claude.exe 없음 — irm https://claude.ai/install.ps1 | iex 먼저 실행" -ForegroundColor Red
    return
}

# STEP 3: 메모리 Junction 생성
Write-Host ""
Write-Host "[STEP 3] 메모리 Junction 생성..." -ForegroundColor Yellow
$src = "C:\Users\$U\.claude\projects\D--\memory"
$dst = "C:\Users\$U\OneDrive\문서\claude-memory"
if (-not (Test-Path -LiteralPath $dst)) {
    Write-Host "  ⚠️ OneDrive 폴더 없음: $dst" -ForegroundColor Yellow
    Write-Host "     PC에서 OneDrive sync 완료까지 대기 후 재실행 필요" -ForegroundColor Yellow
    Write-Host "     또는 OneDrive 웹에서 'claude-memory' 폴더 보이는지 확인" -ForegroundColor Yellow
} else {
    if (Test-Path -LiteralPath $src) {
        $attr = (Get-Item -LiteralPath $src).Attributes
        if ($attr -match 'ReparsePoint') {
            Write-Host "  - 이미 Junction 존재" -ForegroundColor DarkGray
        } else {
            Remove-Item -LiteralPath $src -Recurse -Force
            New-Item -ItemType Directory -Path (Split-Path $src -Parent) -Force | Out-Null
            cmd /c mklink /J "`"$src`"" "`"$dst`"" | Out-Null
            Write-Host "  ✓ Junction 생성: $src" -ForegroundColor Green
        }
    } else {
        New-Item -ItemType Directory -Path (Split-Path $src -Parent) -Force | Out-Null
        cmd /c mklink /J "`"$src`"" "`"$dst`"" | Out-Null
        Write-Host "  ✓ Junction 생성: $src" -ForegroundColor Green
    }
}

# STEP 4: git clone (3개 repo)
Write-Host ""
Write-Host "[STEP 4] git clone..." -ForegroundColor Yellow

# 4a. D:/github (frogsash-website)
if (-not (Test-Path "D:\github\.git")) {
    New-Item -ItemType Directory -Path "D:\github" -Force | Out-Null
    Set-Location D:\github
    git clone https://github.com/anhjjeon1/frogsash-website.git . 2>&1 | Out-Null
    Write-Host "  ✓ D:\github (frogsash-website) cloned" -ForegroundColor Green
} else {
    Write-Host "  - D:\github 이미 존재" -ForegroundColor DarkGray
}

# 4b. D:/github/inspect-spare
if (-not (Test-Path "D:\github\inspect-spare\.git")) {
    Set-Location D:\github
    git clone https://github.com/anhjjeon1/inspect-spare.git inspect-spare 2>&1 | Out-Null
    Write-Host "  ✓ D:\github\inspect-spare cloned" -ForegroundColor Green
} else {
    Write-Host "  - D:\github\inspect-spare 이미 존재" -ForegroundColor DarkGray
}

# 4c. D:/frogcheck
if (-not (Test-Path "D:\frogcheck\.git")) {
    Set-Location D:\
    git clone https://github.com/anhjjeon1/frogcheck.git frogcheck 2>&1 | Out-Null
    Write-Host "  ✓ D:\frogcheck cloned" -ForegroundColor Green
} else {
    Write-Host "  - D:\frogcheck 이미 존재" -ForegroundColor DarkGray
}

# STEP 5: 최종 요약
Write-Host ""
Write-Host "=========================================" -ForegroundColor Cyan
Write-Host "  셋업 완료!" -ForegroundColor Green
Write-Host "=========================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "다음:" -ForegroundColor White
Write-Host "  1. 이 PowerShell 창 닫기" -ForegroundColor White
Write-Host "  2. 새 PowerShell 열기 (Win+X → 터미널)" -ForegroundColor White
Write-Host "  3. 'claude' 명령어로 어디서나 실행 가능" -ForegroundColor White
Write-Host "  4. D:\github, D:\frogcheck 폴더 확인" -ForegroundColor White
Write-Host ""
