# Laptop Setup Script v2 (2026-05-17)
# 노트북에 D 드라이브 없음 → C 드라이브로 변경. 잘못 clone된 폴더 정리 + 재clone

$ErrorActionPreference = 'Continue'
Write-Host ""
Write-Host "=========================================" -ForegroundColor Cyan
Write-Host "  Laptop Setup v2 — C 드라이브 사용" -ForegroundColor Cyan
Write-Host "=========================================" -ForegroundColor Cyan
Write-Host ""

$U = $env:USERNAME

# STEP A: 잘못 clone된 폴더 정리 (C:\Users\anhje 안)
Write-Host "[A] 잘못 clone된 폴더 정리..." -ForegroundColor Yellow
$wrongLocations = @(
    "C:\Users\$U\github",
    "C:\Users\$U\inspect-spare",
    "C:\Users\$U\frogcheck"
)
foreach ($p in $wrongLocations) {
    if (Test-Path -LiteralPath $p) {
        Remove-Item -LiteralPath $p -Recurse -Force
        Write-Host "  ✓ 제거: $p" -ForegroundColor Green
    }
}

# STEP B: C 드라이브에 새로 clone
Write-Host ""
Write-Host "[B] C 드라이브에 git clone..." -ForegroundColor Yellow

# B1. C:\github (frogsash-website)
if (-not (Test-Path "C:\github\.git")) {
    New-Item -ItemType Directory -Path "C:\github" -Force | Out-Null
    Set-Location C:\github
    git clone https://github.com/anhjjeon1/frogsash-website.git . 2>&1 | Out-Null
    if (Test-Path "C:\github\.git") {
        Write-Host "  ✓ C:\github (frogsash-website) cloned" -ForegroundColor Green
    } else {
        Write-Host "  ❌ C:\github clone 실패" -ForegroundColor Red
    }
} else {
    Write-Host "  - C:\github 이미 존재" -ForegroundColor DarkGray
}

# B2. C:\github\inspect-spare
if (-not (Test-Path "C:\github\inspect-spare\.git")) {
    Set-Location C:\github
    git clone https://github.com/anhjjeon1/inspect-spare.git inspect-spare 2>&1 | Out-Null
    if (Test-Path "C:\github\inspect-spare\.git") {
        Write-Host "  ✓ C:\github\inspect-spare cloned" -ForegroundColor Green
    } else {
        Write-Host "  ❌ inspect-spare clone 실패" -ForegroundColor Red
    }
} else {
    Write-Host "  - C:\github\inspect-spare 이미 존재" -ForegroundColor DarkGray
}

# B3. C:\frogcheck
if (-not (Test-Path "C:\frogcheck\.git")) {
    Set-Location C:\
    git clone https://github.com/anhjjeon1/frogcheck.git frogcheck 2>&1 | Out-Null
    if (Test-Path "C:\frogcheck\.git") {
        Write-Host "  ✓ C:\frogcheck cloned" -ForegroundColor Green
    } else {
        Write-Host "  ❌ frogcheck clone 실패" -ForegroundColor Red
    }
} else {
    Write-Host "  - C:\frogcheck 이미 존재" -ForegroundColor DarkGray
}

# STEP C: 최종 검증
Write-Host ""
Write-Host "[C] 최종 검증..." -ForegroundColor Yellow
$claudeExe = "C:\Users\$U\.local\bin\claude.exe"
if (Test-Path $claudeExe) {
    $ver = & $claudeExe --version 2>&1
    Write-Host "  Claude: $ver" -ForegroundColor Green
}

$dirs = @("C:\github", "C:\github\inspect-spare", "C:\frogcheck")
foreach ($d in $dirs) {
    if (Test-Path -LiteralPath $d) {
        $fc = (Get-ChildItem -LiteralPath $d -ErrorAction SilentlyContinue).Count
        Write-Host "  ${d}: $fc items" -ForegroundColor Green
    } else {
        Write-Host "  ${d}: ❌ 없음" -ForegroundColor Red
    }
}

# STEP D: 메모리 Junction (OneDrive 폴더 있을 때만)
Write-Host ""
Write-Host "[D] 메모리 Junction..." -ForegroundColor Yellow
$src = "C:\Users\$U\.claude\projects\D--\memory"
$dst = "C:\Users\$U\OneDrive\문서\claude-memory"
if (-not (Test-Path -LiteralPath $dst)) {
    Write-Host "  ⚠️ OneDrive 'claude-memory' 폴더 없음 (sync 미완)" -ForegroundColor Yellow
    Write-Host "     OneDrive sync 완료 후 다시 시도하세요" -ForegroundColor Yellow
} else {
    if (Test-Path -LiteralPath $src) {
        $attr = (Get-Item -LiteralPath $src).Attributes
        if ($attr -match 'ReparsePoint') {
            Write-Host "  - 이미 Junction 존재" -ForegroundColor DarkGray
        } else {
            Remove-Item -LiteralPath $src -Recurse -Force
            cmd /c mklink /J "`"$src`"" "`"$dst`"" | Out-Null
            Write-Host "  ✓ Junction 생성" -ForegroundColor Green
        }
    } else {
        New-Item -ItemType Directory -Path (Split-Path $src -Parent) -Force | Out-Null
        cmd /c mklink /J "`"$src`"" "`"$dst`"" | Out-Null
        Write-Host "  ✓ Junction 생성" -ForegroundColor Green
    }
}

Write-Host ""
Write-Host "=========================================" -ForegroundColor Cyan
Write-Host "  v2 셋업 완료!" -ForegroundColor Green
Write-Host "=========================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "노트북 작업 경로 (D: 대신 C: 사용):" -ForegroundColor White
Write-Host "  C:\github          (frogsash-website)" -ForegroundColor White
Write-Host "  C:\github\inspect-spare" -ForegroundColor White
Write-Host "  C:\frogcheck" -ForegroundColor White
