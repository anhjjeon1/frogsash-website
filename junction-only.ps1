# Junction 생성 전용 스크립트 (2026-05-17)
# OneDrive sync 완료 후 노트북에서 한 줄로 실행:
#   irm https://raw.githubusercontent.com/anhjjeon1/frogsash-website/main/junction-only.ps1 | iex

$src = "C:\Users\$env:USERNAME\.claude\projects\D--\memory"
$dst = "C:\Users\$env:USERNAME\OneDrive\문서\claude-memory"

Write-Host ""
Write-Host "Junction 생성 시도..." -ForegroundColor Cyan
Write-Host "  src: $src"
Write-Host "  dst: $dst"
Write-Host ""

if (-not (Test-Path -LiteralPath $dst)) {
    Write-Host "❌ OneDrive 'claude-memory' 폴더가 노트북에 없음" -ForegroundColor Red
    Write-Host "   OneDrive sync 완료까지 대기 후 다시 실행" -ForegroundColor Yellow
    Write-Host "   확인: 노트북 탐색기에서 C:\Users\$env:USERNAME\OneDrive\문서\ 열기" -ForegroundColor Yellow
    return
}

if (Test-Path -LiteralPath $src) {
    $attr = (Get-Item -LiteralPath $src).Attributes
    if ($attr -match 'ReparsePoint') {
        Write-Host "ℹ️ 이미 Junction 존재" -ForegroundColor DarkGray
        return
    } else {
        Remove-Item -LiteralPath $src -Recurse -Force
        Write-Host "  ✓ 기존 폴더 제거됨" -ForegroundColor Green
    }
}

$parent = Split-Path $src -Parent
if (-not (Test-Path -LiteralPath $parent)) {
    New-Item -ItemType Directory -Path $parent -Force | Out-Null
}

$result = cmd /c mklink /J "`"$src`"" "`"$dst`"" 2>&1
Write-Host "  $result" -ForegroundColor Green
Write-Host ""

if (Test-Path -LiteralPath $src) {
    $attr = (Get-Item -LiteralPath $src).Attributes
    if ($attr -match 'ReparsePoint') {
        Write-Host "✅ Junction 생성 성공" -ForegroundColor Green
        Write-Host ""
        $count = (Get-ChildItem -LiteralPath $src -ErrorAction SilentlyContinue).Count
        Write-Host "  메모리 파일 수: $count" -ForegroundColor Green
    } else {
        Write-Host "⚠️ 폴더는 생겼지만 Junction 아님" -ForegroundColor Red
    }
} else {
    Write-Host "❌ Junction 생성 실패" -ForegroundColor Red
}
