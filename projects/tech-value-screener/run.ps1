# run.ps1 — 科技股潛在盈利篩選器 快速啟動腳本
# 使用：右鍵 → 用 PowerShell 執行，或在終端機輸入 .\run.ps1

$env:PYTHONUTF8 = "1"
$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $ScriptDir

Write-Host ""
Write-Host "======================================" -ForegroundColor Cyan
Write-Host "  科技股潛在盈利篩選器" -ForegroundColor Cyan
Write-Host "======================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "選擇執行模式：" -ForegroundColor Yellow
Write-Host "  [1] 分析美股清單（~30支，需5-8分鐘）"
Write-Host "  [2] 分析台股清單（~10支，需2-3分鐘）"
Write-Host "  [3] 分析美股＋台股（需10-15分鐘）"
Write-Host "  [4] 自訂股票（快速）"
Write-Host ""
$choice = Read-Host "請輸入選項 (1/2/3/4)"

switch ($choice) {
    "1" {
        Write-Host "▶ 分析美股..." -ForegroundColor Green
        python main.py --market us --output both
    }
    "2" {
        Write-Host "▶ 分析台股..." -ForegroundColor Green
        python main.py --market tw --output both
    }
    "3" {
        Write-Host "▶ 分析全部..." -ForegroundColor Green
        python main.py --market all --output both
    }
    "4" {
        $tickers = Read-Host "輸入股票代碼（空格分隔，台股加.TW，如：MSFT NVDA 2330.TW）"
        Write-Host "▶ 分析中..." -ForegroundColor Green
        python main.py --tickers $tickers.Split(' ') --output both
    }
    default {
        Write-Host "無效選項，執行預設美股分析..." -ForegroundColor Yellow
        python main.py --market us --output both
    }
}

Write-Host ""
Write-Host "完成！按任意鍵關閉..." -ForegroundColor Green
$null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
