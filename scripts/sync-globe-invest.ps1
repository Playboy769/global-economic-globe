<#
.SYNOPSIS
  Copies the dev-source copies of GlobalEco / InvestFrame / CausalFrame (this repo, `app/`)
  onto their deployed mirrors in the separate `globe-invest` repo (`globe-invest/app/`).

.WHY
  globe-invest is deployed to Railway from its own repo — this repo's `app/` folder is
  NOT what ships. Every edit to GlobalEco/InvestFrame/CausalFrame has to be hand-copied into
  globe-invest, and that manual copy step has already caused a real regression once (an
  unrelated "fix sea routes" commit silently dropped the TTF commodity and mislabeled
  Heating Oil as Jet Fuel in the oil-price panel, and it went unnoticed for ~2 weeks).
  This script makes the copy step a single reviewed command instead of manual copy-paste.

.NOTE ON SCOPE
  Only these three pairs are covered — they are the ones with a confirmed clean 1:1
  dev-source -> deployed-mirror relationship:
    app/GlobalEco/index.html    -> globe-invest/app/globe/index.html
    app/InvestFrame/index.html  -> globe-invest/app/invest/index.html
    app/CausalFrame/index.html  -> globe-invest/app/causal/index.html

  globe-invest/app/options, /warning, and /research ALSO look like they started as copies
  of root-level files (options_guide.html, 市場警示雷達.html, MU Analysis/), but as of
  2026-07-01 the warning and research pairs have diverged into genuinely different content
  (not just "behind a few edits" — different structure/theme/language). Do NOT add them to
  this script's automated copy list until a human reconciles which side is authoritative;
  a blind sync would silently overwrite whichever side is actually correct.

.USAGE
  pwsh scripts/sync-globe-invest.ps1          # preview diffs, prompt per file before copying
  pwsh scripts/sync-globe-invest.ps1 -DryRun   # show diffs only, copy nothing
#>

param(
  [switch]$DryRun
)

$root = Split-Path -Parent $PSScriptRoot

$pairs = @(
  @{ Name = 'GlobalEco';   Src = 'app/GlobalEco/index.html';   Dst = 'globe-invest/app/globe/index.html' },
  @{ Name = 'InvestFrame'; Src = 'app/InvestFrame/index.html'; Dst = 'globe-invest/app/invest/index.html' },
  @{ Name = 'CausalFrame'; Src = 'app/CausalFrame/index.html'; Dst = 'globe-invest/app/causal/index.html' }
)

foreach ($pair in $pairs) {
  $srcPath = Join-Path $root $pair.Src
  $dstPath = Join-Path $root $pair.Dst

  Write-Host "`n=== $($pair.Name) ===" -ForegroundColor Cyan

  if (-not (Test-Path $srcPath)) { Write-Host "  MISSING source: $($pair.Src)" -ForegroundColor Red; continue }
  if (-not (Test-Path $dstPath)) { Write-Host "  MISSING dest:   $($pair.Dst)" -ForegroundColor Red; continue }

  $srcHash = (Get-FileHash $srcPath -Algorithm SHA256).Hash
  $dstHash = (Get-FileHash $dstPath -Algorithm SHA256).Hash

  if ($srcHash -eq $dstHash) {
    Write-Host "  Already in sync." -ForegroundColor Green
    continue
  }

  Write-Host "  DIFFERS - dev source ($($pair.Src)) vs deployed mirror ($($pair.Dst))" -ForegroundColor Yellow
  git --no-pager diff --no-index --stat -- $dstPath $srcPath | Out-Host

  if ($DryRun) { continue }

  $answer = Read-Host "  Copy $($pair.Src) -> $($pair.Dst)? [y/N]"
  if ($answer -eq 'y') {
    Copy-Item -Path $srcPath -Destination $dstPath -Force
    Write-Host "  Copied." -ForegroundColor Green
  } else {
    Write-Host "  Skipped." -ForegroundColor DarkGray
  }
}

Write-Host "`nDone. Remember: this only stages the file copy - you still need to" -ForegroundColor Cyan
Write-Host "commit + push BOTH repos (origin and globe-invest) for changes to go live." -ForegroundColor Cyan
