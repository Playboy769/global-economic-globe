<#
.SYNOPSIS
  Copies the dev-source copies of GlobalEco / InvestFrame / CausalFrame (this repo, `app/`)
  onto their deployed mirrors in the separate `globe-invest` repo (`globe-invest/app/`), then
  commits + pushes the change in BOTH repos.

.WHY
  globe-invest is deployed to Railway from its own repo — this repo's `app/` folder is
  NOT what ships. Every edit to GlobalEco/InvestFrame/CausalFrame has to be hand-copied into
  globe-invest, and that manual copy step has already caused a real regression once (an
  unrelated "fix sea routes" commit silently dropped the TTF commodity and mislabeled
  Heating Oil as Jet Fuel in the oil-price panel, and it went unnoticed for ~2 weeks).
  This script makes the copy + both commits + both pushes a single reviewed command instead
  of manual copy-paste followed by four separate git commands.

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

  Commit scope: each repo's commit is built with `git commit -- <paths>`, restricted to just
  the files this script copied — it will NOT sweep up unrelated staged/unstaged changes you
  might already have sitting in either repo (see this repo's Commit hygiene rule).

.USAGE
  pwsh scripts/sync-globe-invest.ps1          # diff, prompt per file, then prompt to commit+push both repos
  pwsh scripts/sync-globe-invest.ps1 -DryRun  # show diffs only, copy/commit/push nothing
#>

param(
  [switch]$DryRun
)

$ErrorActionPreference = 'Stop'

$root = Split-Path -Parent $PSScriptRoot
$giRoot = Join-Path $root 'globe-invest'

$pairs = @(
  @{ Name = 'GlobalEco';   Src = 'app/GlobalEco/index.html';   Dst = 'app/globe/index.html' },
  @{ Name = 'InvestFrame'; Src = 'app/InvestFrame/index.html'; Dst = 'app/invest/index.html' },
  @{ Name = 'CausalFrame'; Src = 'app/CausalFrame/index.html'; Dst = 'app/causal/index.html' }
)

$copied = @()

foreach ($pair in $pairs) {
  $srcPath = Join-Path $root $pair.Src
  $dstPath = Join-Path $giRoot $pair.Dst

  Write-Host "`n=== $($pair.Name) ===" -ForegroundColor Cyan

  if (-not (Test-Path $srcPath)) { Write-Host "  MISSING source: $($pair.Src)" -ForegroundColor Red; continue }
  if (-not (Test-Path $dstPath)) { Write-Host "  MISSING dest:   globe-invest/$($pair.Dst)" -ForegroundColor Red; continue }

  $srcHash = (Get-FileHash $srcPath -Algorithm SHA256).Hash
  $dstHash = (Get-FileHash $dstPath -Algorithm SHA256).Hash

  if ($srcHash -eq $dstHash) {
    Write-Host "  Already in sync." -ForegroundColor Green
    continue
  }

  Write-Host "  DIFFERS - dev source ($($pair.Src)) vs deployed mirror (globe-invest/$($pair.Dst))" -ForegroundColor Yellow
  git --no-pager diff --no-index --stat -- $dstPath $srcPath | Out-Host

  if ($DryRun) { continue }

  $answer = Read-Host "  Copy $($pair.Src) -> globe-invest/$($pair.Dst)? [y/N]"
  if ($answer -eq 'y') {
    Copy-Item -Path $srcPath -Destination $dstPath -Force
    Write-Host "  Copied." -ForegroundColor Green
    $copied += $pair
  } else {
    Write-Host "  Skipped." -ForegroundColor DarkGray
  }
}

if ($DryRun) {
  Write-Host "`nDry run - nothing copied, committed, or pushed." -ForegroundColor Cyan
  exit 0
}

if ($copied.Count -eq 0) {
  Write-Host "`nNothing copied - nothing to commit." -ForegroundColor Cyan
  exit 0
}

$names = ($copied | ForEach-Object { $_.Name }) -join ', '
Write-Host "`nCopied: $names" -ForegroundColor Cyan

$answer = Read-Host "Commit + push BOTH repos (origin here, globe-invest's origin) now? [y/N]"
if ($answer -ne 'y') {
  Write-Host "Skipped commit/push - files are copied on disk but repos are untouched." -ForegroundColor DarkGray
  exit 0
}

$defaultMsg = "Update $names"
$mainMsg = Read-Host "Commit message for THIS repo [$defaultMsg]"
if (-not $mainMsg) { $mainMsg = $defaultMsg }

$srcPaths = $copied | ForEach-Object { $_.Src }
$dstPaths = $copied | ForEach-Object { $_.Dst }

$mainBranch = (git -C $root rev-parse --abbrev-ref HEAD).Trim()
git -C $root commit -m $mainMsg -- $srcPaths | Out-Host
git -C $root push origin $mainBranch | Out-Host
Write-Host "Pushed this repo ($mainBranch)." -ForegroundColor Green

$giBranch = (git -C $giRoot rev-parse --abbrev-ref HEAD).Trim()
git -C $giRoot commit -m "Sync from Claudecode: $names" -- $dstPaths | Out-Host
git -C $giRoot push origin $giBranch | Out-Host
Write-Host "Pushed globe-invest ($giBranch)." -ForegroundColor Green
