<#
.SYNOPSIS
  Syncs the dev-source copy of article_db/index.html (this repo) to the deployed
  article-db-api repo (remote `article-db`, branch `main`), which is what Railway's
  articlebase.up.railway.app actually builds from.

.WHY
  article-db-api is a wholly separate repo/history from this one — pushing article_db/index.html
  changes to `origin` alone does NOT update the live site. Before this script existed the sync
  was fully manual (worktree add -> copy -> commit -> push -> worktree remove), and a bugfix
  once looked shipped (committed + pushed to origin) but never went live because Railway builds
  from article-db-api, not this repo. This script automates that worktree dance into one command.

  The diff/no-diff decision is made by letting `git status`/`git diff` inside the checked-out
  worktree do the comparison (not a hand-rolled hash of a manually-extracted blob) — that way
  line-ending normalization follows the target repo's own git config instead of getting mangled
  by string-processing the blob through PowerShell's pipeline.

.SCOPE
  Only article_db/index.html <-> article-db-api's root index.html. This does NOT commit or push
  anything in this repo (origin) — do that yourself as usual. This script only handles pushing
  the file to the article-db remote.

.USAGE
  pwsh scripts/sync-article-db.ps1          # fetch, diff, prompt before pushing
  pwsh scripts/sync-article-db.ps1 -DryRun  # fetch + show diff only, push nothing
#>

param(
  [switch]$DryRun
)

$ErrorActionPreference = 'Stop'

$root = Split-Path -Parent $PSScriptRoot
$srcPath = Join-Path $root 'article_db/index.html'
$remoteName = 'article-db'
$remoteBranch = 'main'

if (-not (Test-Path $srcPath)) {
  Write-Host "MISSING source: article_db/index.html" -ForegroundColor Red
  exit 1
}

Write-Host "Fetching $remoteName..." -ForegroundColor Cyan
git -C $root fetch $remoteName | Out-Host

$worktreePath = Join-Path $env:TEMP "article-db-sync-$([guid]::NewGuid().ToString('N'))"
Write-Host "Checking out ${remoteName}/${remoteBranch} into temp worktree..." -ForegroundColor Cyan
git -C $root worktree add --detach $worktreePath "${remoteName}/${remoteBranch}" | Out-Host

try {
  Copy-Item -Path $srcPath -Destination (Join-Path $worktreePath 'index.html') -Force
  git -C $worktreePath add index.html
  $status = git -C $worktreePath status --porcelain

  if (-not $status) {
    Write-Host "Already in sync." -ForegroundColor Green
    exit 0
  }

  Write-Host "DIFFERS - article_db/index.html vs ${remoteName}/${remoteBranch}:index.html" -ForegroundColor Yellow
  git -C $worktreePath --no-pager diff --cached --stat | Out-Host

  if ($DryRun) {
    Write-Host "`nDry run - nothing pushed." -ForegroundColor Cyan
    exit 0
  }

  $answer = Read-Host "Push article_db/index.html -> ${remoteName}/${remoteBranch}? [y/N]"
  if ($answer -ne 'y') {
    Write-Host "Skipped." -ForegroundColor DarkGray
    exit 0
  }

  git -C $worktreePath commit -m "Sync from global-economic-globe article_db/index.html" | Out-Host
  $sha = (git -C $worktreePath rev-parse HEAD).Trim()
  git -C $root push $remoteName "${sha}:${remoteBranch}" | Out-Host
  Write-Host "Pushed $sha to ${remoteName}/${remoteBranch}." -ForegroundColor Green
} finally {
  git -C $root worktree remove $worktreePath --force | Out-Host
}
