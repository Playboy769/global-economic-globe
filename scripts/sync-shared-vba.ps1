<#
.SYNOPSIS
  Inject the shared-vba/ modules into the two workbooks that share them.

.DESCRIPTION
  shared-vba/ holds VBA modules used verbatim by BOTH:
    * the standalone SEC-Filing-Fetcher workbook, and
    * the RR4 portfolio workbook's "Company research" integration.

  This script pushes the current shared-vba/*.bas into each workbook's
  VBAProject via the VBIDE object model (CodeModule.AddFromString, UTF-8) --
  the SAME mechanism SEC-Filing-Fetcher/build.ps1 uses. It deliberately does
  NOT go through the VBE "Import File" menu, which decodes .bas with the
  system ANSI codepage and turns modHttp's Chinese error strings (and any
  BOM) into mojibake -- see the vba-bas-ascii-only memory / the mangled
  RR4/Sanner.bas banner.

  Modules synced (5):
    modHttp  modJsonUtil  modPrices  modSECData  modMOPSData

  For SEC-Filing-Fetcher, modHttp/modJsonUtil/modPrices are byte-identical to
  its own copies, so re-injecting them is a no-op in effect; modSECData /
  modMOPSData are the genuinely new pieces. (A later build.ps1 change should
  make it read the 3 leaves from ../shared-vba/ so there is one source of
  truth -- until then this script keeps them aligned.)

.PARAMETER SecXlsm
  Path to the built SEC-Filing-Fetcher .xlsm (default: the SaveAs target in
  SEC-Filing-Fetcher/build.ps1).

.PARAMETER Rr4Xlsm
  Path to the RR4 portfolio .xlsm.

.PARAMETER DryRun
  Report what would change (per module: NEW / REPLACED / UNCHANGED) without
  writing or saving anything.

.NOTES
  Requires: Excel installed, and "Trust access to the VBA project object
  model" enabled (File > Options > Trust Center > Trust Center Settings >
  Macro Settings). BOTH workbooks must be CLOSED before running.
#>
[CmdletBinding()]
param(
    [string]$SecXlsm = "C:\Users\ryan9\OneDrive\桌面\SECFilingFetcher.xlsm",
    [string]$Rr4Xlsm = "C:\Users\ryan9\OneDrive\桌面\Portfolio\Compound RR4 Portfolio 2026 H2.xlsm",
    [switch]$DryRun
)

$ErrorActionPreference = "Stop"
$repoRoot   = Split-Path -Parent $PSScriptRoot
$sharedDir  = Join-Path $repoRoot "shared-vba"
$modules    = @("modHttp", "modJsonUtil", "modPrices", "modSECData", "modMOPSData")
$vbext_ct_StdModule = 1
$xlOpenXMLWorkbookMacroEnabled = 52

# ---- read + lint the shared .bas files up front -------------------------------
$bodies = @{}
foreach ($m in $modules) {
    $p = Join-Path $sharedDir "$m.bas"
    if (-not (Test-Path $p)) { throw "missing $p" }
    $raw = Get-Content $p -Raw -Encoding UTF8

    # modSECData / modMOPSData must stay pure ASCII (see .NOTES). modHttp keeps
    # its Chinese error strings and modJsonUtil a BOM -- those are tolerated
    # because injection is UTF-8-safe, but the two NEW modules must not drift.
    if ($m -in @("modSECData", "modMOPSData")) {
        $bytes = [System.Text.Encoding]::UTF8.GetBytes($raw)
        if ($bytes | Where-Object { $_ -gt 127 }) {
            throw "$m.bas contains non-ASCII bytes -- keep it pure ASCII (rewrite CJK comments to English)."
        }
    }

    # strip the Attribute VB_Name line (AddFromString re-derives it from the component name)
    $lines = ($raw -split "`r`n|`n") | Where-Object { $_ -notmatch '^\s*Attribute\s+VB_Name\s*=' }
    $bodies[$m] = ($lines -join "`r`n").TrimEnd() + "`r`n"
}

function Sync-Workbook {
    param([string]$Path, [string]$Label)

    if (-not (Test-Path $Path)) {
        Write-Warning "[$Label] not found, skipping: $Path"
        return
    }
    Write-Host "=== $Label : $Path ===" -ForegroundColor Cyan

    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.AutomationSecurity = 1   # msoAutomationSecurityLow -- do not run auto-macros on open
    $dirty = $false
    try {
        $wb = $excel.Workbooks.Open($Path)

        try { $null = $wb.VBProject.VBComponents }
        catch {
            throw "cannot reach VBProject for $Label. Enable 'Trust access to the VBA project object model' in Excel's Trust Center."
        }
        $proj = $wb.VBProject

        foreach ($m in $modules) {
            $existing = $null
            foreach ($c in $proj.VBComponents) { if ($c.Name -eq $m) { $existing = $c; break } }

            $newBody = $bodies[$m]
            if ($existing -ne $null) {
                $cm = $existing.CodeModule
                $curBody = ""
                if ($cm.CountOfLines -gt 0) { $curBody = $cm.Lines(1, $cm.CountOfLines) }
                $curN = ($curBody   -replace "`r`n","`n").TrimEnd()
                $newN = ($newBody   -replace "`r`n","`n").TrimEnd()
                if ($curN -eq $newN) {
                    Write-Host ("  {0,-12} UNCHANGED" -f $m)
                    continue
                }
                if ($DryRun) { Write-Host ("  {0,-12} REPLACE (differs)" -f $m) -ForegroundColor Yellow; continue }
                $proj.VBComponents.Remove($existing)
            }
            else {
                if ($DryRun) { Write-Host ("  {0,-12} NEW" -f $m) -ForegroundColor Green; continue }
            }

            $comp = $proj.VBComponents.Add($vbext_ct_StdModule)
            $comp.Name = $m
            $comp.CodeModule.AddFromString($newBody)
            $dirty = $true
            Write-Host ("  {0,-12} {1}" -f $m, $(if ($existing) { "REPLACED" } else { "ADDED" })) -ForegroundColor Green
        }

        if ($DryRun) {
            Write-Host "  (dry run -- nothing written)"
            $wb.Close($false)
        }
        elseif ($dirty) {
            $wb.Save()
            $wb.Close($true)
            Write-Host "  saved." -ForegroundColor Green
        }
        else {
            $wb.Close($false)
            Write-Host "  nothing to do."
        }
    }
    finally {
        $excel.Quit()
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
        [System.GC]::Collect(); [System.GC]::WaitForPendingFinalizers()
    }
}

Write-Host "shared-vba sync  (modules: $($modules -join ', '))" -ForegroundColor White
if ($DryRun) { Write-Host "DRY RUN" -ForegroundColor Yellow }

Sync-Workbook -Path $SecXlsm -Label "SEC-Filing-Fetcher"
Sync-Workbook -Path $Rr4Xlsm -Label "RR4 portfolio"

Write-Host ""
Write-Host "Done. Verify with an Excel COM smoke test (see shared-vba/README.md)," -ForegroundColor White
Write-Host "e.g.  `$excel.Run('modSECData.GetXbrlFinancials','NVDA',4,8)" -ForegroundColor White
