# Builds "Ticker Volatility.xlsm" from a blank workbook.
# Run:  powershell -STA -File VT\build.ps1 -OutPath "C:\path\Ticker Volatility.xlsm"
# Needs Excel with "Trust access to the VBA project object model" enabled.
param([Parameter(Mandatory = $true)][string]$OutPath)

$ErrorActionPreference = 'Stop'
$here = $PSScriptRoot
$bas  = [IO.File]::ReadAllText((Join-Path $here 'modVT.bas'), [Text.Encoding]::ASCII)
$sht  = [IO.File]::ReadAllText((Join-Path $here 'SheetTickerVolatility_Code.txt'), [Text.Encoding]::ASCII)

# the module header line must not be part of the injected text
$lines = $bas -split "`r?`n"
if ($lines[0] -like 'Attribute VB_Name*') { $lines = $lines[1..($lines.Count - 1)] }
$bas = ($lines -join "`r`n")

$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false
$xl.AutomationSecurity = 1
try {
    $wb = $xl.Workbooks.Add()
    while ($wb.Worksheets.Count -gt 1) { $wb.Worksheets.Item($wb.Worksheets.Count).Delete() }
    $ws = $wb.Worksheets.Item(1)
    $ws.Name = 'Ticker Volatility'

    $mod = $wb.VBProject.VBComponents.Add(1)
    $mod.Name = 'modVT'
    $mod.CodeModule.AddFromString($bas)

    $sm = $wb.VBProject.VBComponents.Item($ws.CodeName).CodeModule
    $sm.AddFromString($sht)

    $xl.Run("'" + $wb.Name + "'!InitPage")
    $ws.Activate()
    $ws.Range('B2').Select()
    $xl.ActiveWindow.DisplayGridlines = $false

    # no personal metadata
    $wb.RemovePersonalInformation = $true
    foreach ($n in 'Author', 'Last Author', 'Manager', 'Company', 'Title', 'Subject', 'Comments', 'Keywords') {
        try { $wb.BuiltinDocumentProperties.Item($n).Value = '' } catch {}
    }

    if (Test-Path $OutPath) { Remove-Item $OutPath -Force }
    $wb.SaveAs($OutPath, 52)
    $wb.Close($false)
    Write-Output "saved $OutPath"
}
finally {
    $xl.Quit()
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl)
}
