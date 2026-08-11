$ErrorActionPreference = "Stop"
$base = "C:\Users\ryan9\OneDrive\桌面\Claudecode\SEC-Filing-Fetcher"
$xlOpenXMLWorkbookMacroEnabled = 52

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

try {
    $wb = $excel.Workbooks.Add()

    while ($wb.Worksheets.Count -lt 3) {
        $wb.Worksheets.Add() | Out-Null
    }
    while ($wb.Worksheets.Count -gt 3) {
        $wb.Worksheets.Item($wb.Worksheets.Count).Delete()
    }

    $wsInput = $wb.Worksheets.Item(1)
    $wsFilings = $wb.Worksheets.Item(2)
    $wsWatchlist = $wb.Worksheets.Item(3)

    $wsInput.Name = "Input"
    $wsFilings.Name = "Filings"
    $wsWatchlist.Name = "Watchlist"

    $wsInput.Range("A2").Value2 = "在 A1 輸入美股股票代號 (例如 AAPL)，按 Enter 後自動抓取 10-K/10-Q/8-K/Form 4"
    $wsInput.Range("A1").Font.Bold = $true
    $wsInput.Columns.Item(1).ColumnWidth = 16
    $wsInput.Columns.Item(2).ColumnWidth = 100
    $wsInput.Range("B1").WrapText = $false

    $wsInput.Range("A3").Value2 = "抓取設定（可留白，使用預設值）"
    $wsInput.Range("A3").Font.Bold = $true
    $wsInput.Range("A4").Value2 = "10-K 年數"
    $wsInput.Range("B4").Value2 = 7
    $wsInput.Range("A5").Value2 = "10-Q 季數"
    $wsInput.Range("B5").Value2 = 8
    $wsInput.Range("A6").Value2 = "8-K 篇數"
    $wsInput.Range("B6").Value2 = 10
    $wsInput.Range("A7").Value2 = "Form 4 篇數"
    $wsInput.Range("B7").Value2 = 10

    # Black/orange dark theme, Yu Gothic throughout, applied once here since the
    # Input sheet's static cells (unlike Filings/OtherFilings/etc.) are never
    # cleared or rewritten by the macros, so this persists across runs.
    $inputRange = $wsInput.Range("A1:B7")
    $inputRange.Interior.Color = 0
    $inputRange.Font.Color = 16777215
    $inputRange.Font.Name = "Yu Gothic"
    $wsInput.Range("A1:AN500").Interior.Color = 0
    $wsInput.Range("A3").Font.Color = 42495
    $wsInput.Range("A4:A7").Font.Color = 42495
    $wsInput.Range("A1").Interior.Color = 23220               # dark orange, RGB(180,90,0)
    $wsInput.Range("A1").Font.Color = 16777215
    $wsInput.Range("B1").Font.Color = 16777215

    $borderRange = $wsInput.Range("A1:B7")
    $borderRange.Borders.LineStyle = 1   # xlContinuous
    $borderRange.Borders.Weight = 2      # xlThin
    $borderRange.Borders.Color = 5263440 # dark gray, RGB(80,80,80)

    $headers = @("公司名稱","CIK","表單類別","會計年度(FY)","會計期間(FP)","財報截止日","申報日期","Accession Number","文件連結",`
        "營收 Revenue","淨利 Net Income","稀釋EPS","總資產 Assets","總負債 Liabilities","營運現金流 CFO",`
        "修正版","毛利率 Gross Margin","營業利益率 Operating Margin","研發費用 R&D","SG&A","負債比 Debt Ratio","每股股利 Dividend/Share",`
        "申報索引頁","部門附註連結(如有)")
    for ($i = 0; $i -lt $headers.Count; $i++) {
        $wsFilings.Cells.Item(1, $i + 1).Value2 = $headers[$i]
    }
    $headerRange = $wsFilings.Range($wsFilings.Cells.Item(1,1), $wsFilings.Cells.Item(1, $headers.Count))
    $headerRange.Font.Bold = $true
    $headerRange.Font.Color = 42495
    $headerRange.Font.Name = "Yu Gothic"
    $headerRange.Interior.Color = 0

    # Watchlist: user types tickers into A3:A22 (20 rows), Worksheet_Change
    # (injected below) fills B:L per row. Pre-created here (rather than
    # lazily via GetOrCreateSheet like Dashboard/OtherFilings) because it
    # needs its own event code-behind, which can only be injected into a
    # sheet's CodeName at build time -- a sheet created at runtime by a
    # running macro has no such hook available to it.
    $wsWatchlist.Range("A1").Value2 = "Watchlist — 同業比較 Peer Comparison"
    $wsWatchlist.Range("A1").Font.Bold = $true
    $wsWatchlist.Range("A2").Value2 = "在下方 A 欄輸入股票代號並按 Enter，自動抓取最新一期 10-K 關鍵指標橫向比較（最多 20 檔，逐列增量更新，不會重算其他列）"
    $wlHeaders = @("代號 Ticker","公司名稱","股價","市值(M)","營收成長% YoY","P/E","P/S","EV/EBITDA","ROE","股利殖利率","流動比率","更新時間")
    for ($i = 0; $i -lt $wlHeaders.Count; $i++) {
        $wsWatchlist.Cells.Item(2, $i + 1).Value2 = $wlHeaders[$i]
    }
    $wlHeaderRange = $wsWatchlist.Range($wsWatchlist.Cells.Item(2,1), $wsWatchlist.Cells.Item(2, $wlHeaders.Count))
    $wlHeaderRange.Font.Bold = $true
    $wlHeaderRange.Font.Color = 42495
    $wlHeaderRange.Font.Name = "Yu Gothic"

    $wlAllRange = $wsWatchlist.Range("A1:L22")
    $wlAllRange.Interior.Color = 0
    $wlAllRange.Font.Color = 16777215
    $wlAllRange.Font.Name = "Yu Gothic"
    $wsWatchlist.Range("A1:AN500").Interior.Color = 0
    $wsWatchlist.Range("A1").Interior.Color = 23220              # dark orange, RGB(180,90,0)
    $wsWatchlist.Range("A1").Font.Color = 16777215

    $wlBorderRange = $wsWatchlist.Range($wsWatchlist.Cells.Item(2,1), $wsWatchlist.Cells.Item(22, $wlHeaders.Count))
    $wlBorderRange.Borders.LineStyle = 1
    $wlBorderRange.Borders.Weight = 2
    $wlBorderRange.Borders.Color = 5263440
    $wsWatchlist.Columns.Item(1).ColumnWidth = 10
    $wsWatchlist.Columns.Item(2).ColumnWidth = 24

    Write-Output "STEP: about to access VBProject"
    $vbProj = $wb.VBProject
    Write-Output "STEP: got VBProject, type = $($vbProj.GetType().Name)"

    $vbext_ct_StdModule = 1
    foreach ($modName in @("modJsonUtil","modHttp","modPrices","modTheme","modSEC","modCharts")) {
        $modComp = $vbProj.VBComponents.Add($vbext_ct_StdModule)
        $modComp.Name = $modName
        $raw = Get-Content "$base\$modName.bas" -Raw -Encoding UTF8
        $bodyLines = ($raw -split "`r`n|`n") | Where-Object { $_ -notmatch '^\s*Attribute\s+VB_Name\s*=' }
        $body = ($bodyLines -join "`r`n")
        $modComp.CodeModule.AddFromString($body)
    }

    $inputCodeName = $wsInput.CodeName
    $comp = $vbProj.VBComponents.Item($inputCodeName)
    $code = Get-Content "$base\SheetInput_Code.txt" -Raw -Encoding UTF8
    $comp.CodeModule.AddFromString($code)

    $watchlistCodeName = $wsWatchlist.CodeName
    $wlComp = $vbProj.VBComponents.Item($watchlistCodeName)
    $wlCode = Get-Content "$base\SheetWatchlist_Code.txt" -Raw -Encoding UTF8
    $wlComp.CodeModule.AddFromString($wlCode)

    $savePath = "C:\Users\ryan9\OneDrive\桌面\SECFilingFetcher.xlsm"
    $wb.SaveAs($savePath, $xlOpenXMLWorkbookMacroEnabled)
    $wb.Close($false)
    Write-Output "SUCCESS: $savePath"
}
catch {
    Write-Output "ERROR: $($_.Exception.Message)"
    Write-Output "AT: $($_.InvocationInfo.PositionMessage)"
}
finally {
    $excel.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
}
