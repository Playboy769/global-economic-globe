Attribute VB_Name = "Sanner"
Option Explicit

' ================================================================
'  MARKET SCANNER v2.1
' ================================================================


' ================================================================
'  GROUP TABLE -- sheet "Groups", ListObject "tblGroups"
'  Single source of truth for scanner groups: GetSectorTickers, the
'  keyword cell on "Company research" and the group panel all read it. Columns by position: 1 Market (TW/US) | 2 Group |
'  3 Ticker | 4 Source | 5 Note -- one row per ticker.
' ================================================================
Public Const GROUPS_SHEET      As String = "Groups"
Public Const GROUPS_TABLE      As String = "tblGroups"
Public Const GROUP_LIST_NAME   As String = "GroupList"
Public Const GROUP_LIST_COL    As Long = 8          ' column H of the Groups sheet
' ----------------------------------------------------------------
'  CR PAGE LAYOUT v4 (2026-09-12) - RR4 colours, one horizontal input strip
'    row 1        A  page title, C  scan context (market | group | time)
'    row 2        ONE input strip across the top (DrawCrHeader):
'                   A label / B GROUP <GO>    C label / D TICKER
'                   E label / F MKT (US / TW / blank)
'    row 3        blank
'    row 4        scan table headers (C:K)
'    rows 5-33    scanned tickers, row 35 the SUMMARY line
'    column L     blank spacer between the scan table and the panel
'    column M on  group database panel (RebuildGroupDb) - group name and
'                 ticker count only; the TICKERS column was dropped in v4,
'                 it was always clipped by the column width anyway
'    row 40+      financial deep-dive (CompanyResearchSEC)
'  Double-click a ticker in the scan table -> deep-dive; double-click a
'  group in the panel -> scan. Status text goes to the Excel status bar.
'  v4 also dropped the 1-year sparkline column (it sat in L and never drew
'  anything after the v3 rebuild); ClearSparklines still runs so leftovers
'  from v3 do not sit underneath the panel.
' ----------------------------------------------------------------
Public Const CR_STRIP_ROW      As Long = 2
Public Const GROUP_INPUT_CELL  As String = "B2"
Public Const CR_TICKER_INPUT   As String = "D2"     ' = CompanyResearchSEC.CR_INPUT_CELL
Public Const CR_MARKET_INPUT   As String = "F2"     ' = CompanyResearchSEC.CR_MARKET_CELL

' The scanner owns rows 1..SCAN_LAST_ROW of C:K; rows 36-39 are the spacer
' above the deep-dive band. The summary line lands two rows below the last
' ticker, so 29 tickers (rows 5..33, summary on 35) is the most that fits.
Public Const SCAN_TITLE_ROW    As Long = 1
Public Const SCAN_HDR_ROW      As Long = 4
Public Const SCAN_FIRST_ROW    As Long = 5
Public Const SCAN_LAST_ROW     As Long = 35
Public Const SCAN_MAX_TICKERS  As Long = 29
Public Const SCAN_ROW_H        As Double = 20

' Group database panel beside the scanner (see RebuildGroupDb): four
' side-by-side blocks from column M, each GROUP | N + a gap column,
' confined to rows 1..SCAN_LAST_ROW because the deep-dive clears the band
' below it. The last block takes the TW groups that do not fit in the TW
' block. v4 dropped the third (TICKERS) column of each block - the list was
' clipped by the column width, so it read as noise rather than data.
' DB_FIRST_COL - 1 (column L) is left blank as a spacer, the same way
' column A is on the RR4 page: the panel used to butt straight up against
' the SECTOR column and the two blocks read as one.
Public Const DB_FIRST_COL      As Long = 13         ' M
Public Const DB_BLOCK_COLS     As Long = 3
Public Const DB_BLOCKS         As Long = 4

' Scan table columns (C..K). SECTOR stays last: RebuildGroupDb reads it to
' highlight the groups shown in the current scan.
Public Const SC_TICKER         As Long = 3
Public Const SC_COMPANY        As Long = 4
Public Const SC_PRICE          As Long = 5
Public Const SC_HIGHDIST       As Long = 6
Public Const SC_CHG180         As Long = 7
Public Const SC_R20            As Long = 8
Public Const SC_R55            As Long = 9
Public Const SC_TREND          As Long = 10
Public Const SC_SECTOR         As Long = 11
Public Const SCAN_LAST_COL     As Long = 11

' Normalised EMA bias (user's TDX formula):
'   B = (C - EMA(C,N)) / EMA(C,N) * 100
'   R = B / MAX(HHV(B,M), 0.01) * 100        when B >= 0
'   R = B / MAX(ABS(LLV(B,M)), 0.01) * 100   when B <  0
' so R sits in [-100, 100]; |R| >= 80 (K = 0.8) is the strong zone.
Private Const NB_SHORT         As Long = 20
Private Const NB_LONG          As Long = 55
Private Const NB_LOOKBACK      As Long = 250
Private Const NB_STRONG        As Double = 80

' CHART_DATA_SHEET / CHART_PREFIX are v4 leftovers kept only so the old
' sparkline source sheet and the even older chart wall can still be found
' and cleaned up; nothing writes to them any more.
Public Const CHART_DATA_SHEET  As String = "ScanPrices"
Private Const CHART_PREFIX     As String = "ScanChart_"
Private Const YEAR_BARS        As Long = 250        ' "1 year" window for 1Y HIGH%

' ================================================================
'  MAIN ENTRY
' ================================================================
' Scans tickerList into rows SCAN_FIRST_ROW.. of C:K. tickerGroups (optional,
' indexed like tickerList) labels each row's SECTOR column when several groups
' are merged. At most SCAN_MAX_TICKERS are scanned so the table never runs into
' the SUMMARY line. Rows are sorted by R55, strongest first.
Sub ScanTickers(market As String, Sector As String, tickerList As Variant, tickerGroups As Variant)
    Dim wsRes As Worksheet
    On Error Resume Next
    Set wsRes = ThisWorkbook.Sheets("Company research")
    On Error GoTo 0
    If wsRes Is Nothing Then MsgBox "Sheet 'Company research' not found.", vbExclamation: Exit Sub

    Dim total As Long, lastIdx As Long
    total = UBound(tickerList) - LBound(tickerList) + 1
    lastIdx = LBound(tickerList) + IIf(total > SCAN_MAX_TICKERS, SCAN_MAX_TICKERS, total) - 1

    Application.ScreenUpdating = False

    ' Only the scan block is cleared -- the group panel sits right of it and
    ' the chart wall / deep-dive below it
    Dim scanArea As Range
    ' from the header row down only - row 2 carries the TICKER / MKT input
    ' cells of the strip, which a group scan must not wipe
    Set scanArea = wsRes.Range(wsRes.cells(SCAN_HDR_ROW, SC_TICKER), wsRes.cells(SCAN_LAST_ROW, SCAN_LAST_COL))
    ' Row 1 (title) and row 3 (the blank line under the strip, where the v3
    ' table used to start) are cleared separately - and REPAINTED, because
    ' Range.Clear leaves "no fill", which is white on this black page.
    Dim gapRow As Range
    Set gapRow = wsRes.Range(wsRes.cells(SCAN_TITLE_ROW, SC_TICKER), wsRes.cells(SCAN_TITLE_ROW, SCAN_LAST_COL))
    gapRow.Clear: gapRow.Interior.Color = RGB(0, 0, 0)
    Set gapRow = wsRes.Range(wsRes.cells(SCAN_TITLE_ROW + 2, SC_TICKER), wsRes.cells(SCAN_TITLE_ROW + 2, SCAN_LAST_COL))
    gapRow.Clear: gapRow.Interior.Color = RGB(0, 0, 0)
    Call ClearSparklines(wsRes)
    With scanArea
        .Clear
        .Interior.Color = RGB(0, 0, 0)
        .Font.Color = RGB(200, 200, 200)
        .Font.Name = "Consolas"
        .Font.Size = 9
    End With

    With wsRes.cells(SCAN_TITLE_ROW, SC_TICKER)
        .Value = "MARKET SCANNER  " & ChrW(&H2014) & "  " & market & "  |  " & Sector & "  |  " & Format(Now, "yyyy/mm/dd hh:mm")
        .Font.Color = RR4_ACCENT
        .Font.Bold = True
        .Font.Size = 13
        .HorizontalAlignment = xlLeft
    End With

    Dim hdrs As Variant
    hdrs = Array("TICKER", "COMPANY", "PRICE", "1Y HIGH%", "180D CHG%", _
                 "R" & NB_SHORT, "R" & NB_LONG, "TREND", "SECTOR")
    Dim ci As Integer
    For ci = 0 To UBound(hdrs)
        With wsRes.cells(SCAN_HDR_ROW, SC_TICKER + ci)
            .Value = hdrs(ci)
            .Font.Color = RR4_ACCENT
            .Font.Bold = True
            .Interior.Color = RGB(10, 10, 10)
            .HorizontalAlignment = xlCenter
        End With
    Next ci
    wsRes.rows(SCAN_HDR_ROW).RowHeight = SCAN_ROW_H
    With wsRes.Range(wsRes.cells(SCAN_HDR_ROW, SC_TICKER), wsRes.cells(SCAN_HDR_ROW, SCAN_LAST_COL)).Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Color = RR4_LINE
        .Weight = xlThin
    End With

    Dim nSlots As Long: nSlots = lastIdx - LBound(tickerList) + 1
    Dim rowNum As Long: rowNum = SCAN_FIRST_ROW
    Dim i As Long
    For i = LBound(tickerList) To lastIdx
        Dim tkr As String: tkr = CStr(tickerList(i))
        Application.StatusBar = "Scanning [" & (i - LBound(tickerList) + 1) & "/" & nSlots & "] " & tkr
        DoEvents

        Dim sym As String, nm As String, dts() As Double, cls() As Double, ok As Boolean
        ok = FetchDaily2y(tkr, sym, nm, dts, cls)

        Dim rowBg As Long: rowBg = IIf((rowNum Mod 2) = 1, RGB(10, 10, 10), RGB(22, 22, 22))
        With wsRes.Range(wsRes.cells(rowNum, SC_TICKER), wsRes.cells(rowNum, SCAN_LAST_COL))
            .Interior.Color = rowBg
            .HorizontalAlignment = xlCenter
            .Font.Name = "Consolas"
            .Font.Size = 9
            .VerticalAlignment = xlCenter
        End With
        wsRes.rows(rowNum).RowHeight = SCAN_ROW_H

        With wsRes.cells(rowNum, SC_TICKER)
            .NumberFormat = "@"
            .Value = sym
            .Font.Color = RR4_ACCENT
            .Font.Bold = True
        End With
        With wsRes.cells(rowNum, SC_COMPANY)
            .Value = nm
            .HorizontalAlignment = xlLeft
            .Font.Color = RGB(210, 210, 210)
        End With
        With wsRes.cells(rowNum, SC_SECTOR)
            If IsArray(tickerGroups) Then .Value = tickerGroups(i) Else .Value = Sector
            .HorizontalAlignment = xlLeft
            .Font.Color = RGB(150, 150, 150)
        End With

        If ok Then
            Dim tech As TechIndicators
            tech = ScanTech(cls)
            wsRes.cells(rowNum, SC_PRICE).Value = tech.price
            wsRes.cells(rowNum, SC_PRICE).NumberFormat = "#,##0.00"
            WriteColoredPct wsRes.cells(rowNum, SC_HIGHDIST), tech.DistFromHigh, "0.00%"
            WriteColoredPct wsRes.cells(rowNum, SC_CHG180), tech.Change180D, "0.00%"
            Call WriteNormBias(wsRes.cells(rowNum, SC_R20), NormBias(cls, NB_SHORT, NB_LOOKBACK))
            Call WriteNormBias(wsRes.cells(rowNum, SC_R55), NormBias(cls, NB_LONG, NB_LOOKBACK))
            Call ApplyScanTrend(wsRes.cells(rowNum, SC_TREND), tech)

        Else
            wsRes.cells(rowNum, SC_PRICE).Value = "ERR"
            wsRes.cells(rowNum, SC_PRICE).Font.Color = RGB(255, 102, 0)
            wsRes.cells(rowNum, SC_TREND).Value = "N/A"
        End If
        rowNum = rowNum + 1
    Next i

    ' Sort by R55, strongest first (blank R55 = failed fetch, always last)
    Dim nShown As Long: nShown = rowNum - SCAN_FIRST_ROW
    If nShown > 1 Then
        wsRes.Range(wsRes.cells(SCAN_FIRST_ROW, SC_TICKER), wsRes.cells(rowNum - 1, SCAN_LAST_COL)).Sort _
            Key1:=wsRes.cells(SCAN_FIRST_ROW, SC_R55), Order1:=xlDescending, Header:=xlNo
    End If
    ' close the table off with the same divider the header row has
    If nShown > 0 Then
        With wsRes.Range(wsRes.cells(rowNum - 1, SC_TICKER), wsRes.cells(rowNum - 1, SCAN_LAST_COL)).Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .Color = RR4_LINE
            .Weight = xlThin
        End With
    End If

    Call DrawScanSummary(wsRes, rowNum + 1, market, Sector, tickerList, nShown)

    wsRes.Range(wsRes.cells(SCAN_HDR_ROW, SC_TICKER), wsRes.cells(SCAN_LAST_ROW, SCAN_LAST_COL)).Columns.AutoFit
    wsRes.Columns(SC_TICKER).ColumnWidth = 10
    wsRes.Columns(SC_COMPANY).ColumnWidth = 28
    wsRes.Columns(SC_SECTOR).ColumnWidth = 14

    ' Refresh the group panel so its highlight follows this scan
    Call RebuildGroupDb
    Call DrawCrHeader(wsRes)

    Application.ScreenUpdating = True
    Call NavNotify("SCAN done " & Format(Now, "hh:mm:ss") & " - " & nShown & " tickers" & _
                   IIf(total > nShown, " (first " & nShown & " of " & total & ", table full)", ""))
End Sub

' ================================================================
'  Summary bar
' ================================================================
Private Sub DrawScanSummary(ws As Worksheet, sumRow As Long, _
                             market As String, Sector As String, _
                             tickers As Variant, count As Long)
    Dim cntBull As Long, cntBear As Long, cntNeut As Long
    Dim r As Long
    For r = SCAN_FIRST_ROW To SCAN_FIRST_ROW + count - 1
        Dim tv As String: tv = ws.cells(r, SC_TREND).Value
        If InStr(tv, "BULL") > 0 Then cntBull = cntBull + 1
        If InStr(tv, "BEAR") > 0 Then cntBear = cntBear + 1
        If InStr(tv, "NEUT") > 0 Then cntNeut = cntNeut + 1
    Next r

    ws.rows(sumRow).RowHeight = SCAN_ROW_H
    With ws.Range(ws.cells(sumRow, SC_TICKER), ws.cells(sumRow, SCAN_LAST_COL))
        .Interior.Color = RGB(18, 18, 18)
        .Font.Bold = True
        .VerticalAlignment = xlCenter
    End With
    ws.cells(sumRow, 3).Value = "SUMMARY": ws.cells(sumRow, 3).Font.Color = RR4_ACCENT
    ws.cells(sumRow, 4).Value = "Total: " & count: ws.cells(sumRow, 4).Font.Color = RGB(200, 200, 200)
    ws.cells(sumRow, 5).Value = "BULLISH: " & cntBull: ws.cells(sumRow, 5).Font.Color = RGB(255, 80, 80)
    ws.cells(sumRow, 6).Value = "BEARISH: " & cntBear: ws.cells(sumRow, 6).Font.Color = RGB(0, 210, 100)
    ws.cells(sumRow, 7).Value = "NEUTRAL: " & cntNeut: ws.cells(sumRow, 7).Font.Color = RGB(200, 200, 200)
    If count > 0 Then
        Dim bp As Double: bp = cntBull / count
        ws.cells(sumRow, 8).Value = bp
        ws.cells(sumRow, 8).NumberFormat = "0%"
        ws.cells(sumRow, 8).Font.Color = IIf(bp > 0.5, RGB(255, 80, 80), RGB(0, 210, 100))
    End If
End Sub

' ================================================================
'  Cell helpers
' ================================================================
' Taiwan convention, same as R20/R55 and TREND: up = red, down = green, 0 = gray
Private Sub WriteColoredPct(cell As Range, val As Double, fmt As String)
    cell.Value = val
    cell.NumberFormat = fmt
    If val > 0 Then
        cell.Font.Color = RGB(255, 80, 80)
    ElseIf val < 0 Then
        cell.Font.Color = RGB(0, 210, 100)
    Else
        cell.Font.Color = RGB(150, 150, 150)
    End If
End Sub

Private Sub ApplyScanTrend(cell As Range, tech As TechIndicators)
    cell.Font.Bold = True: cell.HorizontalAlignment = xlCenter
    If tech.Bias5 > 0 And tech.Bias20 > 0 And tech.Bias60 > 0 Then
        cell.Value = "STRONG BULL": cell.Font.Color = RGB(255, 60, 60): cell.Interior.Color = RGB(40, 0, 0)
    ElseIf tech.Bias20 > 0 And tech.Bias60 > 0 Then
        cell.Value = "BULLISH": cell.Font.Color = RGB(255, 120, 120): cell.Interior.Color = RGB(25, 0, 0)
    ElseIf tech.Bias5 < 0 And tech.Bias20 < 0 And tech.Bias60 < 0 Then
        cell.Value = "STRONG BEAR": cell.Font.Color = RGB(0, 255, 100): cell.Interior.Color = RGB(0, 35, 0)
    ElseIf tech.Bias20 < 0 And tech.Bias60 < 0 Then
        cell.Value = "BEARISH": cell.Font.Color = RGB(100, 220, 100): cell.Interior.Color = RGB(0, 20, 0)
    ElseIf tech.Bias20 > 0 Then
        cell.Value = "WEAK BULL": cell.Font.Color = RR4_ACCENT
    ElseIf tech.Bias20 < 0 Then
        cell.Value = "WEAK BEAR": cell.Font.Color = RGB(150, 200, 150)
    Else
        cell.Value = "NEUTRAL": cell.Font.Color = RGB(150, 150, 150)
    End If
End Sub

' ================================================================
'  Export helpers
' ================================================================
Sub ExportTopToWatchlist()
    Dim wsRes As Worksheet
    On Error Resume Next: Set wsRes = ThisWorkbook.Sheets("Company research"): On Error GoTo 0
    If wsRes Is Nothing Then MsgBox "Run a scan first.", vbExclamation: Exit Sub

    Dim topN As Long
    topN = CLng(InputBox("Export top N (table is sorted by R" & NB_LONG & "):", "Export", "10"))
    If topN <= 0 Then Exit Sub

    Dim wsWatch As Worksheet
    On Error Resume Next: Set wsWatch = ThisWorkbook.Sheets("Watchlist"): On Error GoTo 0
    If wsWatch Is Nothing Then
        Set wsWatch = ThisWorkbook.Sheets.Add(After:=wsRes)
        wsWatch.Name = "Watchlist"
    End If

    wsWatch.cells.Clear
    With wsWatch.cells: .Interior.Color = RGB(0, 0, 0): .Font.Color = RGB(200, 200, 200): .Font.Name = "Consolas": .Font.Size = 10: End With
    wsWatch.cells(1, 1).Value = "WATCHLIST " & ChrW(&H2014) & " " & Format(Now, "yyyy/mm/dd hh:mm")
    wsWatch.cells(1, 1).Font.Color = RR4_ACCENT: wsWatch.cells(1, 1).Font.Bold = True

    ' destination column <- scan table column
    Dim headers As Variant: headers = Array("TICKER", "COMPANY", "R" & NB_SHORT, "R" & NB_LONG, "TREND", "180D CHG%")
    Dim srcCols As Variant: srcCols = Array(SC_TICKER, SC_COMPANY, SC_R20, SC_R55, SC_TREND, SC_CHG180)
    Dim ci As Integer
    For ci = 0 To UBound(headers)
        With wsWatch.cells(2, ci + 1): .Value = headers(ci): .Font.Color = RR4_ACCENT: .Font.Bold = True: .Interior.Color = RGB(10, 10, 10): End With
    Next ci

    Dim cc As Long: cc = WorksheetFunction.Min(topN, ScanTickerRows(wsRes))
    Dim r As Long, wr As Long: wr = 3
    For r = SCAN_FIRST_ROW To SCAN_FIRST_ROW + cc - 1
        wsWatch.Range(wsWatch.cells(wr, 1), wsWatch.cells(wr, UBound(headers) + 1)).Interior.Color = IIf((wr Mod 2) = 1, RGB(12, 12, 12), RGB(20, 20, 20))
        For ci = 0 To UBound(srcCols)
            wsWatch.cells(wr, ci + 1).Value = wsRes.cells(r, srcCols(ci)).Value
            wsWatch.cells(wr, ci + 1).NumberFormat = wsRes.cells(r, srcCols(ci)).NumberFormat
        Next ci
        wsWatch.cells(wr, 1).Font.Color = RR4_ACCENT: wsWatch.cells(wr, 1).Font.Bold = True
        wr = wr + 1
    Next r
    wsWatch.Columns("A:F").AutoFit: wsWatch.Columns("B").ColumnWidth = 28
    wsWatch.Activate
    MsgBox "Top " & cc & " exported!", vbInformation
End Sub

Sub ExportScanCSV()
    Dim wsRes As Worksheet
    On Error Resume Next: Set wsRes = ThisWorkbook.Sheets("Company research"): On Error GoTo 0
    If wsRes Is Nothing Then Exit Sub
    Dim path As String: path = ThisWorkbook.path & "\Scan_" & Format(Now, "yyyymmdd_hhmmss") & ".csv"
    Dim fNum As Integer: fNum = FreeFile
    Open path For Output As #fNum
    Dim LR As Long: LR = ScanLastRow(wsRes)
    Dim r As Long, c As Integer, line As String
    For r = SCAN_HDR_ROW To LR
        line = ""
        For c = SC_TICKER To SCAN_LAST_COL
            Dim v As String: v = CStr(wsRes.cells(r, c).Value)
            If InStr(v, ",") > 0 Then v = """" & v & """"
            line = line & IIf(c > 3, ",", "") & v
        Next c
        Print #fNum, line
    Next r
    Close #fNum
    MsgBox "Exported: " & path, vbInformation
End Sub

' ================================================================
'  GROUP TABLE LOOKUPS (sheet Groups / tblGroups)
' ================================================================
Private Function GroupTable() As ListObject
    On Error Resume Next
    Set GroupTable = ThisWorkbook.Worksheets(GROUPS_SHEET).ListObjects(GROUPS_TABLE)
    On Error GoTo 0
End Function

Private Function GroupData() As Variant
    Dim lo As ListObject: Set lo = GroupTable()
    If lo Is Nothing Then Exit Function
    If lo.DataBodyRange Is Nothing Then Exit Function
    GroupData = lo.DataBodyRange.Value
End Function

' Matching key: case-insensitive, half- and full-width spaces ignored
Private Function GroupKey(ByVal s As String) As String
    s = Replace(Replace(s, ChrW(&H3000), ""), " ", "")
    GroupKey = LCase$(s)
End Function

' Tickers of one group, de-duplicated, in table order. market "" = any.
Function GetSectorTickers(market As String, Sector As String) As Variant
    Dim data As Variant: data = GroupData()
    If IsEmpty(data) Then Exit Function

    Dim key As String: key = GroupKey(Sector)
    Dim seen As Object: Set seen = CreateObject("Scripting.Dictionary")
    seen.CompareMode = vbTextCompare
    Dim r As Long, tk As String
    For r = 1 To UBound(data, 1)
        If GroupKey(CStr(data(r, 2))) = key Then
            If Len(market) = 0 Or StrComp(CStr(data(r, 1)), market, vbTextCompare) = 0 Then
                tk = UCase$(Trim$(CStr(data(r, 3))))
                If Len(tk) > 0 Then
                    If Not seen.Exists(tk) Then seen.Add tk, True
                End If
            End If
        End If
    Next r
    If seen.count > 0 Then GetSectorTickers = seen.keys
End Function

' Distinct group names in table order. market "" = all markets.
Function GetGroupNames(market As String) As Variant
    Dim data As Variant: data = GroupData()
    If IsEmpty(data) Then Exit Function

    Dim seen As Object: Set seen = CreateObject("Scripting.Dictionary")
    seen.CompareMode = vbTextCompare
    Dim r As Long, g As String
    For r = 1 To UBound(data, 1)
        g = Trim$(CStr(data(r, 2)))
        If Len(g) > 0 Then
            If Len(market) = 0 Or StrComp(CStr(data(r, 1)), market, vbTextCompare) = 0 Then
                If Not seen.Exists(g) Then seen.Add g, True
            End If
        End If
    Next r
    If seen.count > 0 Then GetGroupNames = seen.keys
End Function

' Market of a group = market of its first row
Private Function GetGroupMarket(groupName As String) As String
    Dim data As Variant: data = GroupData()
    If IsEmpty(data) Then Exit Function
    Dim r As Long
    For r = 1 To UBound(data, 1)
        If GroupKey(CStr(data(r, 2))) = GroupKey(groupName) Then
            GetGroupMarket = UCase$(Trim$(CStr(data(r, 1))))
            Exit Function
        End If
    Next r
End Function

' Rewrites the distinct group names into column GROUP_LIST_COL of the Groups
' sheet and points the GroupList name at them (keyword-cell dropdown source).
Sub RebuildGroupList()
    Dim lo As ListObject: Set lo = GroupTable()
    If lo Is Nothing Then Exit Sub
    Dim ws As Worksheet: Set ws = lo.Parent

    ws.Range(ws.cells(2, GROUP_LIST_COL), ws.cells(ws.rows.count, GROUP_LIST_COL)).ClearContents
    ws.cells(1, GROUP_LIST_COL).Value = "GroupList (auto)"

    Dim grpNames As Variant: grpNames = GetGroupNames("")
    If IsEmpty(grpNames) Then Exit Sub

    Dim i As Long
    For i = 0 To UBound(grpNames)
        ws.cells(2 + i, GROUP_LIST_COL).Value = grpNames(i)
    Next i
    ThisWorkbook.Names.Add Name:=GROUP_LIST_NAME, RefersTo:="='" & ws.Name & "'!" & _
        ws.Range(ws.cells(2, GROUP_LIST_COL), ws.cells(2 + UBound(grpNames), GROUP_LIST_COL)).Address
End Sub

' Keyword cell on "Company research": one or more group names separated by
' commas (the full-width comma and the ideographic comma are accepted too).
' Every name must match a group exactly (GroupKey rules). If any name does
' not, the groups containing it are listed as candidates and nothing runs.
Sub RunGroupKeyword(ByVal raw As String)
    Dim wsRes As Worksheet
    Set wsRes = ThisWorkbook.Worksheets("Company research")

    raw = Replace(Replace(raw, ChrW(&HFF0C), ","), ChrW(&H3001), ",")
    If Len(Trim$(raw)) = 0 Then Exit Sub

    Dim allNames As Variant: allNames = GetGroupNames("")
    If IsEmpty(allNames) Then Call NavNotify("Groups table missing or empty.", True): Exit Sub

    Dim picked As Object: Set picked = CreateObject("Scripting.Dictionary")
    Dim miss As String, cand As String
    Dim parts As Variant: parts = Split(raw, ",")
    Dim i As Long, j As Long, t As String, hit As Boolean
    For i = 0 To UBound(parts)
        t = Trim$(parts(i))
        If Len(t) > 0 Then
            hit = False
            For j = 0 To UBound(allNames)
                If GroupKey(CStr(allNames(j))) = GroupKey(t) Then
                    hit = True
                    If Not picked.Exists(allNames(j)) Then picked.Add allNames(j), True
                    Exit For
                End If
            Next j
            If Not hit Then
                miss = miss & IIf(Len(miss) > 0, ", ", "") & t
                For j = 0 To UBound(allNames)
                    If InStr(1, GroupKey(CStr(allNames(j))), GroupKey(t), vbTextCompare) > 0 Then
                        cand = cand & IIf(Len(cand) > 0, ", ", "") & allNames(j)
                    End If
                Next j
            End If
        End If
    Next i

    If Len(miss) > 0 Then
        Call NavNotify("No exact group [" & miss & "] - " & _
                   IIf(Len(cand) > 0, "candidates: " & cand, "no group name contains that text") & _
                   " - nothing scanned", True)
        Exit Sub
    End If
    If picked.count = 0 Then Exit Sub

    ' Merge tickers across the picked groups; the first group a ticker
    ' appears in labels its row.
    Dim tick As Object: Set tick = CreateObject("Scripting.Dictionary")
    tick.CompareMode = vbTextCompare
    Dim mkt As String, gm As String, g As Variant, tks As Variant, k As Long
    For Each g In picked.keys
        gm = GetGroupMarket(CStr(g))
        If Len(mkt) = 0 Then
            mkt = gm
        ElseIf StrComp(mkt, gm, vbTextCompare) <> 0 Then
            mkt = "MIX"
        End If
        tks = GetSectorTickers("", CStr(g))
        If Not IsEmpty(tks) Then
            For k = 0 To UBound(tks)
                If Not tick.Exists(tks(k)) Then tick.Add tks(k), CStr(g)
            Next k
        End If
    Next g
    If tick.count = 0 Then Call NavNotify("Group has no tickers.", True): Exit Sub

    Dim scanLabel As String: scanLabel = Join(picked.keys, " + ")
    Dim n As Long: n = tick.count

    Call ScanTickers(mkt, scanLabel, tick.keys, tick.items)
End Sub

' Last used scan row in column C, never looking below SCAN_LAST_ROW (the
' deep-dive output further down the sheet also uses column C).
Private Function ScanLastRow(ws As Worksheet) As Long
    If Len(CStr(ws.cells(SCAN_LAST_ROW, 3).Value)) > 0 Then
        ScanLastRow = SCAN_LAST_ROW
    Else
        ScanLastRow = ws.cells(SCAN_LAST_ROW, 3).End(xlUp).row
    End If
End Function

' ================================================================
'  GROUP DATABASE PANEL -- "Company research", right of the scanner
'  Block 0 = every non built-in source (H2 ...), block 1 = built-in
'  US, block 2 = built-in TW, block 3 = built-in TW that overflows
'  block 2. One row per group: name + ticker count. v4 dropped the
'  third column (the comma-joined ticker list): it was always clipped
'  by the fixed column width and it made the panel twice as wide as
'  the scan results it sits next to.
'  Rebuilt after every tblGroups edit and every scan. Groups present
'  in the current scan's SECTOR column are highlighted.
' ================================================================

' Source value the migration gave the original sectors ("built-in")
Private Function BuiltinSource() As String
    BuiltinSource = ChrW(&H5167) & ChrW(&H5EFA)
End Function

Private Function DbBlockOf(ByVal mkt As String, ByVal src As String) As Long
    If StrComp(Trim$(src), BuiltinSource(), vbBinaryCompare) <> 0 Then
        DbBlockOf = 0
    ElseIf StrComp(Trim$(mkt), "US", vbTextCompare) = 0 Then
        DbBlockOf = 1
    Else
        DbBlockOf = 2
    End If
End Function

Sub RebuildGroupDb()
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("Company research")
    On Error GoTo 0
    If ws Is Nothing Then Exit Sub

    Dim lastCol As Long: lastCol = DB_FIRST_COL + DB_BLOCK_COLS * DB_BLOCKS - 2
    ' start one column early so the spacer is cleared and painted too
    With ws.Range(ws.cells(1, DB_FIRST_COL - 1), ws.cells(SCAN_LAST_ROW, lastCol))
        .Clear
        .Interior.Color = RGB(0, 0, 0)
        .Font.Name = "Consolas"
        .Font.Size = 9
        .Font.Color = RGB(200, 200, 200)
        .WrapText = False
        .NumberFormat = "@"
    End With

    ws.Columns(DB_FIRST_COL - 1).ColumnWidth = 3      ' the spacer

    ' Groups visible in the current scan (SECTOR column) get highlighted
    Dim active As Object: Set active = CreateObject("Scripting.Dictionary")
    active.CompareMode = vbTextCompare
    Dim r As Long, v As String
    For r = SCAN_FIRST_ROW To SCAN_LAST_ROW
        v = CStr(ws.cells(r, SC_TICKER).Value)
        If Len(v) > 0 And v <> "SUMMARY" Then
            v = GroupKey(CStr(ws.cells(r, SC_SECTOR).Value))
            If Len(v) > 0 Then
                If Not active.Exists(v) Then active.Add v, True
            End If
        End If
    Next r

    ' Aggregate tblGroups rows into one entry per group, table order
    Dim data As Variant: data = GroupData()
    Dim idx As Object: Set idx = CreateObject("Scripting.Dictionary")
    idx.CompareMode = vbTextCompare
    Dim gName() As String, gBlock() As Long, gCount() As Long, gTick() As String
    Dim n As Long, g As String, tk As String, k As Long
    If Not IsEmpty(data) Then
        ReDim gName(1 To UBound(data, 1)): ReDim gBlock(1 To UBound(data, 1))
        ReDim gCount(1 To UBound(data, 1)): ReDim gTick(1 To UBound(data, 1))
        For r = 1 To UBound(data, 1)
            g = Trim$(CStr(data(r, 2)))
            tk = UCase$(Trim$(CStr(data(r, 3))))
            If Len(g) > 0 Then
                If Not idx.Exists(g) Then
                    n = n + 1
                    idx.Add g, n
                    gName(n) = g
                    gBlock(n) = DbBlockOf(CStr(data(r, 1)), CStr(data(r, 4)))
                End If
                k = idx(g)
                If Len(tk) > 0 Then
                    If InStr(1, "," & gTick(k) & ",", "," & tk & ",", vbTextCompare) = 0 Then
                        gTick(k) = gTick(k) & IIf(Len(gTick(k)) > 0, ",", "") & tk
                        gCount(k) = gCount(k) + 1
                    End If
                End If
            End If
        Next r
    End If

    Dim titles As Variant: titles = Array("H2 / CUSTOM", "US", "TW", "TW (cont.)")
    Dim nameWidths As Variant: nameWidths = Array(18, 36, 32, 32)
    Dim maxRows As Long: maxRows = SCAN_LAST_ROW - SCAN_FIRST_ROW + 1
    Dim b As Long, c0 As Long, rw As Long, shown As Long, total As Long, hot As Boolean

    ' TW groups past the first maxRows continue in the last block
    Dim twSeen As Long
    For k = 1 To n
        If gBlock(k) = 2 Then
            twSeen = twSeen + 1
            If twSeen > maxRows Then gBlock(k) = 3
        End If
    Next k
    For b = 0 To DB_BLOCKS - 1
        c0 = DB_FIRST_COL + b * DB_BLOCK_COLS
        total = 0
        For k = 1 To n
            If gBlock(k) = b Then total = total + 1
        Next k

        With ws.cells(SCAN_TITLE_ROW, c0)
            .Value = titles(b) & "  (" & total & " groups)"
            .Font.Color = RR4_ACCENT
            .Font.Bold = True
        End With
        ws.cells(SCAN_HDR_ROW, c0).Value = "GROUP"
        ws.cells(SCAN_HDR_ROW, c0 + 1).Value = "N"
        With ws.Range(ws.cells(SCAN_HDR_ROW, c0), ws.cells(SCAN_HDR_ROW, c0 + 1))
            .Font.Color = RR4_ACCENT
            .Font.Bold = True
            .Interior.Color = RGB(10, 10, 10)
            .Borders(xlEdgeBottom).LineStyle = xlContinuous
            .Borders(xlEdgeBottom).Color = RR4_LINE
        End With
        ws.Columns(c0).ColumnWidth = nameWidths(b)
        ws.Columns(c0 + 1).ColumnWidth = 5
        If b < DB_BLOCKS - 1 Then ws.Columns(c0 + 2).ColumnWidth = 3

        rw = SCAN_FIRST_ROW: shown = 0
        For k = 1 To n
            If gBlock(k) = b Then
                If total > maxRows And shown = maxRows - 1 Then
                    ws.cells(rw, c0).Value = "+" & (total - shown) & " more -> sheet Groups"
                    ws.cells(rw, c0).Font.Color = RGB(150, 150, 150)
                    Exit For
                End If
                hot = active.Exists(GroupKey(gName(k)))
                With ws.Range(ws.cells(rw, c0), ws.cells(rw, c0 + 1))
                    .Interior.Color = IIf(hot, RGB(55, 28, 0), IIf((rw Mod 2) = 1, RGB(10, 10, 10), RGB(22, 22, 22)))
                    .VerticalAlignment = xlCenter
                End With
                With ws.cells(rw, c0)
                    .Value = gName(k)
                    .Font.Color = IIf(hot, RGB(255, 150, 40), RGB(200, 200, 200))
                    .Font.Bold = hot
                End With
                With ws.cells(rw, c0 + 1)
                    .NumberFormat = "0"
                    .Value = gCount(k)
                    .HorizontalAlignment = xlCenter
                    ' orange = more tickers than one scan can show
                    .Font.Color = IIf(gCount(k) > SCAN_MAX_TICKERS, RR4_ACCENT, RGB(140, 140, 140))
                End With
                ws.rows(rw).RowHeight = SCAN_ROW_H
                rw = rw + 1: shown = shown + 1
            End If
        Next k
    Next b
End Sub

' Double-click on the group database panel: put that row's group into the
' keyword cell, whose Worksheet_Change runs the scan. Returns True when the
' click was on a group (the caller then cancels Excel's edit mode).
Function RunGroupFromDb(ByVal Target As Range) As Boolean
    Dim ws As Worksheet: Set ws = Target.Worksheet
    Dim r As Long, c As Long, offs As Long, g As String
    r = Target.row: c = Target.Column
    If r < SCAN_FIRST_ROW Or r > SCAN_LAST_ROW Then Exit Function
    If c < DB_FIRST_COL Or c > DB_FIRST_COL + DB_BLOCK_COLS * DB_BLOCKS - 2 Then Exit Function
    offs = (c - DB_FIRST_COL) Mod DB_BLOCK_COLS
    If offs = DB_BLOCK_COLS - 1 Then Exit Function              ' gap column
    g = Trim$(CStr(ws.cells(r, c - offs).Value))
    If Len(g) = 0 Or Left$(g, 1) = "+" Then Exit Function      ' empty / "+N more"
    RunGroupFromDb = True
    ws.Range(GROUP_INPUT_CELL).Value = g
End Function

' ================================================================
'  SCAN DATA -- one Yahoo v8 chart call (2y daily) per ticker
' ================================================================
Private Function ScanHttpGet(ByVal url As String) As String
    Dim http As Object
    On Error Resume Next
    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
    If http Is Nothing Then Set http = CreateObject("MSXML2.XMLHTTP")
    http.Open "GET", url, False
    http.SetRequestHeader "User-Agent", "Mozilla/5.0"
    http.Send
    If http.status = 200 Then ScanHttpGet = http.ResponseText
    On Error GoTo 0
End Function

' "key":123.45 (plain number, as in the chart endpoint's meta block)
Private Function ScanMetaNum(json As String, key As String) As Double
    Dim p As Long: p = InStr(json, """" & key & """:")
    If p = 0 Then Exit Function
    p = p + Len(key) + 3
    Dim e As Long: e = p
    Do While e <= Len(json)
        If InStr(",}]", Mid$(json, e, 1)) > 0 Then Exit Do
        e = e + 1
    Loop
    Dim s As String: s = Trim$(Mid$(json, p, e - p))
    If IsNumeric(s) Then ScanMetaNum = val(s)
End Function

' "key":"text"
Private Function ScanMetaStr(json As String, key As String) As String
    Dim tag As String: tag = """" & key & """:"""
    Dim p As Long: p = InStr(json, tag)
    If p = 0 Then Exit Function
    p = p + Len(tag)
    Dim e As Long: e = InStr(p, json, """")
    If e > p Then ScanMetaStr = Mid$(json, p, e - p)
End Function

' Mirrors Attach.GetTechData + GetCompanyName in a single request: numeric
' tickers try .TW then .TWO, nulls are skipped, and the last close is replaced
' by meta regularMarketPrice. dts() are Excel dates (UTC day of each bar).
Private Function FetchDaily2y(ByVal tkr As String, ByRef symOut As String, ByRef nameOut As String, _
                              ByRef dts() As Double, ByRef cls() As Double) As Boolean
    Dim raw As String: raw = UCase$(Trim$(tkr))
    symOut = raw
    If InStr(symOut, ".TW") = 0 And IsNumeric(symOut) Then symOut = symOut & ".TW"
    Const BASE_URL As String = "https://query1.finance.yahoo.com/v8/finance/chart/"
    Dim resp As String
    resp = ScanHttpGet(BASE_URL & symOut & "?interval=1d&range=2y")
    If InStr(resp, """timestamp""") = 0 And IsNumeric(raw) Then
        symOut = raw & ".TWO"
        resp = ScanHttpGet(BASE_URL & symOut & "?interval=1d&range=2y")
    End If

    nameOut = ScanMetaStr(resp, "longName")
    If Len(nameOut) = 0 Then nameOut = ScanMetaStr(resp, "shortName")
    If Len(nameOut) = 0 Then nameOut = symOut

    Dim tsS As Long, tsE As Long, qPos As Long, cS As Long, cE As Long
    tsS = InStr(resp, """timestamp"":[")
    If tsS = 0 Then Exit Function
    tsS = tsS + Len("""timestamp"":[")
    tsE = InStr(tsS, resp, "]")
    qPos = InStr(resp, """quote"":[{")
    If qPos = 0 Then Exit Function
    cS = InStr(qPos, resp, """close"":[")
    If cS = 0 Then Exit Function
    cS = cS + Len("""close"":[")
    cE = InStr(cS, resp, "]")
    If tsE <= tsS Or cE <= cS Then Exit Function

    Dim tsArr() As String, clArr() As String
    tsArr = Split(Mid$(resp, tsS, tsE - tsS), ",")
    clArr = Split(Mid$(resp, cS, cE - cS), ",")
    Dim nMax As Long: nMax = UBound(tsArr)
    If UBound(clArr) < nMax Then nMax = UBound(clArr)
    ReDim dts(0 To nMax): ReDim cls(0 To nMax)

    Dim i As Long, cnt As Long
    For i = 0 To nMax
        If IsNumeric(clArr(i)) And IsNumeric(tsArr(i)) Then
            cls(cnt) = val(clArr(i))
            dts(cnt) = CDbl(DateSerial(1970, 1, 1)) + Int(val(tsArr(i)) / 86400#)
            cnt = cnt + 1
        End If
    Next i
    If cnt = 0 Then Exit Function
    ReDim Preserve dts(0 To cnt - 1): ReDim Preserve cls(0 To cnt - 1)

    Dim mp As Double: mp = ScanMetaNum(resp, "regularMarketPrice")
    If mp > 0 Then cls(cnt - 1) = mp
    FetchDaily2y = True
End Function

' TREND / 180D follow Attach.GetTechData (180D = 126 trading days back).
' 1Y HIGH% is measured against the max close of the last YEAR_BARS bars --
' the same window as the chart wall and its dashed high line.
Private Function ScanTech(cls() As Double) As TechIndicators
    Dim t As TechIndicators
    Dim n As Long: n = UBound(cls) - LBound(cls) + 1
    Dim cur As Double: cur = cls(UBound(cls))
    Dim mx As Double, i As Long, i0 As Long
    i0 = UBound(cls) - YEAR_BARS + 1
    If i0 < LBound(cls) Then i0 = LBound(cls)
    For i = i0 To UBound(cls)
        If cls(i) > mx Then mx = cls(i)
    Next i
    t.price = cur
    If n >= 2 Then
        If cls(UBound(cls) - 1) > 0 Then t.ChangePercent = (cur - cls(UBound(cls) - 1)) / cls(UBound(cls) - 1)
    End If
    If n >= 126 Then
        If cls(UBound(cls) - 125) > 0 Then t.Change180D = (cur - cls(UBound(cls) - 125)) / cls(UBound(cls) - 125)
    End If
    If mx > 0 Then t.DistFromHigh = (cur - mx) / mx
    t.Bias5 = CalculateBias(cls, 5, cur)
    t.Bias20 = CalculateBias(cls, 20, cur)
    t.Bias60 = CalculateBias(cls, 60, cur)
    t.Bias120 = CalculateBias(cls, 120, cur)
    t.Bias240 = CalculateBias(cls, 240, cur)
    t.Success = True
    ScanTech = t
End Function

' Latest normalised bias R (see the NB_* constants). EMA seeded with the
' first close, as TDX does. Empty when the series is too short.
Private Function NormBias(cls() As Double, ByVal n As Long, ByVal m As Long) As Variant
    Dim lo As Long: lo = LBound(cls)
    Dim cnt As Long: cnt = UBound(cls) - lo + 1
    If cnt <= n Then Exit Function

    Dim b() As Double: ReDim b(0 To cnt - 1)
    Dim k As Double: k = 2 / (n + 1)
    Dim ema As Double: ema = cls(lo)
    Dim i As Long
    For i = 0 To cnt - 1
        If i > 0 Then ema = (cls(lo + i) - ema) * k + ema
        If ema <> 0 Then b(i) = (cls(lo + i) - ema) / ema * 100
    Next i

    Dim j0 As Long: j0 = cnt - m
    If j0 < 0 Then j0 = 0
    Dim hh As Double, ll As Double
    hh = b(j0): ll = b(j0)
    For i = j0 To cnt - 1
        If b(i) > hh Then hh = b(i)
        If b(i) < ll Then ll = b(i)
    Next i

    Dim cur As Double: cur = b(cnt - 1)
    If cur >= 0 Then
        NormBias = cur / IIf(hh > 0.01, hh, 0.01) * 100
    Else
        NormBias = cur / IIf(Abs(ll) > 0.01, Abs(ll), 0.01) * 100
    End If
End Function

' R cell: number to 1 dp, coloured with the TREND column's four tiers
' (|R| >= NB_STRONG = the "strong" shades).
Private Sub WriteNormBias(cell As Range, v As Variant)
    cell.HorizontalAlignment = xlCenter
    If IsEmpty(v) Then Exit Sub                    ' blank sorts last
    Dim r As Double: r = v
    cell.Value = r
    cell.NumberFormat = "0.0"
    cell.Font.Bold = True
    If r >= NB_STRONG Then
        cell.Font.Color = RGB(255, 60, 60): cell.Interior.Color = RGB(40, 0, 0)
    ElseIf r > 0 Then
        cell.Font.Color = RGB(255, 120, 120): cell.Interior.Color = RGB(25, 0, 0)
    ElseIf r <= -NB_STRONG Then
        cell.Font.Color = RGB(0, 255, 100): cell.Interior.Color = RGB(0, 35, 0)
    ElseIf r < 0 Then
        cell.Font.Color = RGB(100, 220, 100): cell.Interior.Color = RGB(0, 20, 0)
    Else
        cell.Font.Color = RGB(150, 150, 150)
    End If
End Sub

' Rows of tickers currently in the scan table (stops at the SUMMARY line)
Private Function ScanTickerRows(ws As Worksheet) As Long
    Dim r As Long: r = SCAN_FIRST_ROW
    Do While r <= SCAN_LAST_ROW
        Dim v As String: v = CStr(ws.cells(r, SC_TICKER).Value)
        If Len(v) = 0 Or v = "SUMMARY" Then Exit Do
        r = r + 1
    Loop
    ScanTickerRows = r - SCAN_FIRST_ROW
End Function

' Drop every sparkline on the sheet (and any chart left from the old wall).
Private Sub ClearSparklines(ws As Worksheet)
    On Error Resume Next
    ws.cells.SparklineGroups.Clear
    Dim j As Long
    For j = ws.ChartObjects.count To 1 Step -1
        If Left$(ws.ChartObjects(j).Name, Len(CHART_PREFIX)) = CHART_PREFIX Then ws.ChartObjects(j).Delete
    Next j
    On Error GoTo 0
End Sub

' ================================================================
'  TITLE + INPUT STRIP (row 1-2) and the cleanup of earlier layouts
' ================================================================
Public Sub DrawCrHeader(ws As Worksheet)
    Dim prevEv As Boolean: prevEv = Application.EnableEvents
    Application.EnableEvents = False
    On Error GoTo Fin

    ' Keep whatever is typed in the three input cells - this runs at the END
    ' of a scan too, and wiping the strip there would throw away the group
    ' the user just scanned and the ticker they were about to dive into.
    ' Only once the strip IS the v4 strip, though: on the first run after the
    ' v3 layout, D2 / F2 still hold that version's "COMPANY" / "1Y HIGH%"
    ' header text, and keeping those would plant them in the input cells.
    Dim migrated As Boolean
    migrated = (CStr(ws.cells(CR_STRIP_ROW, 1).Value) = "GROUP <GO>")
    Dim keepG As String, keepT As String, keepM As String
    If migrated Then
        keepG = CStr(ws.Range(GROUP_INPUT_CELL).Value)
        keepT = CStr(ws.Range(CR_TICKER_INPUT).Value)
        keepM = CStr(ws.Range(CR_MARKET_INPUT).Value)
    End If

    ' Leftovers of the earlier layouts: the v2 merged cells / dropdown on A4
    ' and its status block, and the v3 strip that stood vertically in A1:B3
    ' (B3 is inside the cleared band below, B1 is cleared with the title).
    With ws.Range("A1:B33")
        .UnMerge
        .Validation.Delete
    End With
    With ws.Range(ws.cells(SCAN_TITLE_ROW, 1), ws.cells(SCAN_TITLE_ROW, 2))
        .Clear
        .Interior.Color = RGB(0, 0, 0)
    End With
    With ws.Range("A3:B33")
        .Clear
        .Interior.Color = RGB(0, 0, 0)
    End With
    ws.Range("A36:A400").EntireRow.RowHeight = ws.StandardHeight
    ws.Columns(1).ColumnWidth = 20
    ws.Columns(2).ColumnWidth = 24

    ' ---- page title ----
    With ws.cells(SCAN_TITLE_ROW, 1)
        .Value = "COMPANY RESEARCH"
        .Font.Name = "Consolas"
        .Font.Size = 13
        .Font.Bold = True
        .Font.Color = RR4_ACCENT
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
    End With
    ws.rows(SCAN_TITLE_ROW).RowHeight = 22

    ' ---- one horizontal input strip, label / cell / label / cell / ... ----
    With ws.Range(ws.cells(CR_STRIP_ROW, 1), ws.cells(CR_STRIP_ROW, SCAN_LAST_COL))
        .Clear
        .Interior.Color = RGB(0, 0, 0)
        .Font.Name = "Consolas"
        .Font.Size = 10
        .VerticalAlignment = xlCenter
    End With
    ws.rows(CR_STRIP_ROW).RowHeight = 20
    Call CrInputPair(ws, 1, "GROUP <GO>")
    Call CrInputPair(ws, 3, "TICKER")
    Call CrInputPair(ws, 5, "MKT")
    With ws.Range(ws.cells(CR_STRIP_ROW, 1), ws.cells(CR_STRIP_ROW, SCAN_LAST_COL)).Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Color = RR4_LINE
        .Weight = xlThin
    End With

    ws.Range(GROUP_INPUT_CELL).Value = keepG
    ws.Range(CR_TICKER_INPUT).Value = keepT
    ws.Range(CR_MARKET_INPUT).Value = keepM

    ' group dropdown on the GROUP cell (GroupList name, see RebuildGroupList)
    On Error Resume Next
    With ws.Range(GROUP_INPUT_CELL).Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertInformation, Formula1:="=" & GROUP_LIST_NAME
        .ShowError = False
    End With
    On Error GoTo Fin
Fin:
    Application.EnableEvents = prevEv
End Sub

' One "label | input" pair of the strip: label in labelCol, typed cell next to it.
Private Sub CrInputPair(ws As Worksheet, ByVal labelCol As Long, ByVal label As String)
    With ws.cells(CR_STRIP_ROW, labelCol)
        .Value = label
        .Font.Color = RR4_ACCENT
        .Font.Bold = True
        .HorizontalAlignment = xlRight
    End With
    With ws.cells(CR_STRIP_ROW, labelCol + 1)
        .NumberFormat = "@"
        .Interior.Color = RR4_INPUT_BG
        .Font.Color = RR4_INPUT_FG
        .Font.Bold = True
        .HorizontalAlignment = xlLeft
    End With
End Sub
