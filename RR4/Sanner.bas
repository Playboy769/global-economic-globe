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
Public Const GROUP_INPUT_CELL  As String = "A4"     ' merged A4:B4
Public Const GROUP_STATUS_CELL As String = "A5"     ' merged A5:B12

' The scanner owns rows 1..SCAN_LAST_ROW of C:K; row 34 is a spacer above
' the chart wall (CHART_FIRST_ROW). The summary line lands two rows below
' the last ticker, so 29 tickers (rows 3..31, summary on 33) is the most
' that fits.
Public Const SCAN_LAST_ROW     As Long = 33
Public Const SCAN_MAX_TICKERS  As Long = 29

' Group database panel beside the scanner (see RebuildGroupDb): three
' side-by-side blocks from column L, each GROUP | N | TICKERS + a gap
' column, confined to rows 1..SCAN_LAST_ROW because the deep-dive clears
' A36:BH400.
Public Const DB_FIRST_COL      As Long = 12         ' L, right after the scan table (C..K)
Public Const DB_BLOCK_COLS     As Long = 4
Public Const DB_BLOCKS         As Long = 3

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

' Price-chart wall under the scan table: one line chart per scanned ticker,
' 5 per row, rebased to 100 at the start of the window, all on one scale.
' Rows 35..106 are fixed at 15pt so the deep-dive (CompanyResearchSEC, input
' on row 108) always starts below the last chart row.
Public Const CHART_FIRST_ROW   As Long = 35
Public Const CHART_LAST_ROW    As Long = 106
Public Const CHART_DATA_SHEET  As String = "ScanPrices"
Private Const CHART_ROW_HEIGHT As Double = 15
Private Const CHART_PER_ROW    As Long = 5
Private Const CHART_W          As Double = 250
Private Const CHART_H          As Double = 168
Private Const CHART_GAP        As Double = 8
Private Const CHART_PREFIX     As String = "ScanChart_"
Private Const YEAR_BARS        As Long = 250        ' "1 year": chart window and 1Y HIGH%

' ================================================================
'  MAIN ENTRY
' ================================================================
' Scans tickerList into rows 3.. of C:K. tickerGroups (optional, indexed like
' tickerList) labels each row's SECTOR column when several groups are merged.
' At most SCAN_MAX_TICKERS are scanned so the table never reaches row 34.
' Rows are sorted by R55 (descending) and the chart wall below follows that order.
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
    Set scanArea = wsRes.Range(wsRes.cells(1, SC_TICKER), wsRes.cells(SCAN_LAST_ROW, SCAN_LAST_COL))
    With scanArea
        .Clear
        .Interior.Color = RGB(0, 0, 0)
        .Font.Color = RGB(200, 200, 200)
        .Font.Name = "Consolas"
        .Font.Size = 9
    End With

    With wsRes.cells(1, SC_TICKER)
        .Value = "MARKET SCANNER  " & ChrW(&H2014) & "  " & market & "  |  " & Sector & "  |  " & Format(Now, "yyyy/mm/dd hh:mm")
        .Font.Color = RGB(255, 192, 0)
        .Font.Bold = True
        .Font.Size = 13
    End With
    wsRes.Range(wsRes.cells(1, SC_TICKER), wsRes.cells(1, SCAN_LAST_COL)).Interior.Color = RGB(10, 10, 10)

    Dim hdrs As Variant
    hdrs = Array("TICKER", "COMPANY", "PRICE", "1Y HIGH%", "180D CHG%", _
                 "R" & NB_SHORT, "R" & NB_LONG, "TREND", "SECTOR")
    Dim ci As Integer
    For ci = 0 To UBound(hdrs)
        With wsRes.cells(2, SC_TICKER + ci)
            .Value = hdrs(ci)
            .Font.Color = RGB(255, 192, 0)
            .Font.Bold = True
            .Interior.Color = RGB(10, 10, 10)
            .HorizontalAlignment = xlCenter
        End With
    Next ci
    With wsRes.Range(wsRes.cells(2, SC_TICKER), wsRes.cells(2, SCAN_LAST_COL)).Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Color = RGB(255, 192, 0)
        .Weight = xlThin
    End With

    ' Daily series kept for the chart wall, keyed by the symbol shown in TICKER
    Dim serKey As Object: Set serKey = CreateObject("Scripting.Dictionary")
    serKey.CompareMode = vbTextCompare
    Dim nSlots As Long: nSlots = lastIdx - LBound(tickerList) + 1
    Dim serDt() As Variant, serCl() As Variant
    ReDim serDt(1 To nSlots): ReDim serCl(1 To nSlots)
    Dim nSer As Long

    Dim rowNum As Long: rowNum = 3
    Dim i As Long
    For i = LBound(tickerList) To lastIdx
        Dim tkr As String: tkr = CStr(tickerList(i))
        Application.StatusBar = "Scanning [" & (i - LBound(tickerList) + 1) & "/" & nSlots & "] " & tkr
        DoEvents

        Dim sym As String, nm As String, dts() As Double, cls() As Double, ok As Boolean
        ok = FetchDaily2y(tkr, sym, nm, dts, cls)

        Dim rowBg As Long: rowBg = IIf((rowNum Mod 2) = 1, RGB(12, 12, 12), RGB(20, 20, 20))
        With wsRes.Range(wsRes.cells(rowNum, SC_TICKER), wsRes.cells(rowNum, SCAN_LAST_COL))
            .Interior.Color = rowBg
            .HorizontalAlignment = xlCenter
            .Font.Name = "Consolas"
            .Font.Size = 9
        End With
        wsRes.rows(rowNum).RowHeight = 16

        With wsRes.cells(rowNum, SC_TICKER)
            .NumberFormat = "@"
            .Value = sym
            .Font.Color = RGB(255, 192, 0)
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
            WriteColoredPct wsRes.cells(rowNum, SC_HIGHDIST), tech.DistFromHigh, "0.00%", False
            WriteColoredPct wsRes.cells(rowNum, SC_CHG180), tech.Change180D, "0.00%", True
            Call WriteNormBias(wsRes.cells(rowNum, SC_R20), NormBias(cls, NB_SHORT, NB_LOOKBACK))
            Call WriteNormBias(wsRes.cells(rowNum, SC_R55), NormBias(cls, NB_LONG, NB_LOOKBACK))
            Call ApplyScanTrend(wsRes.cells(rowNum, SC_TREND), tech)

            If Not serKey.Exists(sym) Then
                nSer = nSer + 1
                serKey.Add sym, nSer
                serDt(nSer) = dts
                serCl(nSer) = cls
            End If
        Else
            wsRes.cells(rowNum, SC_PRICE).Value = "ERR"
            wsRes.cells(rowNum, SC_PRICE).Font.Color = RGB(255, 102, 0)
            wsRes.cells(rowNum, SC_TREND).Value = "N/A"
        End If
        rowNum = rowNum + 1
    Next i

    ' Sort by R55, strongest first (blank R55 = failed fetch, always last)
    If rowNum > 4 Then
        wsRes.Range(wsRes.cells(3, SC_TICKER), wsRes.cells(rowNum - 1, SCAN_LAST_COL)).Sort _
            Key1:=wsRes.cells(3, SC_R55), Order1:=xlDescending, Header:=xlNo
    End If

    Call DrawScanSummary(wsRes, rowNum + 1, market, Sector, tickerList, rowNum - 3)

    scanArea.Columns.AutoFit
    wsRes.Columns(SC_TICKER).ColumnWidth = 10
    wsRes.Columns(SC_COMPANY).ColumnWidth = 28
    wsRes.Columns(SC_SECTOR).ColumnWidth = 14

    ' Refresh the group panel so its highlight follows this scan
    Call RebuildGroupDb
    Call DrawChartWall(wsRes, serKey, serDt, serCl, rowNum - 3)

    Application.StatusBar = "Scan complete " & ChrW(&H2014) & " " & Format(Now, "hh:mm:ss")
    Application.ScreenUpdating = True
    ' An invisible (automation) instance would hang on a modal box
    If Application.Visible Then
        MsgBox "Scan complete! " & (rowNum - 3) & " stocks scanned." & _
               IIf(total > rowNum - 3, vbLf & "Truncated: " & total & " tickers in the list, only the first " & _
                   (rowNum - 3) & " fit in the scan table.", ""), vbInformation
    End If
End Sub

' ================================================================
'  Summary bar
' ================================================================
Private Sub DrawScanSummary(ws As Worksheet, sumRow As Long, _
                             market As String, Sector As String, _
                             tickers As Variant, count As Long)
    Dim cntBull As Long, cntBear As Long, cntNeut As Long
    Dim r As Long
    For r = 3 To 3 + count - 1
        Dim tv As String: tv = ws.cells(r, SC_TREND).Value
        If InStr(tv, "BULL") > 0 Then cntBull = cntBull + 1
        If InStr(tv, "BEAR") > 0 Then cntBear = cntBear + 1
        If InStr(tv, "NEUT") > 0 Then cntNeut = cntNeut + 1
    Next r

    With ws.Range(ws.cells(sumRow, SC_TICKER), ws.cells(sumRow, SCAN_LAST_COL))
        .Interior.Color = RGB(15, 15, 15)
        .Font.Bold = True
    End With
    ws.cells(sumRow, 3).Value = "SUMMARY": ws.cells(sumRow, 3).Font.Color = RGB(255, 192, 0)
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
Private Sub WriteColoredPct(cell As Range, val As Double, fmt As String, positiveIsGood As Boolean)
    cell.Value = val
    cell.NumberFormat = fmt
    
    ' If the value is exactly 0, color it Gray
    If val = 0 Then
        cell.Font.Color = RGB(150, 150, 150)
        Exit Sub
    End If
    
    ' Determine colors based on the positiveIsGood flag
    If positiveIsGood Then
        If val > 0 Then
            cell.Font.Color = RGB(0, 210, 100)  ' Positive is Good -> Green
        Else
            cell.Font.Color = RGB(255, 80, 80)  ' Negative is Bad -> Red
        End If
    Else
        ' If positiveIsGood is False, we assume Negative is Good
        If val < 0 Then
            cell.Font.Color = RGB(0, 210, 100)  ' Negative is Good -> Green
        Else
            cell.Font.Color = RGB(255, 80, 80)  ' Positive is Bad -> Red
        End If
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
        cell.Value = "WEAK BULL": cell.Font.Color = RGB(255, 192, 0)
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
    wsWatch.cells(1, 1).Font.Color = RGB(255, 192, 0): wsWatch.cells(1, 1).Font.Bold = True

    ' destination column <- scan table column
    Dim headers As Variant: headers = Array("TICKER", "COMPANY", "R" & NB_SHORT, "R" & NB_LONG, "TREND", "180D CHG%")
    Dim srcCols As Variant: srcCols = Array(SC_TICKER, SC_COMPANY, SC_R20, SC_R55, SC_TREND, SC_CHG180)
    Dim ci As Integer
    For ci = 0 To UBound(headers)
        With wsWatch.cells(2, ci + 1): .Value = headers(ci): .Font.Color = RGB(255, 192, 0): .Font.Bold = True: .Interior.Color = RGB(10, 10, 10): End With
    Next ci

    Dim cc As Long: cc = WorksheetFunction.Min(topN, ScanTickerRows(wsRes))
    Dim r As Long, wr As Long: wr = 3
    For r = 3 To 3 + cc - 1
        wsWatch.Range(wsWatch.cells(wr, 1), wsWatch.cells(wr, UBound(headers) + 1)).Interior.Color = IIf((wr Mod 2) = 1, RGB(12, 12, 12), RGB(20, 20, 20))
        For ci = 0 To UBound(srcCols)
            wsWatch.cells(wr, ci + 1).Value = wsRes.cells(r, srcCols(ci)).Value
            wsWatch.cells(wr, ci + 1).NumberFormat = wsRes.cells(r, srcCols(ci)).NumberFormat
        Next ci
        wsWatch.cells(wr, 1).Font.Color = RGB(255, 192, 0): wsWatch.cells(wr, 1).Font.Bold = True
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
    For r = 2 To LR
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
    Dim st As Range: Set st = wsRes.Range(GROUP_STATUS_CELL)

    raw = Replace(Replace(raw, ChrW(&HFF0C), ","), ChrW(&H3001), ",")
    If Len(Trim$(raw)) = 0 Then st.Value = "": Exit Sub

    Dim allNames As Variant: allNames = GetGroupNames("")
    If IsEmpty(allNames) Then st.Value = "Groups table missing or empty.": Exit Sub

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
                        cand = cand & vbLf & "- " & allNames(j)
                    End If
                Next j
            End If
        End If
    Next i

    If Len(miss) > 0 Then
        st.Value = "No exact group: " & miss & vbLf & _
                   IIf(Len(cand) > 0, "Candidates:" & cand, "No group name contains that text.") & _
                   vbLf & "Nothing scanned."
        Exit Sub
    End If
    If picked.count = 0 Then st.Value = "": Exit Sub

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
    If tick.count = 0 Then st.Value = "Group has no tickers.": Exit Sub

    Dim scanLabel As String: scanLabel = Join(picked.keys, " + ")
    Dim n As Long: n = tick.count

    Call ScanTickers(mkt, scanLabel, tick.keys, tick.items)

    If n > SCAN_MAX_TICKERS Then
        st.Value = "Scanned first " & SCAN_MAX_TICKERS & " of " & n & " tickers (truncated: only " & _
                   SCAN_MAX_TICKERS & " fit in the scan table) @ " & Format(Now, "hh:mm") & vbLf & scanLabel
    Else
        st.Value = "Scanned " & n & " tickers @ " & Format(Now, "hh:mm") & vbLf & scanLabel
    End If
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
'  US, block 2 = built-in TW. One row per group: name, ticker count,
'  ticker list (single line, clipped by the fixed column width).
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
    With ws.Range(ws.cells(1, DB_FIRST_COL), ws.cells(SCAN_LAST_ROW, lastCol))
        .Clear
        .Interior.Color = RGB(0, 0, 0)
        .Font.Name = "Consolas"
        .Font.Size = 9
        .Font.Color = RGB(200, 200, 200)
        .WrapText = False
        .NumberFormat = "@"
    End With

    ' Groups visible in the current scan (SECTOR column) get highlighted
    Dim active As Object: Set active = CreateObject("Scripting.Dictionary")
    active.CompareMode = vbTextCompare
    Dim r As Long, v As String
    For r = 3 To SCAN_LAST_ROW
        v = CStr(ws.cells(r, 3).Value)
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

    Dim titles As Variant: titles = Array("H2 / CUSTOM", "US", "TW")
    Dim nameWidths As Variant: nameWidths = Array(18, 36, 32)
    Dim maxRows As Long: maxRows = SCAN_LAST_ROW - 2          ' rows 3..SCAN_LAST_ROW
    Dim b As Long, c0 As Long, rw As Long, shown As Long, total As Long, hot As Boolean
    For b = 0 To DB_BLOCKS - 1
        c0 = DB_FIRST_COL + b * DB_BLOCK_COLS
        total = 0
        For k = 1 To n
            If gBlock(k) = b Then total = total + 1
        Next k

        With ws.cells(1, c0)
            .Value = titles(b) & "  (" & total & " groups)"
            .Font.Color = RGB(255, 192, 0)
            .Font.Bold = True
        End With
        ws.cells(2, c0).Value = "GROUP"
        ws.cells(2, c0 + 1).Value = "N"
        ws.cells(2, c0 + 2).Value = "TICKERS"
        With ws.Range(ws.cells(2, c0), ws.cells(2, c0 + 2))
            .Font.Color = RGB(255, 192, 0)
            .Font.Bold = True
            .Interior.Color = RGB(10, 10, 10)
            .Borders(xlEdgeBottom).LineStyle = xlContinuous
            .Borders(xlEdgeBottom).Color = RGB(255, 192, 0)
        End With
        ws.Columns(c0).ColumnWidth = nameWidths(b)
        ws.Columns(c0 + 1).ColumnWidth = 4
        ws.Columns(c0 + 2).ColumnWidth = 40
        If b < DB_BLOCKS - 1 Then ws.Columns(c0 + 3).ColumnWidth = 2

        rw = 3: shown = 0
        For k = 1 To n
            If gBlock(k) = b Then
                If total > maxRows And shown = maxRows - 1 Then
                    ws.cells(rw, c0).Value = "+" & (total - shown) & " more -> sheet Groups"
                    ws.cells(rw, c0).Font.Color = RGB(150, 150, 150)
                    Exit For
                End If
                hot = active.Exists(GroupKey(gName(k)))
                With ws.Range(ws.cells(rw, c0), ws.cells(rw, c0 + 2))
                    .Interior.Color = IIf(hot, RGB(60, 45, 0), IIf((rw Mod 2) = 1, RGB(12, 12, 12), RGB(20, 20, 20)))
                End With
                With ws.cells(rw, c0)
                    .Value = gName(k)
                    .Font.Color = IIf(hot, RGB(255, 192, 0), RGB(210, 210, 210))
                    .Font.Bold = hot
                End With
                With ws.cells(rw, c0 + 1)
                    .NumberFormat = "0"
                    .Value = gCount(k)
                    .HorizontalAlignment = xlCenter
                    .Font.Color = IIf(gCount(k) > SCAN_MAX_TICKERS, RGB(255, 140, 0), RGB(180, 180, 255))
                End With
                With ws.cells(rw, c0 + 2)
                    .Value = Replace(gTick(k), ",", ", ")
                    .Font.Color = IIf(hot, RGB(230, 200, 120), RGB(150, 150, 150))
                End With
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
    If r < 3 Or r > SCAN_LAST_ROW Then Exit Function
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
    Dim r As Long: r = 3
    Do While r <= SCAN_LAST_ROW
        Dim v As String: v = CStr(ws.cells(r, SC_TICKER).Value)
        If Len(v) = 0 Or v = "SUMMARY" Then Exit Do
        r = r + 1
    Loop
    ScanTickerRows = r - 3
End Function

' ================================================================
'  PRICE-CHART WALL (rows CHART_FIRST_ROW..CHART_LAST_ROW)
'  Source data lives on the hidden sheet CHART_DATA_SHEET, three columns
'  per ticker (Date | Index=100 | Window high), row 1 = ticker, row 2 =
'  headers, row 3.. = data. Charts point at those ranges.
' ================================================================
Private Function ScanDataSheet() As Worksheet
    Dim wd As Worksheet
    On Error Resume Next
    Set wd = ThisWorkbook.Worksheets(CHART_DATA_SHEET)
    On Error GoTo 0
    If wd Is Nothing Then
        Dim back As Object: Set back = ActiveSheet
        Set wd = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.count))
        wd.Name = CHART_DATA_SHEET
        wd.Visible = xlSheetHidden
        If Not back Is Nothing Then back.Activate
    End If
    Set ScanDataSheet = wd
End Function

' A round axis step giving 3..6 intervals over the span
Private Function NiceStep(ByVal span As Double) As Double
    Dim steps As Variant: steps = Array(5, 10, 20, 25, 50, 100, 200, 250, 500, 1000)
    Dim i As Long
    For i = 0 To UBound(steps)
        If span / steps(i) <= 6 Then NiceStep = steps(i): Exit Function
    Next i
    NiceStep = 1000
End Function

Private Sub DrawChartWall(ws As Worksheet, serKey As Object, serDt() As Variant, serCl() As Variant, ByVal nRows As Long)
    Dim j As Long
    For j = ws.ChartObjects.count To 1 Step -1
        If Left$(ws.ChartObjects(j).Name, Len(CHART_PREFIX)) = CHART_PREFIX Then ws.ChartObjects(j).Delete
    Next j
    With ws.Range(ws.cells(CHART_FIRST_ROW - 1, 1), ws.cells(CHART_LAST_ROW + 1, 60))
        .Clear
        .Interior.Color = RGB(0, 0, 0)
    End With
    ws.Range(ws.cells(CHART_FIRST_ROW - 1, 1), ws.cells(CHART_LAST_ROW + 1, 1)).EntireRow.RowHeight = CHART_ROW_HEIGHT
    If nRows <= 0 Or serKey.count = 0 Then Exit Sub

    Dim wd As Worksheet: Set wd = ScanDataSheet()
    wd.cells.Clear

    ' Pass 1: write each plotted ticker's window (table order) and find the
    ' shared value range
    Dim plotTk() As String, plotBase() As Long, plotLen() As Long, plotLast() As Double
    ReDim plotTk(1 To nRows): ReDim plotBase(1 To nRows): ReDim plotLen(1 To nRows): ReDim plotLast(1 To nRows)
    Dim nPlot As Long, r As Long, tk As String, s As Long
    Dim gMin As Double, gMax As Double: gMin = 1E+300: gMax = -1E+300
    For r = 3 To 2 + nRows
        tk = CStr(ws.cells(r, SC_TICKER).Value)
        If serKey.Exists(tk) Then
            s = serKey(tk)
            Dim d As Variant, c As Variant: d = serDt(s): c = serCl(s)
            Dim last As Long: last = UBound(c)
            Dim first As Long: first = last - YEAR_BARS + 1
            If first < LBound(c) Then first = LBound(c)
            If c(first) > 0 And last > first Then
                nPlot = nPlot + 1
                Dim m As Long: m = last - first + 1
                Dim out() As Variant: ReDim out(1 To m + 2, 1 To 3)
                out(1, 1) = tk: out(2, 1) = "Date": out(2, 2) = "Index": out(2, 3) = "High"
                Dim hi As Double: hi = 0
                Dim q As Long
                For q = 0 To m - 1
                    out(q + 3, 1) = d(first + q)
                    out(q + 3, 2) = c(first + q) / c(first) * 100
                    If out(q + 3, 2) > hi Then hi = out(q + 3, 2)
                    If out(q + 3, 2) < gMin Then gMin = out(q + 3, 2)
                Next q
                For q = 0 To m - 1
                    out(q + 3, 3) = hi
                Next q
                If hi > gMax Then gMax = hi
                Dim col0 As Long: col0 = (nPlot - 1) * 3 + 1
                wd.Range(wd.cells(1, col0), wd.cells(m + 2, col0 + 2)).Value = out
                wd.Range(wd.cells(3, col0), wd.cells(m + 2, col0)).NumberFormat = "yyyy-mm-dd"
                plotTk(nPlot) = tk: plotBase(nPlot) = col0: plotLen(nPlot) = m
                plotLast(nPlot) = out(m + 2, 2)
            End If
        End If
    Next r
    If nPlot = 0 Then Exit Sub

    Dim yStep As Double: yStep = NiceStep(gMax - gMin)
    Dim yMin As Double: yMin = Int(gMin / yStep) * yStep
    Dim yMax As Double: yMax = -Int(-gMax / yStep) * yStep

    ' Pass 2: one chart per ticker, CHART_PER_ROW across
    Dim x0 As Double: x0 = ws.cells(1, SC_TICKER).Left
    Dim y0 As Double: y0 = ws.cells(CHART_FIRST_ROW, 1).Top
    Dim p As Long
    For p = 1 To nPlot
        Dim co As ChartObject
        Set co = ws.ChartObjects.Add( _
            x0 + ((p - 1) Mod CHART_PER_ROW) * (CHART_W + CHART_GAP), _
            y0 + ((p - 1) \ CHART_PER_ROW) * (CHART_H + CHART_GAP), CHART_W, CHART_H)
        co.Name = CHART_PREFIX & p
        co.Placement = xlFreeFloating
        Dim xr As Range, vr As Range, hr As Range
        Set xr = wd.Range(wd.cells(3, plotBase(p)), wd.cells(2 + plotLen(p), plotBase(p)))
        Set vr = xr.Offset(0, 1)
        Set hr = xr.Offset(0, 2)
        With co.Chart
            .ChartType = xlLine
            Do While .SeriesCollection.count > 0
                .SeriesCollection(1).Delete
            Loop
            With .SeriesCollection.NewSeries
                .Name = plotTk(p)
                .XValues = xr
                .Values = vr
                .MarkerStyle = xlMarkerStyleNone
                .Format.line.Visible = msoTrue
                .Format.line.Weight = 0.75
                .Format.line.ForeColor.RGB = IIf(plotLast(p) >= 100, RGB(255, 80, 80), RGB(0, 210, 100))
            End With
            With .SeriesCollection.NewSeries
                .Name = "High"
                .XValues = xr
                .Values = hr
                .MarkerStyle = xlMarkerStyleNone
                .Format.line.Visible = msoTrue
                .Format.line.Weight = 0.75
                .Format.line.DashStyle = msoLineDash
                .Format.line.ForeColor.RGB = RGB(160, 160, 160)
            End With
            .HasLegend = False
            .HasTitle = True
            .ChartTitle.text = plotTk(p) & "   " & Format(plotLast(p) - 100, "+0.0;-0.0;0.0") & "%"

            .ChartArea.Format.Fill.Visible = msoTrue
            .ChartArea.Format.Fill.Solid
            .ChartArea.Format.Fill.ForeColor.RGB = RGB(0, 0, 0)
            .ChartArea.Format.line.Visible = msoTrue
            .ChartArea.Format.line.ForeColor.RGB = RGB(60, 60, 60)
            .PlotArea.Format.Fill.Visible = msoTrue
            .PlotArea.Format.Fill.Solid
            .PlotArea.Format.Fill.ForeColor.RGB = RGB(0, 0, 0)

            With .Axes(xlValue)
                .MinimumScale = yMin
                .MaximumScale = yMax
                .MajorUnit = yStep
                .HasMajorGridlines = True
                .MajorGridlines.Format.line.ForeColor.RGB = RGB(45, 45, 45)
                .Format.line.ForeColor.RGB = RGB(90, 90, 90)
                .TickLabels.NumberFormat = "0"
            End With
            With .Axes(xlCategory)
                .CategoryType = xlTimeScale
                .BaseUnit = xlDays
                .MajorUnitScale = xlMonths
                .MajorUnit = 3
                .TickLabels.NumberFormat = "yy/mm"
                .Format.line.ForeColor.RGB = RGB(90, 90, 90)
            End With

            ' All chart text: white Calibri 12
            .ChartArea.Font.Name = "Calibri"
            .ChartArea.Font.Size = 12
            .ChartArea.Font.Color = RGB(255, 255, 255)
            With .ChartTitle.Font
                .Name = "Calibri": .Size = 12: .Color = RGB(255, 255, 255): .Bold = True
            End With
            With .Axes(xlValue).TickLabels.Font
                .Name = "Calibri": .Size = 12: .Color = RGB(255, 255, 255)
            End With
            With .Axes(xlCategory).TickLabels.Font
                .Name = "Calibri": .Size = 12: .Color = RGB(255, 255, 255)
            End With
        End With
    Next p
End Sub



