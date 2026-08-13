Attribute VB_Name = "PortfolioDashboard_v3"
Option Explicit

' ================================================================
'  PORTFOLIO DASHBOARD v3.0 - Bloomberg Terminal Style
' ================================================================

Private Const SH_PORT  As String = "RR4"
Private Const SH_TRANS As String = "Transactions"
Private Const SH_REAL  As String = "Realized"
Private Const SH_HIST  As String = "HistoryLog"

Public g_AutoOn   As Boolean
Public g_NextTime As Date

' ================================================================
'  MAIN ENTRY
' ================================================================
Sub RebuildPortfolioDashboard()
    Dim wsP  As Worksheet
    Dim wsTr As Worksheet
    Dim wsR  As Worksheet
    Set wsP = ThisWorkbook.Sheets(SH_PORT)
    Set wsTr = ThisWorkbook.Sheets(SH_TRANS)
    Set wsR = ThisWorkbook.Sheets(SH_REAL)
    
    If Not g_PriceCache Is Nothing Then
        g_PriceCache.RemoveAll
        g_CacheTime = Now
    End If
    
    ' Backfill missing history rows

    Application.ScreenUpdating = False
    Dim swingRiskMap As Object
    Set swingRiskMap = CreateObject("Scripting.Dictionary")
    Dim srLastRow As Long
    srLastRow = wsP.cells(wsP.Rows.count, 1).End(xlUp).row
    Dim srRow As Long
    For srRow = 5 To srLastRow
        Dim srTicker As String: srTicker = CStr(wsP.cells(srRow, 1).Value)
        Dim srVal    As String: srVal = CStr(wsP.cells(srRow, 17).Value)
        If srTicker <> "" And srVal <> "" Then
            swingRiskMap(srTicker) = srVal
        End If
    Next srRow
    Call ResetSheetStyle(wsP)

    Dim positions As Object
    Set positions = BuildPositions(wsTr)

    Dim exRate As Double
    exRate = GetExRate(wsP)

    Dim posData()    As Variant
    Dim totalMktTWD  As Double
    Dim totalCostTWD As Double
    Dim totalUnrlTWD As Double
    Dim posCount     As Long

    Call CalcPositions(positions, exRate, posData, totalMktTWD, totalCostTWD, totalUnrlTWD, posCount)
    Call CalculateRealizedPnL

    Dim realPnL As Double
    On Error Resume Next
    realPnL = Application.WorksheetFunction.Sum(wsR.Columns("G"))
    On Error GoTo 0

    Dim portBeta As Double
    portBeta = CalcPortBeta(posData)

    Call DrawHeader(wsP, totalMktTWD, totalCostTWD, totalUnrlTWD, realPnL, portBeta, posCount, exRate)
    Call DrawColumnHeaders(wsP)
    Dim lastDataRow As Long
    lastDataRow = WritePositionRows(wsP, posData, totalMktTWD, swingRiskMap)
    Call DrawDisclaimer(wsP, lastDataRow)
    Call LogHistory(totalMktTWD, totalUnrlTWD + realPnL, realPnL)

    Application.ScreenUpdating = True
    Call DrawDeepAnalysis(wsP, posData, totalMktTWD, totalCostTWD, totalUnrlTWD, realPnL)
    Application.StatusBar = "Dashboard updated: " & Format(Now, "hh:mm:ss")
    MsgBox "Dashboard updated!", vbInformation
End Sub

Private Sub LogHistory(totalMkt As Double, totalPnL As Double, realPnL As Double)
    Dim wsH As Worksheet
    On Error Resume Next
    Set wsH = ThisWorkbook.Sheets(SH_HIST)
    On Error GoTo 0
    If wsH Is Nothing Then Exit Sub
    
    Dim nr As Long: nr = wsH.cells(wsH.Rows.count, "A").End(xlUp).row + 1
    If nr > 2 Then
        If Int(CDate(wsH.cells(nr - 1, 1).Value)) = Date Then nr = nr - 1
    End If
    
    Dim priceSPY As Double, priceQQQ As Double, priceTWII As Double
    Dim priceSOX As Double, priceTWOII As Double
    
    priceSPY = GetStockPrice("SPY")
    priceQQQ = GetStockPrice("QQQ")
    priceTWII = GetStockPrice("^TWII")
    priceSOX = GetStockPrice("^SOX")
    priceTWOII = GetStockPrice("^TWOII")
    
    With wsH
        .cells(nr, "A").Value = Now
        .cells(nr, "B").Value = totalMkt
        .cells(nr, "C").Value = totalPnL
        .cells(nr, "D").Value = priceSPY
        .cells(nr, "E").Value = priceQQQ
        .cells(nr, "F").Value = priceTWII
        .cells(nr, "G").Value = realPnL
        .cells(nr, "N").Value = priceSOX
        .cells(nr, "O").Value = priceTWOII
        
        .cells(nr, "P").Formula = "=(N" & nr & "-$N$2)/$N$2"
        .cells(nr, "Q").Formula = "=(O" & nr & "-$O$2)/$O$2"
        
        .cells(nr, "A").NumberFormat = "yyyy/m/d h:mm:ss"
        .Range("B" & nr & ":C" & nr).NumberFormat = "#,##0"
        .Range("D" & nr & ":F" & nr).NumberFormat = "#,##0.00"
        .cells(nr, "G").NumberFormat = "#,##0"
        .cells(nr, "N").NumberFormat = "#,##0.00"
        .cells(nr, "O").NumberFormat = "#,##0.00"
        .cells(nr, "P").NumberFormat = "0.00%"
        .cells(nr, "Q").NumberFormat = "0.00%"
    End With
End Sub
Function GetHistoricalPrice(Ticker As String, targetDate As Date) As Double
    Dim http As Object
    Set http = CreateObject("MSXML2.XMLHTTP")
    
    ' Unix epoch is UTC; TW market is UTC+8 (8h ahead) - opening crosses a day boundary
    Dim p1 As Long, p2 As Long
    p1 = DateDiff("s", #1/1/1970#, targetDate) - 8 * 3600
    p2 = p1 + 86400
    
    ' ^ -- URL encode -- %5E
    Dim safeTicker As String
    safeTicker = Replace(Ticker, "^", "%5E")
    
    Dim url As String
    url = "https://query2.finance.yahoo.com/v8/finance/chart/" & safeTicker & _
          "?period1=" & p1 & "&period2=" & p2 & "&interval=1d"
    
    On Error GoTo ErrHandler
    http.Open "GET", url, False
    http.setRequestHeader "User-Agent", "Mozilla/5.0"
    http.send
    
    Dim resp As String: resp = http.responseText
    
    ' ------------------------------------------------------------
    Dim pos As Long
    pos = InStr(resp, """close"":[")
    If pos = 0 Then GoTo ErrHandler
    pos = pos + 9
    Dim endPos As Long: endPos = InStr(pos, resp, "]")
    Dim closeVal As String: closeVal = Trim(Mid(resp, pos, endPos - pos))
    If Len(closeVal) = 0 Or InStr(closeVal, "null") > 0 Then GoTo ErrHandler
    
    GetHistoricalPrice = CDbl(Split(closeVal, ",")(0))
    Exit Function
    
ErrHandler:
    GetHistoricalPrice = 0
End Function
Sub BackfillHistory()
    Dim wsH As Worksheet
    On Error Resume Next
    Set wsH = ThisWorkbook.Sheets(SH_HIST)
    On Error GoTo 0
    If wsH Is Nothing Then MsgBox "History sheet not found": Exit Sub
    
    Dim lastRow As Long
    lastRow = wsH.cells(wsH.Rows.count, "A").End(xlUp).row
    If lastRow < 2 Then MsgBox "No data to backfill": Exit Sub
    
    Application.ScreenUpdating = False
    
    Dim i As Long, countSOX As Long, countOTC As Long
    For i = 2 To lastRow
        If wsH.cells(i, "A").Value = "" Then GoTo NextRow
        
        ' H: refill SPY Ret%
        wsH.cells(i, "H").Formula = "=(D" & i & "-$D$2)/$D$2"
        wsH.cells(i, "H").NumberFormat = "0.00%"
        
        ' ------------------------------------------------------------
        Dim targetDate As Date
        On Error Resume Next
        targetDate = CDate(Int(CDbl(wsH.cells(i, "A").Value)))
        On Error GoTo 0
        If targetDate = 0 Then GoTo NextRow
        
        Dim price As Double
        
        ' N -- G -- ^ -- ^SOX -- v -- L --
        price = GetHistoricalPrice("^SOX", targetDate)
        If price > 0 Then
            wsH.cells(i, "N").Value = price
            wsH.cells(i, "N").NumberFormat = "#,##0.00"
            wsH.cells(i, "P").Formula = "=(N" & i & "-$N$2)/$N$2"
            wsH.cells(i, "P").NumberFormat = "0.00%"
            countSOX = countSOX + 1
        End If
        
        ' O -- G -- ^ -- ^TWOII -- v -- L --
        price = GetHistoricalPrice("^TWOII", targetDate)
        If price > 0 Then
            wsH.cells(i, "O").Value = price
            wsH.cells(i, "O").NumberFormat = "#,##0.00"
            wsH.cells(i, "Q").Formula = "=(O" & i & "-$O$2)/$O$2"
            wsH.cells(i, "Q").NumberFormat = "0.00%"
            countOTC = countOTC + 1
        End If
        
        Application.StatusBar = "Backfilling row " & i & "/" & lastRow & _
                                 "  SOX:" & countSOX & "  OTC:" & countOTC
NextRow:
    Next i
    
    Application.ScreenUpdating = True
    Application.StatusBar = False
    MsgBox "Backfill complete!" & vbLf & _
           "H = SPY Ret%: refilled (total " & lastRow - 1 & " rows)" & vbLf & _
           "N = ^SOX: updated " & countSOX & " rows" & vbLf & _
           "O = ^TWOII: updated " & countOTC & " rows" & vbLf & _
           "Existing / blank rows skipped automatically"
End Sub

' ================================================================
'  Sheet style reset
' ================================================================
Private Sub ResetSheetStyle(ws As Worksheet)
    ws.cells.Clear
    With ws.cells
        .Interior.Color = RGB(0, 0, 0)
        .Font.Color = RGB(221, 221, 221)
        .Font.Name = "Consolas"
        .Font.Size = 10
        .Font.Bold = False
        .NumberFormat = "General"
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
    End With
    ws.Activate
    ActiveWindow.DisplayGridlines = False

    Dim colWidths As Variant
    colWidths = Array(5, 22, 10, 8, 8, 10, 10, 8, 10, 10, 8, 11, 7, 8, 10, 10)
    ' Column widths (Excel units ? pixels / 7)
    ' A=240px?34  B=580px?83  F=580px?83  others=138px?20
    Dim i As Integer
    For i = 1 To 16
        Select Case i
            Case 1: ws.Columns(i).ColumnWidth = 17          ' A: 240px
            Case 2, 6: ws.Columns(i).ColumnWidth = 40       ' B, F: 580px
            Case Else: ws.Columns(i).ColumnWidth = 11       ' others: 138px
        End Select
    Next i

    Dim r As Long
    For r = 1 To 100
        ws.Rows(r).RowHeight = 18
    Next r
    ws.Rows(1).RowHeight = 28
    ws.Rows(4).RowHeight = 20
End Sub

' ================================================================
'  Header (rows 1-3)
' ================================================================
Private Sub DrawHeader(ws As Worksheet, totalMkt As Double, totalCost As Double, _
                        totalUnrl As Double, realPnL As Double, _
                        portBeta As Double, posCount As Long, exRate As Double)

    Dim unrlPct As Double
    If totalCost > 0 Then unrlPct = totalUnrl / totalCost Else unrlPct = 0

    With ws.cells(1, 1)
        .Value = Format(totalMkt, "$#,##0")
        .Font.Color = RGB(255, 192, 0)
        .Font.Size = 16
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
    End With
    
    ws.cells(1, 2).Value = " <- TOTAL MKT "
    ws.cells(1, 2).Font.Color = RGB(255, 192, 0)
    ws.cells(1, 2).Font.Bold = True
    
    ws.cells(1, 3).Value = "INCEPTION"
    ws.cells(1, 3).Font.Color = RGB(150, 150, 150)
    ws.cells(2, 3).Value = Format(GetInceptionDate(ws), "m/d/yyyy")
    ws.cells(2, 3).Font.Color = RGB(221, 221, 221)

    ws.cells(1, 4).Value = "DAYS"
    ws.cells(1, 4).Font.Color = RGB(150, 150, 150)
    ws.cells(2, 4).Value = Date - GetInceptionDate(ws)
    ws.cells(2, 4).Font.Color = RGB(221, 221, 221)

    ws.cells(1, 5).Value = "STARTING"
    ws.cells(1, 5).Font.Color = RGB(150, 150, 150)
    ws.cells(2, 5).Value = GetStartingCapital(ws)
    ws.cells(2, 5).Font.Color = RGB(221, 221, 221)
    ws.cells(2, 5).NumberFormat = "#,##0"

    ' 2026-08-13: PORT.BETA / POSITIONS moved down to row 2 (see below, where
    ' UNRL PNL%/UNRL PNL used to display) - this row is left intentionally
    ' blank so the widget keeps its original two-row height.
    ws.cells(1, 9).Value = ""
    ws.cells(1, 9).NumberFormat = "General"
    ws.cells(1, 10).Value = ""
    ws.cells(1, 10).NumberFormat = "General"
    ws.cells(1, 11).Value = ""
    ws.cells(1, 11).NumberFormat = "General"
    ws.cells(1, 12).Value = ""
    ws.cells(1, 12).NumberFormat = "General"

    ws.cells(2, 13).Value = "RL PNL"
    ws.cells(2, 13).Font.Color = RGB(150, 150, 150)
    ws.cells(2, 14).Value = realPnL
    ws.cells(2, 14).NumberFormat = "$#,##0"
    ws.cells(2, 14).Font.Color = PnLColor(realPnL)
    ws.cells(2, 14).Font.Bold = True

    Dim startCap As Double: startCap = GetStartingCapital(ws)
    Dim rlPct    As Double
    If startCap > 0 Then rlPct = realPnL / startCap Else rlPct = 0
    ws.cells(2, 15).Value = "RL PNL%"
    ws.cells(2, 15).Font.Color = RGB(150, 150, 150)
    ws.cells(2, 16).Value = rlPct
    ws.cells(2, 16).NumberFormat = "0.00%"
    ws.cells(2, 16).Font.Color = PnLColor(rlPct)
    ws.cells(2, 16).Font.Bold = True

    ws.cells(2, 1).Value = "USD/TWD"
    ws.cells(2, 1).Font.Color = RGB(150, 150, 150)
    ws.cells(2, 2).Value = exRate
    ws.cells(2, 2).NumberFormat = "0.00"
    ws.cells(2, 2).Font.Color = RGB(221, 221, 221)

    ' UNRL PNL%/UNRL PNL calc (unrlPct above, totalUnrl param) intentionally
    ' kept for LogHistory/DrawDeepAnalysis - just no longer displayed here.
    ' POSITIONS and PORT.BETA now occupy this row instead (order swapped:
    ' Positions left, Beta right), same styling as their old row-1 spot.
    ws.cells(2, 9).Value = "POSITIONS"
    ws.cells(2, 9).Font.Color = RGB(150, 150, 150)
    ws.cells(2, 10).Value = posCount
    ws.cells(2, 10).NumberFormat = "General"
    ws.cells(2, 10).Font.Color = RGB(221, 221, 221)
    ws.cells(2, 10).Font.Bold = False

    ws.cells(2, 11).Value = "PORT.BETA"
    ws.cells(2, 11).Font.Color = RGB(150, 150, 150)
    ws.cells(2, 12).Value = portBeta
    ws.cells(2, 12).NumberFormat = "0.00"
    ws.cells(2, 12).Font.Color = RGB(255, 192, 0)
    ws.cells(2, 12).Font.Bold = True

    With ws.Range(ws.cells(3, 1), ws.cells(3, 16))
        .Interior.Color = RGB(255, 192, 0)
        .RowHeight = 3
    End With
End Sub

' ================================================================
'  DEEP ANALYSIS - Sector / Strategy / Currency / Technical
'  Writes to a separate "Analysis" sheet, dark theme
' ================================================================
Sub DrawDeepAnalysis(wsP As Worksheet, posData() As Variant, _
                     totalMkt As Double, totalCost As Double, _
                     totalUnrl As Double, realPnL As Double)

    ' --- Get or create Analysis sheet ---
    Dim wsA As Worksheet
    On Error Resume Next
    Set wsA = ThisWorkbook.Sheets("Analysis")
    On Error GoTo 0
    If wsA Is Nothing Then
        Set wsA = ThisWorkbook.Sheets.Add(After:=wsP)
        wsA.Name = "Analysis"
    End If

    wsA.cells.Clear
    With wsA.cells
        .Interior.Color = RGB(0, 0, 0)
        .Font.Color = RGB(221, 221, 221)
        .Font.Name = "Consolas"
        .Font.Size = 36
    End With
    wsA.Activate
    ActiveWindow.DisplayGridlines = False

    If Not IsArray(posData) Then Exit Sub
    Dim n As Long
    On Error Resume Next
    n = UBound(posData, 1)
    If Err.Number <> 0 Or n < 1 Then Exit Sub
    On Error GoTo 0

    ' ------------------------------------------------------------
    With wsA.cells(1, 1)
        .Value = "DEEP ANALYSIS  -  " & Format(Now, "yyyy/mm/dd hh:mm")
        .Font.Color = RGB(255, 192, 0)
        .Font.Bold = True
        .Font.Size = 36
    End With
    With wsA.Range(wsA.cells(2, 1), wsA.cells(2, 20))
        .Interior.Color = RGB(255, 192, 0)
        .RowHeight = 3
    End With

    ' ------------------------------------------------------------
    Dim totalPnL As Double: totalPnL = totalUnrl + realPnL
    Dim retPct   As Double
    If totalCost > 0 Then retPct = totalPnL / totalCost

    Dim r As Long: r = 4
    wsA.cells(r, 1).Value = "RETURN SUMMARY"
    wsA.cells(r, 1).Font.Color = RGB(255, 192, 0)
    wsA.cells(r, 1).Font.Bold = True
    r = r + 1

    Call WriteKV(wsA, r, "Total Return %", retPct, "0.00%", True): r = r + 1
    Call WriteKV(wsA, r, "Total Cost", totalCost, "$#,##0", False): r = r + 1
    Call WriteKV(wsA, r, "Unrealized PnL", totalUnrl, "$#,##0", True): r = r + 1
    Call WriteKV(wsA, r, "Realized PnL", realPnL, "$#,##0", True): r = r + 1
    Call WriteKV(wsA, r, "Total PnL", totalPnL, "$#,##0", True): r = r + 1

    ' ------------------------------------------------------------
    r = r + 2
    Dim sectorD   As Object: Set sectorD = CreateObject("Scripting.Dictionary")
    Dim stratD    As Object: Set stratD = CreateObject("Scripting.Dictionary")
    Dim currencyD As Object: Set currencyD = CreateObject("Scripting.Dictionary")

    Dim i As Long
    For i = 1 To n
        Dim sec  As String: sec = CStr(posData(i, 7))
        Dim cur  As String: cur = GetCurrencyType(CStr(posData(i, 2)))
        Dim mv   As Double: mv = posData(i, 8)
        Dim stg  As String
        ' strategy stored in posData(i,2) ticker -> need from positions
        ' use sector as proxy if strategy not in posData
        stg = CStr(posData(i, 7))   ' fallback: sector

        If sec <> "" Then If sectorD.Exists(sec) Then sectorD(sec) = sectorD(sec) + mv Else sectorD.Add sec, mv
        If cur <> "" Then If currencyD.Exists(cur) Then currencyD(cur) = currencyD(cur) + mv Else currencyD.Add cur, mv
    Next i

    r = DrawAllocTable(wsA, r, "SECTOR ALLOCATION", sectorD, totalMkt)
    r = r + 2
    r = DrawAllocTable(wsA, r, "CURRENCY ALLOCATION", currencyD, totalMkt)

    ' ------------------------------------------------------------
    r = r + 2
    wsA.cells(r, 1).Value = "TECHNICAL & HOLDING ANALYSIS"
    wsA.cells(r, 1).Font.Color = RGB(255, 192, 0)
    wsA.cells(r, 1).Font.Bold = True
    r = r + 1

    ' Header
    Dim techHdrs As Variant
    techHdrs = Array("TICKER", "NAME", "DAYS", "LAST", "HIGH DIST%", _
                     "DAY CHG%", "BIAS5", "BIAS20", "BIAS60", "BIAS120", "BIAS240", "TREND")
    Dim c As Integer
    For c = 0 To UBound(techHdrs)
        With wsA.cells(r, c + 1)
            .Value = techHdrs(c)
            .Font.Color = RGB(255, 192, 0)
            .Font.Bold = True
            .Interior.Color = RGB(10, 10, 10)
            .HorizontalAlignment = xlCenter
        End With
    Next c
    r = r + 1

    ' Data rows
    For i = 1 To n
        Dim tkr   As String: tkr = CStr(posData(i, 2))
        Dim nm    As String: nm = CStr(posData(i, 3))
        Dim ddays As Long: ddays = posData(i, 5)

        Dim rowBg As Long: rowBg = IIf(i Mod 2 = 1, RGB(15, 15, 15), RGB(22, 22, 22))
        With wsA.Range(wsA.cells(r, 1), wsA.cells(r, 12))
            .Interior.Color = rowBg
            .HorizontalAlignment = xlCenter
        End With

        ' Get tech data from existing function
        Dim tech As TechIndicators
        tech = GetTechData(tkr)

        wsA.cells(r, 1).Value = tkr
        wsA.cells(r, 2).Value = nm
        wsA.cells(r, 2).HorizontalAlignment = xlLeft
        wsA.cells(r, 3).Value = ddays
        wsA.cells(r, 4).Value = posData(i, 11)
        wsA.cells(r, 4).NumberFormat = "#,##0.00"

        If tech.Success Then
            wsA.cells(r, 5).Value = tech.DistFromHigh
            wsA.cells(r, 6).Value = tech.ChangePercent
            wsA.cells(r, 7).Value = tech.Bias5
            wsA.cells(r, 8).Value = tech.Bias20
            wsA.cells(r, 9).Value = tech.Bias60
            wsA.cells(r, 10).Value = tech.Bias120
            wsA.cells(r, 11).Value = tech.Bias240

            ' Colour each metric cell
            Dim mc As Integer
            For mc = 5 To 11
                wsA.cells(r, mc).NumberFormat = "0.00%"
                wsA.cells(r, mc).Font.Color = PnLColor(wsA.cells(r, mc).Value)
            Next mc

            ' Trend
            If tech.Bias20 > 0 And tech.Bias60 > 0 Then
                wsA.cells(r, 12).Value = "BULLISH"
                wsA.cells(r, 12).Font.Color = RGB(255, 80, 80)
            ElseIf tech.Bias20 < 0 And tech.Bias60 < 0 Then
                wsA.cells(r, 12).Value = "BEARISH"
                wsA.cells(r, 12).Font.Color = RGB(0, 210, 100)
            Else
                wsA.cells(r, 12).Value = "NEUTRAL"
                wsA.cells(r, 12).Font.Color = RGB(200, 200, 200)
            End If
            wsA.cells(r, 12).Font.Bold = True
        Else
            wsA.cells(r, 5).Value = "N/A"
        End If
        r = r + 1
    Next i

    wsA.Columns("A:L").AutoFit
    wsP.Activate   ' return to main sheet
End Sub

Sub RebuildActiveXButtons()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("RR4")
    
    ' ------------------------------ Unprotect sheet ------------------------------
    On Error Resume Next
    ws.Unprotect
    On Error GoTo 0
    
    ' ------------------------------ Remove existing shapes -----------------------
    
    Dim shp As Shape
    Dim delList As New Collection
    For Each shp In ws.Shapes
        delList.Add shp.Name
    Next shp
    Dim nm As Variant
    For Each nm In delList
        On Error Resume Next
        ws.Shapes(CStr(nm)).Delete
        On Error GoTo 0
    Next nm
    
    ' ------------------------------ Probe ActiveX availability -------------------
    Dim testOle As OLEObject
    On Error Resume Next
    Set testOle = ws.OLEObjects.Add(ClassType:="Forms.CommandButton.1", _
                                     Left:=10, Top:=5, Width:=50, Height:=20)
    If Err.Number <> 0 Then
        MsgBox "ActiveX is unavailable. Reason: " & Err.Description & vbCr & vbCr & _
               "Please check:" & vbCr & _
               "1. File is saved as .xlsm" & vbCr & _
               "2. Trust Center > Enable ActiveX controls" & vbCr & _
               "3. Worksheet is not protected", vbExclamation
        Exit Sub
    End If
    testOle.Delete
    On Error GoTo 0
    
    ' ------------------------------ Create new buttons ---------------------------
    Dim topPos As Double: topPos = ws.Rows(35).Top

    Dim btnDefs As Variant
    btnDefs = Array( _
        Array("DELETE ALL", 10, topPos, 130, 20, RGB(30, 0, 0), RGB(255, 60, 60)), _
        Array("REGULATE", 150, topPos, 160, 20, RGB(10, 10, 30), RGB(100, 160, 255)), _
        Array("UPDATE", 320, topPos, 100, 20, RGB(0, 25, 0), RGB(0, 220, 100)), _
        Array("ADD TRANSACTION", 430, topPos, 160, 20, RGB(20, 15, 0), RGB(255, 192, 0)), _
        Array("HOLDINGS CORR", 600, topPos, 140, 20, RGB(15, 0, 30), RGB(180, 100, 255)), _
        Array("DRAWDOWN", 750, topPos, 110, 20, RGB(30, 5, 0), RGB(255, 120, 60)), _
        Array("AUTO ON", 870, topPos, 100, 20, RGB(0, 40, 0), RGB(0, 255, 136)), _
        Array("AUTO OFF", 980, topPos, 100, 20, RGB(40, 0, 0), RGB(255, 80, 80)), _
        Array("DEBUG", 1090, topPos, 80, 20, RGB(0, 10, 30), RGB(100, 160, 255)) _
    )
    
    Dim i As Integer
    For i = 0 To UBound(btnDefs)
        Dim b As Variant: b = btnDefs(i)
        Dim ole As OLEObject
        Set ole = ws.OLEObjects.Add( _
            ClassType:="Forms.CommandButton.1", _
            Left:=b(1), Top:=b(2), Width:=b(3), Height:=b(4))
        ole.Name = "CmdBtn" & i
        With ole.Object
            .Caption = b(0)
            .Font.Name = "Consolas"
            .Font.Size = 10
            .Font.Bold = True
            .BackColor = b(5)
            .ForeColor = b(6)
            .BackStyle = 1
        End With
    Next i

End Sub

' ------------------------------------------------------------
Private Function DrawAllocTable(ws As Worksheet, startRow As Long, _
                                  title As String, dict As Object, _
                                  totalMkt As Double) As Long
    Dim r As Long: r = startRow
    ws.cells(r, 1).Value = title
    ws.cells(r, 1).Font.Color = RGB(255, 192, 0)
    ws.cells(r, 1).Font.Bold = True
    r = r + 1

    ' Header
    ws.cells(r, 1).Value = "CATEGORY"
    ws.cells(r, 2).Value = "VALUE"
    ws.cells(r, 3).Value = "WEIGHT"
    With ws.Range(ws.cells(r, 1), ws.cells(r, 3))
        .Font.Bold = True
        .Interior.Color = RGB(10, 10, 10)
        .HorizontalAlignment = xlCenter
        .Font.Color = RGB(255, 192, 0)
    End With
    r = r + 1

    Dim key As Variant
    Dim i   As Long: i = 0
    For Each key In dict.keys
        Dim rowBg As Long: rowBg = IIf(i Mod 2 = 0, RGB(15, 15, 15), RGB(22, 22, 22))
        With ws.Range(ws.cells(r, 1), ws.cells(r, 3))
            .Interior.Color = rowBg
            .HorizontalAlignment = xlCenter
        End With
        ws.cells(r, 1).Value = key
        ws.cells(r, 1).HorizontalAlignment = xlLeft
        ws.cells(r, 2).Value = dict(key)
        ws.cells(r, 2).NumberFormat = "$#,##0"
        If totalMkt > 0 Then
            ws.cells(r, 3).Value = dict(key) / totalMkt
            ws.cells(r, 3).NumberFormat = "0.00%"
        End If
        r = r + 1: i = i + 1
    Next key

    DrawAllocTable = r
End Function
' ================================================================
'  Column header row (row 4)
' ================================================================
Private Sub DrawColumnHeaders(ws As Worksheet)
    Dim headers As Variant
    headers = Array("ISIN", "NAME", "ENTRY DT", "DAYS", "BROKER", _
                    "SECTOR", "NET EXPOS", "SHARES", "ENTRY PX", "LAST", _
                    "% CHG", "UNRL PNL", "WT%", "W.BETA", "Beta 180D", "P.TARGET")
    Dim i As Integer
    For i = 0 To UBound(headers)
        With ws.cells(4, i + 1)
            .Value = headers(i)
            .Font.Color = RGB(255, 192, 0)
            .Font.Bold = True
            .Font.Size = 9
            .Font.Name = "Consolas"
            .Interior.Color = RGB(10, 10, 10)
            .HorizontalAlignment = xlCenter
        End With
    Next i
    
    ' SWING RISK header (col 17)
    With ws.cells(4, 17)
        .Value = "SWING RISK"
        .Font.Color = RGB(255, 192, 0)
        .Font.Bold = True
        .Font.Size = 9
        .Font.Name = "Consolas"
        .Interior.Color = RGB(10, 10, 10)
        .HorizontalAlignment = xlCenter
    End With
    
    ws.Columns(17).ColumnWidth = 20
    With ws.Range(ws.cells(4, 1), ws.cells(4, 17)).Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Color = RGB(255, 192, 0)
        .Weight = xlThin
    End With
End Sub

' ------------------------------------------------------------
Private Function WritePositionRows(ws As Worksheet, posData() As Variant, _
                                    totalMkt As Double, ByVal swingRiskMap As Object) As Long
    WritePositionRows = 5
    If Not IsArray(posData) Then Exit Function
    Dim n As Long
    On Error Resume Next: n = UBound(posData, 1)
    If Err.Number <> 0 Or n < 1 Then Exit Function
    On Error GoTo 0

    Dim i As Long, r As Long: r = 5
    Dim currentBroker As String: currentBroker = ""
    Dim brokerMkt As Double: brokerMkt = 0
    Dim brokerUnrl As Double: brokerUnrl = 0

    For i = 1 To n
        Dim thisBroker As String: thisBroker = CStr(posData(i, 6))

        If thisBroker <> currentBroker Then
            If currentBroker <> "" Then
                Call DrawBrokerSubtotal(ws, r, currentBroker, brokerMkt, brokerUnrl)
                r = r + 1
            End If
            Call DrawBrokerHeader(ws, r, thisBroker)
            r = r + 1
            currentBroker = thisBroker
            brokerMkt = 0: brokerUnrl = 0
        End If

        Call WriteOnePositionRow(ws, r, i, posData, totalMkt, swingRiskMap)
        brokerMkt = brokerMkt + posData(i, 8)
        brokerUnrl = brokerUnrl + posData(i, 13)
        r = r + 1
    Next i

    If currentBroker <> "" Then
        Call DrawBrokerSubtotal(ws, r, currentBroker, brokerMkt, brokerUnrl)
        r = r + 1
    End If

    WritePositionRows = r - 1
End Function

' ------------------------------------------------------------
Private Sub DrawBrokerHeader(ws As Worksheet, r As Long, brokerName As String)
    ws.Rows(r).RowHeight = 22
    With ws.Range(ws.cells(r, 1), ws.cells(r, 17))
        .Interior.Color = RGB(35, 25, 0)
        .Font.Name = "Consolas"
        .Font.Size = 10
        .Font.Bold = True
        .Font.Color = RGB(255, 192, 0)
    End With
    ws.cells(r, 1).Value = "  == " & brokerName & " " & String(28, "=")
    ws.cells(r, 1).HorizontalAlignment = xlLeft
End Sub

' ------------------------------------------------------------
Private Sub DrawBrokerSubtotal(ws As Worksheet, r As Long, brokerName As String, _
                                brokerMkt As Double, brokerUnrl As Double)
    ws.Rows(r).RowHeight = 18
    With ws.Range(ws.cells(r, 1), ws.cells(r, 17))
        .Interior.Color = RGB(25, 18, 0)
    End With
    With ws.cells(r, 1)
        .Value = "  Subtotal  " & brokerName
        .Font.Color = RGB(200, 160, 0)
        .Font.Bold = True
        .Font.Name = "Consolas"
        .Font.Size = 9
        .HorizontalAlignment = xlLeft
    End With
    With ws.cells(r, 7)
        .Value = brokerMkt
        .NumberFormat = "$#,##0"
        .Font.Color = RGB(200, 160, 0)
        .Font.Bold = True
    End With
    With ws.cells(r, 12)
        .Value = brokerUnrl
        .NumberFormat = "$#,##0;-$#,##0"
        .Font.Color = PnLColor(Round(brokerUnrl, 0))
        .Font.Bold = True
    End With
End Sub

' ------------------------------------------------------------
Private Sub WriteOnePositionRow(ws As Worksheet, r As Long, i As Long, _
                                 posData() As Variant, totalMkt As Double, _
                                 swingRiskMap As Object)
    Dim rowBg As Long
    rowBg = IIf(i Mod 2 = 1, RGB(15, 15, 15), RGB(22, 22, 22))

    With ws.Range(ws.cells(r, 1), ws.cells(r, 17))
        .Interior.Color = rowBg
        .Font.Name = "Consolas"
        .Font.Size = 10
        .Font.Color = RGB(210, 210, 210)
        .HorizontalAlignment = xlCenter
    End With
    ws.Rows(r).RowHeight = 18

    Dim tickerCode As String: tickerCode = CStr(posData(i, 1))
    Dim nm         As String: nm = posData(i, 3)
    Dim entryDt    As Date:   entryDt = posData(i, 4)
    Dim days       As Long:   days = posData(i, 5)
    Dim broker     As String: broker = posData(i, 6)
    Dim Sector     As String: Sector = posData(i, 7)
    Dim netExpos   As Double: netExpos = posData(i, 8)
    Dim shares     As Double: shares = posData(i, 9)
    Dim entryPx    As Double: entryPx = posData(i, 10)
    Dim lastPx     As Double: lastPx = posData(i, 11)
    Dim chgPct     As Double: chgPct = posData(i, 12)
    Dim unrlPnl    As Double: unrlPnl = posData(i, 13)
    Dim Beta       As Double: Beta = posData(i, 14)
    Dim beta30d    As Double: beta30d = posData(i, 15)
    Dim pTgt       As Double: pTgt = posData(i, 16)

    Dim wtPct As Double: If totalMkt > 0 Then wtPct = netExpos / totalMkt
    Dim wBeta As Double: wBeta = wtPct * Beta

    ws.cells(r, 1).Value = tickerCode
    ws.cells(r, 1).Font.Color = RGB(255, 192, 0)
    ws.cells(r, 1).Font.Bold = True

    ws.cells(r, 2).Value = nm
    ws.cells(r, 2).HorizontalAlignment = xlLeft
    ws.cells(r, 2).Font.Color = RGB(221, 221, 221)

    ws.cells(r, 3).Value = entryDt
    ws.cells(r, 3).NumberFormat = "m/d/yyyy"

    ws.cells(r, 4).Value = days

    ws.cells(r, 5).Value = broker
    ws.cells(r, 5).Font.Color = RGB(180, 180, 180)

    ws.cells(r, 6).Value = Sector
    ws.cells(r, 6).Font.Color = RGB(180, 180, 180)

    ws.cells(r, 7).Value = netExpos
    ws.cells(r, 7).NumberFormat = "$#,##0"

    ws.cells(r, 8).Value = shares
    ws.cells(r, 8).NumberFormat = "#,##0.##"

    ws.cells(r, 9).Value = entryPx
    ws.cells(r, 9).NumberFormat = "#,##0.00"

    ws.cells(r, 10).Value = lastPx
    ws.cells(r, 10).NumberFormat = "#,##0.00"
    ws.cells(r, 10).Font.Color = RGB(221, 221, 221)

    ws.cells(r, 11).Value = chgPct
    ws.cells(r, 11).NumberFormat = "0.00%"
    ws.cells(r, 11).Font.Bold = True
    ws.cells(r, 11).Font.Color = PnLColorMuted(Round(unrlPnl, 0))

    ws.cells(r, 12).Value = unrlPnl
    ws.cells(r, 12).NumberFormat = "$#,##0;-$#,##0"
    ws.cells(r, 12).Font.Bold = True
    ws.cells(r, 12).Font.Color = PnLColorMuted(Round(unrlPnl, 0))

    ws.cells(r, 13).Value = wtPct
    ws.cells(r, 13).NumberFormat = "0.00%"

    ws.cells(r, 14).Value = wBeta
    ws.cells(r, 14).NumberFormat = "0.000"

    ws.cells(r, 15).Value = beta30d
    ws.cells(r, 15).NumberFormat = "0.000"
    ws.cells(r, 15).Font.Color = RGB(180, 180, 255)

    ws.cells(r, 16).Value = pTgt
    ws.cells(r, 16).NumberFormat = "#,##0"
    ws.cells(r, 16).Font.Color = RGB(255, 192, 0)

    If pTgt > 0 And lastPx > pTgt Then
        ws.Range(ws.cells(r, 1), ws.cells(r, 17)).Interior.Color = RGB(40, 25, 0)
    End If

    If swingRiskMap.Exists(tickerCode) Then
        With ws.cells(r, 17)
            .Value = swingRiskMap(tickerCode)
            .Font.Color = RGB(255, 192, 0)
            .Font.Bold = True
            .HorizontalAlignment = xlCenter
        End With
    End If
End Sub
' ================================================================
'  Disclaimer (bottom rows)
' ================================================================
Private Sub DrawDisclaimer(ws As Worksheet, startRow As Long)
    Dim r As Long: r = startRow + 3

    On Error Resume Next
    ws.Range(ws.cells(r, 1), ws.cells(r, 16)).UnMerge
    ws.Range(ws.cells(r, 1), ws.cells(r, 16)).Merge
    On Error GoTo 0
    ws.cells(r, 1).Value = "On the way in Medium-High  Beta Between 2-2.5 , Remember alwaus do the Eliminate underperformers"
    ws.cells(r, 1).Font.Color = RGB(255, 192, 0)
    ws.cells(r, 1).Font.Italic = True
    ws.cells(r, 1).HorizontalAlignment = xlLeft

    r = r + 1
    On Error Resume Next
    ws.Range(ws.cells(r, 1), ws.cells(r, 16)).UnMerge
    ws.Range(ws.cells(r, 1), ws.cells(r, 16)).Merge
    On Error GoTo 0
    ws.cells(r, 1).Value = "FOCUS ON WHAT MY ACTIONS ARE & DO YOUR OWN WORK ! => WHO SAYS YOU CAN'T FIND BETTER TRADE-IDEAS ? => Narrative + Money Flow + Technicals = Valuation Skyrocket"
    ws.cells(r, 1).Font.Color = RGB(255, 192, 0)
    ws.cells(r, 1).Font.Italic = True
    ws.cells(r, 1).HorizontalAlignment = xlLeft
End Sub

Private Sub CalcPositions(positions As Object, exRate As Double, _
                           ByRef posData() As Variant, _
                           ByRef totalMkt As Double, _
                           ByRef totalCost As Double, _
                           ByRef totalUnrl As Double, _
                           ByRef posCount As Long)
    totalMkt = 0: totalCost = 0: totalUnrl = 0: posCount = 0

    ' ------------------------------------------------------------
    Dim tkv As Variant
    For Each tkv In positions.keys
        Dim d0 As Variant: d0 = positions(tkv)
        If d0(0) > 0 Then
            Dim pp0 As Long: pp0 = InStr(CStr(tkv), "|")
            Dim t0 As String: t0 = IIf(pp0 > 0, Left(CStr(tkv), pp0 - 1), CStr(tkv))
            Dim px0 As Double: px0 = GetStockPrice(t0)
            If px0 > 0 Then
                If GetCurrencyType(t0) = "USD" Then
                    totalMkt = totalMkt + px0 * d0(0) * exRate
                Else
                    totalMkt = totalMkt + px0 * d0(0)
                End If
            End If
        End If
    Next tkv

    ' ------------------------------------------------------------
    Dim validCount As Long: validCount = 0
    For Each tkv In positions.keys
        If positions(tkv)(0) > 0 Then validCount = validCount + 1
    Next tkv
    If validCount = 0 Then Exit Sub

    ' ------------------------------------------------------------
    Dim sortedKeys() As String
    ReDim sortedKeys(0 To validCount - 1)
    Dim ki As Long: ki = 0

    Dim brokerOrder(0 To 1) As String
    brokerOrder(0) = "Default"
    brokerOrder(1) = ChrW(&H570B) & ChrW(&H6CF0) & ChrW(&H4E16) & ChrW(&H83EF)  ' Cathay United

    Dim bi As Long
    For bi = 0 To 1
        Dim bName As String: bName = brokerOrder(bi)
        Dim twStart As Long: twStart = ki

        ' -- broker -- x --
        For Each tkv In positions.keys
            If positions(tkv)(0) > 0 Then
                Dim pp1 As Long: pp1 = InStr(CStr(tkv), "|")
                If pp1 > 0 Then
                    Dim t1 As String: t1 = Left(CStr(tkv), pp1 - 1)
                    Dim b1 As String: b1 = Mid(CStr(tkv), pp1 + 1)
                    If b1 = bName And GetCurrencyType(t1) = "TWD" Then
                        sortedKeys(ki) = CStr(tkv): ki = ki + 1
                    End If
                End If
            End If
        Next tkv

        ' ------------------------------------------------------------
        Dim sa As Long, sb As Long, tmp1 As String
        For sa = twStart To ki - 2
            For sb = sa + 1 To ki - 1
                Dim ppA As Long: ppA = InStr(sortedKeys(sa), "|")
                Dim ppB As Long: ppB = InStr(sortedKeys(sb), "|")
                Dim cA As String: cA = Left(sortedKeys(sa), ppA - 1)
                Dim cB As String: cB = Left(sortedKeys(sb), ppB - 1)
                If IsNumeric(cA) And IsNumeric(cB) Then
                    If val(cA) > val(cB) Then
                        tmp1 = sortedKeys(sa): sortedKeys(sa) = sortedKeys(sb): sortedKeys(sb) = tmp1
                    End If
                End If
            Next sb
        Next sa

        Dim usStart As Long: usStart = ki

        ' -- broker --
        For Each tkv In positions.keys
            If positions(tkv)(0) > 0 Then
                Dim pp2 As Long: pp2 = InStr(CStr(tkv), "|")
                If pp2 > 0 Then
                    Dim t2 As String: t2 = Left(CStr(tkv), pp2 - 1)
                    Dim b2 As String: b2 = Mid(CStr(tkv), pp2 + 1)
                    If b2 = bName And GetCurrencyType(t2) = "USD" Then
                        sortedKeys(ki) = CStr(tkv): ki = ki + 1
                    End If
                End If
            End If
        Next tkv

        ' ------------------------------------------------------------
        Dim tmp2 As String
        For sa = usStart To ki - 2
            For sb = sa + 1 To ki - 1
                Dim ppA2 As Long: ppA2 = InStr(sortedKeys(sa), "|")
                Dim ppB2 As Long: ppB2 = InStr(sortedKeys(sb), "|")
                Dim cA2 As String: cA2 = Left(sortedKeys(sa), ppA2 - 1)
                Dim cB2 As String: cB2 = Left(sortedKeys(sb), ppB2 - 1)
                If cA2 > cB2 Then
                    tmp2 = sortedKeys(sa): sortedKeys(sa) = sortedKeys(sb): sortedKeys(sb) = tmp2
                End If
            Next sb
        Next sa
    Next bi

    ' ------------------------------------------------------------
    Dim addedKeys As Object: Set addedKeys = CreateObject("Scripting.Dictionary")
    Dim si2 As Long
    For si2 = 0 To ki - 1: addedKeys(sortedKeys(si2)) = 1: Next si2
    For Each tkv In positions.keys
        If positions(tkv)(0) > 0 Then
            If Not addedKeys.Exists(CStr(tkv)) Then
                sortedKeys(ki) = CStr(tkv): ki = ki + 1
            End If
        End If
    Next tkv

    ' ------------------------------------------------------------
    ReDim posData(1 To validCount, 1 To 16)
    posCount = 0

    Dim si As Long
    For si = 0 To validCount - 1
        Dim compositeKey As String: compositeKey = sortedKeys(si)
        Dim pipePos As Long: pipePos = InStr(compositeKey, "|")
        Dim tickerStr As String: tickerStr = Left(compositeKey, pipePos - 1)
        Dim brokerPart As String: brokerPart = Mid(compositeKey, pipePos + 1)

        Dim d As Variant: d = positions(compositeKey)
        Dim netQty As Double: netQty = d(0)
        If netQty <= 0 Then GoTo NextSi

        posCount = posCount + 1
        Dim p As Long: p = posCount

        Dim totCost As Double: totCost = d(2)
        Dim entryDt As Date
        If d(5) > 0 Then entryDt = CDate(d(5)) Else entryDt = Date

        Dim AvgCost As Double
        If netQty > 0 Then AvgCost = totCost / netQty

        Dim curType   As String: curType = GetCurrencyType(tickerStr)
        Dim livePrice As Double: livePrice = GetStockPrice(tickerStr)

        Dim mktVal As Double, unrlPnl As Double
        If livePrice > 0 Then
            mktVal = livePrice * netQty
            unrlPnl = mktVal - totCost
        Else
            mktVal = totCost: unrlPnl = 0
        End If

        Dim mktTWD As Double, unrlTWD As Double
        If curType = "USD" Then
            mktTWD = mktVal * exRate: unrlTWD = unrlPnl * exRate
        Else
            mktTWD = mktVal: unrlTWD = unrlPnl
        End If

        totalCost = totalCost + IIf(curType = "USD", totCost * exRate, totCost)
        totalUnrl = totalUnrl + unrlTWD

        Dim chgPct As Double
        If AvgCost > 0 And livePrice > 0 Then chgPct = (livePrice - AvgCost) / AvgCost

        posData(p, 1) = tickerStr
        posData(p, 2) = tickerStr
        posData(p, 3) = GetCompanyName(tickerStr)
        posData(p, 4) = entryDt
        posData(p, 5) = Date - entryDt + 1
        posData(p, 6) = brokerPart
        posData(p, 7) = CStr(d(6))
        posData(p, 8) = mktTWD
        posData(p, 9) = netQty
        posData(p, 10) = AvgCost
        posData(p, 11) = livePrice
        posData(p, 12) = chgPct
        posData(p, 13) = unrlTWD
        posData(p, 14) = d(9)
        posData(p, 15) = d(9)
        posData(p, 16) = d(7)
NextSi:
    Next si
End Sub
Private Function BuildPositions(wsTr As Worksheet) As Object
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")

    Dim lastRow As Long
    lastRow = wsTr.cells(wsTr.Rows.count, "A").End(xlUp).row
    If lastRow < 2 Then Set BuildPositions = dict: Exit Function

    Dim data As Variant
    data = wsTr.Range("B2:N" & lastRow).Value   ' -- M -- N

    Dim i As Long
    For i = 1 To UBound(data, 1)
        Dim Ticker As String: Ticker = CStr(data(i, 2))
        Dim Action As String: Action = CStr(data(i, 3))
        Dim shares As Double: shares = val(data(i, 4))
        Dim netAmt As Double: netAmt = val(data(i, 8))
        Dim Sector As String: Sector = CStr(data(i, 9))
        Dim pTgt   As Double: pTgt = val(data(i, 10))
        Dim strat  As String: strat = CStr(data(i, 11))
        Dim Beta   As Double: Beta = val(data(i, 12))
        Dim broker As String
        On Error Resume Next
        broker = Trim(CStr(data(i, 13)))
        On Error GoTo 0
        If broker = "" Then broker = "Default"

        If Ticker <> "" Then
            Dim posKey As String: posKey = Ticker & "|" & broker
            If Not dict.Exists(posKey) Then
                dict.Add posKey, Array(0#, 0#, 0#, 0#, 0#, 0#, "", 0#, "", 0#, broker)
            End If
            Dim d As Variant: d = dict(posKey)
            If pTgt > 0 Then d(7) = pTgt
            If strat <> "" Then d(8) = strat
            If Beta <> 0 Then d(9) = Beta
            If Sector <> "" Then d(6) = Sector
            d(10) = broker

            Select Case UCase(Action)
                Case "BUY"
                d(0) = d(0) + shares
                d(2) = d(2) + netAmt          ' --  -- ٭ -- G -- Shares
                If d(5) = 0 Then d(5) = CLng(CDate(data(i, 1)))
                Case "SELL"
                    Dim avg As Double
                    If d(0) > 0 Then avg = d(2) / d(0) Else avg = 0
                    d(0) = d(0) - shares
                    d(2) = d(2) - avg * shares
                    If d(0) <= 0.0001 Then d(0) = 0: d(2) = 0: d(5) = 0
                Case "ADJUSTCOST"
                    d(2) = d(2) + netAmt
            End Select
            dict(posKey) = d
        End If
    Next i
    Set BuildPositions = dict
End Function
' ================================================================
'  Portfolio Beta
' ================================================================
Private Function CalcPortBeta(posData() As Variant) As Double
    CalcPortBeta = 0
    If Not IsArray(posData) Then Exit Function
    On Error Resume Next
    Dim n As Long: n = UBound(posData, 1)
    If Err.Number <> 0 Or n < 1 Then Exit Function
    On Error GoTo 0
    
    Dim totalMkt As Double, totalBeta As Double
    Dim i As Long
    For i = 1 To n
        totalMkt = totalMkt + posData(i, 8)
        totalBeta = totalBeta + posData(i, 8) * posData(i, 14)
    Next i
    If totalMkt > 0 Then CalcPortBeta = totalBeta / totalMkt
End Function



' ------------------------------------------------------------
Private Function FindNearestPrice(ByRef dict As Object, ByVal targetDate As Date) As Double
    Dim checkDate As Date
    checkDate = targetDate
    Dim daysBack As Integer
    
    ' ------------------------------------------------------------
    For daysBack = 1 To 10
        checkDate = checkDate - 1
        If dict.Exists(checkDate) Then
            FindNearestPrice = dict(checkDate)
            Exit Function
        End If
    Next daysBack
    
    FindNearestPrice = 0
End Function
' ================================================================
'  Helpers
' ================================================================
Private Function PnLColor(v As Double) As Long
    If v > 0 Then
        PnLColor = RGB(255, 80, 80)
    ElseIf v < 0 Then
        PnLColor = RGB(0, 210, 100)
    Else
        PnLColor = RGB(200, 200, 200)
    End If
End Function

' Muted variant used ONLY by the RR4 holdings rows (columns K/L).
' Profit and flat share one grey; losses get a soft red so the eye lands
' on what is bleeding. Feed it the UNRL PNL value for BOTH columns so a
' short position cannot paint % CHG and PNL in opposite colours.
Private Function PnLColorMuted(v As Double) As Long
    If v < 0 Then
        PnLColorMuted = RGB(255, 130, 130)
    Else
        PnLColorMuted = RGB(200, 200, 200)
    End If
End Function

Private Function GetExRate(ws As Worksheet) As Double
    On Error Resume Next
    GetExRate = val(ws.Range("B2").Value)
    On Error GoTo 0
    If GetExRate <= 0 Then GetExRate = 31.6
End Function

Private Function GetInceptionDate(ws As Worksheet) As Date
    On Error Resume Next
    GetInceptionDate = CDate(ws.Range("InceptionDate").Value)
    On Error GoTo 0
    If GetInceptionDate <= DateSerial(2000, 1, 1) Then GetInceptionDate = DateSerial(2026, 8, 1)
End Function

Private Function GetStartingCapital(ws As Worksheet) As Double
    On Error Resume Next
    GetStartingCapital = val(ws.Range("StartingCapital").Value)
    On Error GoTo 0
    If GetStartingCapital <= 0 Then GetStartingCapital = 600000
End Function

' ================================================================
'  One-time config setup
' ================================================================
Sub SetupPortfolioConfig()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(SH_PORT)

    Dim inDate As String
    inDate = InputBox("Inception Date (yyyy/m/d):", "Config", "2026/8/1")
    If Not IsDate(inDate) Then MsgBox "Invalid date format.": Exit Sub

    Dim startCap As String
    startCap = InputBox("Starting Capital (TWD, e.g. 600000):", "Config", "600000")
    If Not IsNumeric(startCap) Then MsgBox "Invalid number.": Exit Sub

    On Error Resume Next
    ThisWorkbook.names("InceptionDate").Delete
    ThisWorkbook.names("StartingCapital").Delete
    On Error GoTo 0

    ws.Range("S1").Value = CDate(inDate)
    ws.Range("S2").Value = CDbl(startCap)
    ThisWorkbook.names.Add "InceptionDate", ws.Range("S1")
    ThisWorkbook.names.Add "StartingCapital", ws.Range("S2")
    ws.Columns("S").Hidden = True

    MsgBox "Config saved!" & vbCr & "Inception: " & inDate & vbCr & "Capital: " & startCap, vbInformation
End Sub

' ================================================================
'  Auto-refresh
' ================================================================
Sub StartAutoRefresh()
    g_AutoOn = True
    Call RebuildPortfolioDashboard
    g_NextTime = Now + TimeSerial(0, 1, 0)
    Application.OnTime g_NextTime, "AutoRefreshLoop"
    MsgBox "Auto-refresh started (every 60 sec)", vbInformation
End Sub

Sub AutoRefreshLoop()
    If Not g_AutoOn Then Exit Sub
    Call RebuildPortfolioDashboard
    g_NextTime = Now + TimeSerial(0, 1, 0)
    Application.OnTime g_NextTime, "AutoRefreshLoop"
End Sub

Sub StopAutoRefresh()
    g_AutoOn = False
    On Error Resume Next
    Application.OnTime g_NextTime, "AutoRefreshLoop", , False
    On Error GoTo 0
    Application.StatusBar = False
    MsgBox "Auto-refresh stopped.", vbInformation
End Sub
Private Sub WriteKV(ws As Worksheet, r As Long, label As String, _
                    val As Double, fmt As String, colorPnL As Boolean)
    ws.cells(r, 1).Value = label
    ws.cells(r, 1).Font.Color = RGB(150, 150, 150)
    ws.cells(r, 2).Value = val
    ws.cells(r, 2).NumberFormat = fmt
    ws.cells(r, 2).Font.Bold = True
    If colorPnL Then
        ws.cells(r, 2).Font.Color = PnLColor(val)
    Else
        ws.cells(r, 2).Font.Color = RGB(221, 221, 221)
    End If
End Sub


Sub BuildHoldingsCorrelation()
    Application.ScreenUpdating = False

    Dim wsTr As Worksheet
    Set wsTr = ThisWorkbook.Sheets("Transactions")

    Dim positions As Object
    Set positions = BuildPositions(wsTr)

    ' ------------------------------------------------------------
    Dim uniqTickers As Object: Set uniqTickers = CreateObject("Scripting.Dictionary")
    Dim tkv As Variant
    For Each tkv In positions.keys
        If positions(tkv)(0) > 0 Then
            Dim pkStr As String: pkStr = CStr(tkv)
            Dim barPos As Long: barPos = InStr(pkStr, "|")
            Dim pureT As String
            If barPos > 0 Then pureT = Left(pkStr, barPos - 1) Else pureT = pkStr
            pureT = UCase(Trim(pureT))
            If pureT <> "" And Not uniqTickers.Exists(pureT) Then
                uniqTickers.Add pureT, 1
            End If
        End If
    Next tkv

    Dim tickerCount As Long: tickerCount = uniqTickers.count
    If tickerCount < 2 Then
        MsgBox "Need at least 2 holdings.", vbInformation
        Application.ScreenUpdating = True: Exit Sub
    End If

    ' ------------------------------------------------------------
    Dim rawTickers() As String
    ReDim rawTickers(0 To tickerCount - 1)
    Dim ki As Long: ki = 0
    For Each tkv In uniqTickers.keys
        rawTickers(ki) = CStr(tkv): ki = ki + 1
    Next tkv

    ' ------------------------------------------------------------
    ' ------------------------------------------------------------
    Dim tickers() As String
    ReDim tickers(0 To tickerCount - 1)
    Dim sortIdx() As Long
    ReDim sortIdx(0 To tickerCount - 1)

    ' Step 1: -- x -- index -- J
    Dim twCount As Long: twCount = 0
    Dim i As Long
    For i = 0 To tickerCount - 1
        If GetCurrencyType(rawTickers(i)) = "TWD" Then
            sortIdx(twCount) = i
            twCount = twCount + 1
        End If
    Next i

    ' Step 2: fill US tickers next
    Dim usIdx As Long: usIdx = twCount
    For i = 0 To tickerCount - 1
        If GetCurrencyType(rawTickers(i)) = "USD" Then
            sortIdx(usIdx) = i
            usIdx = usIdx + 1
        End If
    Next i

    ' Step 3: sort TW tickers numerically
    Dim a As Long, b As Long, tmpL As Long
    For a = 0 To twCount - 2
        For b = a + 1 To twCount - 1
            Dim codeA As String: codeA = Replace(Replace(UCase(rawTickers(sortIdx(a))), ".TWO", ""), ".TW", "")
            Dim codeB As String: codeB = Replace(Replace(UCase(rawTickers(sortIdx(b))), ".TWO", ""), ".TW", "")
            If IsNumeric(codeA) And IsNumeric(codeB) Then
                If val(codeA) > val(codeB) Then
                    tmpL = sortIdx(a): sortIdx(a) = sortIdx(b): sortIdx(b) = tmpL
                End If
            End If
        Next b
    Next a

    ' Step 4: sort US tickers alphabetically
    For a = twCount To tickerCount - 2
        For b = a + 1 To tickerCount - 1
            If rawTickers(sortIdx(a)) > rawTickers(sortIdx(b)) Then
                tmpL = sortIdx(a): sortIdx(a) = sortIdx(b): sortIdx(b) = tmpL
            End If
        Next b
    Next a

    ' Step 5: apply sorted order
    For i = 0 To tickerCount - 1
        tickers(i) = rawTickers(sortIdx(i))
    Next i

    ' ------------------------------------------------------------
    Dim priceData() As Object
    ReDim priceData(0 To tickerCount - 1)
    For i = 0 To tickerCount - 1
        Application.StatusBar = "Fetching " & tickers(i) & " (" & (i + 1) & "/" & tickerCount & ")..."
        Set priceData(i) = GetHistoryPrices(tickers(i))
    Next i

    ' ------------------------------------------------------------
    Dim commonDates() As Date
    Dim usedCount As Long
    Call CorrGetCommonDates(priceData, tickerCount, commonDates, usedCount)

    If usedCount < 4 Then
        Dim diagMsg As String
        diagMsg = "Insufficient common trading days (" & usedCount & " day(s))." & vbCr & vbCr & _
                 "History row count per ticker:" & vbCr
        Dim dIdx As Long
        For dIdx = 0 To tickerCount - 1
            Dim dcnt As Long: dcnt = 0
            If Not priceData(dIdx) Is Nothing Then dcnt = priceData(dIdx).count
            diagMsg = diagMsg & "  " & tickers(dIdx) & " : " & dcnt & " rows"
            If dcnt = 0 Then diagMsg = diagMsg & "  <-- API returned nothing"
            diagMsg = diagMsg & vbCr
        Next dIdx
        diagMsg = diagMsg & vbCr & "Tip: re-import Attach module, save / close / reopen workbook, retry."
        MsgBox diagMsg, vbInformation, "Holdings Correlation"
        Application.ScreenUpdating = True: Exit Sub
    End If

    ' ------------------------------------------------------------
        ' ------------------------------------------------------------
    Dim retCount As Long: retCount = usedCount - 1
    Dim returns() As Double
    ReDim returns(0 To tickerCount - 1, 0 To retCount - 1)

    Dim j As Long
    For i = 0 To tickerCount - 1
        For j = 0 To retCount - 1
            Dim px1 As Double: px1 = 0
            Dim px2 As Double: px2 = 0
            On Error Resume Next
            px1 = priceData(i)(commonDates(j))
            px2 = priceData(i)(commonDates(j + 1))
            On Error GoTo 0
            If px1 > 0 And px2 > 0 Then returns(i, j) = Log(px2 / px1)
        Next j
    Next i

    ' ------------------------------------------------------------
    Dim corrMatrix() As Double
    ReDim corrMatrix(0 To tickerCount - 1, 0 To tickerCount - 1)
    Dim r As Long, c As Long
    For r = 0 To tickerCount - 1
        For c = 0 To tickerCount - 1
            corrMatrix(r, c) = CorrPearson(returns, r, c, retCount)
        Next c
    Next r

    ' ------------------------------------------------------------
    Call HoldingsCorrRenderSheet(tickers, tickerCount, corrMatrix, usedCount, commonDates)

    Application.ScreenUpdating = True
    Application.StatusBar = "HoldingsCorr updated: " & Format(Now, "hh:mm:ss")
     MsgBox "Holdings correlation matrix updated.", vbInformation
End Sub

' ------------------------------------------------------------
Private Sub HoldingsCorrRenderSheet(tickers() As String, tickerCount As Long, _
                                     corrMatrix() As Double, usedCount As Long, _
                                     commonDates() As Date)
    Dim wsC As Worksheet
    On Error Resume Next: Set wsC = ThisWorkbook.Sheets("HoldingsCorr"): On Error GoTo 0
    If wsC Is Nothing Then
        Set wsC = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.count))
        wsC.Name = "HoldingsCorr"
    End If

    wsC.cells.Clear
    With wsC.cells
        .Interior.Color = RGB(0, 0, 0)
        .Font.Color = RGB(200, 200, 200)
        .Font.Name = "Consolas"
        .Font.Size = 9
    End With
    wsC.Activate
    ActiveWindow.DisplayGridlines = False

    ' ------------------------------------------------------------
    Dim retCount As Long: retCount = usedCount - 1
    With wsC.cells(1, 1)
        .Value = "HOLDINGS CORRELATION  |  " & retCount & "-DAY  |  " & _
                 Format(commonDates(0), "m/d/yy") & " ~ " & _
                 Format(commonDates(usedCount - 1), "m/d/yy") & _
                 "  |  " & Format(Now, "hh:mm:ss")
        .Font.Color = RGB(255, 192, 0)
        .Font.Bold = True
        .Font.Size = 11
    End With
    With wsC.Range(wsC.cells(2, 1), wsC.cells(2, tickerCount + 2))
        .Interior.Color = RGB(255, 192, 0)
        .RowHeight = 3
    End With

    ' ------------------------------------------------------------
    Dim HDR As Long: HDR = 4
    wsC.Columns(1).ColumnWidth = 12
    Dim i As Long
    For i = 0 To tickerCount - 1
        wsC.Columns(i + 2).ColumnWidth = 7
    Next i

    ' ------------------------------------------------------------
    With wsC.cells(HDR, 1)
         .Value = "Row / Col"
        .Font.Color = RGB(100, 100, 100)
        .Interior.Color = RGB(10, 10, 10)
        .HorizontalAlignment = xlCenter
    End With

    For i = 0 To tickerCount - 1
        With wsC.cells(HDR, i + 2)
            .Value = tickers(i)
            .Font.Color = RGB(255, 192, 0)
            .Font.Bold = True
            .Font.Size = 8
            .Interior.Color = RGB(10, 10, 10)
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlBottom
        End With
        wsC.Rows(HDR).RowHeight = 20
    Next i

    ' ------------------------------------------------------------
    Dim j As Long
    For i = 0 To tickerCount - 1
        Dim rn As Long: rn = HDR + 1 + i
        wsC.Rows(rn).RowHeight = 18

        ' ------------------------------------------------------------
        With wsC.cells(rn, 1)
            .Value = tickers(i)
            .Font.Color = RGB(255, 192, 0)
            .Font.Bold = True
            .Font.Size = 9
            .HorizontalAlignment = xlCenter
            .Interior.Color = RGB(10, 10, 10)
        End With

        For j = 0 To tickerCount - 1
            Dim cv As Double: cv = corrMatrix(i, j)
            Dim bgC As Long, fgC As Long

            ' ------------------------------------------------------------
            If i = j Then
                ' ------------------------------------------------------------
                bgC = RGB(40, 28, 0)
                fgC = RGB(255, 192, 0)
            ElseIf cv >= 0.2 Then
                ' ------------------------------------------------------------
                If cv >= 0.7 Then
                    bgC = RGB(100, 0, 0):   fgC = RGB(255, 130, 130)
                ElseIf cv >= 0.4 Then
                    bgC = RGB(70, 0, 0):    fgC = RGB(230, 100, 100)
                Else
                    bgC = RGB(45, 5, 5):    fgC = RGB(210, 80, 80)
                End If
            ElseIf cv <= -0.2 Then
                ' ------------------------------------------------------------
                If cv <= -0.7 Then
                    bgC = RGB(0, 70, 20):   fgC = RGB(80, 255, 140)
                ElseIf cv <= -0.4 Then
                    bgC = RGB(0, 50, 15):   fgC = RGB(60, 220, 110)
                Else
                    bgC = RGB(0, 35, 10):   fgC = RGB(40, 200, 90)
                End If
            Else
                ' ------------------------------------------------------------
                bgC = RGB(18, 18, 18):      fgC = RGB(150, 150, 150)
            End If

            With wsC.cells(rn, j + 2)
                .Value = Round(cv, 2)
                .NumberFormat = "0.00"
                .HorizontalAlignment = xlCenter
                .Font.Bold = True
                .Font.Size = 8
                .Interior.Color = bgC
                .Font.Color = fgC
                ' ------------------------------------------------------------
                With .Borders
                    .LineStyle = xlContinuous
                    .Color = RGB(0, 0, 0)
                    .Weight = xlThin
                End With
            End With
        Next j
    Next i

    ' ------------------------------------------------------------
    Dim LR As Long: LR = HDR + tickerCount + 2
    wsC.cells(LR, 1).Value = "LEGEND"
    wsC.cells(LR, 1).Font.Color = RGB(255, 192, 0)
    wsC.cells(LR, 1).Font.Bold = True

    Dim leg As Variant
    leg = Array( _
        Array(">= 0.70  High +corr", RGB(100, 0, 0), RGB(255, 130, 130)), _
        Array("0.40~0.69 Mid +corr", RGB(70, 0, 0), RGB(230, 100, 100)), _
        Array("0.20~0.39 Low +corr", RGB(45, 5, 5), RGB(210, 80, 80)), _
        Array("-0.19~0.19 Neutral", RGB(18, 18, 18), RGB(150, 150, 150)), _
        Array("-0.20~-0.39 Low -corr", RGB(0, 35, 10), RGB(40, 200, 90)), _
        Array("-0.40~-0.69 Mid -corr", RGB(0, 50, 15), RGB(60, 220, 110)), _
        Array("<= -0.70 High -corr", RGB(0, 70, 20), RGB(80, 255, 140)))

    Dim li As Long
    For li = 0 To UBound(leg)
        With wsC.cells(LR + 1 + li, 1)
            .Value = leg(li)(0)
            .Interior.Color = leg(li)(1)
            .Font.Color = leg(li)(2)
            .Font.Bold = True
            .Font.Size = 9
        End With
        wsC.Rows(LR + 1 + li).RowHeight = 16
    Next li

    ' ------------------------------------------------------------

    wsC.Activate
End Sub
' ------------------------------------------------------------
Private Sub CorrGetCommonDates(priceData() As Object, tickerCount As Long, _
                                ByRef commonDates() As Date, ByRef usedCount As Long)
    Dim allDates As Object: Set allDates = CreateObject("Scripting.Dictionary")
    Dim dkey As Variant

    ' ------------------------------------------------------------
    For Each dkey In priceData(0).keys
        allDates(dkey) = 1
    Next dkey

    ' ------------------------------------------------------------
    Dim i As Long
    For i = 1 To tickerCount - 1
        Dim toRemove() As Variant
        Dim removeCount As Long: removeCount = 0
        ReDim toRemove(0 To allDates.count)
        For Each dkey In allDates.keys
            If Not priceData(i).Exists(dkey) Then
                toRemove(removeCount) = dkey
                removeCount = removeCount + 1
            End If
        Next dkey
        Dim k As Long
        For k = 0 To removeCount - 1
            allDates.Remove toRemove(k)
        Next k
    Next i

    Dim dateCount As Long: dateCount = allDates.count
    If dateCount = 0 Then usedCount = 0: Exit Sub

    ' ------------------------------------------------------------
    Dim sortedDates() As Date
    ReDim sortedDates(0 To dateCount - 1)
    Dim di As Long: di = 0
    For Each dkey In allDates.keys
        sortedDates(di) = CDate(dkey): di = di + 1
    Next dkey

    Dim a As Long, b As Long, tmpD As Date
    For a = 0 To dateCount - 2
        For b = a + 1 To dateCount - 1
            If sortedDates(a) > sortedDates(b) Then
                tmpD = sortedDates(a): sortedDates(a) = sortedDates(b): sortedDates(b) = tmpD
            End If
        Next b
    Next a

    ' ------------------------------------------------------------
    ' 180 trading-day window: 181 price points yield 180 log returns.
    ' Short-history tickers fall back to whatever common days exist.
    Dim wantPx As Long: wantPx = 181
    Dim startIdx As Long: startIdx = IIf(dateCount > wantPx, dateCount - wantPx, 0)
    usedCount = dateCount - startIdx

    ReDim commonDates(0 To usedCount - 1)
    For i = 0 To usedCount - 1
        commonDates(i) = sortedDates(startIdx + i)
    Next i
End Sub

' ------------------------------------------------------------
Private Function CorrPearson(returns() As Double, r1 As Long, r2 As Long, n As Long) As Double
    If n < 2 Then CorrPearson = 0: Exit Function

    Dim sumX As Double, sumY As Double, sumXY As Double
    Dim sumX2 As Double, sumY2 As Double, j As Long

    For j = 0 To n - 1
        Dim x As Double: x = returns(r1, j)
        Dim y As Double: y = returns(r2, j)
        sumX = sumX + x:   sumY = sumY + y
        sumXY = sumXY + x * y
        sumX2 = sumX2 + x * x
        sumY2 = sumY2 + y * y
    Next j

    Dim denom As Double
    denom = Sqr((n * sumX2 - sumX ^ 2) * (n * sumY2 - sumY ^ 2))
    If denom = 0 Then CorrPearson = 0 Else CorrPearson = (n * sumXY - sumX * sumY) / denom
End Function

' ------------------------------------------------------------
Private Sub CorrRenderSheet(tickers() As String, tickerCount As Long, _
                             corrMatrix() As Double, usedCount As Long, _
                             commonDates() As Date)
    Dim wsC As Worksheet
    On Error Resume Next: Set wsC = ThisWorkbook.Sheets("HoldingsCorr"): On Error GoTo 0
    If wsC Is Nothing Then
        Set wsC = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.count))
        wsC.Name = "HoldingsCorr"
    End If

    wsC.cells.Clear
    With wsC.cells
        .Interior.Color = RGB(0, 0, 0)
        .Font.Color = RGB(221, 221, 221)
        .Font.Name = "Consolas"
        .Font.Size = 10
    End With
    wsC.Activate
    ActiveWindow.DisplayGridlines = False

    ' ------------------------------------------------------------
    Dim retCount As Long: retCount = usedCount - 1
    With wsC.cells(1, 1)
        .Value = "CORRELATION MATRIX  |  " & retCount & "-DAY RETURNS  |  " & _
                 Format(commonDates(0), "m/d/yy") & " ~ " & _
                 Format(commonDates(usedCount - 1), "m/d/yy") & _
                 "  |  Updated: " & Format(Now, "hh:mm:ss")
        .Font.Color = RGB(255, 192, 0)
        .Font.Bold = True
        .Font.Size = 12
    End With
    With wsC.Range(wsC.cells(2, 1), wsC.cells(2, tickerCount + 3))
        .Interior.Color = RGB(255, 192, 0)
        .RowHeight = 3
    End With

    ' ------------------------------------------------------------
    Dim HDR As Long: HDR = 4
    With wsC.cells(HDR, 1)
         .Value = "ROW / COL"
        .Font.Color = RGB(150, 150, 150)
        .Interior.Color = RGB(10, 10, 10)
        .HorizontalAlignment = xlCenter
    End With

    Dim i As Long, j As Long
    For i = 0 To tickerCount - 1
        With wsC.cells(HDR, i + 2)
            .Value = tickers(i)
            .Font.Color = RGB(255, 192, 0)
            .Font.Bold = True
            .Interior.Color = RGB(10, 10, 10)
            .HorizontalAlignment = xlCenter
        End With
    Next i

    ' ------------------------------------------------------------
    For i = 0 To tickerCount - 1
        Dim rn As Long: rn = HDR + 1 + i
        wsC.Rows(rn).RowHeight = 20

        With wsC.cells(rn, 1)
            .Value = tickers(i)
            .Font.Color = RGB(255, 192, 0)
            .Font.Bold = True
            .HorizontalAlignment = xlCenter
            .Interior.Color = RGB(10, 10, 10)
        End With

        For j = 0 To tickerCount - 1
            Dim cv As Double: cv = corrMatrix(i, j)
            Dim bgC As Long, fgC As Long

            If i = j Then                                  ' -- 﨤 -- u -- G -- ۤv -- P -- ۤv = 1.00
                bgC = RGB(40, 30, 0):   fgC = RGB(255, 192, 0)
            ElseIf cv >= 0.7 Then                          ' -- ץ --
                bgC = RGB(90, 10, 10):  fgC = RGB(255, 100, 100)
            ElseIf cv >= 0.4 Then                          ' -- ץ --
                bgC = RGB(55, 20, 0):   fgC = RGB(230, 140, 70)
            ElseIf cv >= 0.1 Then                          ' -- C -- ץ --
                bgC = RGB(22, 22, 22):  fgC = RGB(200, 200, 160)
            ElseIf cv >= -0.1 Then                         ' ------------------------------------------------------------
                bgC = RGB(15, 15, 15):  fgC = RGB(170, 170, 170)
            ElseIf cv >= -0.4 Then                         ' -- C -- ׭t --
                bgC = RGB(0, 22, 28):   fgC = RGB(80, 200, 200)
            Else                                           ' -- ׭t --
                bgC = RGB(0, 45, 20):   fgC = RGB(0, 210, 100)
            End If

            With wsC.cells(rn, j + 2)
                .Value = Round(cv, 2)
                .NumberFormat = "0.00"
                .HorizontalAlignment = xlCenter
                .Font.Bold = True
                .Interior.Color = bgC
                .Font.Color = fgC
            End With
        Next j
    Next i

    ' ------------------------------------------------------------
    Dim LR As Long: LR = HDR + tickerCount + 3
    wsC.cells(LR, 1).Value = "COLOR LEGEND"
    wsC.cells(LR, 1).Font.Color = RGB(255, 192, 0)
    wsC.cells(LR, 1).Font.Bold = True

    Dim leg As Variant
    leg = Array( _
        Array("? 0.70   HIGH POSITIVE", RGB(90, 10, 10), RGB(255, 100, 100)), _
        Array("0.40 ~ 0.69   MODERATE POSITIVE", RGB(55, 20, 0), RGB(230, 140, 70)), _
        Array("0.10 ~ 0.39   LOW POSITIVE", RGB(22, 22, 22), RGB(200, 200, 160)), _
        Array("-0.09 ~ 0.09   NEUTRAL", RGB(15, 15, 15), RGB(170, 170, 170)), _
        Array("-0.10 ~ -0.39   LOW NEGATIVE", RGB(0, 22, 28), RGB(80, 200, 200)), _
        Array("? -0.40   HIGH NEGATIVE", RGB(0, 45, 20), RGB(0, 210, 100)) _
    )

    Dim li As Long
    For li = 0 To UBound(leg)
        With wsC.cells(LR + 1 + li, 1)
            .Value = leg(li)(0)
            .Interior.Color = leg(li)(1)
            .Font.Color = leg(li)(2)
            .Font.Bold = True
            .HorizontalAlignment = xlLeft
        End With
        wsC.Rows(LR + 1 + li).RowHeight = 18
    Next li

    ' ------------------------------------------------------------
    wsC.Columns(1).ColumnWidth = 20
    Dim col As Long
    For col = 2 To tickerCount + 1
        wsC.Columns(col).ColumnWidth = IIf(tickerCount > 10, 8, 10)
    Next col

    wsC.Activate
End Sub


