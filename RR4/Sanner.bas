Attribute VB_Name = "Sanner"
Option Explicit

' ================================================================
'  MARKET SCANNER v2.1
' ================================================================

Private Const W_BIAS20  As Double = 0.25
Private Const W_BIAS60  As Double = 0.25
Private Const W_BIAS5   As Double = 0.1
Private Const W_DISTHI  As Double = 0.15
Private Const W_CHG180  As Double = 0.15
Private Const W_BIAS120 As Double = 0.1

' ================================================================
'  MAIN ENTRY
' ================================================================
Sub RunCompanyResearch(market As String, Sector As String)
    Dim wsRes As Worksheet
    On Error Resume Next
    Set wsRes = ThisWorkbook.Sheets("Company research")
    On Error GoTo 0
    If wsRes Is Nothing Then MsgBox "Sheet 'Company research' not found.", vbExclamation: Exit Sub

    Dim tickerList As Variant
    tickerList = GetSectorTickers(market, Sector)
    If IsEmpty(tickerList) Then
        MsgBox "No tickers found for: " & market & " / " & Sector, vbExclamation
        Exit Sub
    End If

    Application.ScreenUpdating = False

    With wsRes
        .Range("C:S").Clear
        .Range("C:S").Interior.Color = RGB(0, 0, 0)
        .Range("C:S").Font.Color = RGB(200, 200, 200)
        .Range("C:S").Font.Name = "Consolas"
        .Range("C:S").Font.Size = 9
    End With

    With wsRes.cells(1, 3)
        .Value = "MARKET SCANNER  —  " & market & "  |  " & Sector & "  |  " & Format(Now, "yyyy/mm/dd hh:mm")
        .Font.Color = RGB(255, 192, 0)
        .Font.Bold = True
        .Font.Size = 13
    End With
    wsRes.Range(wsRes.cells(1, 3), wsRes.cells(1, 19)).Interior.Color = RGB(10, 10, 10)

    Dim hdrs As Variant
    hdrs = Array("TICKER", "COMPANY", "SECTOR", "PRICE", _
                 "HIGH DIST%", "180D CHG%", "DAY CHG%", _
                 "BIAS5", "BIAS20", "BIAS60", "BIAS120", "BIAS240", _
                 "PE", "FWD PE", "PEG", "TREND", "SCORE")

    Dim ci As Integer
    For ci = 0 To UBound(hdrs)
        With wsRes.cells(2, ci + 3)
            .Value = hdrs(ci)
            .Font.Color = RGB(255, 192, 0)
            .Font.Bold = True
            .Interior.Color = RGB(10, 10, 10)
            .HorizontalAlignment = xlCenter
        End With
    Next ci
    With wsRes.Range(wsRes.cells(2, 3), wsRes.cells(2, 19)).Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Color = RGB(255, 192, 0)
        .Weight = xlThin
    End With

    Dim rowNum As Long: rowNum = 3
    Dim i As Long

    For i = LBound(tickerList) To UBound(tickerList)
        Dim tkr As String: tkr = CStr(tickerList(i))
        Application.StatusBar = "Scanning [" & i + 1 & "/" & (UBound(tickerList) + 1) & "] " & tkr
        DoEvents

        Dim tech As TechIndicators
        tech = GetTechData(tkr)
        Dim nm As String: nm = GetCompanyName(tkr)

        Dim rowBg As Long: rowBg = IIf((rowNum Mod 2) = 1, RGB(12, 12, 12), RGB(20, 20, 20))
        With wsRes.Range(wsRes.cells(rowNum, 3), wsRes.cells(rowNum, 19))
            .Interior.Color = rowBg
            .HorizontalAlignment = xlCenter
            .Font.Name = "Consolas"
            .Font.Size = 9
        End With
        wsRes.Rows(rowNum).RowHeight = 16

        With wsRes.cells(rowNum, 3)
            .Value = tkr
            .Font.Color = RGB(255, 192, 0)
            .Font.Bold = True
        End With

        wsRes.cells(rowNum, 4).Value = nm
        wsRes.cells(rowNum, 4).HorizontalAlignment = xlLeft
        wsRes.cells(rowNum, 4).Font.Color = RGB(210, 210, 210)
        wsRes.cells(rowNum, 5).Value = Sector
        wsRes.cells(rowNum, 5).Font.Color = RGB(150, 150, 150)

        If tech.Success Then
            wsRes.cells(rowNum, 6).Value = tech.price
            wsRes.cells(rowNum, 6).NumberFormat = "#,##0.00"

            WriteColoredPct wsRes.cells(rowNum, 7), tech.DistFromHigh, "0.00%", False
            WriteColoredPct wsRes.cells(rowNum, 8), tech.Change180D, "0.00%", True
            WriteColoredPct wsRes.cells(rowNum, 9), tech.ChangePercent, "0.00%", True
            WriteColoredPct wsRes.cells(rowNum, 10), tech.Bias5, "0.00%", True
            WriteColoredPct wsRes.cells(rowNum, 11), tech.Bias20, "0.00%", True
            WriteColoredPct wsRes.cells(rowNum, 12), tech.Bias60, "0.00%", True
            WriteColoredPct wsRes.cells(rowNum, 13), tech.Bias120, "0.00%", True
            WriteColoredPct wsRes.cells(rowNum, 14), tech.Bias240, "0.00%", True

            ' ── PE / FWD PE / PEG — fetch directly ──────────────
            Dim pe As Double, fpe As Double, peg As Double
            Call FetchFundamentals(tkr, pe, fpe, peg)

            If pe > 0 Then
                wsRes.cells(rowNum, 15).Value = pe
                wsRes.cells(rowNum, 15).NumberFormat = "0.0"
                wsRes.cells(rowNum, 15).Font.Color = RGB(180, 180, 255)
            Else
                wsRes.cells(rowNum, 15).Value = "N/A"
                wsRes.cells(rowNum, 15).Font.Color = RGB(100, 100, 100)
            End If

            If fpe > 0 Then
                wsRes.cells(rowNum, 16).Value = fpe
                wsRes.cells(rowNum, 16).NumberFormat = "0.0"
                wsRes.cells(rowNum, 16).Font.Color = RGB(180, 180, 255)
            Else
                wsRes.cells(rowNum, 16).Value = "N/A"
                wsRes.cells(rowNum, 16).Font.Color = RGB(100, 100, 100)
            End If

            If peg > 0 Then
                wsRes.cells(rowNum, 17).Value = peg
                wsRes.cells(rowNum, 17).NumberFormat = "0.00"
                wsRes.cells(rowNum, 17).Font.Color = IIf(peg < 1.5, RGB(0, 220, 100), RGB(200, 150, 100))
            Else
                wsRes.cells(rowNum, 17).Value = "N/A"
                wsRes.cells(rowNum, 17).Font.Color = RGB(100, 100, 100)
            End If

            Call ApplyScanTrend(wsRes.cells(rowNum, 18), tech)

            Dim score As Double: score = CalcMomentumScore(tech)
            With wsRes.cells(rowNum, 19)
                .Value = score
                .NumberFormat = "0"
                .Font.Bold = True
                Call ApplyScoreColor(wsRes.cells(rowNum, 19), score)
            End With
        Else
            wsRes.cells(rowNum, 6).Value = "ERR"
            wsRes.cells(rowNum, 6).Font.Color = RGB(255, 102, 0)
            wsRes.cells(rowNum, 18).Value = "N/A"
            wsRes.cells(rowNum, 19).Value = 0
        End If
        rowNum = rowNum + 1
    Next i

    ' Sort by Score
    If rowNum > 4 Then
        wsRes.Range(wsRes.cells(3, 3), wsRes.cells(rowNum - 1, 19)).Sort _
            Key1:=wsRes.cells(3, 19), Order1:=xlDescending, Header:=xlNo
    End If

    ' Summary
    Call DrawScanSummary(wsRes, rowNum + 1, market, Sector, tickerList, rowNum - 3)

    ' Column widths
    wsRes.Columns("C:S").AutoFit
    wsRes.Columns("C").ColumnWidth = 8    ' Ticker: 1/3 of default (~8 units)
    wsRes.Columns("D").ColumnWidth = 28   ' Company name

    Application.StatusBar = "Scan complete — " & Format(Now, "hh:mm:ss")
    Application.ScreenUpdating = True
    MsgBox "Scan complete! " & (rowNum - 3) & " stocks scanned.", vbInformation
End Sub

' ================================================================
'  Dedicated fundamentals fetcher — handles both JSON formats
'  {"raw": 25.3} AND direct "key": 25.3
' ================================================================
Sub FetchFundamentals(Ticker As String, ByRef pe As Double, ByRef fpe As Double, ByRef peg As Double)
    pe = 0: fpe = 0: peg = 0

    Ticker = UCase(Trim(Ticker))
    If InStr(Ticker, ".TW") = 0 And InStr(Ticker, ".TWO") = 0 And IsNumeric(Ticker) Then
        Ticker = Ticker & ".TW"
    End If

    Dim http As Object
    On Error Resume Next
    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
    If http Is Nothing Then Set http = CreateObject("MSXML2.XMLHTTP")
    On Error GoTo 0

    ' Try quoteSummary price module first (most reliable for PE)
    Dim url As String
    url = "https://query1.finance.yahoo.com/v10/finance/quoteSummary/" & Ticker & _
          "?modules=summaryDetail,defaultKeyStatistics"

    Dim resp As String
    On Error Resume Next
    With http
        .Open "GET", url, False
        .setRequestHeader "User-Agent", "Mozilla/5.0"
        .send
        If .status = 200 Then resp = .responseText
    End With
    On Error GoTo 0

    If Len(resp) < 50 Then Exit Sub

    ' Clean whitespace
    resp = Replace(Replace(Replace(resp, " ", ""), vbCr, ""), vbLf, "")

    ' Extract with both formats
    pe = ExtractNum(resp, "trailingPE")
    fpe = ExtractNum(resp, "forwardPE")
    peg = ExtractNum(resp, "pegRatio")

    ' Fallback: try v8 chart endpoint regularMarketPE
    If pe = 0 Then
        Dim url2 As String
        url2 = "https://query1.finance.yahoo.com/v8/finance/chart/" & Ticker & "?interval=1d&range=1d"
        Dim resp2 As String
        On Error Resume Next
        With http
            .Open "GET", url2, False
            .setRequestHeader "User-Agent", "Mozilla/5.0"
            .send
            If .status = 200 Then resp2 = .responseText
        End With
        On Error GoTo 0
        If Len(resp2) > 50 Then
            resp2 = Replace(Replace(Replace(resp2, " ", ""), vbCr, ""), vbLf, "")
            pe = ExtractNum(resp2, "trailingPE")
            fpe = ExtractNum(resp2, "forwardPE")
        End If
    End If
End Sub

' Handles both {"raw":25.3} and direct 25.3 formats
Private Function ExtractNum(json As String, key As String) As Double
    ExtractNum = 0

    ' Format 1: "key":{"raw":25.3
    Dim s1 As String: s1 = """" & key & """:{""raw"":"
    Dim pos As Long: pos = InStr(json, s1)
    If pos > 0 Then
        Dim vs As Long: vs = pos + Len(s1)
        Dim ve As Long: ve = InStr(vs, json, ",")
        Dim ve2 As Long: ve2 = InStr(vs, json, "}")
        If ve = 0 Or (ve2 > 0 And ve2 < ve) Then ve = ve2
        If ve > vs Then
            Dim tmp As String: tmp = Mid(json, vs, ve - vs)
            If IsNumeric(tmp) And CDbl(tmp) > 0 Then ExtractNum = CDbl(tmp): Exit Function
        End If
    End If

    ' Format 2: "key":25.3
    Dim s2 As String: s2 = """" & key & """"
    pos = InStr(json, s2)
    Do While pos > 0
        Dim afterKey As Long: afterKey = pos + Len(s2)
        ' Skip until colon
        Dim colonPos As Long: colonPos = InStr(afterKey, json, ":")
        If colonPos > afterKey + 3 Then Exit Do  ' colon too far, skip
        If colonPos > 0 Then
            Dim valStart As Long: valStart = colonPos + 1
            ' Skip opening brace if present (means it's an object, not direct)
            If Mid(json, valStart, 1) = "{" Then
                pos = InStr(pos + 1, json, s2)
                GoTo NextOccurrence
            End If
            Dim valEnd As Long: valEnd = InStr(valStart, json, ",")
            Dim valEnd2 As Long: valEnd2 = InStr(valStart, json, "}")
            If valEnd = 0 Or (valEnd2 > 0 And valEnd2 < valEnd) Then valEnd = valEnd2
            If valEnd > valStart Then
                Dim tv As String: tv = Mid(json, valStart, valEnd - valStart)
                tv = Replace(tv, """", "")
                If IsNumeric(tv) And CDbl(tv) > 0 Then ExtractNum = CDbl(tv): Exit Function
            End If
        End If
NextOccurrence:
        pos = InStr(pos + 1, json, s2)
    Loop
End Function

' ================================================================
'  Summary bar
' ================================================================
Private Sub DrawScanSummary(ws As Worksheet, sumRow As Long, _
                             market As String, Sector As String, _
                             tickers As Variant, count As Long)
    Dim cntBull As Long, cntBear As Long, cntNeut As Long
    Dim r As Long
    For r = 3 To 3 + count - 1
        Dim tv As String: tv = ws.cells(r, 18).Value
        If InStr(tv, "BULL") > 0 Then cntBull = cntBull + 1
        If InStr(tv, "BEAR") > 0 Then cntBear = cntBear + 1
        If InStr(tv, "NEUT") > 0 Then cntNeut = cntNeut + 1
    Next r

    With ws.Range(ws.cells(sumRow, 3), ws.cells(sumRow, 19))
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
'  Momentum Score
' ================================================================
Private Function CalcMomentumScore(tech As TechIndicators) As Double
    Dim s As Double
    s = s + NormScore(tech.Bias20, -0.3, 0.3) * 25
    s = s + NormScore(tech.Bias60, -0.3, 0.3) * 25
    s = s + NormScore(tech.Bias5, -0.15, 0.15) * 10
    s = s + NormScore(-tech.DistFromHigh, 0, 0.4) * 15
    s = s + NormScore(tech.Change180D, -0.5, 1) * 15
    s = s + NormScore(tech.Bias120, -0.3, 0.3) * 10
    CalcMomentumScore = WorksheetFunction.Max(0, WorksheetFunction.Min(100, s))
End Function

Private Function NormScore(val As Double, minV As Double, maxV As Double) As Double
    If maxV = minV Then NormScore = 0.5: Exit Function
    NormScore = (val - minV) / (maxV - minV)
    If NormScore < 0 Then NormScore = 0
    If NormScore > 1 Then NormScore = 1
End Function

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

Private Sub ApplyScoreColor(cell As Range, score As Double)
    Select Case True
        Case score >= 75: cell.Font.Color = RGB(255, 60, 60): cell.Interior.Color = RGB(40, 0, 0)
        Case score >= 55: cell.Font.Color = RGB(255, 160, 80): cell.Interior.Color = RGB(25, 10, 0)
        Case score >= 40: cell.Font.Color = RGB(200, 200, 200): cell.Interior.Color = RGB(20, 20, 20)
        Case score >= 25: cell.Font.Color = RGB(100, 200, 120): cell.Interior.Color = RGB(0, 15, 5)
        Case Else: cell.Font.Color = RGB(0, 210, 100): cell.Interior.Color = RGB(0, 25, 10)
    End Select
End Sub

' ================================================================
'  Export helpers
' ================================================================
Sub ExportTopToWatchlist()
    Dim wsRes As Worksheet
    On Error Resume Next: Set wsRes = ThisWorkbook.Sheets("Company research"): On Error GoTo 0
    If wsRes Is Nothing Then MsgBox "Run a scan first.", vbExclamation: Exit Sub

    Dim topN As Long
    topN = CLng(InputBox("Export top N:", "Export", "10"))
    If topN <= 0 Then Exit Sub

    Dim wsWatch As Worksheet
    On Error Resume Next: Set wsWatch = ThisWorkbook.Sheets("Watchlist"): On Error GoTo 0
    If wsWatch Is Nothing Then
        Set wsWatch = ThisWorkbook.Sheets.Add(After:=wsRes)
        wsWatch.Name = "Watchlist"
    End If

    wsWatch.cells.Clear
    With wsWatch.cells: .Interior.Color = RGB(0, 0, 0): .Font.Color = RGB(200, 200, 200): .Font.Name = "Consolas": .Font.Size = 10: End With
    wsWatch.cells(1, 1).Value = "WATCHLIST — " & Format(Now, "yyyy/mm/dd hh:mm"): wsWatch.cells(1, 1).Font.Color = RGB(255, 192, 0): wsWatch.cells(1, 1).Font.Bold = True

    Dim headers As Variant: headers = Array("TICKER", "COMPANY", "SCORE", "TREND", "BIAS20", "BIAS60", "PE", "FWD PE")
    Dim ci As Integer
    For ci = 0 To UBound(headers)
        With wsWatch.cells(2, ci + 1): .Value = headers(ci): .Font.Color = RGB(255, 192, 0): .Font.Bold = True: .Interior.Color = RGB(10, 10, 10): End With
    Next ci

    Dim LR As Long: LR = wsRes.cells(wsRes.Rows.count, 3).End(xlUp).row
    Dim cc As Long: cc = WorksheetFunction.Min(topN, LR - 2)
    Dim wr As Long: wr = 3
    Dim r As Long
    For r = 3 To 3 + cc - 1
        wsWatch.Range(wsWatch.cells(wr, 1), wsWatch.cells(wr, 8)).Interior.Color = IIf((wr Mod 2) = 1, RGB(12, 12, 12), RGB(20, 20, 20))
        wsWatch.cells(wr, 1).Value = wsRes.cells(r, 3).Value: wsWatch.cells(wr, 1).Font.Color = RGB(255, 192, 0): wsWatch.cells(wr, 1).Font.Bold = True
        wsWatch.cells(wr, 2).Value = wsRes.cells(r, 4).Value
        wsWatch.cells(wr, 3).Value = wsRes.cells(r, 19).Value
        wsWatch.cells(wr, 4).Value = wsRes.cells(r, 18).Value
        wsWatch.cells(wr, 5).Value = wsRes.cells(r, 11).Value: wsWatch.cells(wr, 5).NumberFormat = "0.00%"
        wsWatch.cells(wr, 6).Value = wsRes.cells(r, 12).Value: wsWatch.cells(wr, 6).NumberFormat = "0.00%"
        wsWatch.cells(wr, 7).Value = wsRes.cells(r, 15).Value
        wsWatch.cells(wr, 8).Value = wsRes.cells(r, 16).Value
        wr = wr + 1
    Next r
    wsWatch.Columns("A:H").AutoFit: wsWatch.Columns("B").ColumnWidth = 28
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
    Dim LR As Long: LR = wsRes.cells(wsRes.Rows.count, 3).End(xlUp).row
    Dim r As Long, c As Integer, line As String
    For r = 2 To LR
        line = ""
        For c = 3 To 19
            Dim v As String: v = CStr(wsRes.cells(r, c).Value)
            If InStr(v, ",") > 0 Then v = """" & v & """"
            line = line & IIf(c > 3, ",", "") & v
        Next c
        Print #fNum, line
    Next r
    Close #fNum
    MsgBox "Exported: " & path, vbInformation
End Sub

Sub OpenMarketScanner()
    frmMarketScanner.Show
End Sub

Sub ColorizeRange(rng As Range)
    Dim cell As Range
    For Each cell In rng
        ' 外層判斷：確認是否為數字
        If IsNumeric(cell.Value) Then
            
            ' 內層判斷：改為標準的區塊寫法 (Block If)
            If cell.Value > 0 Then
                cell.Font.Color = RGB(255, 100, 100)
            ElseIf cell.Value < 0 Then
                cell.Font.Color = RGB(100, 255, 100)
            End If ' <--- 補上內層 If 的結尾
            
        End If ' <--- 原本外層 IsNumeric 的結尾
    Next cell
End Sub

' ================================================================
'  SECTOR TICKERS — EXPANDED
' ================================================================
Function GetSectorTickers(market As String, Sector As String) As Variant
    Dim arr As Variant: arr = Empty

    If market = "TW" Then
        Select Case Sector
            Case "01. 半導體 - 晶圓代工 (Foundry)"
                arr = Array("2330", "2303", "5347", "6770", "3443", "4961", "3034", "6669")
            Case "02. 半導體 - IC設計 (IC Design)"
                arr = Array("2454", "2379", "3034", "3035", "4961", "2458", "8016", "3529", "6643", "4966", "3661", "6547", "5274", "3130", "6770", "2385", "3044", "4919", "6770", "3017")
            Case "03. 半導體 - 封測 (Packaging)"
                arr = Array("3711", "2449", "6239", "6271", "8150", "3450", "2412", "6257", "6269", "3189", "6274", "3016")
            Case "04. 半導體 - IP/ASIC (Silicon IP)"
                arr = Array("3661", "3443", "3035", "3529", "6669", "6643", "6533", "4966", "3665", "6643")
            Case "05. 電腦週邊 - AI 伺服器 (AI Server)"
                arr = Array("2382", "3231", "2356", "6669", "2317", "2376", "3017", "3697", "6213", "2399", "3704", "6456", "3023", "6197", "3005")
            Case "06. 電腦週邊 - 品牌與代工 (OEM/ODM)"
                arr = Array("2357", "2324", "2353", "2301", "4938", "3231", "2382", "2356", "6669", "2317")
            Case "07. 電腦週邊 - 工業電腦 (IPC)"
                arr = Array("2395", "6414", "8050", "6245", "3540", "5443", "6185", "3548")
            Case "08. 通信網路 -網通設備 (Networking)"
                arr = Array("2345", "3704", "5388", "6285", "3596", "2332", "4906", "4943", "3518", "6415", "3526", "2465", "6456")
            Case "09. 光電業 - 面板與光學 (Panel & Optics)"
                arr = Array("2409", "3481", "3008", "3406", "3376", "6706", "3189", "2384", "3044", "6274", "3296", "6443", "4919")
            Case "10. 電子零組件 - PCB (印刷電路板)"
                arr = Array("3037", "2368", "2313", "4958", "6269", "8046", "3044", "3189", "4938", "2382", "6257", "3042", "8011", "6274")
            Case "11. 電子零組件 - 被動元件 (Passive)"
                arr = Array("2327", "2492", "2456", "3026", "2438", "3557", "6269", "2478", "3018", "2485")
            Case "12. 電子零組件 - 連接器 (Connectors)"
                arr = Array("3533", "3605", "3023", "2345", "6598", "3588", "2308", "6289")
            Case "13. 電子零組件 - 電源供應 (Power Supply)"
                arr = Array("2308", "6278", "6412", "2301", "3483", "6449", "3591", "1537", "4977", "3577")
            Case "14. 金融 - 金控 (FHC)"
                arr = Array("2881", "2882", "2891", "2886", "2884", "2892", "2885", "2880", "2883", "2890", "2887", "5880", "2801")
            Case "15. 金融 - 銀行與證券 (Bank/Securities)"
                arr = Array("5880", "2834", "2887", "6005", "2897", "2845", "2838", "2849", "5876", "2849", "6015", "6016")
            Case "16. 航運 (Container Shipping)"
                arr = Array("2603", "2609", "2615", "2617", "5608", "2208")
            Case "17. 傳產 - 散裝與航空 (Bulk & Aviation)"
                arr = Array("2606", "2637", "2618", "2610", "5608", "2634", "2605", "2601", "2608", "2614")
            Case "18. 鋼鐵 (Steel)"
                arr = Array("2002", "2014", "2006", "2027", "2031", "2023", "2059", "2062", "2069", "2038", "2034")
            Case "19. 塑膠 (Plastics)"
                arr = Array("1301", "1303", "1326", "1304", "1308", "1310", "1313", "2915", "1307", "1711")
            Case "20. 紡織 (Textiles)"
                arr = Array("1402", "1476", "1477", "1417", "1409", "1414", "1463", "1464", "1474", "1475")
            Case "21. 水泥 (Cement)"
                arr = Array("1101", "1102", "1103", "1104", "1108", "1109", "1110")
            Case "22. 電機機械"
                arr = Array("1503", "1504", "1513", "1519", "4566", "1514", "1516", "1528", "1560", "2308", "1536", "1590")
            Case "23. 電器電纜"
                arr = Array("1605", "1609", "1603", "1608", "1612", "1614", "1615", "1616", "1618", "1626")
            Case "24. 生技醫療 (Biotech)"
                arr = Array("1795", "4123", "4128", "4743", "6446", "4108", "4144", "4157", "4174", "4116", "1701", "4155", "4174", "6547", "4103")
            Case "25. 化學工業 (Chemicals)"
                arr = Array("1722", "1710", "4770", "1701", "1737", "1773", "1750", "1785", "1730", "1742")
            Case "26. 汽車工業 (Auto)"
                arr = Array("2201", "2204", "2207", "1319", "2206", "2209", "2212", "2213", "2219", "2230", "6191")
            Case "27. 建材營造 (Construction)"
                arr = Array("2501", "2542", "2548", "5522", "2538", "2543", "2547", "2560", "2514", "2524", "2530")
            Case "28. 食品工業 (Food)"
                arr = Array("1216", "1210", "1227", "1215", "1218", "1229", "1231", "1232", "1233", "1234", "1236", "1256")
            Case "29. 觀光餐旅 (Tourism)"
                arr = Array("2707", "2727", "2723", "5706", "2706", "2712", "2718", "2722", "2726", "2731")
            Case "30. 綠能環保 (Green Energy)"
                arr = Array("9933", "6869", "8341", "3708", "6414", "1560", "3576", "6690", "3536", "3622", "1563", "3577")
        End Select

    ElseIf market = "US" Then
        Select Case Sector
        Case "01. Tech - Mega Cap (七巨頭)"
                arr = Array("AAPL", "MSFT", "GOOGL", "GOOG", "AMZN", "NVDA", "META", "TSLA", "AVGO", "ORCL", "ADBE", "NFLX", "AMD", "INTC", "QCOM")
            Case "02. Tech - Software Infrastructure (軟體基建)"
                arr = Array("MSFT", "ORCL", "ADBE", "PLTR", "PANW", "CRWD", "SNOW", "FTNT", "NET", "ZS", "OKTA", "S", "DDOG", "GTLB", "ESTC", "MDB", "CFLT")
            Case "03. Tech - Software Application (應用軟體)"
                arr = Array("CRM", "INTU", "NOW", "UBER", "SAP", "ADP", "WDAY", "CDNS", "ADSK", "DDOG", "ROP", "FICO", "HUBS", "SHOP", "SQ", "BILL", "PCTY")
            Case "04. Tech - Semiconductors (半導體)"
                arr = Array("NVDA", "AVGO", "TSM", "AMD", "QCOM", "TXN", "MU", "ADI", "LRCX", "AMAT", "INTC", "NXPI", "MRVL", "MCHP", "ON", "KLAC", "ASML", "MPWR")
            Case "05. Tech - Consumer Electronics (消費電子)"
                arr = Array("AAPL", "SONY", "HPQ", "DELL", "APH", "TEL", "GLW", "STX", "WDC", "NTAP", "PSTG", "ROKU", "SMCI", "LOGI")
            Case "06. Tech - Info Services (資訊服務)"
                arr = Array("IBM", "ACN", "FI", "IT", "FIS", "CTSH", "BR", "EPAM", "GLOB", "INFY", "LDOS", "SAIC", "BAH", "CACI")
            Case "07. Financial - Banks Diversified (綜合銀行)"
                arr = Array("JPM", "BAC", "WFC", "C", "HSBC", "USB", "PNC", "TFC", "FITB", "RF", "HBAN", "KEY", "CFG", "MTB", "ZION", "CMA", "EWBC")
            Case "08. Financial - Credit Services (信貸服務)"
                arr = Array("V", "MA", "AXP", "PYPL", "COF", "SYF", "DFS", "SQ", "FIS", "FI", "GPN", "AFRM", "SOFI", "UPST")
            Case "09. Financial - Asset Management (資產管理)"
                arr = Array("BLK", "BX", "KKR", "GS", "MS", "APO", "BK", "STT", "AMP", "IVZ", "TROW", "BEN", "VIRT", "LPLA", "MORN")
            Case "10. Financial - Insurance (保險)"
                arr = Array("BRK-B", "PGR", "CB", "MMC", "AON", "TRV", "ALL", "AFL", "MET", "PRU", "AIG", "L", "UNM", "HIG", "RNR", "RE", "WRB", "CINF")
            Case "11. Healthcare - Drug Manufacturers (製藥)"
                arr = Array("LLY", "JNJ", "MRK", "ABBV", "PFE", "NVS", "AZN", "BMY", "RHHBY", "SNY", "GSK", "TEVA", "PRGO", "VTRS", "JAZZ")
            Case "12. Healthcare - Biotechnology (生技)"
                arr = Array("AMGN", "VRTX", "GILD", "REGN", "MRNA", "BIIB", "ALNY", "IONS", "EXEL", "SRPT", "NTLA", "BEAM", "CRSP", "EDIT")
            Case "13. Healthcare - Medical Devices (醫材)"
                arr = Array("ABT", "MDT", "SYK", "ISRG", "BSX", "EW", "BDX", "ZBH", "STE", "HOLX", "TFX", "NVCR", "ATRC", "SWAV", "INSP", "NARI")
            Case "14. Healthcare - Plans (健保)"
                arr = Array("UNH", "ELV", "CI", "CVS", "HUM", "CNC", "MOH", "OSCR", "CLOV")
            Case "15. Cyclical - Auto Manufacturers (汽車)"
                arr = Array("TSLA", "TM", "F", "GM", "HMC", "RACE", "RIVN", "LCID", "NIO", "LI", "XPEV", "STLA", "APTV", "MGA", "BWA", "LEA")
            Case "16. Cyclical - Internet Retail (電商)"
                arr = Array("AMZN", "BABA", "PDD", "JD", "EBAY", "DASH", "BKNG", "ABNB", "EXPE", "W", "CHWY")
            Case "17. Cyclical - Restaurants (餐飲)"
                arr = Array("MCD", "SBUX", "CMG", "YUM", "QSR", "DRI", "TXRH", "SHAK", "WING", "JACK", "WEN", "CAKE", "EAT", "BLMN")
            Case "18. Cyclical - Travel & Leisure (旅遊)"
                arr = Array("MAR", "HLT", "CCL", "RCL", "LVS", "WYNN", "MGM", "CZR", "NCLH", "UAL", "DAL", "LUV", "AAL", "ALK", "JBLU")
           Case "19. Defensive - Discount Stores (零售)"
                arr = Array("WMT", "COST", "TGT", "DG", "DLTR", "TJX", "ROST", "NKE", "LULU", "PVH", "VFC", "RL", "SKX", "CROX", "YETI")
            Case "20. Defensive - Beverages (飲料)"
                arr = Array("KO", "PEP", "MNST", "KDP", "STZ", "BUD", "SAM", "TAP", "ABEV", "CELH", "FIZZ")
            Case "21. Defensive - Household (家用品)"
                arr = Array("PG", "CL", "KMB", "EL", "K", "GIS", "CPB", "SJM", "MKC", "HRL", "CAG", "BG", "INGR", "POST", "LANC")
            Case "22. Comm - Internet Content (網路內容)"
                arr = Array("GOOGL", "META", "NFLX", "DASH", "PINS", "SNAP", "RDDT", "MTCH", "IAC", "ZG", "ANGI", "CDW")
            Case "23. Comm - Telecom (電信)"
                arr = Array("TMUS", "VZ", "T", "CMCSA", "CHTR", "LUMN", "IDT", "OOMA", "BAND")
            Case "24. Comm - Entertainment (娛樂)"
                arr = Array("DIS", "WBD", "LYV", "EA", "TTWO", "NFLX", "PARA", "FOX", "FOXA", "OMC", "IPG", "DKNG", "PENN")
            Case "25. Industrials - Aerospace (航太軍工)"
                arr = Array("GE", "RTX", "BA", "LMT", "GD", "NOC", "LHX", "TDG", "HEI", "KTOS", "RKLB", "IRDM", "BWXT", "CW", "HII", "MOOG")
            Case "26. Industrials - Logistics (物流)"
                arr = Array("UPS", "FDX", "UNP", "CSX", "NSC", "XPO", "SAIA", "JBHT", "ODFL", "CHRW", "LSTR", "GXO", "RXO", "MRTN", "KNX")
            Case "27. Industrials - Farm & Heavy (重工)"
                arr = Array("CAT", "DE", "ETN", "ITW", "PCAR", "EMR", "PH", "CMI", "IR", "ROK", "AME", "FTV", "GNRC", "GGG", "GTLS", "TTC")
            Case "28. Energy - Oil & Gas (石油天然氣)"
                arr = Array("XOM", "CVX", "SHEL", "TTE", "COP", "SLB", "EOG", "MPC", "PSX", "VLO", "OXY", "DVN", "MRO", "APA", "HAL", "BKR", "CTRA", "EQT")
            Case "29. Basic Materials (原物料)"
                arr = Array("LIN", "SHW", "FCX", "SCCO", "NEM", "APD", "ECL", "DD", "PPG", "RPM", "FMC", "MOS", "CF", "HUN", "CTVA", "AXTA")
            Case "30. Real Estate (REITs)"
                arr = Array("PLD", "AMT", "EQIX", "PSA", "O", "SPG", "CCI", "DLR", "WELL", "AVB", "EQR", "ESS", "UDR", "MAA", "NNN", "WPC", "VICI", "GLPI", "IIPR")
            Case "31. ETFs - Major Indices (指數與區域)"
                arr = Array("SPY", "VOO", "QQQ", "VTI", "VT", "IWM", "VUG", "VTV", "SCHG", "VGT", "XLK", "XLF", "XLV", "XLE", "XLY", "XLP", "XLI", "XLU", "XLRE", "XLB", "XLC", "SMH", "SOXX", "ARKK", "GDX", "ICLN", "TAN", "JETS", "XBI")
        End Select
    End If

    GetSectorTickers = arr
End Function


