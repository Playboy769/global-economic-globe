Attribute VB_Name = "RR5_Core"
    Option Explicit

    ' --- Module-level variables (must be at top of module) ---
    Private NextRefreshTime As Date

    ' =====================================================================
    ' BLACK-SCHOLES FUNCTIONS
    ' Can be called from worksheet cells: =BS_Delta(S,K,T,r,sigma,"Call")
    ' =====================================================================

    Public Function NormCDF(x As Double) As Double
        ' Standard normal CDF. Previously a broken hybrid of Hart approximation
        ' and Excel's Norm_S_Dist that gave wrong results for x >= 0 (mixed PDF
        ' into a CDF formula). Now uses Excel's built-in CDF directly.
        NormCDF = Application.WorksheetFunction.Norm_S_Dist(x, True)
    End Function

    Public Function BS_Price(S As Double, k As Double, T As Double, r As Double, sigma As Double, optType As String) As Double
        If T <= 0 Then
            If UCase(optType) = "CALL" Then
                BS_Price = Application.WorksheetFunction.Max(0, S - k)
            Else
                BS_Price = Application.WorksheetFunction.Max(0, k - S)
            End If
            Exit Function
        End If
        If sigma <= 0 Or S <= 0 Or k <= 0 Then BS_Price = 0: Exit Function
        Dim d1 As Double, d2 As Double
        d1 = (Log(S / k) + (r + 0.5 * sigma ^ 2) * T) / (sigma * Sqr(T))
        d2 = d1 - sigma * Sqr(T)
        If UCase(optType) = "CALL" Then
            BS_Price = S * NormCDF(d1) - k * Exp(-r * T) * NormCDF(d2)
        Else
            BS_Price = k * Exp(-r * T) * NormCDF(-d2) - S * NormCDF(-d1)
        End If
    End Function

    Public Function BS_Delta(S As Double, k As Double, T As Double, r As Double, sigma As Double, optType As String) As Double
        If T <= 0 Or sigma <= 0 Or S <= 0 Or k <= 0 Then
            If UCase(optType) = "CALL" Then BS_Delta = IIf(S > k, 1, 0) Else BS_Delta = IIf(S < k, -1, 0)
            Exit Function
        End If
        Dim d1 As Double
        d1 = (Log(S / k) + (r + 0.5 * sigma ^ 2) * T) / (sigma * Sqr(T))
        If UCase(optType) = "CALL" Then
            BS_Delta = NormCDF(d1)
        Else
            BS_Delta = NormCDF(d1) - 1
        End If
    End Function

    Public Function BS_Gamma(S As Double, k As Double, T As Double, r As Double, sigma As Double) As Double
        If T <= 0 Or sigma <= 0 Or S <= 0 Or k <= 0 Then BS_Gamma = 0: Exit Function
        Dim d1 As Double
        d1 = (Log(S / k) + (r + 0.5 * sigma ^ 2) * T) / (sigma * Sqr(T))
        BS_Gamma = Exp(-0.5 * d1 ^ 2) / (Sqr(2 * 3.14159265358979) * S * sigma * Sqr(T))
    End Function

    Public Function BS_Theta(S As Double, k As Double, T As Double, r As Double, sigma As Double, optType As String) As Double
        If T <= 0 Or sigma <= 0 Or S <= 0 Or k <= 0 Then BS_Theta = 0: Exit Function
        Dim d1 As Double, d2 As Double
        d1 = (Log(S / k) + (r + 0.5 * sigma ^ 2) * T) / (sigma * Sqr(T))
        d2 = d1 - sigma * Sqr(T)
        Dim term1 As Double
        term1 = -S * Exp(-0.5 * d1 ^ 2) / (Sqr(2 * 3.14159265358979) * 2 * Sqr(T)) * sigma
        If UCase(optType) = "CALL" Then
            BS_Theta = (term1 - r * k * Exp(-r * T) * NormCDF(d2)) / 365
        Else
            BS_Theta = (term1 + r * k * Exp(-r * T) * NormCDF(-d2)) / 365
        End If
    End Function

    Public Function BS_Vega(S As Double, k As Double, T As Double, r As Double, sigma As Double) As Double
        If T <= 0 Or sigma <= 0 Or S <= 0 Or k <= 0 Then BS_Vega = 0: Exit Function
        Dim d1 As Double
        d1 = (Log(S / k) + (r + 0.5 * sigma ^ 2) * T) / (sigma * Sqr(T))
        BS_Vega = S * Sqr(T) * Exp(-0.5 * d1 ^ 2) / (Sqr(2 * 3.14159265358979) * 100)
    End Function

    Public Function BS_Rho(S As Double, k As Double, T As Double, r As Double, sigma As Double, optType As String) As Double
        If T <= 0 Or sigma <= 0 Or S <= 0 Or k <= 0 Then BS_Rho = 0: Exit Function
        Dim d1 As Double, d2 As Double
        d1 = (Log(S / k) + (r + 0.5 * sigma ^ 2) * T) / (sigma * Sqr(T))
        d2 = d1 - sigma * Sqr(T)
        If UCase(optType) = "CALL" Then
            BS_Rho = k * T * Exp(-r * T) * NormCDF(d2) / 100
        Else
            BS_Rho = -k * T * Exp(-r * T) * NormCDF(-d2) / 100
        End If
    End Function

    Public Function BS_IV(mktPx As Double, S As Double, k As Double, T As Double, r As Double, optType As String) As Double
        ' Newton-Raphson implied volatility solver (50 iterations).
        ' Bug B5 fix: returns 0 on non-convergence so caller can detect failure
        ' instead of using a noisy final sigma value.
        If T <= 0 Or S <= 0 Or k <= 0 Or mktPx <= 0 Then BS_IV = 0: Exit Function
        Dim sigma As Double, price As Double, vega As Double
        Dim converged As Boolean: converged = False
        sigma = 0.3  ' initial guess
        Dim i As Integer
        For i = 1 To 50
            price = BS_Price(S, k, T, r, sigma, optType)
            vega = BS_Vega(S, k, T, r, sigma) * 100   ' un-scale from /100
            If Abs(vega) < 0.000001 Then Exit For
            sigma = sigma - (price - mktPx) / vega
            If sigma < 0.001 Then sigma = 0.001
            If sigma > 10 Then sigma = 10
            If Abs(price - mktPx) < 0.0001 Then
                converged = True
                Exit For
            End If
        Next i
        ' Final convergence check: tolerate 1c absolute or 0.1% relative error
        If Not converged Then
            Dim finalPx As Double
            finalPx = BS_Price(S, k, T, r, sigma, optType)
            If Abs(finalPx - mktPx) > 0.01 And Abs(finalPx - mktPx) / mktPx > 0.001 Then
                BS_IV = 0
                Exit Function
            End If
        End If
        BS_IV = sigma
    End Function

    Public Function Warrant_Delta(S As Double, k As Double, T As Double, r As Double, sigma As Double, convRatio As Double, wType As String) As Double
        ' wType: "Call Warrant" or "Put Warrant" or "Bull Cert" or "Bear Cert"
        Dim optT As String
        If InStr(UCase(wType), "CALL") > 0 Or InStr(UCase(wType), "BULL") > 0 Then
            optT = "Call"
        Else
            optT = "Put"
        End If
        If convRatio = 0 Then Warrant_Delta = 0: Exit Function
        Warrant_Delta = BS_Delta(S, k, T, r, sigma, optT) / convRatio
    End Function

    ' =====================================================================
    ' OPERATIONAL SUBROUTINES
    ' =====================================================================

    Private Function GetConfig(wsC As Worksheet, key As String) As Variant
        Dim i As Long
        For i = 1 To 20
            If wsC.Cells(i, 1).Value = key Then
                GetConfig = wsC.Cells(i, 2).Value
                Exit Function
            End If
        Next i
        GetConfig = ""
    End Function

    ' --- Public wrappers so UserForms / other modules can read & update _Config ---
    Public Function GetConfigValue(key As String) As Variant
        ' Returns "" if key not found.
        Dim wsC As Worksheet
        Set wsC = ThisWorkbook.Sheets("_Config")
        GetConfigValue = GetConfig(wsC, key)
    End Function

    Public Sub SetConfigValue(key As String, val As Variant)
        ' Updates a row in _Config (col B) where col A matches key.
        ' Does nothing if key not found.
        Dim wsC As Worksheet
        Set wsC = ThisWorkbook.Sheets("_Config")
        Dim i As Long
        For i = 1 To 20
            If wsC.Cells(i, 1).Value = key Then
                wsC.Cells(i, 2).Value = val
                Exit Sub
            End If
        Next i
    End Sub

    Private Sub LogDebug(wsD As Worksheet, macroName As String, status As String, msg As String)
        If wsD Is Nothing Then Exit Sub   ' guard: DebugLog sheet may not exist
        On Error Resume Next
        Dim nextRow As Long
        nextRow = wsD.Cells(wsD.Rows.Count, 1).End(xlUp).row + 1
        wsD.Cells(nextRow, 1).Value = Now()
        wsD.Cells(nextRow, 2).Value = macroName
        wsD.Cells(nextRow, 3).Value = status
        wsD.Cells(nextRow, 4).Value = msg
        On Error GoTo 0
    End Sub

    Public Sub RefreshGreeks()
        ' Updates Delta/Gamma/Theta/Vega/Rho for all options and warrants.
        ' Called by timer or manually. Reads IV and spot price from input cells.
        '
        ' Fix: _Config OPT_SEC_START=30 points to the header row (row 30).
        ' Auto-advance past the header if the STRIKE cell is non-numeric.
        On Error GoTo ErrHandler
        Dim wsR As Worksheet, wsC As Worksheet, wsD As Worksheet
        Set wsR = ThisWorkbook.Sheets("RR5")
        Set wsC = ThisWorkbook.Sheets("_Config")
        Set wsD = ThisWorkbook.Sheets("DebugLog")

        Dim r As Double
        r = CDbl(GetConfig(wsC, "BS_RISK_FREE_RATE"))
        If r = 0 Then r = 0.02

        Dim optStart As Long, optCount As Long
        optStart = CLng(GetConfig(wsC, "OPT_SEC_START"))
        optCount = CLng(GetConfig(wsC, "OPT_INST_COUNT"))

        ' Auto-advance past header row(s). Bug B6 fix: was single-step (only
        ' advanced once), so misconfigured _Config that pointed > 1 row above
        ' data would silently write Greeks to header cells. Loop until we land
        ' on numeric STRIKE or hit a safety cap.
        Dim hdrGuard As Long: hdrGuard = 0
        Do While Not IsNumeric(wsR.Cells(optStart, 9).Value) And hdrGuard < 5
            optStart = optStart + 1
            hdrGuard = hdrGuard + 1
        Loop

        ' Options section: rows optStart to optStart+optCount-1  (rows 31-70)
        Dim i As Long, row As Long
        For i = 0 To optCount - 1
            row = optStart + i
            Dim contract As String
            contract = wsR.Cells(row, 1).Value
            If contract = "" Then GoTo NextOpt
            ' Extra guard: skip any non-data row (e.g. if config is off by more)
            If Not IsNumeric(wsR.Cells(row, 9).Value) Then GoTo NextOpt

            Dim S As Double, k As Double, dte As Double, iv As Double
            Dim oType As String
            S = wsR.Cells(row, 11).Value      ' K  UNDL LAST
            k = wsR.Cells(row, 9).Value       ' I  STRIKE
            dte = wsR.Cells(row, 5).Value     ' E  DTE (days)
            iv = wsR.Cells(row, 19).Value     ' S  IV (fraction 0-1)
            oType = wsR.Cells(row, 8).Value   ' H  TYPE (Call/Put)

            If S <= 0 Or k <= 0 Or dte < 0 Then GoTo NextOpt

            Dim T As Double
            T = dte / 365

            ' Reverse-solve IV from LAST PREM if IV is missing/zero.
            ' Stores the implied vol back into col S so subsequent calls reuse it.
            If iv <= 0 Then
                Dim lastPrem As Double
                lastPrem = wsR.Cells(row, 15).Value    ' O LAST PREM
                If lastPrem > 0 Then
                    Dim solvedIV As Double
                    solvedIV = BS_IV(lastPrem, S, k, T, r, oType)
                    If solvedIV > 0.001 And solvedIV < 5 Then
                        iv = solvedIV
                        wsR.Cells(row, 19).Value = iv     ' write back to S IV
                    End If
                End If
            End If
            If iv <= 0 Then GoTo NextOpt

            wsR.Cells(row, 20).Value = BS_Delta(S, k, T, r, iv, oType)  ' T DELTA
            wsR.Cells(row, 21).Value = BS_Gamma(S, k, T, r, iv)         ' U GAMMA
            wsR.Cells(row, 22).Value = BS_Theta(S, k, T, r, iv, oType)  ' V THETA/DAY
            wsR.Cells(row, 23).Value = BS_Vega(S, k, T, r, iv)          ' W VEGA
            wsR.Cells(row, 24).Value = BS_Rho(S, k, T, r, iv, oType)    ' X RHO
NextOpt:
        Next i

        ' Warrants section
        Dim wrtStart As Long, wrtCount As Long
        wrtStart = CLng(GetConfig(wsC, "WRT_SEC_START"))
        wrtCount = CLng(GetConfig(wsC, "WRT_INST_COUNT"))

        ' Auto-advance past warrant header row(s). Same trick as options.
        ' Warrant STRIKE column is K (col 11). If header text, advance up to 3.
        Dim wHg As Long: wHg = 0
        Do While Not IsNumeric(wsR.Cells(wrtStart, 11).Value) And wHg < 3
            wrtStart = wrtStart + 1: wHg = wHg + 1
        Loop

        For i = 0 To wrtCount - 1
            row = wrtStart + i
            Dim wCode As String
            wCode = wsR.Cells(row, 1).Value
            If wCode = "" Then GoTo NextWrt

            Dim wS As Double, wK As Double, wDte As Double, wIV As Double
            Dim convRatio As Double, wType As String
            wS = wsR.Cells(row, 10).Value         ' UNDL LAST
            wK = wsR.Cells(row, 11).Value         ' STRIKE
            wDte = wsR.Cells(row, 5).Value        ' DTE
            wIV = wsR.Cells(row, 19).Value        ' IV
            convRatio = wsR.Cells(row, 12).Value  ' CONV RATIO
            wType = wsR.Cells(row, 7).Value       ' TYPE

            If wS <= 0 Or wK <= 0 Or wIV <= 0 Or convRatio <= 0 Then GoTo NextWrt

            Dim wT As Double
            wT = wDte / 365

            ' Normalise warrant type to "Call"/"Put" for BS_Theta sign handling
            Dim wOptT As String
            If InStr(UCase(wType), "CALL") > 0 Or InStr(UCase(wType), "BULL") > 0 Then
                wOptT = "Call"
            Else
                wOptT = "Put"
            End If

            ' All Greeks are scaled by /convRatio to express "per warrant unit"
            wsR.Cells(row, 20).Value = Warrant_Delta(wS, wK, wT, r, wIV, convRatio, wType)
            wsR.Cells(row, 24).Value = BS_Gamma(wS, wK, wT, r, wIV) / convRatio         ' X GAMMA
            wsR.Cells(row, 25).Value = BS_Theta(wS, wK, wT, r, wIV, wOptT) / convRatio  ' Y THETA/DAY
            wsR.Cells(row, 26).Value = BS_Vega(wS, wK, wT, r, wIV) / convRatio          ' Z VEGA

            ' Gearing needs LAST PX (col O) -- guarded separately since a fresh
            ' position may have S/K/IV/convRatio filled before FetchAllPrices runs.
            Dim wLastPx As Double: wLastPx = wsR.Cells(row, 15).Value   ' O LAST PX
            If wLastPx > 0 Then
                Dim wDelta As Double: wDelta = wsR.Cells(row, 20).Value      ' T DELTA (just written above)
                wsR.Cells(row, 27).Value = wS / (wLastPx * convRatio)  ' AA NOM GEARING  = UNDL / (LastPx*ConvRatio)
                wsR.Cells(row, 28).Value = wS * wDelta / wLastPx       ' AB EFF GEARING  = UNDL*Delta / LastPx
            Else
                wsR.Cells(row, 27).Value = ""
                wsR.Cells(row, 28).Value = ""
            End If
NextWrt:
        Next i

        LogDebug wsD, "RefreshGreeks", "OK", "Greeks updated at " & Now()
        Exit Sub
ErrHandler:
        LogDebug wsD, "RefreshGreeks", "ERROR", Err.Description
    End Sub

    Public Sub DailySnapshot()
        ' Appends one row to HistoryLog and Greeks_History with today's portfolio state.
        '
        ' HistoryLog columns:
        '   A=Date  B=TotalMarketValue  C=TotalCumulativePnL
        '   D=NetDelta  E=NetGamma  F=NetTheta  G=NetVega
        '   H=Price(TX)  I=Ret%(TX)  J=Price(SPY)  K=Ret%(SPY)
        '   L=RealizedPnL  M=PortRet%  N=PortRetChg%  O=IV_Avg  P=Summary
        Dim wsR As Worksheet, wsH As Worksheet, wsG As Worksheet, wsD As Worksheet
        On Error Resume Next
        Set wsD = ThisWorkbook.Sheets("DebugLog")
        Set wsG = ThisWorkbook.Sheets("Greeks_History")   ' optional
        On Error GoTo 0
        On Error GoTo ErrHandler
        Set wsR = ThisWorkbook.Sheets("RR5")
        Set wsH = ThisWorkbook.Sheets("HistoryLog")

        ' Target row: if today's row already exists, OVERWRITE it (refresh)
        ' instead of skipping. This keeps one row per day but lets repeated
        ' saves during the same day update MV / Greeks / PnL in place.
        Dim lastRowH As Long
        lastRowH = wsH.Cells(wsH.Rows.Count, 1).End(xlUp).row

        Dim isUpdate As Boolean: isUpdate = False
        If lastRowH >= 2 Then
            If IsDate(wsH.Cells(lastRowH, 1).Value) Then
                If DateValue(wsH.Cells(lastRowH, 1).Value) = Date Then isUpdate = True
            End If
        End If

        Dim nextH As Long, nextG As Long, prevRow As Long
        If isUpdate Then
            nextH = lastRowH        ' overwrite today's row
            prevRow = lastRowH - 1  ' compare against the prior day
        Else
            nextH = lastRowH + 1    ' append a new row
            prevRow = lastRowH      ' compare against the last (older) row
        End If

        If Not wsG Is Nothing Then
            Dim lastG As Long: lastG = wsG.Cells(wsG.Rows.Count, 1).End(xlUp).row
            If isUpdate And lastG >= 2 Then nextG = lastG Else nextG = lastG + 1
        End If

        ' --- Read current RR5 summary values ---
        ' Force a recalc first: ShowTransactionForm runs with Calculation=Manual,
        ' so the summary formulas (B1 MV, L1 Delta, ...) would otherwise still
        ' hold pre-FetchAllPrices values and we'd log a stale / zero snapshot.
        On Error Resume Next
        wsR.Calculate
        On Error GoTo ErrHandler
        Dim curMV    As Double: curMV    = ToDbl(wsR.Range("B1").Value)  ' Total Market Value
        Dim netDelta As Double: netDelta = ToDbl(wsR.Range("L1").Value)  ' Net Delta
        Dim netGamma As Double: netGamma = ToDbl(wsR.Range("N1").Value)  ' Net Gamma
        Dim netTheta As Double: netTheta = ToDbl(wsR.Range("J2").Value)  ' Net Theta/Day
        Dim netVega  As Double: netVega  = ToDbl(wsR.Range("L2").Value)  ' Net Vega
        Dim rlPnL    As Double: rlPnL    = ToDbl(wsR.Range("H2").Value)  ' Realized PnL
        Dim unrlPnL  As Double: unrlPnL  = ToDbl(wsR.Range("F2").Value)  ' Unrealized PnL

        ' --- Previous row values for % change ---
        Dim prevMV  As Double: prevMV  = 0
        Dim prevTX  As Double: prevTX  = 0
        Dim prevSPY As Double: prevSPY = 0
        If prevRow >= 2 Then
            If IsNumeric(wsH.Cells(prevRow, 2).Value)  Then prevMV  = CDbl(wsH.Cells(prevRow, 2).Value)
            If IsNumeric(wsH.Cells(prevRow, 8).Value)  Then prevTX  = CDbl(wsH.Cells(prevRow, 8).Value)
            If IsNumeric(wsH.Cells(prevRow, 10).Value) Then prevSPY = CDbl(wsH.Cells(prevRow, 10).Value)
        End If

        ' --- Fetch benchmark prices (silently skip on network failure) ---
        Dim txPx As Double, spyPx As Double
        Application.StatusBar = "DailySnapshot: fetching ^TWII..."
        On Error Resume Next
        txPx  = FetchPrice("^TWII")
        spyPx = FetchPrice("SPY")
        On Error GoTo ErrHandler
        Application.StatusBar = False

        ' --- IV average across active options + warrants (col S=19) ---
        Dim ivSum As Double, ivCnt As Long, i As Long, ivV As Variant
        For i = 19 To 26   ' Options
            If Trim(CStr(wsR.Cells(i, 1).Value)) <> "" Then
                ivV = wsR.Cells(i, 19).Value
                If IsNumeric(ivV) And CDbl(ivV) > 0 Then ivSum = ivSum + CDbl(ivV): ivCnt = ivCnt + 1
            End If
        Next i
        For i = 31 To 38   ' Warrants
            If Trim(CStr(wsR.Cells(i, 1).Value)) <> "" Then
                ivV = wsR.Cells(i, 19).Value
                If IsNumeric(ivV) And CDbl(ivV) > 0 Then ivSum = ivSum + CDbl(ivV): ivCnt = ivCnt + 1
            End If
        Next i
        Dim ivAvg As Double
        If ivCnt > 0 Then ivAvg = ivSum / ivCnt

        ' --- PortRetChg%: day-over-day MV change ---
        Dim portRetChg As Double
        If prevMV > 0 And curMV > 0 Then portRetChg = (curMV - prevMV) / prevMV

        ' --- PortRet%: cumulative since first snapshot (row 2) ---
        Dim portRet As Double
        If lastRowH >= 2 And IsNumeric(wsH.Cells(2, 2).Value) Then
            Dim firstMV As Double: firstMV = CDbl(wsH.Cells(2, 2).Value)
            If firstMV > 0 And curMV > 0 Then portRet = (curMV - firstMV) / firstMV
        End If

        ' --- Summary text ---
        Dim posCount As Variant: posCount = wsR.Range("P1").Value
        Dim summaryTxt As String
        summaryTxt = "Pos=" & IIf(IsNumeric(posCount), CStr(CLng(posCount)), "?") & _
                     " MV=" & Format(curMV, "#,##0") & _
                     " UnrlPnL=" & IIf(unrlPnL >= 0, "+", "") & Format(unrlPnL, "#,##0") & _
                     " RlPnL=" & IIf(rlPnL >= 0, "+", "") & Format(rlPnL, "#,##0") & _
                     IIf(ivAvg > 0, " IV=" & Format(ivAvg, "0.0%"), "")

        ' HistoryLog row -- all 16 columns
        With wsH
            .Cells(nextH, 1).Value  = Now()                    ' A Date
            .Cells(nextH, 2).Value  = curMV                    ' B TotalMarketValue
            .Cells(nextH, 3).Value  = unrlPnL + rlPnL          ' C TotalCumulativePnL
            .Cells(nextH, 4).Value  = netDelta                  ' D NetDelta
            .Cells(nextH, 5).Value  = netGamma                  ' E NetGamma
            .Cells(nextH, 6).Value  = netTheta                  ' F NetTheta
            .Cells(nextH, 7).Value  = netVega                   ' G NetVega
            If txPx  > 0 Then .Cells(nextH, 8).Value  = txPx   ' H Price(TX)
            If txPx  > 0 And prevTX  > 0 Then _
                .Cells(nextH, 9).Value  = (txPx  - prevTX)  / prevTX   ' I Ret%(TX)
            If spyPx > 0 Then .Cells(nextH, 10).Value = spyPx  ' J Price(SPY)
            If spyPx > 0 And prevSPY > 0 Then _
                .Cells(nextH, 11).Value = (spyPx - prevSPY) / prevSPY  ' K Ret%(SPY)
            .Cells(nextH, 12).Value = rlPnL                     ' L RealizedPnL
            .Cells(nextH, 13).Value = portRet                   ' M PortRet%
            .Cells(nextH, 14).Value = portRetChg                ' N PortRetChg%
            If ivAvg > 0 Then .Cells(nextH, 15).Value = ivAvg  ' O IV_Avg
            .Cells(nextH, 16).Value = summaryTxt                ' P Summary

            ' Number formats
            .Cells(nextH, 1).NumberFormat  = "yyyy/m/d hh:mm"
            .Cells(nextH, 2).NumberFormat  = "#,##0"
            .Cells(nextH, 3).NumberFormat  = "#,##0;(#,##0)"
            .Cells(nextH, 4).NumberFormat  = "+0.0000;-0.0000;0"
            .Cells(nextH, 5).NumberFormat  = "0.000000"
            .Cells(nextH, 6).NumberFormat  = "0.0000"
            .Cells(nextH, 7).NumberFormat  = "0.0000"
            .Cells(nextH, 8).NumberFormat  = "#,##0.00"
            .Cells(nextH, 9).NumberFormat  = "0.00%"
            .Cells(nextH, 10).NumberFormat = "#,##0.00"
            .Cells(nextH, 11).NumberFormat = "0.00%"
            .Cells(nextH, 12).NumberFormat = "#,##0;(#,##0)"
            .Cells(nextH, 13).NumberFormat = "0.00%"
            .Cells(nextH, 14).NumberFormat = "+0.00%;-0.00%;0.00%"
            .Cells(nextH, 15).NumberFormat = "0.00%"
        End With

        ' Greeks_History row (optional sheet)
        If Not wsG Is Nothing Then
            wsG.Cells(nextG, 1).Value = Now()
            wsG.Cells(nextG, 3).Value = netDelta
            wsG.Cells(nextG, 4).Value = netGamma
            wsG.Cells(nextG, 5).Value = netTheta
            wsG.Cells(nextG, 6).Value = netVega
        End If

        LogDebug wsD, "DailySnapshot", "OK", _
            IIf(isUpdate, "UPDATED row " & nextH, "APPENDED row " & nextH) & _
            " | MV=" & curMV & " TX=" & txPx & " SPY=" & spyPx & _
            " IV_Avg=" & Format(ivAvg, "0.0%")
        Exit Sub
ErrHandler:
        Application.StatusBar = False
        LogDebug wsD, "DailySnapshot", "ERROR", Err.Description
    End Sub

    Public Sub ExpiryAlert()
        ' Scans DTE columns across all three sections and alerts if any
        ' <= DTE_WARN_DAYS. Bug B2 fix: row ranges now read from _Config
        ' (FUT/OPT/WRT_SEC_START + INST_COUNT) instead of hardcoded.
        On Error GoTo ErrHandler
        Dim wsR As Worksheet, wsC As Worksheet, wsD As Worksheet
        Set wsR = ThisWorkbook.Sheets("RR5")
        Set wsC = ThisWorkbook.Sheets("_Config")
        Set wsD = ThisWorkbook.Sheets("DebugLog")

        Dim warnDays As Long
        warnDays = CLng(GetConfig(wsC, "DTE_WARN_DAYS"))
        If warnDays = 0 Then warnDays = 5

        Dim alertMsg As String
        alertMsg = ""

        ' Load section ranges from _Config
        Dim futStart As Long, futCount As Long
        Dim optStart As Long, optCount As Long
        Dim wrtStart As Long, wrtCount As Long
        futStart = CLng(GetConfig(wsC, "FUT_SEC_START")): futCount = CLng(GetConfig(wsC, "FUT_INST_COUNT"))
        optStart = CLng(GetConfig(wsC, "OPT_SEC_START")): optCount = CLng(GetConfig(wsC, "OPT_INST_COUNT"))
        wrtStart = CLng(GetConfig(wsC, "WRT_SEC_START")): wrtCount = CLng(GetConfig(wsC, "WRT_INST_COUNT"))

        ' Advance past header row(s) if the start cell isn't actual data.
        ' STRIKE col differs by section: futures has no strike (use A=contract),
        ' options strike is col 9, warrants strike is col 11.
        Dim hg As Long
        hg = 0: Do While Len(CStr(wsR.Cells(futStart, 1).Value)) = 0 And IsEmpty(wsR.Cells(futStart, 4).Value) And hg < 3
            futStart = futStart + 1: hg = hg + 1
        Loop
        hg = 0: Do While Not IsNumeric(wsR.Cells(optStart, 9).Value) And hg < 5
            optStart = optStart + 1: hg = hg + 1
        Loop
        hg = 0: Do While Not IsNumeric(wsR.Cells(wrtStart, 11).Value) And hg < 5
            wrtStart = wrtStart + 1: hg = hg + 1
        Loop

        Call ScanExpirySection(wsR, futStart, futCount, "FUTURES", warnDays, alertMsg)
        Call ScanExpirySection(wsR, optStart, optCount, "OPTIONS", warnDays, alertMsg)
        Call ScanExpirySection(wsR, wrtStart, wrtCount, "WARRANTS", warnDays, alertMsg)

        If alertMsg <> "" Then
            MsgBox "EXPIRY ALERT - positions expiring within " & warnDays & " days:" & vbCrLf & vbCrLf & alertMsg, _
                   vbExclamation, "RR5 Expiry Alert"
        End If
        LogDebug wsD, "ExpiryAlert", "OK", IIf(alertMsg = "", "No alerts", "ALERTED: " & alertMsg)
        Exit Sub
ErrHandler:
        LogDebug wsD, "ExpiryAlert", "ERROR", Err.Description
    End Sub

    Private Sub ScanExpirySection(wsR As Worksheet, startRow As Long, count As Long, _
                                  label As String, warnDays As Long, ByRef alertMsg As String)
        Dim i As Long, dte As Variant, contract As String
        For i = startRow To startRow + count - 1
            dte = wsR.Cells(i, 5).Value
            contract = CStr(wsR.Cells(i, 1).Value)
            If contract = "" Then GoTo NextI
            If IsNumeric(dte) Then
                If CLng(dte) > 0 And CLng(dte) <= warnDays Then
                    alertMsg = alertMsg & label & ": " & contract & " DTE=" & dte & vbCrLf
                End If
            End If
NextI:
        Next i
    End Sub

    ' --- Timer-based auto-refresh (NextRefreshTime declared at top of module) ---

    ' =================================================================
    ' ONE-SHOT: Update _Config to compact 8-rows-per-section layout
    ' =================================================================
    '   Futures  data rows 7-14   (header 6,  subtotal 15)
    '   Options  data rows 19-26  (banner 17, header 18, subtotal 27)
    '   Warrants data rows 31-38  (banner 29, header 30, subtotal 39)
    '   GrandTotal row 41
    Public Sub Setup_CompactLayout_Config()
        On Error GoTo Fail
        Dim wsC As Worksheet
        Set wsC = ThisWorkbook.Sheets("_Config")

        Call SetConfigValue("FUT_SEC_START", 6)   ' header row; ExpiryAlert auto-advances
        Call SetConfigValue("FUT_INST_COUNT", 8)
        Call SetConfigValue("OPT_SEC_START", 18)
        Call SetConfigValue("OPT_INST_COUNT", 8)
        Call SetConfigValue("WRT_SEC_START", 30)
        Call SetConfigValue("WRT_INST_COUNT", 8)

        MsgBox "_Config updated to compact layout:" & vbCrLf & _
               "  FUT_SEC_START=6   FUT_INST_COUNT=8   (data 7-14)" & vbCrLf & _
               "  OPT_SEC_START=18  OPT_INST_COUNT=8   (data 19-26)" & vbCrLf & _
               "  WRT_SEC_START=30  WRT_INST_COUNT=8   (data 31-38)", _
               vbInformation, "Compact Layout Config"
        Exit Sub
Fail:
        MsgBox "Setup_CompactLayout_Config failed: " & Err.Description, vbCritical
    End Sub

    Public Sub AutoRefresh_Start()
        Dim wsC As Worksheet
        On Error Resume Next
        Set wsC = ThisWorkbook.Sheets("_Config")
        Dim mins As Long
        mins = CLng(wsC.Cells(7, 2).Value)  ' RefreshInterval_Min row
        If mins <= 0 Then mins = 30
        NextRefreshTime = Now() + TimeSerial(0, mins, 0)
        Application.OnTime NextRefreshTime, "RR5_Core.AutoRefresh_Tick"
    End Sub

    Public Sub AutoRefresh_Stop()
        On Error Resume Next
        Application.OnTime NextRefreshTime, "RR5_Core.AutoRefresh_Tick", , False
    End Sub

    Public Sub AutoRefresh_Tick()
        Call RefreshGreeks
        Call ExpiryAlert
        ' Pull aggregates from Realized / IVTracker / HistoryLog / Greeks_History
        ' into RR5 rows 3-4. Defensive: if RR5_Dashboard module not present
        ' yet, skip silently so timer still runs.
        On Error Resume Next
        Application.Run "RR5_Dashboard.DashboardRefresh"
        On Error GoTo 0
        Call AutoRefresh_Start   ' reschedule
    End Sub

    ' =====================================================================
    ' TRANSACTION ENTRY FORM HELPERS
    ' =====================================================================

    Public Function GenerateTxID() As String
        ' Format: T-YYYYMMDD-HHMMSS
        GenerateTxID = "T-" & Format(Now(), "yyyymmdd-hhnnss")
    End Function

    Public Sub AppendTransaction(rowData As Variant)
        ' Appends a variable-length 1-based array to the Transactions sheet,
        ' then auto-chains:  RebuildPositionsFromTransactions -> FetchAllPrices
        ' (the latter is invoked transitively by Rebuild's last step).
        '
        ' Transactions sheet has 24 columns:
        '  1=ID  2=Date  3=Ticker  4=Action  5=Type   6=Strike  7=Expiry
        '  8=Qty 9=Mult  10=Price  11=Fee    12=Tax   13=NetAmt
        ' 14=Sector 15=Target 16=Strategy 17=IV 18=Delta 19=Beta 20=Broker
        ' 21=Underlying 22=Gamma 23=Theta 24=Conv_Ratio
        Dim wsT As Worksheet, wsD As Worksheet
        On Error Resume Next
        Set wsD = ThisWorkbook.Sheets("DebugLog")   ' optional; LogDebug guards Nothing
        On Error GoTo 0
        On Error GoTo ErrHandler
        Set wsT = ThisWorkbook.Sheets("Transactions")

        Dim nextRow As Long
        nextRow = wsT.Cells(wsT.Rows.Count, 1).End(xlUp).row + 1
        If nextRow < 2 Then nextRow = 2

        Dim i As Long
        For i = 1 To UBound(rowData)   ' dynamic: handles 20-col or 23-col arrays
            ' Col 3 (Ticker) -- force text format so leading zeros survive
            ' (e.g. warrant code "066669" wouldn't drop to 66669)
            If i = 3 Then wsT.Cells(nextRow, i).NumberFormat = "@"
            wsT.Cells(nextRow, i).Value = rowData(i)
        Next i

        LogDebug wsD, "AppendTransaction", "OK", "Row " & nextRow & " | " & rowData(1)

        ' Chain: rebuild positions -> auto fetch prices -> refresh greeks.
        ' RebuildPositionsFromTransactions itself calls FetchAllPrices(silent:=True)
        ' which calls RefreshGreeks, so one call cascades the whole pipeline.
        Call RebuildPositionsFromTransactions
        Exit Sub
ErrHandler:
        LogDebug wsD, "AppendTransaction", "ERROR", Err.Description
        MsgBox "Error saving transaction: " & Err.Description, vbCritical
    End Sub

    Public Sub ShowTransactionForm()
        ' Entry point: opens the UserForm. Called by toolbar button or Ctrl+Shift+T.
        '
        ' We freeze calc + events while the form is open, otherwise volatile
        ' worksheet formulas (option pricing, third-party broker add-ins, etc.)
        ' can fire repeatedly while the user is still mid-entry and spam errors.
        On Error GoTo NoForm

        Dim oldCalc As XlCalculation, oldEvt As Boolean, oldScr As Boolean
        oldCalc = Application.Calculation
        oldEvt  = Application.EnableEvents
        oldScr  = Application.ScreenUpdating
        Application.Calculation   = xlCalculationManual
        Application.EnableEvents  = False
        Application.ScreenUpdating = True   ' keep UI responsive

        frmTransaction.Show vbModal          ' blocks Excel until form closes

        ' Restore state -- always run, even if user cancelled
        Application.Calculation   = oldCalc
        Application.EnableEvents  = oldEvt
        Application.ScreenUpdating = oldScr
        ' Force one clean recalc now that everything is back to normal
        Application.CalculateFull
        Exit Sub
NoForm:
        Application.Calculation   = xlCalculationAutomatic
        Application.EnableEvents  = True
        Application.ScreenUpdating = True
        MsgBox "frmTransaction UserForm not found." & vbCrLf & _
               "Please create it via Alt+F11 -> Insert -> UserForm." & vbCrLf & _
               "See RR5_UserForm_Guide.md for instructions.", vbExclamation
    End Sub

    ' =====================================================================
    ' PASTE THE FOLLOWING INTO ThisWorkbook MODULE (not here):
    '
    ' Private Sub Workbook_Open()
    '     Call RR5_Core.ExpiryAlert
    '     Call RR5_Core.RefreshGreeks
    '     Call RR5_Core.AutoRefresh_Start
    '     Application.OnKey "^+t", "RR5_Core.ShowTransactionForm"   ' Ctrl+Shift+T
    ' End Sub
    '
    ' Private Sub Workbook_BeforeClose(Cancel As Boolean)
    '     Call RR5_Core.AutoRefresh_Stop
    '     Application.OnKey "^+t"   ' release shortcut
    ' End Sub
    ' =====================================================================

    ' =====================================================================
    ' REBUILD POSITIONS FROM TRANSACTIONS
    ' =====================================================================
    '
    ' Reads ALL rows in the Transactions sheet, groups by position key
    ' (Ticker + Type + Strike + Expiry), nets Buy/Sell quantities, and
    ' writes active (non-zero) positions to the three RR5 sections:
    '
    '   Future                 => Futures  rows  7-26  (cols A/C-D/F/H-K/R)
    '   Call / Put             => Options  rows 31-70  (cols A/C-D/F/H-N/S-T/AB)
    '   Call Warrant / Put Warrant
    '   Bull Cert / Bear Cert  => Warrants rows 75-104 (cols A/C-D/F-G/K-N/S-T)
    '   Equity / ETF           => skipped (no section in RR5)
    '
    ' Action handling:
    '   Buy        => +qty to net; accumulates weighted-average cost
    '   Sell       => -qty from net
    '   AdjustCost => overrides entry price to transaction price
    '   Expire /
    '   Exercise   => zeros out position (closed)
    '
    ' NOTE: After running, manually update in RR5:
    '   LAST / UNDL LAST  (current market prices)
    '   CONV RATIO        (warrants: conversion ratio)
    '   MARGIN REQ        (futures: per-contract margin)
    '   Then re-run RefreshGreeks to update all Greeks.
    '
    Public Sub RebuildPositionsFromTransactions()
        Dim wsT As Worksheet, wsR As Worksheet, wsD As Worksheet
        On Error Resume Next
        Set wsD = ThisWorkbook.Sheets("DebugLog")   ' optional; LogDebug guards Nothing
        On Error GoTo 0
        On Error GoTo ErrHandler
        Set wsT = ThisWorkbook.Sheets("Transactions")
        Set wsR = ThisWorkbook.Sheets("RR5")

        ' -- Step 1: Read all Transactions ----------------------------------
        ' Transactions sheet -- 24 columns:
        '  1=Transaction_ID  2=Date       3=Ticker     4=Action   5=Type
        '  6=Strike          7=Expiry     8=Shares_Contracts      9=Multiplier
        ' 10=Price          11=Fee       12=Tax       13=Net_Amount
        ' 14=Sector         15=Target    16=Strategy_Tag
        ' 17=IV_At_Entry    18=Delta_At_Entry 19=Beta  20=Broker
        ' 21=Underlying     22=Gamma     23=Theta      24=Conv_Ratio

        Dim lastRow As Long
        lastRow = wsT.Cells(wsT.Rows.Count, 1).End(xlUp).row
        If lastRow < 2 Then
            MsgBox "Transactions sheet has no data.", vbInformation
            Exit Sub
        End If

        ' Dictionary: posKey => 18-element Variant array (indices 0-17)
        '  0=posKey   1=ticker     2=txType  3=strike    4=expiry
        '  5=netQty   6=buyCostSum 7=buyQtySum
        '  8=multiplier 9=broker  10=entryDate
        ' 11=strategy 12=iv       13=delta
        ' 14=underlying 15=gamma  16=theta  17=convRatio  18=issuer
        Dim dict As Object
        Set dict = CreateObject("Scripting.Dictionary")

        Dim e(18) As Variant  ' reused for each new entry

        Dim i As Long
        For i = 2 To lastRow
            Dim txDate As Variant
            Dim ticker As String, action As String, txType As String
            Dim strike As Variant, expiry As Variant
            Dim qty As Double, txMult As Double, price As Double
            Dim strategy As String, iv As Variant, delta As Variant
            Dim broker As String, UnderLying As String
            Dim gamma As Variant, theta As Variant, convRatio As Variant
            Dim issuer As String

            txDate     = wsT.Cells(i, 2).Value
            ticker     = Trim(CStr(wsT.Cells(i, 3).Value))
            action     = UCase(Trim(CStr(wsT.Cells(i, 4).Value)))
            txType     = Trim(CStr(wsT.Cells(i, 5).Value))
            strike     = wsT.Cells(i, 6).Value
            expiry     = wsT.Cells(i, 7).Value
            qty        = CDbl(wsT.Cells(i, 8).Value)
            txMult     = CDbl(wsT.Cells(i, 9).Value)
            If txMult = 0 Then txMult = 1
            price      = CDbl(wsT.Cells(i, 10).Value)
            strategy   = Trim(CStr(wsT.Cells(i, 16).Value))
            iv         = wsT.Cells(i, 17).Value
            delta      = wsT.Cells(i, 18).Value
            broker     = Trim(CStr(wsT.Cells(i, 20).Value))
            UnderLying = Trim(CStr(wsT.Cells(i, 21).Value))  ' col U
            gamma      = wsT.Cells(i, 22).Value              ' col V
            theta      = wsT.Cells(i, 23).Value              ' col W
            convRatio  = wsT.Cells(i, 24).Value              ' col X Conv_Ratio (warrants)
            issuer     = Trim(CStr(wsT.Cells(i, 25).Value))  ' col Y Issuer (warrants)

            If ticker = "" Then GoTo SkipRow

            ' Build position key by instrument type
            Dim posKey As String
            Select Case txType
                Case "Future"
                    posKey = "FUT|" & ticker & "|" & CStr(expiry)
                Case "Call", "Put"
                    posKey = "OPT|" & ticker & "|" & txType & "|" & CStr(strike) & "|" & CStr(expiry)
                Case "Call Warrant", "Put Warrant", "Bull Cert", "Bear Cert"
                    posKey = "WRT|" & ticker
                Case Else
                    GoTo SkipRow   ' Equity / ETF / unknown -- no section in RR5
            End Select

            ' Initialise dict entry on first encounter
            If Not dict.Exists(posKey) Then
                Erase e
                e(0) = posKey:  e(1) = ticker:       e(2) = txType
                e(3) = strike:  e(4) = expiry
                e(5) = 0:       e(6) = 0:             e(7) = 0
                e(8) = txMult:  e(9) = broker:        e(10) = txDate
                e(11) = strategy: e(12) = iv:         e(13) = delta
                e(14) = UnderLying: e(15) = gamma:    e(16) = theta
                e(17) = convRatio
                e(18) = issuer
                dict(posKey) = e
            End If

            Dim p() As Variant
            p = dict(posKey)

            ' --- Position state update --------------------------------------
            ' p(5) = signed netQty (+ long, - short)
            ' p(6) = avgCost of currently OPEN side (positive, 0 if flat)
            ' p(7) = 0 (unused after the round-trip bug fix; previously was
            '         cumulative buy qty which incorrectly persisted across
            '         full close/reopen cycles)
            Dim curNet As Double:  curNet = CDbl(p(5))
            Dim curAvg As Double:  curAvg = CDbl(p(6))
            Dim resQty As Double

            Select Case action
                Case "BUY"
                    If curNet >= 0 Then
                        ' Add to (or open) long
                        If curNet = 0 Then
                            curAvg = price
                            p(10) = txDate
                        Else
                            curAvg = (curNet * curAvg + qty * price) / (curNet + qty)
                        End If
                        curNet = curNet + qty
                    Else
                        ' Cover short (curNet < 0)
                        If qty >= -curNet Then
                            ' Full cover; residual reopens long
                            resQty = qty + curNet     ' qty - |short|
                            curNet = 0: curAvg = 0
                            p(10) = Empty
                            If resQty > 0 Then
                                curNet = resQty
                                curAvg = price
                                p(10) = txDate
                            End If
                        Else
                            ' Partial cover; avg short cost unchanged
                            curNet = curNet + qty
                        End If
                    End If

                Case "SELL"
                    If curNet > 0 Then
                        ' Close long
                        If qty >= curNet Then
                            ' Full close; residual opens short
                            resQty = qty - curNet
                            curNet = 0: curAvg = 0
                            p(10) = Empty
                            If resQty > 0 Then
                                curNet = -resQty
                                curAvg = price
                                p(10) = txDate
                            End If
                        Else
                            ' Partial close; avg long cost unchanged
                            curNet = curNet - qty
                        End If
                    Else
                        ' Add to (or open) short
                        If curNet = 0 Then
                            curAvg = price
                            p(10) = txDate
                        Else
                            curAvg = (-curNet * curAvg + qty * price) / (-curNet + qty)
                        End If
                        curNet = curNet - qty
                    End If

                Case "ADJUSTCOST"
                    If curNet <> 0 Then curAvg = price

                Case "EXPIRE", "EXERCISE"
                    curNet = 0: curAvg = 0
                    p(10) = Empty
            End Select

            p(5) = curNet
            p(6) = curAvg
            p(7) = 0   ' deprecated

            ' Keep most-recent non-empty metadata
            If strategy <> "" Then p(11) = strategy
            If CStr(iv) <> "" And CStr(iv) <> "0" Then p(12) = iv
            If CStr(delta) <> "" And CStr(delta) <> "0" Then p(13) = delta
            If broker <> "" Then p(9) = broker
            If txMult > 0 Then p(8) = txMult
            If UnderLying <> "" Then p(14) = UnderLying
            If CStr(gamma) <> "" And CStr(gamma) <> "0" Then p(15) = gamma
            If CStr(theta) <> "" And CStr(theta) <> "0" Then p(16) = theta
            ' Conv Ratio: keep first non-zero value (fixed per warrant code)
            If CStr(convRatio) <> "" And CDbl(convRatio) > 0 Then
                If CStr(p(17)) = "" Or CDbl(p(17)) = 0 Then p(17) = convRatio
            End If
            If issuer <> "" Then p(18) = issuer
            dict(posKey) = p
SkipRow:
        Next i

        ' -- Step 2: Clear input cells in all 3 sections -------------------
        ' Only input columns are cleared; formula columns (E,M,N,O,Q,S,V etc.)
        ' are left untouched so they recalculate automatically.

        Dim colIdx As Variant

        ' COMPACT LAYOUT (8 rows per section):
        '   Futures  rows  7-14 (header=6, subtotal=15)
        '   Options  rows 19-26 (banner=17, header=18, subtotal=27)
        '   Warrants rows 31-38 (banner=29, header=30, subtotal=39)
        '   GrandTotal at row 41

        ' Futures input cols: A B C D F G H I J K L P R T U W
        Dim futInp As Variant
        futInp = Array(1, 2, 3, 4, 6, 7, 8, 9, 10, 11, 12, 16, 18, 20, 21, 23)
        For Each colIdx In futInp
            wsR.Range(wsR.Cells(7, CLng(colIdx)), wsR.Cells(14, CLng(colIdx))).ClearContents
        Next colIdx

        ' Options input cols: A B C D F G H I J K L M N O S T U V W X AB
        Dim optInp As Variant
        optInp = Array(1, 2, 3, 4, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 19, 20, 21, 22, 23, 24, 28)
        For Each colIdx In optInp
            wsR.Range(wsR.Cells(19, CLng(colIdx)), wsR.Cells(26, CLng(colIdx))).ClearContents
        Next colIdx

        ' Warrants input cols: A B C D F G H I J K L M N O S T
        Dim wrtInp As Variant
        wrtInp = Array(1, 2, 3, 4, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 19, 20)
        For Each colIdx In wrtInp
            wsR.Range(wsR.Cells(31, CLng(colIdx)), wsR.Cells(38, CLng(colIdx))).ClearContents
        Next colIdx

        ' -- Step 3: Write active positions --------------------------------
        Dim futRow As Long, optRow As Long, wrtRow As Long
        futRow = 7: optRow = 19: wrtRow = 31
        Dim cntFut As Long, cntOpt As Long, cntWrt As Long
        cntFut = 0: cntOpt = 0: cntWrt = 0

        ' Bug B3 fix: sort dict keys for stable display order across rebuilds.
        ' FUT section sorted by ticker -> expiry; OPT by underlying -> expiry ->
        ' type (Call before Put) -> strike; WRT by ticker.
        Dim sortedKeys() As String
        sortedKeys = SortPositionKeys(dict)

        Dim k As Variant
        Dim ki As Long
        For ki = LBound(sortedKeys) To UBound(sortedKeys)
            k = sortedKeys(ki)
            Dim pos() As Variant
            pos = dict(k)

            Dim netQty As Double
            netQty = CDbl(pos(5))
            If Abs(netQty) < 0.000001 Then GoTo SkipPos   ' closed/expired

            ' Avg cost of currently OPEN position (resets on round-trip close).
            ' pos(6) now stores avgCost directly -- not cumulative / cumQty.
            Dim entryPx As Double
            entryPx = CDbl(pos(6))

            Dim ks As String: ks = CStr(k)

            If Left(ks, 4) = "FUT|" Then
                ' -- Futures -----------------------------------------------
                If futRow > 14 Then GoTo SkipPos   ' section full (8 rows, 7-14)
                wsR.Cells(futRow, 1).Value = pos(1)                        ' A CONTRACT
                wsR.Cells(futRow, 3).Value = pos(10)                       ' C ENTRY DT
                wsR.Cells(futRow, 4).Value = pos(4)                        ' D EXPIRY
                wsR.Cells(futRow, 6).Value = pos(9)                        ' F BROKER
                wsR.Cells(futRow, 8).Value = IIf(netQty > 0, "Long", "Short")  ' H DIRECTION
                wsR.Cells(futRow, 9).Value = Abs(netQty)                   ' I CONTRACTS
                wsR.Cells(futRow, 10).Value = CDbl(pos(8))                 ' J MULTIPLIER
                wsR.Cells(futRow, 11).Value = entryPx                      ' K ENTRY PX
                wsR.Cells(futRow, 18).Value = 1                            ' R DELTA = 1
                futRow = futRow + 1: cntFut = cntFut + 1

            ElseIf Left(ks, 4) = "OPT|" Then
                ' -- Options -----------------------------------------------
                If optRow > 26 Then GoTo SkipPos   ' section full (8 rows, 19-26)
                ' CONTRACT label: e.g. "TXO Call 18000"
                wsR.Cells(optRow, 1).Value = pos(1) & " " & pos(2) & " " & CStr(pos(3))
                wsR.Cells(optRow, 3).Value = pos(10)                       ' C ENTRY DT
                wsR.Cells(optRow, 4).Value = pos(4)                        ' D EXPIRY
                wsR.Cells(optRow, 6).Value = pos(9)                        ' F BROKER
                wsR.Cells(optRow, 8).Value = pos(2)                        ' H TYPE (Call/Put)
                wsR.Cells(optRow, 9).Value = CDbl(pos(3))                  ' I STRIKE
                wsR.Cells(optRow, 10).Value = pos(14)                      ' J UNDERLYING (name)
                ' CONTRACTS is signed: positive=long, negative=short
                ' NET DELTA formula =IF(H="Put",-L*M*ABS(T),L*M*T) handles sign correctly
                wsR.Cells(optRow, 12).Value = netQty                       ' L CONTRACTS (signed)
                wsR.Cells(optRow, 13).Value = CDbl(pos(8))                 ' M MULTIPLIER
                wsR.Cells(optRow, 14).Value = entryPx                      ' N ENTRY PREM
                wsR.Cells(optRow, 19).Value = pos(12)                      ' S IV
                wsR.Cells(optRow, 20).Value = pos(13)                      ' T DELTA
                wsR.Cells(optRow, 21).Value = pos(15)                      ' U GAMMA
                wsR.Cells(optRow, 22).Value = pos(16)                      ' V THETA/DAY
                wsR.Cells(optRow, 28).Value = pos(11)                      ' AB STRATEGY TAG
                optRow = optRow + 1: cntOpt = cntOpt + 1

            ElseIf Left(ks, 4) = "WRT|" Then
                ' -- Warrants ----------------------------------------------
                If wrtRow > 38 Then GoTo SkipPos   ' section full (8 rows, 31-38)
                ' Pad warrant code to 6 digits (preserves leading zeros lost
                ' when Transactions stored ticker as Number, e.g. 66669 -> 066669)
                Dim wCodeStr As String: wCodeStr = CStr(pos(1))
                If IsNumeric(wCodeStr) And Len(wCodeStr) <= 6 Then
                    wCodeStr = Format(Val(wCodeStr), "000000")
                End If
                wsR.Cells(wrtRow, 1).NumberFormat = "@"
                wsR.Cells(wrtRow, 1).Value = wCodeStr                       ' A CODE
                wsR.Cells(wrtRow, 3).Value  = pos(10)                      ' C ENTRY DT
                wsR.Cells(wrtRow, 4).Value  = pos(4)                       ' D EXPIRY
                wsR.Cells(wrtRow, 6).Value  = pos(9)                       ' F BROKER
                wsR.Cells(wrtRow, 7).Value  = pos(2)                       ' G TYPE
                wsR.Cells(wrtRow, 8).Value  = CStr(pos(18))                ' H ISSUER
                wsR.Cells(wrtRow, 9).Value  = pos(14)                      ' I UNDERLYING (name)
                wsR.Cells(wrtRow, 11).Value = CDbl(pos(3))                 ' K STRIKE
                ' L CONV RATIO -- required for Warrant_Delta = BS_Delta / convRatio
                If CStr(pos(17)) <> "" And CDbl(pos(17)) > 0 Then
                    wsR.Cells(wrtRow, 12).Value = CDbl(pos(17))            ' L CONV RATIO
                End If
                wsR.Cells(wrtRow, 13).Value = Abs(netQty)                  ' M SHARES (always +)
                wsR.Cells(wrtRow, 14).Value = entryPx                      ' N ENTRY PX
                wsR.Cells(wrtRow, 19).Value = pos(12)                      ' S IV
                wsR.Cells(wrtRow, 20).Value = pos(13)                      ' T DELTA
                wrtRow = wrtRow + 1: cntWrt = cntWrt + 1
            End If
SkipPos:
        Next ki

        ' -- Step 4: Fetch live prices + Greeks + Realize -----------------
        ' FetchAllPrices internally calls:
        '   RefreshGreeks               (per-leg Greeks from BS)
        '   RealizeClosedTransactions   (derive realized PnL into Realized)
        ' so one call covers all three downstream steps.
        Call FetchAllPrices(silent:=True)

        ' -- Step 5: Append today's snapshot to HistoryLog -----------------
        ' Runs after FetchAllPrices so UNDL LAST / Greeks / RealizedPnL are
        ' all current before the row is written. DailySnapshot skips if a row
        ' for today already exists (idempotent).
        Call DailySnapshot

        ' -- Step 6: Refresh combined RR4+RR5 performance sheet + CPLog -----
        ' Runs right after DailySnapshot so it reuses the wsR.Calculate that
        ' DailySnapshot just forced (RR5 summary cells are guaranteed fresh).
        ' Defensive: CombinedPerf module may not be imported into every copy
        ' of this workbook, so a missing module must not abort the rebuild.
        On Error Resume Next
        Application.Run "CombinedPerf.RefreshCombinedPerf"
        On Error GoTo ErrHandler

        LogDebug wsD, "RebuildPositionsFromTransactions", "OK", _
            "Fut=" & cntFut & " Opt=" & cntOpt & " Wrt=" & cntWrt

        ' Count realized rows for the summary box
        Dim rlCount As Long
        On Error Resume Next
        rlCount = ThisWorkbook.Sheets("Realized"). _
            Cells(ThisWorkbook.Sheets("Realized").Rows.Count, 1).End(xlUp).Row - 1
        On Error GoTo ErrHandler
        If rlCount < 0 Then rlCount = 0

        MsgBox "Rebuild + Fetch + Realize + Snapshot complete!" & vbCrLf & vbCrLf & _
               "Open positions:" & vbCrLf & _
               "  Futures : " & cntFut & vbCrLf & _
               "  Options : " & cntOpt & vbCrLf & _
               "  Warrants: " & cntWrt & vbCrLf & vbCrLf & _
               "Realized (from Transactions): " & rlCount & " row(s)" & vbCrLf & vbCrLf & _
               "Auto-filled:" & vbCrLf & _
               "  - UNDL LAST / LAST / LAST PREM  (TPEx + TWSE + Yahoo)" & vbCrLf & _
               "  - Futures MARGIN REQ            (TickerMap)" & vbCrLf & _
               "  - Greeks (Delta/Gamma/Theta/Vega/Rho)" & vbCrLf & _
               "  - Realized PnL                  (BUY/SELL/EXPIRE matching)" & vbCrLf & _
               "  - HistoryLog snapshot           (TX/SPY/IV_Avg/Summary)" & vbCrLf & _
               "  - CombinedPerf + CPLog row      (RR4+RR5 combined)", _
               vbInformation, "RR5 Rebuild"
        Exit Sub

ErrHandler:
        LogDebug wsD, "RebuildPositionsFromTransactions", "ERROR", _
            "Row " & i & " | " & Err.Description
        MsgBox "Rebuild failed at row " & i & ": " & Err.Description, vbCritical
    End Sub

    ' =====================================================================
    ' REALIZE CLOSED TRANSACTIONS
    ' ---------------------------------------------------------------------
    ' Walks Transactions chronologically per position key (Ticker / Type /
    ' Strike / Expiry). Detects closure events and writes one row per
    ' realized batch to the Realized worksheet.
    '
    '   BUY    -> if currently SHORT, COVER (realize short PnL); excess opens long
    '   SELL   -> if currently LONG,  CLOSE (realize long PnL);  excess opens short
    '   EXPIRE -> close ALL remaining at price (price=0 means worthless)
    '   EXERCISE -> close ALL remaining at price (price = intrinsic value)
    '
    ' Long-side  PnL per unit = (sell_price - avg_long_cost)
    ' Short-side PnL per unit = (avg_short_cost - cover_price)
    ' Multiplied by qty * multiplier and converted to TWD (col I) when underlying is USD.
    '
    ' Idempotent: clears Realized rows 2+ before re-deriving from full Transactions.
    ' =====================================================================
    Public Sub RealizeClosedTransactions(Optional silent As Boolean = False)
        On Error GoTo Fail

        Dim wsT As Worksheet, wsRl As Worksheet, wsR As Worksheet, wsD As Worksheet
        Set wsT = ThisWorkbook.Sheets("Transactions")
        Set wsRl = ThisWorkbook.Sheets("Realized")
        Set wsR = ThisWorkbook.Sheets("RR5")
        Set wsD = ThisWorkbook.Sheets("DebugLog")

        Dim usdRate As Double
        usdRate = ToDbl(wsR.Range("D1").Value)
        If usdRate <= 0 Then usdRate = 31.6

        Dim lastRow As Long
        lastRow = wsT.Cells(wsT.Rows.Count, 1).End(xlUp).Row
        If lastRow < 2 Then
            If Not silent Then MsgBox "Transactions sheet is empty.", vbInformation
            Exit Sub
        End If

        ' --- Read all transactions into array ---
        Dim txns As Variant
        txns = wsT.Range("A2:X" & lastRow).Value
        Dim nTx As Long: nTx = UBound(txns, 1)

        ' --- Sort by date (col 2). Simple bubble sort -- small N. ---
        Dim ord() As Long: ReDim ord(1 To nTx)
        Dim i As Long, j As Long
        For i = 1 To nTx: ord(i) = i: Next i
        For i = 1 To nTx - 1
            For j = 1 To nTx - i
                Dim dA As Date, dB As Date
                On Error Resume Next
                dA = CDate(txns(ord(j), 2))
                dB = CDate(txns(ord(j + 1), 2))
                On Error GoTo Fail
                If dA > dB Then
                    Dim tmp As Long: tmp = ord(j): ord(j) = ord(j + 1): ord(j + 1) = tmp
                End If
            Next j
        Next i

        ' --- Per-position state dict ---
        ' state array layout (Variant array, 0-based):
        '   0=netQty (signed: + long, - short)
        '   1=avgCost (always positive)
        '   2=openDate (Date or Empty)
        '   3=entryIV    4=entryDelta
        '   5=ticker     6=txType    7=strike    8=expiry    9=broker
        '   10=strategy  11=isUSD (Boolean)
        Dim posState As Object: Set posState = CreateObject("Scripting.Dictionary")
        Dim outRow As Long: outRow = 2
        Dim realizedCount As Long: realizedCount = 0

        ' --- Clear existing Realized rows 2+ ---
        Dim rlLast As Long
        rlLast = wsRl.Cells(wsRl.Rows.Count, 1).End(xlUp).Row
        If rlLast >= 2 Then wsRl.Range("A2:T" & rlLast).ClearContents

        For i = 1 To nTx
            Dim r As Long: r = ord(i)
            Dim txDate As Date
            On Error Resume Next: txDate = CDate(txns(r, 2)): On Error GoTo Fail
            If txDate = 0 Then GoTo NextTx

            Dim ticker As String: ticker = Trim(CStr(txns(r, 3)))
            If ticker = "" Then GoTo NextTx
            ' Restore leading zeros if numeric warrant code lost them
            If IsNumeric(ticker) And Len(ticker) < 6 Then
                Dim numVal As Long: numVal = CLng(Val(ticker))
                If numVal >= 1000 And numVal <= 999999 Then
                    ticker = Format(numVal, "000000")
                End If
            End If

            Dim action As String: action = UCase(Trim(CStr(txns(r, 4))))
            Dim txType As String: txType = Trim(CStr(txns(r, 5)))
            Dim strike As Variant: strike = txns(r, 6)
            Dim expiry As Variant: expiry = txns(r, 7)
            Dim qty As Double: qty = ToDbl(txns(r, 8))
            Dim mult As Double: mult = ToDbl(txns(r, 9))
            If mult = 0 Then mult = 1
            Dim price As Double: price = ToDbl(txns(r, 10))
            Dim strategy As String: strategy = Trim(CStr(txns(r, 16)))
            Dim ivAtEntry As Double: ivAtEntry = ToDbl(txns(r, 17))
            Dim deltaAtEntry As Double: deltaAtEntry = ToDbl(txns(r, 18))
            Dim broker As String: broker = Trim(CStr(txns(r, 20)))
            Dim underly As String: underly = Trim(CStr(txns(r, 21)))

            If qty <= 0 Then GoTo NextTx

            ' --- Build position key ---
            Dim posKey As String
            Select Case txType
                Case "Future"
                    posKey = "FUT|" & ticker & "|" & CStr(expiry) & "|" & broker
                Case "Call", "Put"
                    posKey = "OPT|" & ticker & "|" & txType & "|" & CStr(strike) & "|" & CStr(expiry) & "|" & broker
                Case "Call Warrant", "Put Warrant", "Bull Cert", "Bear Cert"
                    posKey = "WRT|" & ticker & "|" & broker
                Case Else
                    posKey = "EQTY|" & ticker & "|" & broker
            End Select

            ' --- Determine USD vs TWD ---
            Dim isUSD As Boolean
            isUSD = NotTaiwanUnderlying(underly, ticker)

            ' --- Init state if new key ---
            If Not posState.Exists(posKey) Then
                Dim newS(11) As Variant
                newS(0) = 0#: newS(1) = 0#: newS(2) = Empty
                newS(3) = 0#: newS(4) = 0#
                newS(5) = ticker: newS(6) = txType
                newS(7) = strike: newS(8) = expiry
                newS(9) = broker: newS(10) = strategy
                newS(11) = isUSD
                posState(posKey) = newS
            End If

            Dim st As Variant: st = posState(posKey)
            Dim netQty As Double: netQty = CDbl(st(0))
            Dim avgCost As Double: avgCost = CDbl(st(1))

            Select Case action
                Case "BUY"
                    If netQty < 0 Then
                        ' COVER short position
                        Dim coverQty As Double
                        coverQty = qty
                        If coverQty > -netQty Then coverQty = -netQty

                        Dim pnlCover As Double
                        pnlCover = (avgCost - price) * coverQty * mult
                        Call WriteRealizedRow(wsRl, outRow, st, coverQty, _
                                              avgCost, price, pnlCover, _
                                              txDate, ivAtEntry, deltaAtEntry, _
                                              usdRate, strategy, "Cover")
                        outRow = outRow + 1
                        realizedCount = realizedCount + 1

                        netQty = netQty + coverQty
                        qty = qty - coverQty
                        If Abs(netQty) < 0.000001 Then
                            netQty = 0: avgCost = 0
                            st(2) = Empty: st(3) = 0#: st(4) = 0#: st(10) = ""
                        End If
                    End If
                    If qty > 0 Then
                        ' Adds to (or opens) long
                        If netQty = 0 Then
                            avgCost = price
                            st(2) = txDate: st(3) = ivAtEntry: st(4) = deltaAtEntry
                            st(10) = strategy
                        Else
                            avgCost = (netQty * avgCost + qty * price) / (netQty + qty)
                        End If
                        netQty = netQty + qty
                    End If

                Case "SELL"
                    If netQty > 0 Then
                        ' CLOSE long position
                        Dim closeQty As Double
                        closeQty = qty
                        If closeQty > netQty Then closeQty = netQty

                        Dim pnlClose As Double
                        pnlClose = (price - avgCost) * closeQty * mult
                        Call WriteRealizedRow(wsRl, outRow, st, closeQty, _
                                              avgCost, price, pnlClose, _
                                              txDate, ivAtEntry, deltaAtEntry, _
                                              usdRate, strategy, "Close")
                        outRow = outRow + 1
                        realizedCount = realizedCount + 1

                        netQty = netQty - closeQty
                        qty = qty - closeQty
                        If Abs(netQty) < 0.000001 Then
                            netQty = 0: avgCost = 0
                            st(2) = Empty: st(3) = 0#: st(4) = 0#: st(10) = ""
                        End If
                    End If
                    If qty > 0 Then
                        ' Adds to (or opens) short
                        If netQty = 0 Then
                            avgCost = price
                            st(2) = txDate: st(3) = ivAtEntry: st(4) = deltaAtEntry
                            st(10) = strategy
                        Else
                            ' netQty < 0 already
                            avgCost = (-netQty * avgCost + qty * price) / (-netQty + qty)
                        End If
                        netQty = netQty - qty
                    End If

                Case "EXPIRE", "EXERCISE"
                    Dim closePx As Double: closePx = price   ' user supplies 0 for worthless, or intrinsic
                    If netQty > 0 Then
                        Dim pnlExpL As Double
                        pnlExpL = (closePx - avgCost) * netQty * mult
                        Call WriteRealizedRow(wsRl, outRow, st, netQty, _
                                              avgCost, closePx, pnlExpL, _
                                              txDate, 0#, 0#, _
                                              usdRate, strategy, action)
                        outRow = outRow + 1
                        realizedCount = realizedCount + 1
                    ElseIf netQty < 0 Then
                        Dim pnlExpS As Double
                        pnlExpS = (avgCost - closePx) * (-netQty) * mult
                        Call WriteRealizedRow(wsRl, outRow, st, -netQty, _
                                              avgCost, closePx, pnlExpS, _
                                              txDate, 0#, 0#, _
                                              usdRate, strategy, action)
                        outRow = outRow + 1
                        realizedCount = realizedCount + 1
                    End If
                    netQty = 0: avgCost = 0
                    st(2) = Empty: st(3) = 0#: st(4) = 0#: st(10) = ""

                Case "ADJUSTCOST"
                    ' Override avg cost to current price; doesn't realize
                    If netQty <> 0 Then avgCost = price

            End Select

            st(0) = netQty: st(1) = avgCost
            posState(posKey) = st

NextTx:
        Next i

        LogDebug wsD, "RealizeClosedTransactions", "OK", _
            "Realized rows written: " & realizedCount

        If Not silent Then
            MsgBox realizedCount & " realized transaction(s) derived from Transactions.", _
                   vbInformation, "Realize Closed Transactions"
        End If
        Exit Sub
Fail:
        LogDebug wsD, "RealizeClosedTransactions", "ERROR", Err.Description
        If Not silent Then MsgBox "RealizeClosedTransactions failed: " & Err.Description, vbCritical
    End Sub

    ' ---------------------------------------------------------------------
    ' Helper: write one realized row to Realized worksheet.
    ' Cols: A TICKER  B TYPE  C STRATEGY  D PNL  E SHARES  F AVG COST
    '       G NET AMT H STRATEGY TAG  I PNL(TWD)  J DATE  K BROKER
    '       L STRIKE  M EXPIRY  N DELTA ENTRY  O DELTA EXIT
    '       P IV ENTRY  Q IV EXIT  R THETA COLLECTED  S HOLD DAYS  T Notes
    ' ---------------------------------------------------------------------
    Private Sub WriteRealizedRow(wsRl As Worksheet, rowN As Long, st As Variant, _
                                  closeQty As Double, avgCost As Double, _
                                  closePx As Double, pnlRaw As Double, _
                                  closeDate As Date, ivExit As Double, _
                                  deltaExit As Double, usdRate As Double, _
                                  strategy As String, noteTag As String)

        Dim ticker As String: ticker = CStr(st(5))
        Dim txType As String: txType = CStr(st(6))
        Dim strike As Variant: strike = st(7)
        Dim expiry As Variant: expiry = st(8)
        Dim broker As String: broker = CStr(st(9))
        Dim stratStored As String: stratStored = CStr(st(10))
        Dim openDate As Variant: openDate = st(2)
        Dim ivEntry As Double: ivEntry = ToDbl(st(3))
        Dim deltaEntry As Double: deltaEntry = ToDbl(st(4))
        Dim isUSD As Boolean: isUSD = CBool(st(11))

        Dim pnlTWD As Double
        If isUSD Then
            pnlTWD = pnlRaw * usdRate
        Else
            pnlTWD = pnlRaw
        End If

        Dim netAmt As Double: netAmt = closeQty * closePx
        Dim holdDays As Long
        If IsDate(openDate) Then holdDays = closeDate - CDate(openDate)
        Dim displayStrat As String
        If stratStored <> "" Then displayStrat = stratStored Else displayStrat = strategy

        With wsRl
            ' Force ticker col to text so leading zeros stick (warrant codes)
            .Cells(rowN, 1).NumberFormat = "@"
            .Cells(rowN, 1).Value = ticker
            .Cells(rowN, 2).Value = txType
            .Cells(rowN, 3).Value = displayStrat
            .Cells(rowN, 4).Value = pnlRaw
            .Cells(rowN, 5).Value = closeQty
            .Cells(rowN, 6).Value = avgCost
            .Cells(rowN, 7).Value = netAmt
            .Cells(rowN, 8).Value = displayStrat
            .Cells(rowN, 9).Value = pnlTWD
            .Cells(rowN, 10).Value = closeDate
            .Cells(rowN, 11).Value = broker
            .Cells(rowN, 12).Value = strike
            .Cells(rowN, 13).Value = expiry
            .Cells(rowN, 14).Value = deltaEntry
            .Cells(rowN, 15).Value = deltaExit
            .Cells(rowN, 16).Value = ivEntry
            .Cells(rowN, 17).Value = ivExit
            ' Col 18 R THETA COLLECTED -- estimate omitted (would need full theta trail)
            .Cells(rowN, 19).Value = holdDays
            .Cells(rowN, 20).Value = noteTag & IIf(isUSD, " (USD)", "")

            .Cells(rowN, 4).NumberFormat = "#,##0.00;(#,##0.00)"
            .Cells(rowN, 6).NumberFormat = "#,##0.00"
            .Cells(rowN, 7).NumberFormat = "#,##0.00"
            .Cells(rowN, 9).NumberFormat = "#,##0;(#,##0)"
            .Cells(rowN, 10).NumberFormat = "yyyy/m/d"
            .Cells(rowN, 13).NumberFormat = "yyyy/m/d"
            .Cells(rowN, 14).NumberFormat = "0.0000"
            .Cells(rowN, 15).NumberFormat = "0.0000"
            .Cells(rowN, 16).NumberFormat = "0.00%"
            .Cells(rowN, 17).NumberFormat = "0.00%"
        End With
    End Sub

    Private Function NotTaiwanUnderlying(underly As String, ticker As String) As Boolean
        ' True if the position is denominated in USD (CBOE / US stocks).
        ' Heuristic: Taiwan if underlying is TWII/TAIEX/TXO or numeric 4/6-digit.
        Dim u As String: u = UCase(Trim(underly))
        Dim t As String: t = UCase(Trim(ticker))
        If u = "TWII" Or u = "TAIEX" Or u = "TXO" Or u = "^TWII" Then
            NotTaiwanUnderlying = False
            Exit Function
        End If
        If t = "TXO" Then NotTaiwanUnderlying = False: Exit Function
        ' Strip .TW / .TWO and test numeric
        Dim uStripped As String: uStripped = Replace(Replace(u, ".TWO", ""), ".TW", "")
        If IsNumeric(uStripped) And (Len(uStripped) = 4 Or Len(uStripped) = 6) Then
            NotTaiwanUnderlying = False
            Exit Function
        End If
        Dim tStripped As String: tStripped = Replace(Replace(t, ".TWO", ""), ".TW", "")
        If IsNumeric(tStripped) And (Len(tStripped) = 4 Or Len(tStripped) = 6) Then
            NotTaiwanUnderlying = False
            Exit Function
        End If
        NotTaiwanUnderlying = True   ' default: USD
    End Function

    Private Function ToDbl(v As Variant) As Double
        If IsNumeric(v) Then ToDbl = CDbl(v) Else ToDbl = 0
    End Function

    ' Bug B3 helper: produce a stable sort order for position dict keys.
    ' Section priority FUT < OPT < WRT; within each section, sort by sort key
    ' built from (ticker, expiry, type, strike) padded for lexicographic order.
    Private Function SortPositionKeys(dict As Object) As String()
        Dim n As Long: n = dict.Count
        Dim keys() As String
        Dim sortVals() As String
        ReDim keys(0 To n - 1)
        ReDim sortVals(0 To n - 1)

        Dim i As Long: i = 0
        Dim k As Variant
        For Each k In dict.Keys
            keys(i) = CStr(k)
            Dim p() As Variant: p = dict(k)
            Dim sectionRank As String
            If Left(CStr(k), 4) = "FUT|" Then
                sectionRank = "1"
            ElseIf Left(CStr(k), 4) = "OPT|" Then
                sectionRank = "2"
            ElseIf Left(CStr(k), 4) = "WRT|" Then
                sectionRank = "3"
            Else
                sectionRank = "9"
            End If
            Dim ticker As String: ticker = PadRight(CStr(p(1)), 12)
            Dim expStr As String
            If IsDate(p(4)) Then
                expStr = Format(CDate(p(4)), "yyyymmdd")
            Else
                expStr = "99999999"
            End If
            Dim typeRank As String
            Select Case CStr(p(2))
                Case "Call":         typeRank = "1"
                Case "Put":          typeRank = "2"
                Case "Call Warrant": typeRank = "1"
                Case "Put Warrant":  typeRank = "2"
                Case "Bull Cert":    typeRank = "3"
                Case "Bear Cert":    typeRank = "4"
                Case Else:           typeRank = "5"
            End Select
            Dim strikeStr As String
            If IsNumeric(p(3)) Then
                strikeStr = Format(CDbl(p(3)), "0000000.00")
            Else
                strikeStr = "0000000.00"
            End If
            sortVals(i) = sectionRank & "|" & ticker & "|" & expStr & "|" & typeRank & "|" & strikeStr
            i = i + 1
        Next k

        ' Simple insertion sort (n is small -- at most ~40 positions)
        Dim a As Long, b As Long, tmpS As String, tmpK As String
        For a = 0 To n - 2
            For b = a + 1 To n - 1
                If sortVals(a) > sortVals(b) Then
                    tmpS = sortVals(a): sortVals(a) = sortVals(b): sortVals(b) = tmpS
                    tmpK = keys(a):     keys(a) = keys(b):         keys(b) = tmpK
                End If
            Next b
        Next a

        SortPositionKeys = keys
    End Function

    Private Function PadRight(s As String, n As Long) As String
        If Len(s) >= n Then
            PadRight = Left(s, n)
        Else
            PadRight = s & String(n - Len(s), " ")
        End If
    End Function

    ' =====================================================================
    ' LIVE QUOTE FETCH (Yahoo Finance public chart endpoint)
    ' =====================================================================
    '
    ' Workflow:
    '   1. AddTransactions   -> Transactions sheet
    '   2. RebuildPositionsFromTransactions -> RR5 sheet (UNDL LAST blank)
    '   3. FetchAllPrices    -> auto-fills UNDL LAST / LAST PREM / LAST PX
    '                           and immediately re-runs RefreshGreeks
    '   4. Greeks (Delta/Gamma/Theta/Vega/Rho) appear in RR5 cols T-X
    '

    Public Function FetchPrice(ticker As String) As Double
        ' Pulls regularMarketPrice from Yahoo Finance chart API.
        ' Returns 0 on any failure (network, ticker not found, parse error).
        ' Bug B7 fix: send User-Agent header so Yahoo doesn't 403 us.
        On Error GoTo Fail
        If Trim(ticker) = "" Then Exit Function

        Dim req As Object, url As String, resp As String
        Set req = CreateObject("MSXML2.XMLHTTP")
        url = "https://query1.finance.yahoo.com/v8/finance/chart/" & ticker
        req.Open "GET", url, False
        req.setRequestHeader "User-Agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36"
        req.send
        resp = req.responseText

        ' Locate the JSON key "regularMarketPrice":<number>
        Dim p As Long, q As Long, s As String
        p = InStr(resp, """regularMarketPrice"":")
        If p = 0 Then Exit Function
        p = p + Len("""regularMarketPrice"":")
        q = InStr(p, resp, ",")
        If q = 0 Then q = InStr(p, resp, "}")
        If q <= p Then Exit Function
        s = Mid(resp, p, q - p)
        FetchPrice = Val(s)
        Exit Function
Fail:
        FetchPrice = 0
    End Function

    Private Function YahooOptionTicker(underlying As String, expiry As Date, _
                                       optType As String, strike As Double) As String
        ' Yahoo Finance option ticker format:
        '   SYMBOL + yymmdd + [C|P] + 8-digit strike * 1000
        '   e.g.  AAPL250117C00150000  =  AAPL 2025-01-17 Call $150
        ' Works for US options only; Yahoo does NOT carry TAIFEX options.
        Dim sym As String: sym = UCase(Trim(underlying))
        Dim dt As String:  dt = Format(expiry, "yymmdd")
        Dim cp As String:  cp = IIf(UCase(optType) = "CALL", "C", "P")
        Dim sk As String:  sk = Format(strike * 1000, "00000000")
        YahooOptionTicker = sym & dt & cp & sk
    End Function

    ' =================================================================
    ' TAIFEX OPTION PRICE FETCH (via FinMind public API, EOD settlement)
    ' -----------------------------------------------------------------
    ' FinMind dataset 'TaiwanOptionDaily' returns a full day of option
    ' rows for one option_id (e.g. "TXO"). Free, no auth required.
    '
    ' Schema per row: contract_date / strike_price / call_put /
    '                 open / max / min / close / volume / trading_session
    '
    ' Contract code rules for TXO:
    '   Monthly   = 3rd Wednesday of month -> "YYYYMM"
    '   Weekly    = other Wednesdays       -> "YYYYMM" & "W" & wedIdx+1
    '               (W3 never exists -- that's the monthly slot)
    '
    ' Auto-falls back up to 5 business days if today's data not posted.
    ' Per-day responses are cached in a Static dict for the session.
    ' =================================================================
    Public Function FetchTaifexOptionPrice(optionID As String, expiry As Date, _
                                            strike As Double, optType As String, _
                                            Optional tradeDate As Variant) As Double
        On Error GoTo Fail
        Static cache As Object
        If cache Is Nothing Then Set cache = CreateObject("Scripting.Dictionary")

        Dim startDt As Date
        If IsMissing(tradeDate) Then
            startDt = PreviousBusinessDay(Date)
        Else
            startDt = CDate(tradeDate)
        End If

        Dim contractCode As String
        contractCode = ExpiryToTaifexContractCode(expiry)
        If contractCode = "" Then Exit Function

        Dim cpLower As String
        If InStr(UCase(optType), "CALL") > 0 Then
            cpLower = "call"
        ElseIf InStr(UCase(optType), "PUT") > 0 Then
            cpLower = "put"
        Else
            Exit Function
        End If

        ' Try up to 5 business days backward (handles holidays / late posting)
        Dim tryDt As Date: tryDt = startDt
        Dim attempt As Long
        For attempt = 1 To 5
            Dim dateStr As String: dateStr = Format(tryDt, "yyyy-mm-dd")
            Dim cacheKey As String
            cacheKey = UCase(optionID) & "|" & dateStr

            Dim respText As String
            If cache.Exists(cacheKey) Then
                respText = cache(cacheKey)
            Else
                Dim url As String
                url = "https://api.finmindtrade.com/api/v4/data" & _
                      "?dataset=TaiwanOptionDaily" & _
                      "&data_id=" & UCase(optionID) & _
                      "&start_date=" & dateStr & _
                      "&end_date=" & dateStr
                Dim req As Object
                Set req = CreateObject("MSXML2.XMLHTTP")
                req.Open "GET", url, False
                req.setRequestHeader "User-Agent", "Mozilla/5.0"
                req.send
                If req.Status = 200 Then
                    respText = req.responseText
                    cache(cacheKey) = respText
                Else
                    respText = ""
                End If
            End If

            ' Skip "no data" responses (empty data array)
            If Len(respText) > 0 And InStr(respText, """data"":[]") = 0 Then
                Dim cl As Double
                cl = ParseFinMindOptionClose(respText, contractCode, strike, cpLower)
                If cl > 0 Then
                    FetchTaifexOptionPrice = cl
                    Exit Function
                End If
            End If

            tryDt = PreviousBusinessDay(tryDt - 1)
        Next attempt

        FetchTaifexOptionPrice = 0
        Exit Function
Fail:
        FetchTaifexOptionPrice = 0
    End Function

    Public Function ExpiryToTaifexContractCode(expiry As Date) As String
        ' Returns "YYYYMM" for monthly contract (3rd Wed) or
        ' "YYYYMMW{n}" for weekly. Returns "" if expiry is not a Wednesday.
        If Weekday(expiry, vbMonday) <> 3 Then
            ' Not a Wednesday -- TAIFEX options expire on Wed; fall back
            ' to monthly code so caller still has something to query.
            ExpiryToTaifexContractCode = Format(expiry, "yyyymm")
            Exit Function
        End If

        Dim y As Integer, m As Integer
        y = Year(expiry): m = Month(expiry)

        ' Index of this Wednesday within the month (0-based)
        Dim wedIdx As Integer: wedIdx = -1
        Dim cur As Date: cur = DateSerial(y, m, 1)
        Dim lastDay As Date: lastDay = DateSerial(y, m + 1, 1) - 1
        Dim count As Integer: count = 0
        Do While cur <= lastDay
            If Weekday(cur, vbMonday) = 3 Then
                If cur = expiry Then wedIdx = count
                count = count + 1
            End If
            cur = cur + 1
        Loop

        If wedIdx = 2 Then
            ' 3rd Wednesday = monthly contract
            ExpiryToTaifexContractCode = Format(expiry, "yyyymm")
        Else
            ' Weekly: W1/W2/W4/W5 (W3 never exists)
            ExpiryToTaifexContractCode = Format(expiry, "yyyymm") & "W" & (wedIdx + 1)
        End If
    End Function

    Private Function ParseFinMindOptionClose(jsonText As String, contractCode As String, _
                                              strike As Double, cpLower As String) As Double
        ' Walks each {...} record block in the 'data' array one at a time,
        ' ensuring contract_date / strike_price / call_put / close /
        ' trading_session are all read from the SAME record.
        '
        ' Returns the LAST TRADED close price (used for MTM / unrealized P&L,
        ' matching what brokers display as the "deal price" / last traded price).
        ' Settlement_price (the official settlement price) is intentionally NOT
        ' used here -- it's the official daily mark for margin, not the
        ' realizable liquidation price.
        '
        ' Session preference: "position" (daytime regular) > "after_market".
        ' FinMind labels: "position" = 08:45-13:45, "after_market" = 15:00-04:50.

        Dim startP As Long, endP As Long
        startP = InStr(jsonText, """data"":[")
        If startP = 0 Then Exit Function
        startP = startP + Len("""data"":[")
        endP = InStrRev(jsonText, "]")
        If endP <= startP Then Exit Function
        Dim body As String: body = Mid(jsonText, startP, endP - startP)
        If Len(body) = 0 Then Exit Function

        Dim strikeKey As String
        If CDbl(CLng(strike)) = strike Then
            strikeKey = """strike_price"":" & CLng(strike)
        Else
            strikeKey = """strike_price"":" & strike
        End If

        Dim bestPrice As Double: bestPrice = -1
        Dim bestRank As Long:   bestRank = -1     ' 2 = close+position, 1 = close+after_market

        Dim pos As Long: pos = 1
        Do While pos <= Len(body)
            Dim recStart As Long: recStart = InStr(pos, body, "{")
            If recStart = 0 Then Exit Do
            Dim recEnd As Long: recEnd = InStr(recStart, body, "}")
            If recEnd = 0 Then Exit Do
            Dim rec As String: rec = Mid(body, recStart, recEnd - recStart + 1)

            If InStr(rec, """contract_date"":""" & contractCode & """") = 0 Then GoTo NextRec
            If InStr(rec, """call_put"":""" & cpLower & """") = 0 Then GoTo NextRec
            Dim hit As Boolean: hit = False
            If InStr(rec, strikeKey & ".0,") > 0 Then hit = True
            If Not hit And InStr(rec, strikeKey & ".0}") > 0 Then hit = True
            If Not hit And InStr(rec, strikeKey & ",") > 0 Then hit = True
            If Not hit And InStr(rec, strikeKey & "}") > 0 Then hit = True
            If Not hit Then GoTo NextRec

            ' --- Session ---
            Dim sess As String: sess = ""
            Dim sp As Long
            sp = InStr(rec, """trading_session"":""")
            If sp > 0 Then
                sp = sp + Len("""trading_session"":""")
                Dim eq As Long: eq = InStr(sp, rec, """")
                If eq > sp Then sess = Mid(rec, sp, eq - sp)
            End If
            Dim isPosition As Boolean
            isPosition = (sess = "position") Or (sess = "regular")

            ' --- close (last traded price) ---
            Dim cVal As Double: cVal = ExtractJsonNumber(rec, """close"":")
            If cVal <= 0 Then GoTo NextRec

            Dim thisRank As Long: thisRank = IIf(isPosition, 2, 1)
            If thisRank > bestRank Then
                bestRank = thisRank
                bestPrice = cVal
            End If
NextRec:
            pos = recEnd + 1
        Loop

        If bestPrice > 0 Then ParseFinMindOptionClose = bestPrice
    End Function

    Private Function ExtractJsonNumber(rec As String, key As String) As Double
        ' Finds key (e.g. """close"":") and extracts the following numeric value
        ' until comma or closing brace. Returns 0 if not found or unparseable.
        Dim p As Long, q As Long
        p = InStr(rec, key)
        If p = 0 Then Exit Function
        p = p + Len(key)
        q = InStr(p, rec, ",")
        If q = 0 Then q = InStr(p, rec, "}")
        If q <= p Then Exit Function
        ExtractJsonNumber = Val(Mid(rec, p, q - p))
    End Function

    ' =================================================================
    ' DIAGNOSTIC: Inspect what FinMind returns for the position in O31
    ' -----------------------------------------------------------------
    ' Run from Alt+F8 or Immediate window. Prints to Debug.Print (Ctrl+G
    ' Immediate window). Shows contract_code derivation, URL hit, total
    ' records, how many matched each criterion, sample matched record,
    ' and final returned price.
    ' =================================================================
    ' =================================================================
    ' DIAGNOSTIC: Inspect why a warrant row didn't fetch / compute
    ' Pass the actual row number (default 31 in compact layout)
    ' =================================================================
    Public Sub Debug_WarrantFetch(Optional wrtRow As Long = 31)
        Dim wsR As Worksheet: Set wsR = ThisWorkbook.Sheets("RR5")
        Dim wCode As String: wCode = Trim(CStr(wsR.Cells(wrtRow, 1).Value))

        Debug.Print "=== Debug_WarrantFetch (row " & wrtRow & ") ==="
        Debug.Print "Code raw  : '" & wCode & "'  (len=" & Len(wCode) & ")"
        If wCode = "" Then
            Debug.Print "EMPTY CODE -- this row has no warrant. Did you re-run RebuildPositionsFromTransactions?"
            Exit Sub
        End If

        Dim wPad As String
        If IsNumeric(wCode) Then
            wPad = Format(Val(wCode), "000000")
        Else
            wPad = wCode
        End If
        Debug.Print "Code pad  : '" & wPad & "'"

        ' Prints ALL sources (not just the one that wins) so a wrong-price
        ' report can be diagnosed by comparing them side by side -- e.g. a
        ' recycled warrant code where Yahoo still serves a stale quote from
        ' a previous, already-expired warrant that used the same code.
        Dim pxTPEx As Double, pxTWSE As Double, pxYahooTW As Double, pxYahooTWO As Double
        pxTPEx = FetchTPExWarrantPrice(wPad)
        Debug.Print wPad & " (TPEx OTC)  -> " & pxTPEx
        pxTWSE = FetchTWSEClose(wPad)
        Debug.Print wPad & " (TWSE)      -> " & pxTWSE
        pxYahooTW = FetchPrice(wPad & ".TW")
        Debug.Print wPad & ".TW  (Yahoo) -> " & pxYahooTW
        pxYahooTWO = FetchPrice(wPad & ".TWO")
        Debug.Print wPad & ".TWO (Yahoo) -> " & pxYahooTWO
        Debug.Print "Live priority order: TPEx -> TWSE -> Yahoo .TW -> Yahoo .TWO"
        Debug.Print "If TPEx/TWSE disagree with Yahoo, trust TPEx/TWSE (Yahoo can serve a stale price for a recycled warrant code)."

        Debug.Print ""
        Debug.Print "RR5 row " & wrtRow & " current cells:"
        Debug.Print "  A CODE       : " & wsR.Cells(wrtRow, 1).Value
        Debug.Print "  G TYPE       : " & wsR.Cells(wrtRow, 7).Value
        Debug.Print "  I UNDERLYING : " & wsR.Cells(wrtRow, 9).Value
        Debug.Print "  J UNDL LAST  : " & wsR.Cells(wrtRow, 10).Value
        Debug.Print "  K STRIKE     : " & wsR.Cells(wrtRow, 11).Value
        Debug.Print "  L CONV RATIO : " & wsR.Cells(wrtRow, 12).Value
        Debug.Print "  M SHARES     : " & wsR.Cells(wrtRow, 13).Value
        Debug.Print "  N ENTRY PX   : " & wsR.Cells(wrtRow, 14).Value
        Debug.Print "  O LAST PX    : " & wsR.Cells(wrtRow, 15).Value
        Debug.Print "  S IV         : " & wsR.Cells(wrtRow, 19).Value
        Debug.Print "  T DELTA      : " & wsR.Cells(wrtRow, 20).Value
        Debug.Print "  X GAMMA      : " & wsR.Cells(wrtRow, 24).Value
        Debug.Print "  Y THETA/DAY  : " & wsR.Cells(wrtRow, 25).Value
        Debug.Print "  Z VEGA       : " & wsR.Cells(wrtRow, 26).Value
        Debug.Print "  AA NOM GEAR  : " & wsR.Cells(wrtRow, 27).Value
        Debug.Print "  AB EFF GEAR  : " & wsR.Cells(wrtRow, 28).Value

        Debug.Print ""
        Debug.Print "RefreshGreeks needs UNDL LAST + STRIKE + IV + CONV RATIO all > 0."
        Debug.Print "If any is 0/blank, Greeks won't compute."
    End Sub

    Public Sub Debug_TaifexFetch(Optional rrRow As Long = 31)
        On Error GoTo Fail
        Dim wsR As Worksheet: Set wsR = ThisWorkbook.Sheets("RR5")

        Dim contract As String, expiry As Variant, strike As Variant
        Dim optType As String, underlying As String
        contract  = CStr(wsR.Cells(rrRow, 1).Value)
        expiry    = wsR.Cells(rrRow, 4).Value
        strike    = wsR.Cells(rrRow, 9).Value
        optType   = Trim(CStr(wsR.Cells(rrRow, 8).Value))
        underlying = Trim(CStr(wsR.Cells(rrRow, 10).Value))

        Debug.Print "=== Debug_TaifexFetch (row " & rrRow & ") ==="
        Debug.Print "Contract  : " & contract
        Debug.Print "Underlying: " & underlying
        Debug.Print "Expiry    : " & expiry & "  (" & Format(CDate(expiry), "yyyy-mm-dd ddd") & ")"
        Debug.Print "Strike    : " & strike
        Debug.Print "Type      : " & optType

        Dim taifexID As String: taifexID = MapUnderlyingToTaifexOptionID(underlying)
        Dim contractCode As String: contractCode = ExpiryToTaifexContractCode(CDate(expiry))
        Debug.Print "Option ID : " & taifexID
        Debug.Print "Contract$ : " & contractCode

        Dim tradeDt As Date: tradeDt = PreviousBusinessDay(Date)
        Dim dateStr As String: dateStr = Format(tradeDt, "yyyy-mm-dd")
        Dim url As String
        url = "https://api.finmindtrade.com/api/v4/data?dataset=TaiwanOptionDaily" & _
              "&data_id=" & taifexID & "&start_date=" & dateStr & "&end_date=" & dateStr
        Debug.Print "URL       : " & url

        Dim req As Object: Set req = CreateObject("MSXML2.XMLHTTP")
        req.Open "GET", url, False
        req.setRequestHeader "User-Agent", "Mozilla/5.0"
        req.send
        Debug.Print "HTTP      : " & req.Status & "  (resp len " & Len(req.responseText) & ")"

        Dim resp As String: resp = req.responseText
        Dim startP As Long, endP As Long
        startP = InStr(resp, """data"":[")
        If startP > 0 Then
            startP = startP + Len("""data"":[")
            endP = InStrRev(resp, "]")
            Dim body As String: body = Mid(resp, startP, endP - startP)
            Debug.Print "Body len  : " & Len(body)

            Dim cpLower As String
            If InStr(UCase(optType), "CALL") > 0 Then cpLower = "call" Else cpLower = "put"

            Dim strikeKey As String
            strikeKey = """strike_price"":" & CLng(CDbl(strike))

            Dim pos As Long: pos = 1
            Dim totalRecs As Long, cMatch As Long, sMatch As Long, pMatch As Long
            Dim fullMatch As Long, sampleContract As String

            Do While pos <= Len(body)
                Dim rs As Long: rs = InStr(pos, body, "{")
                If rs = 0 Then Exit Do
                Dim re As Long: re = InStr(rs, body, "}")
                If re = 0 Then Exit Do
                Dim rec As String: rec = Mid(body, rs, re - rs + 1)
                totalRecs = totalRecs + 1

                Dim hC As Boolean: hC = InStr(rec, """contract_date"":""" & contractCode & """") > 0
                Dim hP As Boolean: hP = InStr(rec, """call_put"":""" & cpLower & """") > 0
                Dim hS As Boolean
                hS = (InStr(rec, strikeKey & ".0,") > 0) Or _
                     (InStr(rec, strikeKey & ".0}") > 0) Or _
                     (InStr(rec, strikeKey & ",") > 0) Or _
                     (InStr(rec, strikeKey & "}") > 0)

                If hC Then
                    cMatch = cMatch + 1
                    If sampleContract = "" Then sampleContract = rec
                End If
                If hS Then sMatch = sMatch + 1
                If hP Then pMatch = pMatch + 1
                If hC And hP And hS Then
                    fullMatch = fullMatch + 1
                    Debug.Print "  >>> FULL MATCH #" & fullMatch & ": " & rec
                End If
                pos = re + 1
            Loop

            Debug.Print "Records   : " & totalRecs
            Debug.Print "  contract_date match : " & cMatch
            Debug.Print "  strike_price match  : " & sMatch
            Debug.Print "  call_put match      : " & pMatch
            Debug.Print "  ALL THREE match     : " & fullMatch
            If fullMatch = 0 And sampleContract <> "" Then
                Debug.Print "First contract-only match (strike/CP didn't match):"
                Debug.Print "  " & Left(sampleContract, 500)
            End If
        End If

        Dim px As Double
        px = FetchTaifexOptionPrice(taifexID, CDate(expiry), CDbl(strike), optType)
        Debug.Print ""
        Debug.Print "FetchTaifexOptionPrice -> " & px
        Exit Sub
Fail:
        Debug.Print "ERROR: " & Err.Description
    End Sub

    Private Function PreviousBusinessDay(d As Date) As Date
        ' Simple Mon-Fri filter; does not account for TW holidays.
        ' FetchTaifexOptionPrice retries 5 days, so holidays self-correct
        ' as long as a real trading day exists within 5 business days.
        Dim r As Date: r = d
        Do While Weekday(r, vbMonday) > 5
            r = r - 1
        Loop
        PreviousBusinessDay = r
    End Function

    Private Function MapUnderlyingToTaifexOptionID(underlying As String) As String
        ' Map RR5 J UNDERLYING value to FinMind option_id.
        '   "TWII" / "TAIEX" / "TXO" -> "TXO"  (Taiwan index options)
        '   Individual 4-digit Taiwan codes -> stock-option series mapping
        '     (FinMind uses stock code for individual stock options, e.g. "2330")
        ' Returns "" if no mapping (caller should skip TAIFEX path).
        Dim u As String: u = UCase(Trim(underlying))
        Select Case u
            Case "TWII", "TAIEX", "TXO", "^TWII": MapUnderlyingToTaifexOptionID = "TXO"
            Case Else
                If IsNumeric(u) And Len(u) = 4 Then
                    MapUnderlyingToTaifexOptionID = u
                Else
                    MapUnderlyingToTaifexOptionID = ""
                End If
        End Select
    End Function

    Public Function LookupFuturesMargin(product As String) As Double
        ' Walks TickerMap (col B = Code, col D = Margin_TWD) for a matching product.
        ' Returns 0 if not found.
        On Error Resume Next
        Dim wsM As Worksheet
        Set wsM = ThisWorkbook.Sheets("TickerMap")
        Dim p As String: p = UCase(Trim(product))
        If p = "" Then Exit Function
        Dim i As Long
        For i = 2 To 20
            If UCase(Trim(CStr(wsM.Cells(i, 2).Value))) = p Then
                If IsNumeric(wsM.Cells(i, 4).Value) Then
                    LookupFuturesMargin = CDbl(wsM.Cells(i, 4).Value)
                End If
                Exit Function
            End If
        Next i
    End Function

    Private Function NormalizeTicker(t As String) As String
        ' Map common Taiwan / US tickers to Yahoo Finance format.
        Dim tt As String: tt = UCase(Trim(t))
        Select Case tt
            Case "TWII", "TAIEX":  NormalizeTicker = "^TWII"
            Case "TXO":            NormalizeTicker = "^TWII"   ' TXO underlying is TAIEX
            Case "SPX":            NormalizeTicker = "^GSPC"
            Case "NDX":            NormalizeTicker = "^NDX"
            Case "DJI":            NormalizeTicker = "^DJI"
            Case "VIX":            NormalizeTicker = "^VIX"
            Case "RUT":            NormalizeTicker = "^RUT"
            Case Else
                ' Taiwan 4-digit stock codes (2330 -> 2330.TW)
                ' Taiwan warrants are typically 6-digit (.TW or .TWO)
                If IsNumeric(tt) And (Len(tt) = 4 Or Len(tt) = 6) Then
                    NormalizeTicker = tt & ".TW"
                Else
                    NormalizeTicker = tt   ' assume US ticker
                End If
        End Select
    End Function

    Public Sub FetchAllPrices(Optional silent As Boolean = False)
        ' Walks all three RR5 sections, fetches Yahoo quotes, fills price cells,
        ' then re-runs RefreshGreeks so Greeks pick up the new UNDL LAST.
        '
        ' silent=True: skip the final MsgBox. Used when called from
        ' RebuildPositionsFromTransactions which has its own summary box.
        On Error GoTo ErrHandler
        Dim wsR As Worksheet, wsD As Worksheet
        Set wsR = ThisWorkbook.Sheets("RR5")
        Set wsD = ThisWorkbook.Sheets("DebugLog")

        Application.StatusBar = "Fetching live prices..."

        Dim okN As Long, failN As Long, i As Long
        Dim tkr As String, px As Double

        ' --- Futures (rows 7-14): A=CONTRACT, G=PRODUCT, L=LAST, P=MARGIN REQ ---
        For i = 7 To 14
            tkr = Trim(CStr(wsR.Cells(i, 1).Value))
            If tkr <> "" Then
                Application.StatusBar = "Fetching Futures " & tkr & "..."
                px = FetchPrice(NormalizeTicker(tkr))
                If px > 0 Then
                    wsR.Cells(i, 12).Value = px    ' L LAST
                    okN = okN + 1
                Else
                    failN = failN + 1
                End If
                ' Auto-fill MARGIN REQ from TickerMap via PRODUCT code.
                ' Bug B4 fix: consolidate duplicate If/ElseIf branches that both
                ' did the same lookup. Fill if cell is empty/text OR numeric 0.
                Dim prod As String, mgn As Double
                Dim cellP As Variant
                prod = Trim(CStr(wsR.Cells(i, 7).Value))   ' G PRODUCT
                If prod <> "" Then
                    cellP = wsR.Cells(i, 16).Value          ' P MARGIN REQ
                    If Not IsNumeric(cellP) Or CDbl(cellP) = 0 Then
                        mgn = LookupFuturesMargin(prod)
                        If mgn > 0 Then
                            wsR.Cells(i, 16).Value = mgn
                            okN = okN + 1
                        End If
                    End If
                End If
            End If
        Next i

        ' --- Options (rows 19-26): J=UNDERLYING, K=UNDL LAST, O=LAST PREM ---
        ' UNDL LAST: fetched from underlying ticker (always tries).
        ' LAST PREM: TAIFEX TXO via FinMind, US options via Yahoo.
        For i = 19 To 26
            Dim contract As String
            contract = Trim(CStr(wsR.Cells(i, 1).Value))
            If contract <> "" Then
                tkr = Trim(CStr(wsR.Cells(i, 10).Value))   ' J UNDERLYING name
                If tkr <> "" Then
                    px = 0
                    Dim optUndlC As String
                    optUndlC = Replace(Replace(UCase(Trim(tkr)), ".TWO", ""), ".TW", "")
                    ' Yahoo Finance first
                    Application.StatusBar = "Fetching Opt UNDL " & tkr & " (Yahoo)..."
                    px = FetchPrice(NormalizeTicker(tkr))
                    ' TPEx -> TWSE as backup for TW numeric codes
                    If px <= 0 And IsNumeric(optUndlC) And Len(optUndlC) >= 4 And Len(optUndlC) <= 6 Then
                        Application.StatusBar = "Fetching Opt UNDL " & optUndlC & " (TPEx OTC)..."
                        px = FetchTPExStockPrice(optUndlC)
                        If px <= 0 Then
                            Application.StatusBar = "Fetching Opt UNDL " & optUndlC & " (TWSE)..."
                            px = FetchTWSEClose(optUndlC)
                        End If
                    End If
                    If px > 0 Then
                        wsR.Cells(i, 11).Value = px   ' K UNDL LAST
                        okN = okN + 1
                    Else
                        failN = failN + 1
                    End If

                    ' --- Option premium fetch ---
                    ' Route by underlying:
                    '   TWII / 4-digit TW code -> FinMind TAIFEX (EOD close)
                    '   Everything else        -> Yahoo Finance US option chain
                    Dim oStrike As Variant, oExpiry As Variant, oType As String
                    oStrike = wsR.Cells(i, 9).Value      ' I STRIKE
                    oExpiry = wsR.Cells(i, 4).Value      ' D EXPIRY
                    oType   = Trim(CStr(wsR.Cells(i, 8).Value))   ' H TYPE
                    Dim hasOptInputs As Boolean
                    hasOptInputs = (IsNumeric(oStrike) And IsDate(oExpiry) And _
                                    (oType = "Call" Or oType = "Put"))

                    Dim taifexID As String
                    taifexID = MapUnderlyingToTaifexOptionID(tkr)

                    If taifexID <> "" And hasOptInputs Then
                        ' TAIFEX path via FinMind
                        Application.StatusBar = "Fetching TAIFEX " & taifexID & _
                            " " & oType & " " & oStrike & " " & Format(oExpiry, "yyyy-mm-dd") & "..."
                        Dim taifexPx As Double
                        taifexPx = FetchTaifexOptionPrice(taifexID, CDate(oExpiry), _
                                                          CDbl(oStrike), oType)
                        If taifexPx > 0 Then
                            wsR.Cells(i, 15).Value = taifexPx    ' O LAST PREM
                            okN = okN + 1
                        Else
                            failN = failN + 1
                        End If
                    ElseIf taifexID = "" And hasOptInputs Then
                        ' US option via Yahoo
                        Dim optTkr As String, premPx As Double
                        optTkr = YahooOptionTicker(tkr, CDate(oExpiry), _
                                                   oType, CDbl(oStrike))
                        Application.StatusBar = "Fetching Opt PREM " & optTkr & "..."
                        premPx = FetchPrice(optTkr)
                        If premPx > 0 Then
                            wsR.Cells(i, 15).Value = premPx   ' O LAST PREM
                            okN = okN + 1
                        End If
                    End If
                End If
            End If
        Next i

        ' --- Warrants (rows 31-38): I=UNDERLYING, J=UNDL LAST;  A=CODE, O=LAST PX ---
        ' Price priority: TPEx OTC  ->  TWSE openapi/STOCK_DAY. (Yahoo fallback for
        ' the warrant's own price was removed -- see note below.)
        For i = 31 To 38
            Dim wCode As String
            wCode = Trim(CStr(wsR.Cells(i, 1).Value))
            If wCode <> "" Then
                ' --- Underlying spot price --- (unaffected by this fix: ordinary
                ' stock tickers aren't recycled the way warrant codes are, so Yahoo
                ' stays the primary source here.)
                tkr = Trim(CStr(wsR.Cells(i, 9).Value))    ' I UNDERLYING name/code
                If tkr <> "" Then
                    px = 0
                    Dim wUndlC As String
                    wUndlC = Replace(Replace(UCase(Trim(tkr)), ".TWO", ""), ".TW", "")
                    ' Yahoo Finance first
                    Application.StatusBar = "Fetching UNDL " & tkr & " (Yahoo)..."
                    px = FetchPrice(NormalizeTicker(tkr))
                    ' TPEx -> TWSE as backup for TW numeric codes
                    If px <= 0 And IsNumeric(wUndlC) And Len(wUndlC) >= 4 And Len(wUndlC) <= 6 Then
                        Application.StatusBar = "Fetching UNDL " & wUndlC & " (TPEx OTC)..."
                        px = FetchTPExStockPrice(wUndlC)
                        If px <= 0 Then
                            Application.StatusBar = "Fetching UNDL " & wUndlC & " (TWSE)..."
                            px = FetchTWSEClose(wUndlC)
                        End If
                    End If
                    If px > 0 Then
                        wsR.Cells(i, 10).Value = px   ' J UNDL LAST
                        okN = okN + 1
                    Else
                        failN = failN + 1
                    End If
                End If

                Dim wPad As String
                If IsNumeric(wCode) Then
                    wPad = Format(Val(wCode), "000000")
                Else
                    wPad = wCode
                End If

                ' --- Warrant name -> col B (fetch when blank; done before the
                ' status tags below so a fresh name doesn't overwrite a tag) ---
                If Trim(CStr(wsR.Cells(i, 2).Value)) = "" Then
                    Application.StatusBar = "Fetching Wrt name " & wPad & "..."
                    Dim wrtName As String
                    wrtName = FetchWarrantName(wPad)
                    If wrtName <> "" Then wsR.Cells(i, 2).Value = wrtName
                End If

                ' --- Warrant own market price ---
                ' NOTE: FinMind was evaluated as a potential primary source here but was
                ' skipped. Every FinMind dataset that includes a warrant closing price
                ' (TaiwanStockInfoWithWarrantSummary, TaiwanStockWarrantTradingDailyReport)
                ' requires a paid Sponsor-tier token, and even those either key by the
                ' underlying mother-stock code (not the warrant's own code) or return
                ' per-broker trade prices instead of an official daily close. The one
                ' Free-tier warrant dataset (TaiwanStockInfoWithWarrant) has no price
                ' field at all. Confirmed 2026-07-01 against https://finmind.github.io/llms-full.txt.
                '
                ' Bug fix (2026-07-15): warrant codes are RECYCLED by issuers after a
                ' warrant expires (the same 6-digit code gets reassigned to a new
                ' warrant later), and Yahoo Finance's ticker DB can keep serving a
                ' stale quote left over from a previous, already-expired warrant that
                ' used the same code -- e.g. code 051812 read back as an old 1.69
                ' instead of the current listing's real 2.71. That quote looks like a
                ' successful fetch (px>0) and silently never updates again. TPEx/TWSE
                ' key off the live exchange record for that code so they're immune to
                ' this; Yahoo is NOT safe as a fallback here and has been removed. If
                ' TPEx+TWSE both miss the code, it's now treated as a real failure
                ' (row marked STALE below) instead of guessing with Yahoo.
                Dim wDteRaw As Variant: wDteRaw = wsR.Cells(i, 5).Value   ' E DTE
                Dim wIsExpired As Boolean
                wIsExpired = IsNumeric(wDteRaw) And CDbl(wDteRaw) <= 0

                If wIsExpired Then
                    ' Past expiry: don't attempt a live price at all (this is
                    ' exactly the scenario Yahoo gets wrong via code recycling).
                    ' LAST PX stays frozen at whatever it last held; row is tagged.
                    Call MarkWarrantStatus(wsR, i, "EXPIRED")
                Else
                    px = 0
                    Application.StatusBar = "Fetching Wrt " & wPad & " (TPEx OTC)..."
                    px = FetchTPExWarrantPrice(wPad)
                    If px <= 0 Then
                        Application.StatusBar = "Fetching Wrt " & wPad & " (TWSE)..."
                        px = FetchTWSEClose(wPad)
                    End If
                    If px > 0 Then
                        wsR.Cells(i, 15).Value = px   ' O LAST PX
                        okN = okN + 1
                        Call ClearWarrantStatus(wsR, i)
                    Else
                        failN = failN + 1
                        Call MarkWarrantStatus(wsR, i, "STALE")
                    End If
                End If
            End If
        Next i

        Application.StatusBar = False

        LogDebug wsD, "FetchAllPrices", "OK", _
            "Updated=" & okN & " Failed=" & failN
        ' Now that UNDL LAST is populated, recalc Greeks
        Call RefreshGreeks
        ' Derive realized PnL from Transactions (silent: aggregate count
        ' included in our own MsgBox / caller's MsgBox)
        Call RealizeClosedTransactions(silent:=True)

        If Not silent Then
            ' Count realized rows for summary
            Dim rlC As Long
            On Error Resume Next
            rlC = ThisWorkbook.Sheets("Realized"). _
                Cells(ThisWorkbook.Sheets("Realized").Rows.Count, 1).End(xlUp).Row - 1
            On Error GoTo ErrHandler
            If rlC < 0 Then rlC = 0

            MsgBox "Quote fetch complete." & vbCrLf & vbCrLf & _
                   "  Updated: " & okN & " cell(s)" & vbCrLf & _
                   "  Failed : " & failN & " cell(s)" & vbCrLf & vbCrLf & _
                   "Greeks refreshed." & vbCrLf & _
                   "Realized rows derived: " & rlC, _
                   vbInformation, "RR5 FetchAllPrices"
        End If
        Exit Sub
ErrHandler:
        Application.StatusBar = False
        LogDebug wsD, "FetchAllPrices", "ERROR", Err.Description
        ' Don't show error msgbox in silent mode -- the caller (e.g.
        ' RebuildPositionsFromTransactions) will report aggregate status
        If Not silent Then
            MsgBox "FetchAllPrices failed: " & Err.Description, vbCritical
        End If
    End Sub

    ' =====================================================================
    ' WARRANT STATUS MARKERS (visual flags used by FetchAllPrices)
    ' ---------------------------------------------------------------------
    ' EXPIRED : DTE<=0 -- price fetch was skipped entirely this run, LAST PX
    '           is frozen at whatever it last held.
    ' STALE   : still active, but TPEx+TWSE both failed this run -- LAST PX
    '           is whatever it held before (not refreshed, not necessarily
    '           wrong).
    ' Tags are appended to the Name cell (col B -- display-only, not read by
    ' any formula or other macro) and the LAST PX cell is highlighted.
    ' ClearWarrantStatus only touches cells this module previously marked
    ' (matched by our own marker colors), so it never clobbers formatting it
    ' didn't set.
    ' Tags are plain ASCII ("[EXPIRED]" / "[STALE]") on purpose -- VBE
    ' imports .bas files using the system ANSI codepage, not UTF-8, so
    ' literal multi-byte Chinese characters written into this file by an
    ' external editor get mangled on import/paste into the VBA project.
    ' =====================================================================
    Private Sub MarkWarrantStatus(wsR As Worksheet, row As Long, status As String)
        Dim cellPx As Range: Set cellPx = wsR.Cells(row, 15)   ' O LAST PX
        Dim tag As String
        Select Case status
            Case "EXPIRED"
                cellPx.Interior.Color = RGB(90, 90, 90)        ' gray
                cellPx.Font.Color = RGB(220, 220, 220)
                tag = " [EXPIRED]"
            Case "STALE"
                cellPx.Interior.Color = RGB(153, 102, 0)       ' amber
                cellPx.Font.Color = RGB(255, 255, 255)
                tag = " [STALE]"
            Case Else
                Exit Sub
        End Select

        Dim nameCell As Range: Set nameCell = wsR.Cells(row, 2)
        Dim baseName As String: baseName = CStr(nameCell.Value)
        baseName = Replace(baseName, " [EXPIRED]", "")
        baseName = Replace(baseName, " [STALE]", "")
        nameCell.Value = baseName & tag
    End Sub

    Private Sub ClearWarrantStatus(wsR As Worksheet, row As Long)
        Dim cellPx As Range: Set cellPx = wsR.Cells(row, 15)   ' O LAST PX
        If cellPx.Interior.Color = RGB(90, 90, 90) Or cellPx.Interior.Color = RGB(153, 102, 0) Then
            cellPx.Interior.ColorIndex = xlColorIndexNone
            cellPx.Font.ColorIndex = xlColorIndexAutomatic
        End If

        Dim nameCell As Range: Set nameCell = wsR.Cells(row, 2)
        Dim baseName As String: baseName = CStr(nameCell.Value)
        Dim cleaned As String
        cleaned = Replace(baseName, " [EXPIRED]", "")
        cleaned = Replace(cleaned, " [STALE]", "")
        If cleaned <> baseName Then nameCell.Value = cleaned
    End Sub

    ' =====================================================================
    ' ONE-SHOT SETUP: Add GAMMA / THETA / VEGA columns to RR5 Warrants
    ' =====================================================================
    '
    ' Run this ONCE (Alt+F8 -> Setup_WarrantsGreeksColumns -> Run).
    ' Idempotent -- safe to re-run.
    '
    ' What it does:
    '   1. Adds headers in row 74:  X=GAMMA  Y=THETA/DAY  Z=VEGA
    '   2. Formats data cells X75:Z104 with 4-decimal precision
    '   3. Adds SUMPRODUCT subtotals in row 105 (cols X/Y/Z)
    '   4. Patches Summary bar (row 1/2) so Net Gamma/Theta/Vega include warrants
    '
    Public Sub Setup_WarrantsGreeksColumns()
        On Error GoTo ErrHandler
        Dim wsR As Worksheet
        Set wsR = ThisWorkbook.Sheets("RR5")

        ' -- 1. Headers in row 74 (cols X, Y, Z) ---------------------------
        Dim hdrSpec As Variant
        hdrSpec = Array(Array(24, "GAMMA"), _
                        Array(25, "THETA/DAY"), _
                        Array(26, "VEGA"))
        Dim h As Variant
        For Each h In hdrSpec
            With wsR.Cells(74, CLng(h(0)))
                .Value = h(1)
                .Font.Name = "Consolas"
                .Font.Size = 9
                .Font.Bold = True
                .Font.Color = RGB(255, 165, 0)     ' orange like other headers
                .Interior.Color = RGB(0, 0, 0)     ' black bg
                .HorizontalAlignment = xlCenter
                .Borders.LineStyle = xlContinuous
                .Borders.Color = RGB(47, 79, 79)
            End With
            wsR.Columns(CLng(h(0))).ColumnWidth = 12
        Next h

        ' -- 2. Format data rows 75-104 (cols X/Y/Z) -----------------------
        With wsR.Range(wsR.Cells(75, 24), wsR.Cells(104, 26))
            .NumberFormat = "0.0000;(0.0000);""-"""
            .Font.Name = "Consolas"
            .Font.Size = 10
            .Font.Color = RGB(255, 215, 0)         ' gold (formula style)
            .Interior.Color = RGB(0, 0, 0)
            .Borders.LineStyle = xlContinuous
            .Borders.Color = RGB(47, 79, 79)
        End With

        ' -- 3. Subtotal row 105 -------------------------------------------
        ' Each Greek aggregated as SUMPRODUCT(SHARES * Greek_per_warrant)
        ' to give portfolio-level impact.
        wsR.Cells(105, 24).Formula = _
            "=SUMPRODUCT((A75:A104<>"""")*M75:M104*X75:X104)"
        wsR.Cells(105, 25).Formula = _
            "=SUMPRODUCT((A75:A104<>"""")*M75:M104*Y75:Y104)"
        wsR.Cells(105, 26).Formula = _
            "=SUMPRODUCT((A75:A104<>"""")*M75:M104*Z75:Z104)"
        With wsR.Range(wsR.Cells(105, 24), wsR.Cells(105, 26))
            .Font.Name = "Consolas"
            .Font.Bold = True
            .Font.Color = RGB(255, 255, 255)
            .Interior.Color = RGB(0, 0, 0)
            .NumberFormat = "#,##0.00;(#,##0.00);""-"""
            .Borders.LineStyle = xlContinuous
            .Borders.Color = RGB(47, 79, 79)
        End With

        ' -- 4. Patch Summary bar to include warrants ----------------------
        ' L1  NET DELTA: original formula was buggy (used IV*ConvRatio for warrants).
        '                Correct: SHARES * Warrant_Delta-per-warrant
        wsR.Range("L1").Formula = _
            "=IFERROR(SUBTOTAL(9,S7:S26)+SUBTOTAL(9,Y31:Y70)" & _
            "+SUMPRODUCT((A75:A104<>"""")*M75:M104*T75:T104),0)"
        ' N1  NET GAMMA   :  options (SUMPRODUCT * SHARES * MULT) + warrants X105
        wsR.Range("N1").Formula = _
            "=IFERROR(SUMPRODUCT((H31:H70<>"""")*U31:U70*L31:L70*M31:M70)+X105,0)"
        ' J2  NET THETA/DAY: options SUBTOTAL(Z31:Z70) + warrants Y105
        wsR.Range("J2").Formula = _
            "=IFERROR(SUBTOTAL(9,Z31:Z70)+Y105,0)"
        ' L2  NET VEGA     : options SUBTOTAL(AA31:AA70) + warrants Z105
        wsR.Range("L2").Formula = _
            "=IFERROR(SUBTOTAL(9,AA31:AA70)+Z105,0)"

        ' -- 5. Done -------------------------------------------------------
        MsgBox "Warrants Greeks columns installed:" & vbCrLf & vbCrLf & _
               "  X = GAMMA      (per warrant)" & vbCrLf & _
               "  Y = THETA/DAY  (per warrant)" & vbCrLf & _
               "  Z = VEGA       (per warrant)" & vbCrLf & vbCrLf & _
               "Row 105: portfolio-level SUMPRODUCT subtotals." & vbCrLf & _
               "Summary bar patched:" & vbCrLf & _
               "  L1 NET DELTA  -- bug fixed (was IV*ConvRatio, now SHARES*Delta)" & vbCrLf & _
               "  N1 NET GAMMA  -- now includes warrants" & vbCrLf & _
               "  J2 NET THETA  -- now includes warrants" & vbCrLf & _
               "  L2 NET VEGA   -- now includes warrants" & vbCrLf & vbCrLf & _
               "Run RefreshGreeks to populate values.", _
               vbInformation, "RR5 Setup"
        Exit Sub
ErrHandler:
        MsgBox "Setup failed: " & Err.Description, vbCritical
    End Sub

    ' =====================================================================
    ' ONE-SHOT SETUP: Add Gearing (leverage) columns to RR5 Warrants
    ' =====================================================================
    '
    ' Run this ONCE (Alt+F8 -> Setup_WarrantGearingColumns -> Run).
    ' Idempotent -- safe to re-run.
    ' Compact layout: warrants header row 30, data rows 31-38.
    '
    ' What it does:
    '   1. Adds headers in row 30:  AA=NOM GEARING  AB=EFF GEARING
    '   2. Formats data cells AA31:AB38 (2-decimal)
    '   3. Values are computed by RefreshGreeks -- run it after this setup.
    '
    Public Sub Setup_WarrantGearingColumns()
        On Error GoTo ErrHandler
        Dim wsR As Worksheet
        Set wsR = ThisWorkbook.Sheets("RR5")

        Dim hdrSpec As Variant
        hdrSpec = Array(Array(27, "NOM GEARING"), _
                        Array(28, "EFF GEARING"))
        Dim h As Variant
        For Each h In hdrSpec
            With wsR.Cells(30, CLng(h(0)))
                .Value = h(1)
                .Font.Name = "Consolas"
                .Font.Size = 9
                .Font.Bold = True
                .Font.Color = RGB(255, 165, 0)     ' orange like other headers
                .Interior.Color = RGB(0, 0, 0)     ' black bg
                .HorizontalAlignment = xlCenter
                .Borders.LineStyle = xlContinuous
                .Borders.Color = RGB(47, 79, 79)
            End With
            wsR.Columns(CLng(h(0))).ColumnWidth = 12
        Next h

        With wsR.Range(wsR.Cells(31, 27), wsR.Cells(38, 28))
            .NumberFormat = "0.00;(0.00);""-"""
            .Font.Name = "Consolas"
            .Font.Size = 10
            .Font.Color = RGB(255, 215, 0)         ' gold (formula style)
            .Interior.Color = RGB(0, 0, 0)
            .Borders.LineStyle = xlContinuous
            .Borders.Color = RGB(47, 79, 79)
        End With

        MsgBox "Warrant Gearing columns installed:" & vbCrLf & vbCrLf & _
               "  AA = NOM GEARING  (UNDL LAST / (LAST PX * CONV RATIO))" & vbCrLf & _
               "  AB = EFF GEARING  (UNDL LAST * DELTA / LAST PX)" & vbCrLf & vbCrLf & _
               "Run RefreshGreeks to populate values.", _
               vbInformation, "RR5 Setup"
        Exit Sub
ErrHandler:
        MsgBox "Setup failed: " & Err.Description, vbCritical
    End Sub

    ' =====================================================================
    ' FINMIND STOCK / WARRANT PRICE (Taiwan EOD close, TWSE + OTC)
    ' ---------------------------------------------------------------------
    ' Uses FinMind TaiwanStockPrice dataset which covers both exchanges with
    ' no suffix needed. More reliable than Yahoo for warrants because:
    '   - Official exchange closing price (not last tick)
    '   - Handles OTC-listed warrants Yahoo misidentifies as TWSE
    '   - Free, no auth required
    ' Retries up to 3 previous business days for post-holiday gaps.
    ' Called as primary source in FetchAllPrices warrant section.
    ' =====================================================================
    Public Function FetchFinMindStockPrice(ticker As String) As Double
        On Error GoTo Fail
        If Trim(ticker) = "" Then Exit Function
        Dim startDt As Date: startDt = PreviousBusinessDay(Date)
        Dim attempt As Long
        For attempt = 1 To 3
            Dim dateStr As String: dateStr = Format(startDt, "yyyy-mm-dd")
            Dim url As String
            url = "https://api.finmindtrade.com/api/v4/data" & _
                  "?dataset=TaiwanStockPrice" & _
                  "&data_id=" & Trim(ticker) & _
                  "&start_date=" & dateStr & _
                  "&end_date=" & dateStr
            Dim req As Object
            Set req = CreateObject("MSXML2.XMLHTTP")
            req.Open "GET", url, False
            req.setRequestHeader "User-Agent", "Mozilla/5.0"
            req.send
            If req.Status = 200 Then
                Dim resp As String: resp = req.responseText
                If InStr(resp, """data"":[]") = 0 And Len(resp) > 20 Then
                    Dim cl As Double: cl = ExtractJsonNumber(resp, """close"":")
                    If cl > 0 Then
                        FetchFinMindStockPrice = cl
                        Exit Function
                    End If
                End If
            End If
            startDt = PreviousBusinessDay(startDt - 1)
        Next attempt
Fail:
        FetchFinMindStockPrice = 0
    End Function

    ' =====================================================================
    ' TWSE STOCK / WARRANT CLOSE  (openapi.twse.com.tw STOCK_DAY_ALL, cached)
    ' -----------------------------------------------------------------
    ' Primary: bulk download from openapi.twse.com.tw - covers all TWSE
    '   listed stocks AND warrants with named JSON fields.
    ' Fallback: rwd/zh/afterTrading/STOCK_DAY for 6-digit warrant codes
    '   not present in STOCK_DAY_ALL (parses the closing-price field from data array).
    ' Cache: one download per calendar day to avoid repeated large fetches.
    ' =====================================================================
    Public Function FetchTWSEClose(stockCode As String) As Double
        On Error GoTo Fail
        If Trim(stockCode) = "" Then Exit Function

        Static twseCache As Object
        If twseCache Is Nothing Then Set twseCache = CreateObject("Scripting.Dictionary")
        Dim today As String: today = Format(Date, "yyyymmdd")

        Dim resp As String
        If Not twseCache.Exists(today) Then
            Dim req As Object
            Set req = CreateObject("MSXML2.XMLHTTP")
            req.Open "GET", "https://openapi.twse.com.tw/v1/exchangeReport/STOCK_DAY_ALL", False
            req.setRequestHeader "User-Agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64)"
            req.setRequestHeader "Accept", "application/json"
            req.send
            If req.Status = 200 Then
                twseCache(today) = req.responseText
            Else
                twseCache(today) = ""
            End If
        End If
        resp = CStr(twseCache(today))

        Dim cl As Double
        If Len(resp) > 10 Then
            Dim codeKey As String: codeKey = """Code"":""" & Trim(stockCode) & """"
            Dim idx As Long: idx = InStr(resp, codeKey)
            If idx > 0 Then
                Dim objEnd As Long: objEnd = InStr(idx, resp, "}")
                If objEnd > idx Then
                    Dim objStr As String: objStr = Mid(resp, idx, objEnd - idx + 1)
                    Dim s As String
                    s = JsonStr(objStr, "ClosingPrice")
                    If s = "" Then s = JsonStr(objStr, "Close")
                    If s <> "" Then
                        s = Replace(s, ",", "")
                        If IsNumeric(s) Then cl = CDbl(s)
                    End If
                    If cl <= 0 Then cl = ExtractJsonNumber(objStr, """ClosingPrice"":")
                    If cl <= 0 Then cl = ExtractJsonNumber(objStr, """Close"":")
                End If
            End If
        End If
        If cl > 0 Then FetchTWSEClose = cl: Exit Function

        ' Fallback: rwd STOCK_DAY endpoint (reliable for 6-digit warrant codes)
        Dim dt As Date: dt = Date
        Dim att As Long
        For att = 1 To 2
            Dim dateStr As String
            dateStr = Format(DateSerial(Year(dt), Month(dt), 1), "yyyymmdd")
            Dim req2 As Object
            Set req2 = CreateObject("MSXML2.XMLHTTP")
            req2.Open "GET", "https://www.twse.com.tw/rwd/zh/afterTrading/STOCK_DAY" & _
                      "?response=json&date=" & dateStr & "&stockNo=" & Trim(stockCode), False
            req2.setRequestHeader "User-Agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64)"
            req2.setRequestHeader "Referer", "https://www.twse.com.tw/"
            req2.send
            If req2.Status = 200 Then
                resp = req2.responseText
                If InStr(resp, """stat"":""OK""") > 0 Then
                    ' data[][6] = closing price; find last inner record before "]]"
                    Dim dStart As Long: dStart = InStr(resp, """data"":[")
                    If dStart > 0 Then
                        dStart = dStart + Len("""data"":[")
                        Dim dEnd As Long: dEnd = InStr(dStart, resp, "]]")
                        If dEnd > dStart Then
                            Dim lastBrk As Long: lastBrk = InStrRev(resp, "[", dEnd)
                            If lastBrk > dStart Then
                                Dim recStr As String
                                recStr = Mid(resp, lastBrk, dEnd - lastBrk + 1)
                                Dim closeF As String
                                closeF = ParseQuotedField(recStr, 6)
                                closeF = Replace(closeF, ",", "")
                                If IsNumeric(closeF) Then
                                    FetchTWSEClose = CDbl(closeF)
                                    Exit Function
                                End If
                            End If
                        End If
                    End If
                End If
            End If
            dt = DateSerial(Year(dt), Month(dt) - 1, 1)
        Next att
Fail:
        FetchTWSEClose = 0
    End Function

    ' =====================================================================
    ' TPEX OTC STOCK CLOSE  (tpex_mainboard_daily_close_quotes, cached)
    ' Returns 0 when stock is TWSE-listed (not in OTC list).
    ' =====================================================================
    Public Function FetchTPExStockPrice(stockCode As String) As Double
        On Error GoTo Fail
        If Trim(stockCode) = "" Then Exit Function

        Static tpexStkCache As Object
        If tpexStkCache Is Nothing Then Set tpexStkCache = CreateObject("Scripting.Dictionary")
        Dim today As String: today = Format(Date, "yyyymmdd")

        Dim resp As String
        If Not tpexStkCache.Exists(today) Then
            Dim req As Object
            Set req = CreateObject("MSXML2.XMLHTTP")
            req.Open "GET", "https://www.tpex.org.tw/openapi/v1/tpex_mainboard_daily_close_quotes", False
            req.setRequestHeader "User-Agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64)"
            req.setRequestHeader "Referer", "https://www.tpex.org.tw/openapi/"
            req.setRequestHeader "Accept", "application/json"
            req.send
            If req.Status = 200 Then
                tpexStkCache(today) = req.responseText
            Else
                tpexStkCache(today) = ""
            End If
        End If
        resp = CStr(tpexStkCache(today))
        If Len(resp) < 10 Then Exit Function

        FetchTPExStockPrice = LookupTPExCode(resp, Trim(stockCode))
        Exit Function
Fail:
        FetchTPExStockPrice = 0
    End Function

    ' =====================================================================
    ' TPEX OTC WARRANT CLOSE  (tpex_warrant_daily_quts, cached)
    ' Returns 0 when warrant is TWSE-listed (not in OTC list).
    ' =====================================================================
    Public Function FetchTPExWarrantPrice(wCode As String) As Double
        On Error GoTo Fail
        If Trim(wCode) = "" Then Exit Function

        Static tpexWrtCache As Object
        If tpexWrtCache Is Nothing Then Set tpexWrtCache = CreateObject("Scripting.Dictionary")
        Dim today As String: today = Format(Date, "yyyymmdd")

        Dim resp As String
        If Not tpexWrtCache.Exists(today) Then
            Dim req As Object
            Set req = CreateObject("MSXML2.XMLHTTP")
            req.Open "GET", "https://www.tpex.org.tw/openapi/v1/tpex_warrant_daily_quts", False
            req.setRequestHeader "User-Agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64)"
            req.setRequestHeader "Referer", "https://www.tpex.org.tw/openapi/"
            req.setRequestHeader "Accept", "application/json"
            req.send
            If req.Status = 200 Then
                tpexWrtCache(today) = req.responseText
            Else
                tpexWrtCache(today) = ""
            End If
        End If
        resp = CStr(tpexWrtCache(today))
        If Len(resp) < 10 Then Exit Function

        FetchTPExWarrantPrice = LookupTPExCode(resp, Trim(wCode))
        Exit Function
Fail:
        FetchTPExWarrantPrice = 0
    End Function

    ' -----------------------------------------------------------------------
    ' LookupTPExCode: shared helper for TPEx JSON bulk responses.
    ' Tries "Code" and "SecuritiesCompanyCode" field names; extracts
    ' ClosingPrice (quoted string or raw number, comma-thousand handled).
    ' -----------------------------------------------------------------------
    Private Function LookupTPExCode(resp As String, code As String) As Double
        On Error Resume Next
        Dim codeKey As String: codeKey = """Code"":""" & code & """"
        Dim idx As Long: idx = InStr(resp, codeKey)
        If idx = 0 Then
            codeKey = """SecuritiesCompanyCode"":""" & code & """"
            idx = InStr(resp, codeKey)
        End If
        If idx = 0 Then Exit Function

        Dim objStart As Long: objStart = InStrRev(resp, "{", idx)
        If objStart = 0 Then objStart = idx
        Dim objEnd As Long: objEnd = InStr(idx, resp, "}")
        If objEnd = 0 Then Exit Function
        Dim obj As String: obj = Mid(resp, objStart, objEnd - objStart + 1)

        Dim s As String
        s = JsonStr(obj, "ClosingPrice")
        If s = "" Then s = JsonStr(obj, "Close")
        If s = "" Then s = JsonStr(obj, "ClosePrice")
        If s <> "" Then
            s = Replace(s, ",", "")
            If IsNumeric(s) Then LookupTPExCode = CDbl(s): Exit Function
        End If
        Dim cl As Double
        cl = ExtractJsonNumber(obj, """ClosingPrice"":")
        If cl <= 0 Then cl = ExtractJsonNumber(obj, """Close"":")
        LookupTPExCode = cl
    End Function

    ' -----------------------------------------------------------------------
    ' ParseQuotedField: extract the nth (0-based) quoted value from a
    ' comma-separated list where values may be quoted strings containing
    ' commas (e.g. "36,000" in TWSE STOCK_DAY data arrays).
    ' -----------------------------------------------------------------------
    Private Function ParseQuotedField(s As String, fieldIdx As Long) As String
        Dim count As Long: count = 0
        Dim p As Long: p = 1
        Do While p <= Len(s)
            If Mid(s, p, 1) = """" Then
                Dim q As Long: q = InStr(p + 1, s, """")
                If q = 0 Then Exit Do
                If count = fieldIdx Then
                    ParseQuotedField = Mid(s, p + 1, q - p - 1)
                    Exit Function
                End If
                count = count + 1
                p = q + 1
            Else
                p = p + 1
            End If
        Loop
    End Function

    ' -----------------------------------------------------------------------
    ' Debug_TPExPrices: verify TPEx API field names and coverage.
    ' Run from Alt+F8.  Prints HTTP status + first matched object for each
    ' endpoint so you can confirm ClosingPrice / Code field names.
    ' -----------------------------------------------------------------------
    Public Sub Debug_TPExPrices(Optional stockCode As String = "2454", _
                                 Optional wCode As String = "")
        Dim req As Object, resp As String, idx As Long
        Debug.Print "=== Debug_TPExPrices ==="

        ' --- 1. OTC stock price ---
        Set req = CreateObject("MSXML2.XMLHTTP")
        req.Open "GET", "https://www.tpex.org.tw/openapi/v1/tpex_mainboard_daily_close_quotes", False
        req.setRequestHeader "User-Agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64)"
        req.setRequestHeader "Accept", "application/json"
        req.send
        Debug.Print "[1] tpex_mainboard_daily_close_quotes  HTTP=" & req.Status & "  len=" & Len(req.responseText)
        resp = req.responseText
        idx = InStr(resp, """Code"":""" & stockCode & """")
        If idx = 0 Then idx = InStr(resp, """SecuritiesCompanyCode"":""" & stockCode & """")
        If idx > 0 Then
            Debug.Print "    " & stockCode & " found: " & Mid(resp, InStrRev(resp, "{", idx), InStr(idx, resp, "}") - InStrRev(resp, "{", idx) + 1)
        Else
            Debug.Print "    " & stockCode & " NOT found in OTC list (may be TWSE-listed)"
        End If
        Debug.Print "    FetchTPExStockPrice(" & stockCode & ") = " & FetchTPExStockPrice(stockCode)

        ' --- 2. OTC warrant price ---
        Set req = CreateObject("MSXML2.XMLHTTP")
        req.Open "GET", "https://www.tpex.org.tw/openapi/v1/tpex_warrant_daily_quts", False
        req.setRequestHeader "User-Agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64)"
        req.setRequestHeader "Accept", "application/json"
        req.send
        Debug.Print "[2] tpex_warrant_daily_quts  HTTP=" & req.Status & "  len=" & Len(req.responseText)
        resp = req.responseText
        ' Show first object to reveal field names
        Dim firstObj As Long: firstObj = InStr(resp, "{")
        Dim firstObjEnd As Long: firstObjEnd = InStr(firstObj, resp, "}")
        If firstObj > 0 And firstObjEnd > firstObj Then
            Debug.Print "    First object: " & Mid(resp, firstObj, firstObjEnd - firstObj + 1)
        End If
        If wCode <> "" Then
            idx = InStr(resp, """Code"":""" & wCode & """")
            If idx > 0 Then
                Debug.Print "    " & wCode & " found: " & Mid(resp, InStrRev(resp, "{", idx), InStr(idx, resp, "}") - InStrRev(resp, "{", idx) + 1)
            Else
                Debug.Print "    " & wCode & " NOT found (likely TWSE-listed)"
            End If
            Debug.Print "    FetchTPExWarrantPrice(" & wCode & ") = " & FetchTPExWarrantPrice(wCode)
        End If

        ' --- 3. TWSE bulk ---
        Set req = CreateObject("MSXML2.XMLHTTP")
        req.Open "GET", "https://openapi.twse.com.tw/v1/exchangeReport/STOCK_DAY_ALL", False
        req.setRequestHeader "User-Agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64)"
        req.setRequestHeader "Accept", "application/json"
        req.send
        Debug.Print "[3] TWSE STOCK_DAY_ALL  HTTP=" & req.Status & "  len=" & Len(req.responseText)
        resp = req.responseText
        firstObj = InStr(resp, "{")
        firstObjEnd = InStr(firstObj, resp, "}")
        If firstObj > 0 And firstObjEnd > firstObj Then
            Debug.Print "    First object: " & Mid(resp, firstObj, firstObjEnd - firstObj + 1)
        End If
        If wCode <> "" Then
            idx = InStr(resp, """Code"":""" & wCode & """")
            If idx > 0 Then
                Debug.Print "    " & wCode & " TWSE found, FetchTWSEClose=" & FetchTWSEClose(wCode)
            Else
                Debug.Print "    " & wCode & " not in STOCK_DAY_ALL -> will use STOCK_DAY rwd fallback"
                Debug.Print "    FetchTWSEClose(" & wCode & ") = " & FetchTWSEClose(wCode)
            End If
        End If
    End Sub

    ' =====================================================================
    ' WARRANT NAME FETCH
    ' ---------------------------------------------------------------------
    ' Primary:  Yahoo Finance chart API -> meta.longName / meta.shortName
    '           Tries .TW (TWSE) then .TWO (OTC/TPEx).
    ' Fallback: FinMind TaiwanStockInfo -> stock_name field
    ' Returns "" on any failure so caller can skip writing.
    ' =====================================================================
    ' =====================================================================
    ' WARRANT NAME FETCH  (3 sources, in order)
    ' 1. Yahoo Finance v7/quote  -> shortName  (best for TW warrants)
    ' 2. FinMind TaiwanStockInfo -> stock_name (needs date param)
    ' 3. TWSE searchSecurities   -> name       (TWSE-listed only)
    ' =====================================================================
    ' =====================================================================
    ' WARRANT NAME FETCH  (2 sources confirmed working)
    '
    ' Method 1: TWSE afterTrading/STOCK_DAY  [TWSE-listed warrants]
    '   URL: /rwd/zh/afterTrading/STOCK_DAY?response=json&date=YYYYMMDD&stockNo=CODE
    '   Title format: "115/06 051812 TSMC-YuantaSecurities-5A-Call04 daily trading info"
    '   Parse: token immediately after the code in the title string.
    '   Tries current month first, then previous month as fallback.
    '
    ' Method 2: TPEx tpex_warrant_issue  [OTC-listed warrants, bulk JSON]
    '   Searches the bulk list for "Code":"wCode" then extracts Name field.
    ' =====================================================================
    Public Function FetchWarrantName(wCode As String) As String
        On Error GoTo Fail
        If Trim(wCode) = "" Then Exit Function

        Dim req As Object, resp As String, nm As String

        ' --- Method 1: TWSE STOCK_DAY (listed warrants) ---
        Dim attempt As Long
        Dim dt As Date: dt = Date
        For attempt = 1 To 2
            Dim dateStr As String
            dateStr = Format(DateSerial(Year(dt), Month(dt), 1), "yyyymmdd")
            Set req = CreateObject("MSXML2.XMLHTTP")
            req.Open "GET", "https://www.twse.com.tw/rwd/zh/afterTrading/STOCK_DAY" & _
                     "?response=json&date=" & dateStr & "&stockNo=" & wCode, False
            req.setRequestHeader "User-Agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64)"
            req.setRequestHeader "Referer", "https://www.twse.com.tw/"
            req.send
            If req.Status = 200 Then
                resp = req.responseText
                If InStr(resp, """stat"":""OK""") > 0 Then
                    nm = JsonStr(resp, "title")
                    If nm <> "" Then
                        Dim cTag As String: cTag = wCode & " "
                        Dim cPos As Long: cPos = InStr(nm, cTag)
                        If cPos > 0 Then
                            Dim afterCode As String
                            afterCode = Mid(nm, cPos + Len(cTag))
                            Dim spPos As Long: spPos = InStr(afterCode, " ")
                            If spPos > 1 Then
                                FetchWarrantName = Trim(Left(afterCode, spPos - 1))
                                Exit Function
                            End If
                        End If
                    End If
                End If
            End If
            ' Retry with previous month
            dt = DateSerial(Year(dt), Month(dt) - 1, 1)
        Next attempt

        ' --- Method 2: TPEx tpex_warrant_issue (OTC-listed warrants, bulk JSON) ---
        Set req = CreateObject("MSXML2.XMLHTTP")
        req.Open "GET", "https://www.tpex.org.tw/openapi/v1/tpex_warrant_issue", False
        req.setRequestHeader "User-Agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64)"
        req.setRequestHeader "Referer", "https://www.tpex.org.tw/openapi/"
        req.setRequestHeader "Accept", "application/json"
        req.send
        If req.Status = 200 Then
            resp = req.responseText
            Dim codeKey As String: codeKey = """Code"":""" & wCode & """"
            Dim codeIdx As Long: codeIdx = InStr(resp, codeKey)
            If codeIdx > 0 Then
                ' Extract Name from the same JSON object (comes after Code in schema)
                Dim objEnd As Long: objEnd = InStr(codeIdx, resp, "}")
                Dim objStr As String
                objStr = Mid(resp, codeIdx, objEnd - codeIdx + 1)
                nm = JsonStr(objStr, "Name")
                If nm <> "" Then
                    FetchWarrantName = nm
                    Exit Function
                End If
            End If
        End If
        Exit Function
Fail:
        FetchWarrantName = ""
    End Function

    ' -----------------------------------------------------------------------
    ' JsonStr: extract string value for a given field name from raw JSON.
    ' Handles both  "key":"value"  and  "key": "value"  (space after colon).
    ' Returns "" if field not found or value is empty.
    ' -----------------------------------------------------------------------
    Private Function JsonStr(json As String, fieldName As String) As String
        Dim needle As String: needle = """" & fieldName & """"
        Dim p As Long: p = InStr(json, needle)
        If p = 0 Then Exit Function
        p = p + Len(needle)
        ' skip : and optional whitespace
        Do While p <= Len(json)
            Dim ch As String: ch = Mid(json, p, 1)
            If ch = ":" Or ch = " " Or ch = Chr(9) Then
                p = p + 1
            Else
                Exit Do
            End If
        Loop
        ' expect opening quote
        If Mid(json, p, 1) <> """" Then Exit Function
        p = p + 1
        ' find closing quote (not escaped)
        Dim q As Long: q = p
        Do While q <= Len(json)
            If Mid(json, q, 1) = """" And (q = 1 Or Mid(json, q - 1, 1) <> "\") Then
                Exit Do
            End If
            q = q + 1
        Loop
        If q > p Then JsonStr = Mid(json, p, q - p)
    End Function

    ' -----------------------------------------------------------------------
    ' Debug_WarrantName: run from Alt+F8 to see exactly what each API returns.
    ' Change wCode to the warrant code you want to test.
    ' -----------------------------------------------------------------------
    ' -----------------------------------------------------------------------
    ' Debug_WarrantName: run from Alt+F8 to verify name fetch for any code.
    ' -----------------------------------------------------------------------
    Public Sub Debug_WarrantName(Optional wCode As String = "051812")
        Dim req As Object, resp As String
        Debug.Print "=== Debug_WarrantName: " & wCode & " ==="

        ' --- Test 1: TWSE STOCK_DAY (primary, confirmed working) ---
        Dim dateStr As String: dateStr = Format(DateSerial(Year(Date), Month(Date), 1), "yyyymmdd")
        Set req = CreateObject("MSXML2.XMLHTTP")
        req.Open "GET", "https://www.twse.com.tw/rwd/zh/afterTrading/STOCK_DAY" & _
                 "?response=json&date=" & dateStr & "&stockNo=" & wCode, False
        req.setRequestHeader "User-Agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64)"
        req.setRequestHeader "Referer", "https://www.twse.com.tw/"
        req.send
        Debug.Print "[1] TWSE STOCK_DAY HTTP=" & req.Status & "  date=" & dateStr
        resp = req.responseText
        Dim nm As String: nm = JsonStr(resp, "title")
        Debug.Print "    title -> [" & nm & "]"
        Debug.Print "    FetchWarrantName result -> [" & FetchWarrantName(wCode) & "]"

        ' --- Test 2: TPEx tpex_warrant_issue (OTC warrants, bulk search) ---
        Set req = CreateObject("MSXML2.XMLHTTP")
        req.Open "GET", "https://www.tpex.org.tw/openapi/v1/tpex_warrant_issue", False
        req.setRequestHeader "User-Agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64)"
        req.setRequestHeader "Referer", "https://www.tpex.org.tw/openapi/"
        req.setRequestHeader "Accept", "application/json"
        req.send
        Debug.Print "[2] TPEx tpex_warrant_issue HTTP=" & req.Status & "  len=" & Len(req.responseText)
        Dim codeKey As String: codeKey = """Code"":""" & wCode & """"
        Dim codeIdx As Long: codeIdx = InStr(req.responseText, codeKey)
        If codeIdx > 0 Then
            Debug.Print "    FOUND at " & codeIdx & ": " & Mid(req.responseText, codeIdx, 150)
        Else
            Debug.Print "    Code " & wCode & " NOT found (likely TWSE-listed, not OTC)"
        End If
    End Sub
