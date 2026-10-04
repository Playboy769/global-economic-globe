# CONVERGENCE — collapse modSEC / modMOPS onto modSECData / modMOPSData

**Status: NOT APPLIED.** The standalone `SEC-Filing-Fetcher` workbook still carries its own
inline copies of the fetch/parse helpers *and* the new `modSECData` / `modMOPSData` sit
alongside them. That duplicated state is **safe and fully functional** — there is no name
clash (every shared name in the *Data modules is `Private`; the two public readers are renamed
`XbrlValue` / `XbrlFyFp`). Converging removes the duplication so there is one source of truth.

Do this **only after** a COM smoke test confirms `modSECData.GetXbrlFinancials` and
`modMOPSData.GetTwFinancials` / `GetTwMonthlyRevenue` return correct data (see `README.md`),
and re-run the standalone workbook's `FetchSECFilings` / `FetchMOPSFilings` afterward to
confirm no regression. Everything here is version-controlled — `git checkout` the two `.bas`
to revert.

---

## 1. `SEC-Filing-Fetcher/modSEC.bas`

### 1a. Delete these Private helpers (now live in `modSECData.bas`)
`RegexEscape`, `GetFilings`, `ScanRecentBlock`, `BuildFilingDict`, `BuildConceptMap`,
`MergeConceptMaps`, `MergeConceptMapsFallback`.

**Keep** `LookupConceptValue` and `LookupFyFp` as-is (still `Public`) — `modCharts` calls
them unqualified from ~30 sites.

### 1b. `ResolveCIK` is also called by `CheckForNewFilings` (not just `FetchSECFilings`)
Two options:
- **A (smaller diff):** keep `ResolveCIK` + `RegexEscape` in `modSEC.bas`, don't delete them
  in 1a. Only the concept-map building collapses (1c). `CheckForNewFilings` untouched.
- **B (fuller):** add `Public Function XbrlResolveCik(inputText, ByRef cik, ByRef title,
  ByRef ticker) As Boolean` to `modSECData` (thin wrapper over its Private `ResolveCIK`),
  point `CheckForNewFilings` at `modSECData.XbrlResolveCik`, then delete `modSEC`'s
  `ResolveCIK` + `RegexEscape`.

Recommend **A** first pass.

### 1c. Rewire `FetchSECFilings` (modSEC.bas ~lines 20–142)
Replace the block from `ResolveCIK(...)` through the last `Set mapAccountsPayable = ...` with:

```vba
Dim x As Object
Set x = modSECData.GetXbrlFinancials(inputText, wantK, wantQ)
If Not x("ok") Then
    SetStatus wsIn, x("error")
    GoTo CleanExit
End If

Dim cik As String, ticker As String, entityName As String
cik = x("cik") : ticker = x("ticker") : entityName = x("entityName")

Dim filings10K As Collection, filings10Q As Collection
Set filings10K = x("filings10K")
Set filings10Q = x("filings10Q")

' still need 8-K / Form 4 for the "other filings" block -- fetch those separately
Dim filings8K As Collection, filingsForm4 As Collection, dummyK As Collection, dummyQ As Collection
Call GetFilings(cik, 0, 0, want8K, want4, dummyK, dummyQ, filings8K, filingsForm4)  ' <-- see note

Dim maps As Object : Set maps = x("maps")
Dim mapRevenue As Object : Set mapRevenue = maps("Revenue")
Dim mapNetIncome As Object : Set mapNetIncome = maps("NetIncome")
' ... one line per map key (26 total) ...
Dim mapAccountsPayable As Object : Set mapAccountsPayable = maps("AccountsPayable")

Dim allMaps As Variant : allMaps = x("allMaps")
```

**Note on 8-K/Form 4:** `GetXbrlFinancials` calls `GetFilings` with `want8K=0, want4=0`, so it
does *not* return 8-K/Form 4. The standalone Filings sheet still shows them. Either:
- keep `modSEC`'s own `GetFilings` (option A above) and call it for `0,0,want8K,want4`, or
- add `res("filings8K")` / `res("filingsForm4")` to `GetXbrlFinancials` (call `GetFilings`
  with the real `want8K`/`want4`) and read them from `x(...)`. Cleaner — do this and drop
  the separate call.

Everything after line 142 (`WriteFilingRow` loop, `ApplyDarkTheme`, `BuildDashboard` etc.)
is unchanged — it consumes the `map*` locals, which now come from `x("maps")`.

### 1d. Verify
`Debug > Compile VBAProject`, then COM: set `Input!A1 = "NVDA"`, let `Worksheet_Change` run,
diff the `Filings` sheet + `Dashboard` charts against a capture taken before this change.

---

## 2. `SEC-Filing-Fetcher/modMOPS.bas`

### 2a. Delete these Private helpers (now in `modMOPSData.bas`)
`MetricKeys`, `FetchQuarterFacts`, `MopsReportUrl`, `QuarterOf`, `QuarterBounds`,
`ParseIxFacts`, `ExtractAttr`, `ExtractIxNonNumeric`, `YMToDate`, `CtxInstant`,
`CtxDuration`, `GetFact`, `ExtractTWMetrics`, `SumIfNumeric`, `SafeDivVariant`, `NewMapSet`,
`AccumulateMetricsIntoMaps`, `BuildTWFilingDict`, `BuildTWRowSpec`, `FetchOneMonthRevenue`,
`HttpGetBig5`, plus `Const MOPS_USER_AGENT`.

**Keep:** `SafeNum`, `SafeSub`, `SafeDaysRatio` (used by `WriteTWFilingRow` / watchlist, not
moved), `WriteTWFilingRow`, `WriteTWHeaders`, `ClearTWFilingsSheet`, `BlankOutDividendRow`,
`FetchTWPriceHistory`, `TWLatestPrice`, `RefreshTWWatchlistRow`, `ClearTWWatchlistRow`,
`SetTWStatus`, `GetTWSetting`.

> `RefreshTWWatchlistRow` also uses `FetchQuarterFacts` / `ExtractTWMetrics` etc. If you
> delete those from `modMOPS`, either (A) keep them, or (B) add public wrappers on
> `modMOPSData` and repoint the watchlist. Recommend keeping them for the first pass — only
> collapse `FetchMOPSFilings`'s quarter-walk + `BuildMonthlyRevenue`'s fetch loop.

### 2b. Rewire `FetchMOPSFilings` (modMOPS.bas ~lines 68–147)
Replace the quarter-decrementing loop with:

```vba
Dim x As Object
Set x = modMOPSData.GetTwFinancials(coId, wantQ)
If Not x("ok") Then
    SetTWStatus wsIn, x("error")
    GoTo CleanExit
End If

Dim entityName As String : entityName = x("entityName")
Dim filingsAnnual As Collection : Set filingsAnnual = x("filingsAnnual")
Dim filingsQuarterly As Collection : Set filingsQuarterly = x("filingsQuarterly")
Dim maps As Object : Set maps = x("maps")

' annualRows / quarterlyRows: WriteTWFilingRow takes a spec dict, and
' x("annualRows") / x("quarterlyRows") ARE those specs (BuildTWRowSpec output).
Dim annualRows As Collection : Set annualRows = x("annualRows")
Dim quarterlyRows As Collection : Set quarterlyRows = x("quarterlyRows")
```

The `WriteTWFilingRow` passes, the `modCharts.BuildDashboard` calls (which read `maps("...")`)
and `BuildMonthlyRevenue` call are unchanged.

### 2c. Rewire `BuildMonthlyRevenue` (modMOPS.bas ~lines 957–1075)
Replace the fetch loop + insertion sort + recompute (everything before the `wsRaw`/chart
block at ~line 1077) with:

```vba
Dim mr As Object
Set mr = modMOPSData.GetTwMonthlyRevenue(coId, months)
If Not mr("ok") Then Exit Sub
Dim labelArr As Variant : labelArr = mr("labels")
Dim revArr As Variant   : revArr   = mr("rev")
Dim yoyArr As Variant   : yoyArr   = mr("yoy")
Dim n As Long : n = UBound(labelArr) - LBound(labelArr) + 1
```

Keep everything from `Set wsRaw = modCharts.GetOrCreateSheet(...)` onward as-is.

### 2d. Verify
`Debug > Compile`, COM: set `Input!B9 = "2395"` (a real listed co_id), let it run, diff
`TW_Filings` + `TW_Dashboard` against a pre-change capture.

---

## 3. `SEC-Filing-Fetcher/build.ps1` — read leaves from shared-vba (optional, tidiest)

Change the module loop (~line 159) so `modHttp` / `modJsonUtil` / `modPrices` are read from
`..\shared-vba\` and add `modSECData` / `modMOPSData` to the list; then delete
`SEC-Filing-Fetcher/modHttp.bas` etc. so there is exactly one copy of each.

```powershell
$sharedRoot = Split-Path -Parent $base   # repo root
foreach ($modName in @("modJsonUtil","modHttp","modPrices","modSECData","modMOPSData",
                       "modTheme","modSEC","modCharts","modMOPS")) {
    $src = if ($modName -in @("modHttp","modJsonUtil","modPrices","modSECData","modMOPSData")) {
        Join-Path $sharedRoot "shared-vba\$modName.bas"
    } else {
        "$base\$modName.bas"
    }
    ...
}
```
