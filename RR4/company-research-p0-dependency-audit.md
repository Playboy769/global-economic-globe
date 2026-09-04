# P0 探勘：modSEC / modMOPS 相依盤點

> 對應 `RR4/company-research-sec-merge-plan.md` 的 P0。純靜態盤點（讀 .bas，未開 .xlsm、未改任何程式）。
> 行號以 `SEC-Filing-Fetcher/` 下各 .bas 現況為準（modSEC 1350 行、modMOPS 1245 行）。

---

## TL;DR — 三個會改到計畫的發現

1. **modSEC / modMOPS 不能原樣搬進 RR4。** 兩者都在 `FetchSECFilings` / `FetchMOPSFilings`
   內硬呼叫 `modCharts.*`（BuildDashboard / BuildQuarterlyDashboard / ExportTickerSnapshot
   …）與 `modTheme.ApplyDarkTheme`，而且 **`modCharts` 反過來呼叫 `modSEC.LookupConceptValue`
   / `LookupFyFp`** —— modSEC ⇄ modCharts 是既有的雙向相依。把 modSEC.bas import 進 RR4
   而不帶 modCharts + modTheme，會在編譯期就出現「Sub or Function not defined」。
   → 計畫 §3.2 的「headless 接縫」不是可選項，是**必要條件**：RR4 端只能拿到
   **抽乾淨的取數函式**（新 `modSECData` / `modMOPSData` 模組，或 trimmed 版），
   不是完整的 `modSEC.bas` / `modMOPS.bas`。

2. **`Sheets("Input")` 相依比想像淺，可完全參數化切掉。**
   modSEC 只從 `Input!A1`（US ticker）、`Input!B4:B7`（wantK/wantQ/want8K/want4）讀；
   modMOPS 只從 `Input!B9`（台股代號）、`Input!B10`（wantQ）讀；兩者用 `Input!B1` 當狀態列。
   全部可改成函式參數（ticker, wantAnnual, wantQuarter, months）+ 由呼叫端自己顯示狀態。

3. **硬撞名只有一個：`ApplyDarkTheme`。** `modTheme.ApplyDarkTheme(ws, range, [hdr])`（3 參）
   vs RR4 `CommodityTermStructure.ApplyDarkTheme()`（0 參，`CommodityTermStructure.bas:416`）
   —— 兩個都是 Public，同專案內會「Ambiguous name detected」。其餘 SEC 子集的 Public 名稱
   與 RR4 現有 Public 名稱**零衝突**（`HttpGet` 例外：RR4 兩處都是 Private，不撞但並存 3 份）。
   Type / Public Const / 模組名 全部無衝突。

---

## 1. SEC-Filing-Fetcher 模組相依 DAG（內部）

```
modHttp       leaf   ── WinHttp；宣告 Public Const SEC_USER_AGENT（modHttp.bas:6）
modJsonUtil   leaf   ── 純字串解析
modTheme      leaf   ── Excel 格式化 / 軸刻度數學（NiceStep, FitValueAxis）
modPrices     ───────> modHttp（HttpGet）, modJsonUtil（ExtractJsonValueRaw, SplitJsonArrayElements）
modCharts     ───────> modTheme, modJsonUtil, **modSEC（LookupConceptValue / LookupFyFp）**   ← 迴圈
modSEC        ───────> modHttp, modJsonUtil, modPrices, modTheme, **modCharts**                ← 迴圈
modMOPS       ───────> modPrices, modTheme, modCharts；自帶 HttpGetBig5（Private, modMOPS.bas:1201）
```

> modSEC ⇄ modCharts 迴圈：`FetchSECFilings` 呼 `modCharts.BuildDashboard`；`BuildDashboard`
> 內部逐 filing 呼 `modSEC.LookupConceptValue`（modCharts.bas:495,504,746,764–829,864–868…）。
> 這個迴圈是既有設計，不是 bug；重點是它讓「只搬 modSEC 不搬 modCharts」不可行。

---

## 2. modSEC.bas — 外部相依逐項

### 2.1 工作表 / 活頁簿參照
| 行 | 參照 | 用途 | headless 化 |
|---|---|---|---|
| 7, 274 | `ThisWorkbook.Sheets("Input")` | 讀 A1 / B4–B7 / 寫 B1 狀態 | 改參數 + 呼叫端顯示狀態 |
| 8, 313 | `ThisWorkbook.Sheets("Filings")` | 寫申報清單表 | RR4 不要這張表；`WriteFilingRow` 迴圈留在獨立活頁簿 |
| 13 | `wsIn.Range("A1")` | US ticker 輸入 | → `ByVal tickerOrCik As String` |
| 27–30 | `GetSetting(wsIn,"B4".."B7", …)` | wantK=7 / wantQ=10 / want8K=10 / want4=10 | → 參數；預設沿用 |
| 260, 1052 | `wsIn.Range("B1")` | 狀態 / 錯誤上一步 | 呼叫端負責 |
| 235, 241, 246 | `ThisWorkbook` 傳進 `BuildDashboard` / `BuildQuarterlyDashboard` / `ExportTickerSnapshot` | 畫圖 + 匯出 xlsx | **RR4 不帶；headless 函式不含這段** |
| 219 | `ApplyDarkTheme(wsOut, wsOut.UsedRange, 1)` | Filings 表上色 | 同上，留在獨立活頁簿 |

### 2.2 跨模組呼叫（全部 unqualified）
| 目標模組 | 被呼叫的 proc | 出現行（節選） |
|---|---|---|
| modHttp | `HttpGet` | 43, 295, 410, 476, 499, 952, 1109 |
| modJsonUtil | `ExtractJsonString` / `ExtractJsonValueRaw` / `SplitJsonArrayElements` / `JsonUnescape` | 46, 50–52, 297–306, 326–327, 479–524, 532–559, 572–586, 683–745, 1109–1117, 1316–1327 |
| modPrices | `FetchDailyPrices` | 234 |
| modPrices | `LatestPrice` | 1192（在 `RefreshWatchlistRow`） |
| modTheme | `ApplyDarkTheme` | 219 |
| modCharts | `BuildDashboard` / `BuildQuarterlyDashboard` / `ExportTickerSnapshot` | 235, 241, 246 |

### 2.3 無以下相依
UserForm（`frm*`）、ActiveX / `OLEObjects`、`Application.Run` / `CallByName`、named range、
`ActiveSheet` / `ActiveCell`、模組層級 Public 變數。全無。

### 2.4 外部 HTTP 端點
- `https://data.sec.gov/api/xbrl/companyfacts/CIK<cik>.json`（43, 1109）— XBRL 全量
- `https://data.sec.gov/submissions/CIK<cik>.json`（295, 476）— 申報索引
- `https://data.sec.gov/submissions/<older-file>.json`（499）— 分頁的舊申報
- `https://www.sec.gov/files/company_tickers.json`（410）— ticker→CIK 對照
- `https://www.sec.gov/Archives/edgar/data/…`（924–971）— 文件 / index / FilingSummary URL 組字串
- UA 一律 `modHttp` 的 `SEC_USER_AGENT`（modHttp.bas:6）

### 2.5 concept-map 建置（headless 化的核心資產，`FetchSECFilings` 內）
- 一次 `HttpGet` companyfacts（43）→ `factsRaw` / `usgaapRaw` / `deiRaw`（50–52）
- 25+ 個 `BuildConceptMap(usgaapRaw, Array(...tag候選...), unitKey)`（64–142），
  含三處 `MergeConceptMaps` / `MergeConceptMapsFallback`（多類股 shares、debt 合併總額 fallback）
- 申報清單 `GetFilings(cik, wantK, wantQ, want8K, want4, …)`（34）→ 4 個 Collection
- 之後 `WriteFilingRow`（163–172）逐 filing 用 `LookupConceptValue` / `LookupFyFp` 取值 + 算衍生指標
  （DSO、利息保障、CCC 各腳…）— 這段約 786–895 行，是「每期每指標」的推導邏輯

---

## 3. modMOPS.bas — 外部相依逐項

### 3.1 工作表 / 活頁簿參照
| 行 | 參照 | 用途 | headless 化 |
|---|---|---|---|
| 55 | `ThisWorkbook.Sheets("Input")` | 讀 B9 / B10、寫 B1 | 改參數 |
| 56 | `modCharts.GetOrCreateSheet(ThisWorkbook,"TW_Filings")` | 建/取 TW 申報表 | 留在獨立活頁簿 |
| 61 | `wsIn.Range("B9")` | 台股代號 | → `ByVal coId As String` |
| 69 | `GetTWSetting(wsIn,"B10",10)` | wantQ | → 參數 |
| 243, 1231 | `wsIn.Range("B1")` | 狀態 | 呼叫端 |
| 202,216,232,254,257,1078,1089 | `ThisWorkbook` / `wb.Sheets(...)` 進 modCharts | TW_Dashboard / TW_RawData / TW_QuarterlySnapshot 畫圖 | **RR4 不帶** |
| 229 | `BuildMonthlyRevenue(wsIn, coId, entityName, wantQ*3+6)` | 月營收（抓 + 寫表 + 畫圖） | 需拆 fetch / render，見 §5 |

### 3.2 跨模組呼叫（qualified）
| 目標 | proc | 行 |
|---|---|---|
| modCharts | `GetOrCreateSheet` | 56, 1078, 1089 |
| modCharts | `BuildDashboard` / `BuildQuarterlyDashboard` / `ExportTickerSnapshot` | 202, 216, 232 |
| modCharts | `WriteRawBlock` / `FinishRawDataSheet` / `ClearAutoSeries` / `ArrayMinMax` | 1085–1140 |
| modTheme | `ApplyDarkTheme` / `ApplyDarkThemeToChart` / `FitValueAxis` | 185, 1129, 1132–1133 |
| modPrices | `FetchDailyPrices` / `LatestPrice` | 767–769, 929–930 |

### 3.3 自帶 HTTP
`HttpGetBig5`（Private, 1201）— 月營收頁是 Big5 編碼，不能用 modHttp.HttpGet。
t164sb01 合併財報頁的抓取在 `FetchQuarterFacts`（285）— **P1 要確認它走 modHttp 還是自帶**
（grep 未見 modMOPS 直接呼叫 `HttpGet(`；可能全部走 `HttpGetBig5`）。

### 3.4 無 UserForm / ActiveX / named range / Application.Run。

### 3.5 外部 HTTP 端點
- `https://mopsov.twse.com.tw/server-java/t164sb01?...REPORT_ID=C`（308，經 `MopsReportUrl`）— ix-XBRL 合併財報
- `https://mopsov.twse.com.tw/nas/t21/{sii|otc}/t21sc03_<民國年>_<月>.html`（1155）— 月營收
- Yahoo chart endpoint（`FetchTWPriceHistory` 762，`<coId>.TW` / `.TWO`，經 modPrices）

### 3.6 取數核心（headless 化資產）
- 季度 walk 迴圈 96–141：`FetchQuarterFacts` → `ExtractTWMetrics(facts, qStart, qEnd, qEnd)` →
  `AccumulateMetricsIntoMaps(maps, …)` + `BuildTWFilingDict` + `BuildTWRowSpec`
- `maps` = `NewMapSet()`（571）：25 個指標 key（`MetricKeys()` 42–46：Revenue/GrossProfit/
  OperatingIncome/NetIncome/Eps/Assets/Liabilities/Cfo/Rd/Sga/Inventory/Ar/CurrentAssets/
  CurrentLiabilities/LongTermDebt/ShortTermDebt/StockholdersEquity/CapEx/Cash/Da/Cogs/
  InterestExpense/EffectiveTaxRate/Shares/AccountsPayable）
- **`quarterlyRows` / `annualRows`（`BuildTWRowSpec` 產出）本身就是每期算好的指標 dict**
  → RR4 renderer 直接讀，不必自己重算
- 月營收：`FetchOneMonthRevenue(coId, rocYear, mo, "sii"→"otc", thisRev, cumRev)`（1151，ByRef 回傳）
  的迴圈（957–1000+），含 CLAUDE.md 要求的「單月 = 本月累計 − 上月累計」交叉驗證

---

## 4. 五個 helper 模組 — RR4 需要哪些

| 模組 | RR4 是否需要 | 理由 |
|---|---|---|
| `modHttp` | **是** | modSEC / modPrices 的 HTTP 全靠它；含 `SEC_USER_AGENT` |
| `modJsonUtil` | **是** | modSEC / modPrices 的 JSON 解析 |
| `modPrices` | **是** | modSEC（估值比率用 `LatestPrice`）、modMOPS（`FetchDailyPrices` 台股價） |
| `modTheme` | **視情況** | 只有「畫表上色」需要（`ApplyDarkTheme`）；若 RR4 renderer 自己上色可不帶。`FitValueAxis` / `ApplyDarkThemeToChart` 只有畫圖用到 → 本版不需要 |
| `modCharts` | **否**（本版） | 圖表；計畫明確排除。但見 §2.2 / §3.2：完整 modSEC/modMOPS 會硬呼叫它 → 只能帶「抽乾淨的取數函式」 |

---

## 5. Headless 接縫 — 建議的抽取邊界

> 目標：RR4 import 的東西 **零 `modCharts` 相依、零 `Sheets("Input"/"Filings")` 相依**。
> 作法：新增 `modSECData.bas` / `modMOPSData.bas`（放 `shared-vba/`），把下列邏輯搬過去、
> 參數化；原 `modSEC.bas` / `modMOPS.bas` 保留在獨立活頁簿，改成呼叫這些新函式再自己寫表畫圖。

### 5.1 `modSECData.GetXbrlFinancials(tickerOrCik, wantAnnual, wantQuarter) As Object`
搬 `FetchSECFilings` 的 **行 20–142**（`ResolveCIK` → `GetFilings` → companyfacts `HttpGet`
→ 全部 `BuildConceptMap` / `Merge*`），回傳 Dictionary：
```
{ "cik","ticker","entityName",
  "filings10K": Collection, "filings10Q": Collection,   ' 每筆含 accn/filingDate/reportDate/form
  "maps": { "Revenue":Obj, "GrossProfit":Obj, ... } }   ' 直接沿用現有 map 物件
```
RR4 renderer：對 `filings10Q`（新到舊）逐筆、對每個要顯示的指標呼 `LookupConceptValue(maps(k),
f("accn"), f("reportDate"), f("form"))` + `LookupFyFp(...)` 取期別標籤。
衍生指標（DSO / 利息保障 / 毛利率…）二選一：
- (a) RR4 renderer 自己用原始 map 算（邏輯簡單者）；
- (b) 另把 `WriteFilingRow` 的每期推導（modSEC.bas 786–895）抽成
  `Public Function ComputeUsFilingMetrics(f, maps...) As Object` 共用。**建議 (b)**，避免兩套算法漂移。

### 5.2 `modMOPSData.GetTwFinancials(coId, quarters) As Object`
搬 `FetchMOPSFilings` 的 **行 75–147**（季度 walk 迴圈 + `NewMapSet`），回傳：
```
{ "entityName",
  "filingsQuarterly":Collection, "filingsAnnual":Collection,
  "quarterlyRows":Collection, "annualRows":Collection,   ' BuildTWRowSpec 產出，已算好
  "maps": {...25 keys...} }
```
RR4 renderer 直接把 `quarterlyRows` / `annualRows` 攤成表。

### 5.3 `modMOPSData.GetTwMonthlyRevenue(coId, months) As Object`
搬 `BuildMonthlyRevenue` 的抓取迴圈（modMOPS.bas 957–~1000，到 `results` dict 填完為止），
**不含** 1078 之後的 RawData / 畫圖。回傳 `{ "labels":Array, "rev":Array, "yoy":Array }`
（含累計反算交叉驗證）。RR4 renderer 自己寫進 Company research 下半區。

---

## 6. Private → Public 升級清單

> 在 `modSECData` / `modMOPSData` 方案下，這些 helper 若被搬進新模組就**不必**升級
> （同模組內可呼叫）。只有「留在原 modSEC/modMOPS、但新模組要呼叫」的才需升級。
> 下表按「新模組直接搬走」為前提列出**必須一起搬**的 Private helper：

### modSEC → modSECData 要一起搬的 Private
`ResolveCIK`(408) · `GetFilings`(457) + `ScanRecentBlock`(506) + `BuildFilingDict`(550) ·
`BuildConceptMap`(565) · `MergeConceptMaps`(612) · `MergeConceptMapsFallback`(643) ·
`GetSetting`(394，或直接改參數丟掉) · （若採 §5.1(b)）`WriteFilingRow` 的推導段 + 其用到的
`SafeNumber`(995)/`SafeRatio`(1007)/`SafeSubtract`(1017)/`SafeDaysRatio`(1283)/
`PriorYearGrowth`(1299)/`SafeMultiply`(1344)
已是 Public、直接用：`LookupConceptValue`(667) · `LookupFyFp`(717)

### modMOPS → modMOPSData 要一起搬的 Private
`FetchQuarterFacts`(285) + `MopsReportUrl`(307) · `QuarterOf`(311) · `QuarterBounds`(316) ·
`ParseIxFacts`(333) + `ExtractAttr`(382) + `ExtractIxNonNumeric`(397) · `YMToDate`(418) ·
`CtxInstant`(422) · `CtxDuration`(426) · `GetFact`(433) · `ExtractTWMetrics`(453) ·
`SumIfNumeric`(528)/`SafeDivVariant`(543)/`SafeNum`(553)/`SafeSub`(561) · `NewMapSet`(571) ·
`AccumulateMetricsIntoMaps`(588) · `BuildTWFilingDict`(629) · `BuildTWRowSpec`(672) ·
`MetricKeys`(42) · `FetchTWPriceHistory`(762) · `SafeDaysRatio`(935) ·
（月營收）`FetchOneMonthRevenue`(1151) + `HttpGetBig5`(1201) + `GetTWSetting`(1235，或丟掉)

> ⚠️ 若不搬、改成「原模組 helper 升 Public 給新模組呼叫」，升級時**一律加前綴**
> （`SEC_` / `MOPS_`），否則 `SafeNum` / `SafeSub` / `SafeDaysRatio` 這種通用名很可能
> 跟獨立活頁簿其他模組或未來的 RR4 撞名。**建議直接搬，不要跨模組升 Public。**

---

## 7. RR4 撞名矩陣（SEC 子集 Public 名稱 vs RR4 現有）

| SEC 子集 Public 名稱 | RR4 現有同名 | 判定 | 處置 |
|---|---|---|---|
| `ApplyDarkTheme`（modTheme, 3 參） | `CommodityTermStructure.ApplyDarkTheme()`（0 參, Public, `CommodityTermStructure.bas:416`）| **硬撞名** | 本版不帶 modTheme 就沒事；若要帶，改 CTS 那顆為 `ApplyCTSDarkTheme`（RR4-only，影響 `CommodityTermStructure.bas` 內部呼叫點）|
| `HttpGet`（modHttp, Public） | `TW_Coverage_Parser.HttpGet`（Private, :847）、`TaiwanPriceFetcher.HttpGet`（Private, :171）| 不撞（RR4 兩處皆 module-scoped）| 不動；備註專案內並存 3 份 HTTP 實作，列入計畫 §10「後續去重」 |
| `SEC_USER_AGENT`（modHttp, Public Const）| 無 | 不撞 | — |
| `LookupConceptValue` / `LookupFyFp` / `FetchSECFilings` / `CheckForNewFilings` / `RefreshWatchlistRow` / `ClearWatchlistRow` | 無 | 不撞 | — |
| `FetchMOPSFilings` / `RefreshTWWatchlistRow` / `ClearTWWatchlistRow` | 無 | 不撞 | — |
| `FetchDailyPrices` / `LatestPrice` / `NearestPriceOnOrBefore`（modPrices）| 無（RR4 用 `GetStockPrice` / `GetTWStockPrice` / `GetHistoryPrices` / `FetchPriceArray`）| 不撞 | — |
| `ExtractJsonValueRaw` / `ExtractJsonString` / `JsonUnescape` / `SplitJsonArrayElements` / `FindKeyValueStart`（modJsonUtil）| 無（RR4 用 `JsonExtract` / `JsonExtractTrend` / `ExtractJsonArray`(Priv) / `ExtractNum`(Priv)）| 不撞 | — |
| `OrangePalette` / `ApplyGridBorder` / `ApplyDarkThemeToChart` / `FitValueAxis`（modTheme）| 無 | 不撞 | 本版不帶 |
| `ArrayMinMax` / `GetOrCreateSheet` / `WriteRawBlock` / `FinishRawDataSheet` / `ClearAutoSeries` / `BuildDashboard` / `BuildQuarterlyDashboard` / `ExportTickerSnapshot`（modCharts）| 無 | 不撞 | 本版不帶 |
| 新增 `GetXbrlFinancials` / `GetTwFinancials` / `GetTwMonthlyRevenue` / `ComputeUsFilingMetrics` | 無 | 不撞 | — |

**Type / Enum**：SEC 子集 5 個模組**無任何** `Public Type` / `Enum`（grep 確認）。
RR4 唯一的 `Public Type` 是 `Attach.TechIndicators`（`Attach.bas:9`）→ 不撞。

**Public Const**：SEC 子集只有 `modHttp.SEC_USER_AGENT`。RR4 的 Const 全是 `Private Const`（module-scoped）→ 不撞。

**模組名**：SEC 子集 = modHttp / modJsonUtil / modPrices / modTheme / modCharts / modSEC / modMOPS
（＋新 modSECData / modMOPSData）。RR4 = Attach / CommodityTermStructure / CorelationMatrix /
DrawDownShadow / PortfolioDashboard_v3 / Sanner / SharedFunctions_Price / TW_Coverage_Parser /
TaiwanPriceFetcher / TickerInsight / modvolatility。**零重疊。**

---

## 8. P1 要開 .xlsm / 實測才能確認的殘留項

1. `Input` 工作表的設定格實際內容與格式（B4–B7 / B10 是否真有值、B1 是否被其他碼讀）——
   靜態看不出來，開 `SECFilingFetcher (17).xlsm` 確認。
2. modMOPS 的 t164sb01 抓取（`FetchQuarterFacts`）到底走 `HttpGetBig5` 還是別的 —— 讀 285–333 行確認。
3. `modCharts` / `modTheme` 是否有模組層級 Public 變數（非 Const）——本次 grep 只掃 Const/Type/Enum；
   `Public <var> As <type>` 未逐一確認（低風險）。
4. RR4 全域變數 `g_PriceCache` / `g_CacheTime`（`PortfolioDashboard_v3.bas:31`）宣告在哪個模組、
   是否 Public —— 確認 SEC 子集沒有同名模組層級變數（低風險，grep 未見）。
5. `TechIndicators` UDT 完整欄位（`Attach.bas:9` 起）—— RR4 renderer 若要跟掃描區共用結構時才需要。

---

## 9. 對 `company-research-sec-merge-plan.md` 的修訂建議

- **§3.1 / §4.1**：`shared-vba/` 內容改為 **`modHttp` / `modJsonUtil` / `modPrices`
  ＋ 新 `modSECData` / `modMOPSData`**（後兩者是抽乾淨的取數層）。**不放** 完整
  `modSEC.bas` / `modMOPS.bas`（那兩支留在獨立活頁簿，改成呼叫 `modSECData` / `modMOPSData`）。
  `modTheme` / `modCharts` 都不進 `shared-vba` 的 RR4 import 清單。
- **§3.2**：把「headless 接縫」從「建議」升為「P1 前置必做」，理由＝§1 發現 1（編譯期硬相依）。
- **§3.3**：撞名處置維持「改 CTS 的 `ApplyDarkTheme`」——但因為本版不帶 `modTheme`，
  這條其實**本版不會觸發**，標為「僅在日後帶入 modTheme（加圖）時處理」。
- **§5 階段表 P1**：拆成 P1a「在獨立活頁簿內把 modSEC/modMOPS 重構出 `modSECData`/`modMOPSData`
  接縫，跑一次 `FetchSECFilings`/`FetchMOPSFilings` 確認不退化」→ P1b「把 3+2 個模組放進
  `shared-vba/`」。
- **§6 同步**：`sync-shared-vba.ps1` 的檔案清單 = 上面 5 個（modHttp/modJsonUtil/modPrices/
  modSECData/modMOPSData）。獨立活頁簿仍要 import 這 5 個 + 它自己的 modSEC/modMOPS/modCharts/modTheme。

---

*本檔為靜態盤點結果，未改動任何程式碼。*
