# Company Research x SEC-Filing-Fetcher 融合計畫

> 目標：把 SEC-Filing-Fetcher 的 **XBRL 財務數據抓取** 與 **MOPS 台股財報線** 做進
> RR4 活頁簿 `Compound RR4 Portfolio 2026 H2.xlsm` 的 **「Company research」工作表**，
> 採「上下兩區並存」——上半維持現有 MARKET SCANNER，下半新增「選定個股的財務深掘」。
> 全程 Excel VBA，獨立活頁簿 `SECFilingFetcher (17).xlsm` **保留、平行維護**。
>
> 交付形式：本 md（中度深度）。定案後再逐階段實作。

---

## 0. 實作進度（2026-09-03，本節取代下方 §3.1 / §4.1 的過期敘述）

P0 探勘與 P1 前半已完成。P0 揪出一個改寫計畫的發現：**modSEC ⇄ modCharts 是雙向相依**
（modSEC 呼 modCharts 畫圖、modCharts 呼 modSEC 的 `LookupConceptValue`），所以
**不能**把完整 `modSEC.bas` / `modMOPS.bas` import 進 RR4——只能給 RR4「抽乾淨的取數層」。
細節見 `RR4/company-research-p0-dependency-audit.md`。

因此 `shared-vba/` 的實際內容改為：

| 檔案 | 角色 | 狀態 |
|---|---|---|
| `shared-vba/modHttp.bas` | HTTP（含 `SEC_USER_AGENT`）| SEC-Filing-Fetcher 逐位元組複製 ✅ |
| `shared-vba/modJsonUtil.bas` | 手刻 JSON | 逐位元組複製 ✅ |
| `shared-vba/modPrices.bas` | Yahoo 日價 | 逐位元組複製 ✅ |
| `shared-vba/modSECData.bas` | **新** · US headless：`GetXbrlFinancials` / `XbrlValue` / `XbrlFyFp` | 已寫，純 ASCII，10 個 Private helper 與 `modSEC.bas` 逐行相同（Python 比對）· **未經 Excel 編譯/執行驗證** |
| `shared-vba/modMOPSData.bas` | **新** · TW headless：`GetTwFinancials` / `GetTwMonthlyRevenue` | 已寫，純 ASCII，21 個 Private helper 與 `modMOPS.bas` 逐行相同 · **未驗證** |
| `shared-vba/README.md` | 共用契約 + 同步 + 驗證方式 | ✅ |
| `scripts/sync-shared-vba.ps1` | `CodeModule.AddFromString` 注入兩個活頁簿（**非** VBE Import）| ✅ |
| `RR4/CompanyResearchSEC.bas` | RR4 下半區 renderer + `RunDeepDive` 進入點 | 已寫，純 ASCII，17/17 balanced · **未驗證** |
| `RR4/SheetCompanyResearch_Code.txt` | 「Company research」工作表 `Worksheet_Change`（監看 B34 / D34）| ✅ |

`modSECData` / `modMOPSData` 的 Public surface 刻意極小、且與 `modSEC` / `modMOPS` 同名者
全部 `Private`（兩個公開讀取器改名 `XbrlValue` / `XbrlFyFp`），所以可跟完整
`modSEC` / `modMOPS` **同專案共存不撞名**——§3.3 那個 `ApplyDarkTheme` 硬撞名因為
本版不帶 `modTheme`，**本版不會觸發**，留到日後加圖再處理。

**已實測（2026-09-03，COM in-memory，未存檔）**：
- `modSECData.GetXbrlFinancials("NVDA",4,8)` → ok，NVIDIA CORP，8 季，FY/FP 標籤正確，
  Revenue/NI/EPS 內部一致。
- `modMOPSData.GetTwFinancials("2395",6)` → ok，研華，毛利率 37.7% / 營益率 19.6% 對得上。
- `modMOPSData.GetTwMonthlyRevenue("2395",15)` → ok，「單月=累計差」＋YoY 正確。
- `CompanyResearchSEC.RunDeepDive`（對 RR4 活頁簿的**拋棄式複本**測，正本沒動）→ US／TW
  兩張財務表 + 衍生區 + 台股月營收區都畫出來。修掉一個測出來的 bug：`DsoRow` 的
  非短路 `And` 撞 `CDbl("")`（台股 Q4 欄 Revenue 空白）。已修。
- 已知/既有（與獨立工具一致，**非本次引入**）：NVDA CapEx 空白（單一 tag
  `PaymentsToAcquirePropertyPlantAndEquipment` 對不到 NVDA 申報）；NVDA FY2025Q1 EPS 5.98
  是分割前原始揭露值；台股年中季 CFO/CapEx 空白（MOPS 現金流只標 YTD context）。

**尚未做**：
1. **放進正本 RR4 活頁簿**：(a) 關掉 RR4 → 跑 `scripts/sync-shared-vba.ps1` + import
   `CompanyResearchSEC.bas` + 貼 `SheetCompanyResearch_Code.txt` 到工作表碼 → `Debug > Compile`；
   或 (b) 讓我注入你開著的那份、你自己 `Ctrl+S`。之後 `B34` 打 `NVDA` / `2395` 測端到端。
2. **收斂**：把 `SEC-Filing-Fetcher/modSEC.bas` / `modMOPS.bas` 改成呼叫 `*Data` 模組——
   步驟見 `shared-vba/CONVERGENCE.md`。**刻意還沒做**：動到你在用的工具、需要對它跑回歸；
   目前「兩份程式並存」安全可運作。
3. commit：建議兩個 scoped commit（① `shared-vba/` + `scripts/sync-shared-vba.ps1` +
   CLAUDE.md directory-map 那行；② `RR4/CompanyResearchSEC.bas` +
   `RR4/SheetCompanyResearch_Code.txt` + 計畫/稽核文件）。

---

## 1. 範圍界定

### 帶進來（IN）
- **US 標的**：`modSEC` 的 XBRL concept-map 路徑 → 逐季 / 逐年 營收、毛利、營業利益、
  EPS、現金流等（`BuildConceptMap` / `MergeConceptMaps` / `LookupConceptValue` /
  `LookupFyFp`）。含既有的 **Q4 = 全年 − 前三季** 反推處理（EDGAR company-concept
  沒有離散 Q4 duration，`modSEC` 內已處理，不要重寫）。
- **TW 標的**：`modMOPS` 的合併財報（`t164sb01` ix-XBRL → `ParseIxFacts` /
  `ExtractTWMetrics`）＋ 月營收（`BuildMonthlyRevenue` / `FetchOneMonthRevenue`，
  含「單月＝本月累計 − 上月累計」交叉驗證，已在 modMOPS 內，不要重寫）。

### 不帶（OUT，這一版）
- 申報清單本身（10-K / 10-Q / 8-K / Form 4 列表 + EDGAR 連結）——`FetchSECFilings`
  寫 `Filings` 表那條線不搬。
- `modCharts` 的圖表物件（股價 + EPS + 逐指標趨勢圖、`ExportTickerSnapshot`、
  `BuildDashboard`）。**先只做表格**。
- Watchlist / TWWatchlist 自動刷新機制（`RefreshWatchlistRow` /
  `RefreshTWWatchlistRow` 的多列監看）。

> ⚠️ 若日後要加 `modCharts` 圖，必須另建 `RawData` 工作表當圖表資料來源
> （見 CLAUDE.md「SEC-Filing-Fetcher 慣例」——圖表 series 指向 RawData 範圍，
> 不硬寫數字）。本版把數字直接寫在 Company research 下半區，是因為沒有圖、
> 不需要那層 series 綁定；加圖時這個決定要回頭改。

---

## 2. 現況盤點

### 2.1 RR4 側 —「Company research」工作表今天做什麼
| 項目 | 內容 |
|---|---|
| 模組 | `RR4/Sanner.bas`（模組名 `Sanner`，拼錯的 "Scanner"） |
| 進入點 | `Sub RunCompanyResearch(market As String, Sector As String)` |
| 觸發 | `OpenMarketScanner` → `frmMarketScanner.Show`（UserForm 選 market + sector）|
| 流程 | `GetSectorTickers(market, Sector)` → 逐檔 `Attach.GetTechData` +
  `Sanner.FetchFundamentals`（PE / FwdPE / PEG）→ `CalcMomentumScore` → 上色 →
  `DrawScanSummary` |
| 版面 | 第 1 列 banner；欄 A~S：TICKER / COMPANY / SECTOR / PRICE / HIGH DIST% /
  180D CHG% / DAY CHG% / BIAS5..240 / PE / FWD PE / PEG / TREND / SCORE |
| 附屬 | `ExportTopToWatchlist`、`ExportScanCSV`、`ColorizeRange` |
| 相依 | `Attach`（Yahoo 抓價 / GetTechData / GetCompanyName）、`TechIndicators` Type |

掃描列數少（一個 sector 通常 ≤ ~30 列），**第 32 列以下整片空白** → 下半區可用。

### 2.2 SEC-Filing-Fetcher 側 — 模組清單
| 模組 | 角色 | Public 進入點 | 內部（Private）關鍵 |
|---|---|---|---|
| `modSEC` | EDGAR / XBRL | `FetchSECFilings`、`CheckForNewFilings`、`RefreshWatchlistRow`、`LookupConceptValue`、`LookupFyFp` | `ResolveCIK`、`GetFilings`、`ScanRecentBlock`、`BuildConceptMap`、`MergeConceptMaps`、`WriteFilingRow` |
| `modMOPS` | 台股 MOPS | `FetchMOPSFilings`、`RefreshTWWatchlistRow`、`ClearTWWatchlistRow` | `FetchQuarterFacts`、`ParseIxFacts`、`ExtractTWMetrics`、`BuildMonthlyRevenue`、`FetchOneMonthRevenue`、`HttpGetBig5` |
| `modCharts` | 圖表 + RawData | `BuildDashboard`、`BuildQuarterlyDashboard`、`ExportTickerSnapshot`、`WriteRawBlock`、`FinishRawDataSheet`、`ClearAutoSeries`、`GetOrCreateSheet`、`ArrayMinMax` | 圖表繪製一票 |
| `modHttp` | HTTP | `HttpGet`（WinHTTP）| — |
| `modJsonUtil` | 手刻 JSON | `ExtractJsonValueRaw`、`ExtractJsonString`、`SplitJsonArrayElements`、`JsonUnescape`、`FindKeyValueStart` | `ExtractBalanced`、`ExtractQuotedRaw` |
| `modPrices` | Yahoo 日價 | `FetchDailyPrices`、`LatestPrice`、`NearestPriceOnOrBefore` | — |
| `modTheme` | 深色主題 / 軸 | `OrangePalette`、`ApplyDarkTheme`、`ApplyGridBorder`、`ApplyDarkThemeToChart`、`FitValueAxis` | `NiceStep`、`ColorSeries` |
| 工作表碼 | — | `SheetInput`（A1=US ticker→FetchSEC；B9=TW co_id→FetchMOPS）、`SheetWatchlist`、`SheetTWWatchlist` 的 `Worksheet_Change` | — |

### 2.3 這一版實際需要的子集
`modHttp` + `modJsonUtil` + `modSEC` + `modMOPS`（＋ `modPrices` 供估值比率的
價格對齊，＋ `modTheme` 若 `modSEC`/`modMOPS` 的列樣式有呼叫到）。**`modCharts` 不要。**

> Phase 0 要先確認：`modSEC` / `modMOPS` 內部有沒有呼叫 `modTheme.*` 或
> `modCharts.GetOrCreateSheet`。有的話那些也得進 shared-vba（或把呼叫點抽掉）。

---

## 3. 關鍵決策（採「你評估後建議」）

### 3.1 模組合併策略 → **共用模組 + fetch/render 接縫**
建 repo 根目錄 **`shared-vba/`** 資料夾，放「抓取 + 解析」層的 .bas。
`SECFilingFetcher.xlsm` 與 RR4 兩個 VBAProject **各自 import 同一份檔案**。

理由：
- 使用者要「平行維護」——若 RR4 各拿一份拷貝，兩份 XBRL 解析邏輯會分頭改，
  就是 CLAUDE.md 裡 globe-invest / article-db「兩 repo 手動同步」踩過的坑再演一次。
- 一份來源、兩邊 import，靠 build 腳本灌，比靠記性可靠。

### 3.2 為什麼需要「fetch/render 接縫」
`modSEC` 真正有用的 XBRL 邏輯（`GetFilings`、`BuildConceptMap`…）**全是 Private，
而且和 `Sheets("Input")` 讀設定、`Sheets("Filings")` 寫結果綁死**。RR4 這邊
不要那兩張表、也不要申報清單，所以不能直接呼叫 `FetchSECFilings`。

作法：在 `modSEC` / `modMOPS` 各開一個 **headless 公開函式**，回傳資料而不寫格子：

```
' modSEC.bas (shared-vba)
Public Function GetXbrlFinancials(ByVal tickerOrCik As String, _
                                  ByVal wantAnnual As Long, _
                                  ByVal wantQuarter As Long) As Object
    ' 回傳 Dictionary: "periods" -> 陣列(FYxxQn 標籤)
    '                 "revenue"/"grossProfit"/"opIncome"/"eps"/... -> 對齊陣列
    ' 內部重用現有 ResolveCIK / GetFilings / BuildConceptMap / MergeConceptMaps
    '           / LookupConceptValue / LookupFyFp，一行不改抓取邏輯
End Function
```

```
' modMOPS.bas (shared-vba)
Public Function GetTwFinancials(ByVal coId As String, ByVal quarters As Long) As Object
Public Function GetTwMonthlyRevenue(ByVal coId As String, ByVal months As Long) As Object
    ' 重用 FetchQuarterFacts / ExtractTWMetrics / FetchOneMonthRevenue
    ' 保留「單月 = 累計差」交叉驗證
End Function
```

- 既有 `FetchSECFilings` / `FetchMOPSFilings`（獨立活頁簿在用的那條）**改成呼叫
  這些新函式再自己寫表** → 獨立活頁簿行為不變，邏輯收斂成一份。
- 把用到的 Private helper 視需要升為 `Public`（或 `Friend`）。

> 折衷替代（若不想動 `FetchSECFilings`）：只加新函式、Private helper 升 Public、
> 舊寫表流程原封不動。省事但兩條路徑會各自演化。**建議走上面那條收斂版。**

### 3.3 名稱衝突處理
把 shared-vba 模組 import 進 RR4 後，同一 VBAProject 內兩個 **Public 同名** 程序
會直接編譯失敗（"Ambiguous name detected"）。實測比對結果：

| 名稱 | shared-vba 這邊 | RR4 既有 | 衝突？ | 處置 |
|---|---|---|---|---|
| `ApplyDarkTheme` | `modTheme` Public `Sub(ws, range, [hdr])` | `CommodityTermStructure` Public `Sub()`（無參數）| **是，硬衝突** | 把 `modTheme` 的改名 `ApplyDarkThemeRange`，更新其呼叫點（都在 `modCharts`，本版沒進 RR4，所以只需在獨立活頁簿改）；或把 CTS 那顆改 `ApplyCTSDarkTheme`（RR4-only，blast radius 較小）→ **建議改 CTS 那顆** |
| `HttpGet` | `modHttp` Public | `TW_Coverage_Parser` / `TaiwanPriceFetcher` 皆 **Private** | 否（module-scoped）| 不動；但專案內會並存 3 套 HTTP，記進「後續去重」 |
| `SafeNum` | `modMOPS` Private | — | 否 | 不動 |
| `WriteHeaders` | `modSEC` Private | `TW_Coverage_Parser` Private | 否 | 不動 |

其他重複但不撞名（維護異味，非本版必修）：JSON 解析（`modJsonUtil` vs
`Attach.JsonExtract` / `Sanner.ExtractNum` / `PortfolioDashboard_v3.ExtractJsonArray`）、
Yahoo 抓價（`modPrices` vs `Attach.GetStockPrice` / `FetchYahooData`）、
台股抓價（`modMOPS` 內 vs `TaiwanPriceFetcher` / `SharedFunctions_Price`）。

### 3.4 純 ASCII 鐵則
shared-vba 的 .bas 與任何 RR4 .bas 修改 **一律純 ASCII，不寫中文字面量**
（VBE 匯入用系統 ANSI 碼頁，中文會變 mojibake）。現成證據：`Sanner.bas` 第 43 行
`"MARKET SCANNER  ?X  "` 的破折號已經爛掉。中文標題一律走「英文寫死 +（可選）
另設中文對照表 / 由 RR4 既有 i18n 機制取字」。

---

## 4. 目標架構

### 4.1 檔案佈局
```
shared-vba/                     <- 新資料夾
  modHttp.bas                   從 SEC-Filing-Fetcher 移入（正本）
  modJsonUtil.bas
  modPrices.bas
  modTheme.bas                  （若 modSEC/modMOPS 有相依才進）
  modSEC.bas                    + 新 GetXbrlFinancials()
  modMOPS.bas                   + 新 GetTwFinancials() / GetTwMonthlyRevenue()
  README.md                     說明「兩活頁簿共用，改這裡、兩邊都要重灌」

SEC-Filing-Fetcher/
  modCharts.bas                 留在原地（獨立活頁簿專用）
  build.ps1                     改成也從 ../shared-vba/ 拉檔
  ...（Input/Watchlist 工作表碼不變）

RR4/
  CompanyResearchSEC.bas        <- 新模組：RR4 端的 renderer + 進入點
  SheetCompanyResearch_Code.txt <- 新：Company research 工作表的 Worksheet_Change 鏡像
  Sanner.bas                    幾乎不動（見 4.3）
  company-research-sec-merge-plan.md   <- 本檔
```

### 4.2 「Company research」工作表版面（上下兩區）
```
 列 1                MARKET SCANNER banner            <- Sanner 既有
 列 2..N             掃描結果表（欄 A~S）             <- Sanner 既有，N 動態
 列 N+1              掃描 SUMMARY 列                  <- Sanner 既有
 ...（留白緩衝）
 列 34   B34 = [ 輸入 ticker / 台股代號 ]  D34 = [US|TW|AUTO]   F34 = 狀態訊息
 列 36                "FINANCIALS  -  <entityName> (<ticker>)"   區塊標題
 列 37..~55           逐季 / 逐年 財務表（US: XBRL / TW: MOPS 合併財報）
 列 ~57..~70          台股月營收表（僅 TW；近 N 個月，單月 + 累計 + YoY%）
```
- 下半區起點 **固定列**（例：34），與掃描區用空白列隔開。掃描區若某次超過 33 列
  （sector 成分股很多），`CompanyResearchSEC` 先把下半區清掉再寫，或把起點參數化
  讀掃描結束列 + 緩衝。**建議固定列 + 溢位保護**。
- 深色主題沿用工作表現況（純黑底 + 琥珀色 heading），數字表用 `modTheme`
  （改名後的）列樣式；缺值一律寫 `-` 不留空、不寫 0。
- 每張表旁一條 ⚑ 來源註記（哪個 API / 哪份 filing）——比照 CLAUDE.md
  earnings 框架「每表必註記」。

### 4.3 資料流
```
使用者在 B34 打字
  -> Company research 工作表 Worksheet_Change
       偵測 Target 命中 B34（且非空）
       市場判定：D34=AUTO 時，全數字 -> TW co_id；含英文字母 -> US ticker
  -> CompanyResearchSEC.RunDeepDive(ticker, market)
       market=US: shared-vba modSEC.GetXbrlFinancials(ticker, wantAnnual, wantQ)
       market=TW: shared-vba modMOPS.GetTwFinancials(coId, quarters)
                  + modMOPS.GetTwMonthlyRevenue(coId, months)
  -> CompanyResearchSEC 把回傳 Dictionary 寫進列 36+ 的表格 + 上色 + 註記
  -> 失敗：F34 寫錯誤訊息，不丟例外對話框
```
US / TW 分流與 SEC-Filing-Fetcher `Input!A1`（US）vs `B9`（TW）的既有慣例一致。

---

## 5. 分階段實施

| 階段 | 內容 | 產出 | 可獨立驗證 |
|---|---|---|---|
| **P0 探勘** | 讀完 `modSEC` / `modMOPS` 全文，列出：① 對 `modTheme` / `modCharts` / `Sheets("Input")` / `Sheets("Filings")` 的每一處相依；② 要升 Public 的 Private helper 清單；③ `ThisWorkbook` / `ActiveWorkbook` 硬引用點 | 探勘筆記（附在本檔或另存）| — |
| **P1 建 shared-vba** | 把 `modHttp` `modJsonUtil` `modPrices`（＋ `modTheme` 若需）`modSEC` `modMOPS` 移到 `shared-vba/`；解 `ApplyDarkTheme` 撞名（改 CTS 那顆）；改 `SEC-Filing-Fetcher/build.ps1` 從 `../shared-vba/` 拉檔；重灌獨立活頁簿、跑一次 `FetchSECFilings` + `FetchMOPSFilings` 確認**沒退化** | shared-vba/ 6~7 檔；build.ps1 更新 | 獨立活頁簿 US + TW 各抓一檔，結果與改動前一致 |
| **P2 fetch/render 接縫** | `modSEC` 加 `GetXbrlFinancials`；`modMOPS` 加 `GetTwFinancials` / `GetTwMonthlyRevenue`；把 `FetchSECFilings` / `FetchMOPSFilings` 改成呼叫新函式再自己寫表（邏輯收斂）；Private helper 升 Public | shared-vba 更新 | 獨立活頁簿再跑一次，結果不變；新函式用 Excel COM 直接呼叫、印回傳 Dictionary |
| **P3 RR4 renderer** | 新 `RR4/CompanyResearchSEC.bas`：`RunDeepDive` + 表格 writer + 上色 + 註記 + 溢位清除；純 ASCII | CompanyResearchSEC.bas | 在 RR4 手動 `RunDeepDive("NVDA","US")` / `RunDeepDive("2395","TW")`，看列 36+ 出表 |
| **P4 觸發接線** | Company research 工作表碼加 `Worksheet_Change`（監看 B34）；鏡像存 `RR4/SheetCompanyResearch_Code.txt`；把 shared-vba + CompanyResearchSEC import 進 RR4 VBAProject；`Debug > Compile` 過 | 工作表碼 + import 清單 | 在 B34 打 `NVDA` / `2395`，自動出表；D34 切 US/TW 手動覆寫可用 |
| **P5 平行維護機制** | 定案 shared-vba 灌兩邊的方式（見 §6）；更新 CLAUDE.md（新資料夾進 directory map、新增同步規則）；README | scripts + CLAUDE.md 更新 | 改 shared-vba 一處 → 兩活頁簿重灌 → 都動到 |

Commit 切法（依 CLAUDE.md commit hygiene，一 concern 一 commit）：
P1 一個、P2 一個、P3+P4（RR4 整合）一個、P5（同步機制 + CLAUDE.md）一個。

---

## 6. 同步 / build 策略（shared-vba ↔ 兩活頁簿）

RR4 目前**沒有 build 腳本**，靠手動在 VBE `File > Import File`。三個選項：

| 選項 | 作法 | 優 | 缺 |
|---|---|---|---|
| **A. RR4 也做 build.ps1** | 仿 `SEC-Filing-Fetcher/build.ps1`，列出 RR4 全部模組 + `shared-vba/*` 一起注入 | 一鍵重灌、可 Excel COM 驗證 | 要維護 RR4 完整模組清單；RR4 模組多（12+） |
| **B. 只補 shared-vba 同步腳本** | `scripts/sync-shared-vba.ps1`：把 `shared-vba/*.bas` 注入兩個 .xlsm 的對應模組（用 VBProject.VBComponents 移除再匯入）| 改動小；RR4 其他模組維持手動 | 仍需手動處理 RR4 其他模組；要開啟「信任 VBA 專案物件模型」 |
| **C. 純文件化** | `shared-vba/README.md` 寫死「改完在兩活頁簿各 re-import 這幾個檔」 | 零腳本 | 靠記性，正是要避免的 |

**建議 B**：`scripts/sync-shared-vba.ps1`，只管 shared-vba 那 6~7 檔、灌進兩個
活頁簿；驗證沿用 CLAUDE.md「SEC-Filing-Fetcher 慣例」那套 Excel COM
（`$excel.AutomationSecurity = 1`、設輸入格觸發 `Worksheet_Change`、約 90 秒、
讀回格子）。`build.ps1` 印 SUCCESS 不代表 VBA 跑得起來，一定要 COM 實跑。
（本版無圖，`Chart.Export` 0 位元組那個坑不會遇到。）

---

## 7. 風險與緩解

| 風險 | 影響 | 緩解 |
|---|---|---|
| `modSEC` / `modMOPS` 對 `Sheets("Input")` 相依比預期深（不只讀設定）| P2 接縫變大改 | P0 先全文盤點；新函式參數化 wantAnnual/wantQuarter/months，不讀 Input |
| Private helper 升 Public 後在獨立活頁簿撞到別的名稱 | 獨立活頁簿編譯失敗 | 升級時加前綴（`SEC_` / `MOPS_`）而非裸名 |
| `ApplyDarkTheme` 之外還有沒抓到的撞名（Type、Const、Enum）| RR4 編譯失敗 | P4 `Debug > Compile` 前先跑一次名稱比對腳本；Type `TechIndicators` 明確不重定義 |
| 中文字面量混進 shared-vba | 兩活頁簿 UI 全變亂碼 | ASCII lint（grep 非 ASCII）納入 sync 腳本 |
| Company research 掃描區某次 > 33 列，蓋到下半區 | 版面壞 | `RunDeepDive` 開頭先 `Range("A34:S200").Clear`；或起點讀 SUMMARY 列 + 緩衝 |
| EDGAR / MOPS 改版或限流 | 抓不到資料 | 沿用 `modSEC` / `modMOPS` 既有錯誤處理；F34 顯示狀態，不彈例外 |
| 獨立活頁簿與 RR4 的 shared-vba 版本漂移 | 兩邊行為分歧 | sync 腳本一次灌兩邊；README 標「正本在 shared-vba/」 |

---

## 8. 驗證方式（每階段）

1. **獨立活頁簿不退化**（P1 / P2）：COM 開 `SECFilingFetcher (17).xlsm`，設
   `Input!A1 = "NVDA"` 觸發、`Input!B9 = "2395"` 觸發，比對 `Filings` 表輸出
   與改動前逐格一致。
2. **新函式**（P2）：COM 直接 `Application.Run("modSEC.GetXbrlFinancials", "NVDA", 3, 8)`，
   dump 回傳 Dictionary 的 keys / 陣列長度 / 抽樣值，對照 SEC 官網數字。
3. **RR4 renderer**（P3）：COM 開 RR4，`Application.Run("CompanyResearchSEC.RunDeepDive", "NVDA", "US")`，
   讀 Company research 列 36+，確認表格、上色、註記、缺值顯示 `-`。
4. **端到端**（P4）：COM 設 `Sheets("Company research").Range("B34") = "2395"`，
   等 `Worksheet_Change` 跑完，讀回月營收表 + 合併財報表；再測 `NVDA` 走 US 分支。
5. **回歸**：跑一次 `Sanner.RunCompanyResearch` 確認上半區掃描完全沒受影響。

---

## 9. CLAUDE.md 待更新項（P5 收尾）

- `## Directory map`：新增 `shared-vba/` —「SEC-Filing-Fetcher 與 RR4 共用的
  VBA 抓取層，正本在此，改完兩活頁簿都要 re-import；同步用
  `scripts/sync-shared-vba.ps1`」。
- `### 8. 非部署工具／資料` 表：`SEC 抓取工具（VBA）` 那列補註「XBRL / MOPS
  抓取層已抽到 `shared-vba/`，RR4 Company research 分頁共用」。
- `#### SEC-Filing-Fetcher 慣例`：補一句「headless 取數走
  `modSEC.GetXbrlFinancials` / `modMOPS.GetTwFinancials`，不要再靠 `Sheets("Input")`
  觸發流程；寫表 / 寫圖仍走舊路徑」。
- `scripts/` 條目：新增 `sync-shared-vba.ps1` 說明。

---

## 10. 未決 / 後續（不擋本版）

- **圖表**：要 `modCharts` 那批（股價 + EPS + 逐指標，`FitValueAxis`）時，必須先建
  `RawData` 工作表當 series 來源（CLAUDE.md 慣例），下半區改成「表 + 圖」雙欄。
- **HTTP / JSON / 抓價去重**：專案內會並存多套；獨立一個 refactor commit 再處理，
  不混進本次融合。
- **月營收「季報後最新月」標註**：若之後 Company research 要餵 earnings 框架，
  比照 CLAUDE.md 台股月營收輔助判斷，把「尚未被任何財報涵蓋的最新月」標出來。
- **AU / 其他市場**：本版只 US + TW，其他市場走 `-` / 不適用。

---

*本檔為實作計畫，非投資建議。*
