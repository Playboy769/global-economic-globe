# RR4 VBA Modules — Aggregate

> 由 `RR4/*.bas` 自動彙整產生 · 2026-08-06 23:00
>
> 來源 11 個模組 · 9236 行 · 355 KB · 110 Sub / 91 Function

## 這份文件的範圍

內容**完全來自 `.bas` 匯出檔**，逐字節保留原始縮排與字串內容，可隨時重新產生。
原始檔編碼不統一（UTF-8 / CP950 混雜），此處一律正規化為 UTF-8。

> ⚠️ **不含**下列 VBA 物件 —— 它們沒有 `.bas` 副本，目前只存在於舊快照
> `RR4 Module md Aggregate.txt`（2026-06-08，由 PDF 轉出，縮排與字串空白已損毀，僅供閱讀不可回貼）：
>
> - 工作表 code-behind：`工作表3`、`工作表6`、`工作表7`、`工作表10`、`工作表11`
>   （`CmdBtn0~5_Click` 按鈕接線、`Worksheet_Change` 自動觸發）
> - UserForm：`frmTransaction`、`frmDeleteTransaction`、`frmMarketScanner`
> - 未匯出模組：`CampaignSheet`、`Test`

---

## 模組總覽

| # | 模組 | VB_Name | 行數 | 大小 | 原始編碼 | 成員 |
|---:|---|---|---:|---:|---|---|
| 1 | [`Attach.bas`](#1-attachbas) | `Attach` | 1447 | 55.3 KB | UTF-8 | 16 Sub · 19 Fn |
| 2 | [`CommodityTermStructure.bas`](#2-commoditytermstructurebas) | `CommodityTermStructure` | 1433 | 55.8 KB | UTF-8 | 18 Sub · 14 Fn |
| 3 | [`CorelationMatrix.bas`](#3-corelationmatrixbas) | `CorelationMatrix` | 309 | 11.9 KB | CP950 (Big5) | 3 Sub · 2 Fn |
| 4 | [`DrawDownShadow.bas`](#4-drawdownshadowbas) | `DrawDownShadow` | 478 | 16.7 KB | ASCII | 1 Sub |
| 5 | [`modvolatility.bas`](#5-modvolatilitybas) | `modvolatility` | 317 | 11.2 KB | CP950 (Big5) | 2 Sub · 1 Fn |
| 6 | [`PortfolioDashboard_v3.bas`](#6-portfoliodashboard_v3bas) | `PortfolioDashboard_v3` | 1860 | 68 KB | UTF-8 | 22 Sub · 12 Fn |
| 7 | [`Sanner.bas`](#7-sannerbas) | `Sanner` | 644 | 32 KB | CP950 (Big5) | 10 Sub · 4 Fn |
| 8 | [`SharedFunctions_Price.bas`](#8-sharedfunctions_pricebas) | `SharedFunctions_Price` | 186 | 6.4 KB | CP950 (Big5) | 1 Sub · 3 Fn |
| 9 | [`TaiwanPriceFetcher.bas`](#9-taiwanpricefetcherbas) | `TaiwanPriceFetcher` | 196 | 7.5 KB | UTF-8 | 2 Sub · 6 Fn |
| 10 | [`TickerInsight.bas`](#10-tickerinsightbas) | `TickerInsight` | 1179 | 42.7 KB | ASCII | 20 Sub · 11 Fn |
| 11 | [`TW_Coverage_Parser.bas`](#11-tw_coverage_parserbas) | `TW_Coverage_Parser` | 1148 | 45.4 KB | CP950 (Big5) | 15 Sub · 20 Fn |

---

## API 索引（全域，依名稱排序）

| 名稱 | 型別 | 可見性 | 模組 | 行 |
|---|---|---|---|---:|
| `AnalyzeAndOutput` | Sub | Public | `modvolatility.bas` | 61 |
| `ApplyDarkTheme` | Sub | Public | `CommodityTermStructure.bas` | 410 |
| `ApplyMatrixCellTheme` | Sub | Private | `CommodityTermStructure.bas` | 391 |
| `ApplyScanTrend` | Sub | Private | `Sanner.bas` | 381 |
| `ApplyScoreColor` | Sub | Private | `Sanner.bas` | 400 |
| `ApplySheetDarkBg` | Sub | Private | `TW_Coverage_Parser.bas` | 274 |
| `ApplyStructure` | Sub | Private | `CommodityTermStructure.bas` | 1307 |
| `AutoFitSheets` | Sub | Private | `TW_Coverage_Parser.bas` | 1107 |
| `AutoRefreshLoop` | Sub | Public | `PortfolioDashboard_v3.bas` | 1283 |
| `BackfillHistory` | Sub | Public | `PortfolioDashboard_v3.bas` | 169 |
| `BuildCorrelationMatrix` | Sub | Public | `CorelationMatrix.bas` | 2 |
| `BuildDrawdownShadow` | Sub | Public | `DrawDownShadow.bas` | 30 |
| `BuildFIFOHistory` | Sub | Private | `TickerInsight.bas` | 606 |
| `BuildHoldingsCorrelation` | Sub | Public | `PortfolioDashboard_v3.bas` | 1313 |
| `BuildPositions` | Function | Private | `PortfolioDashboard_v3.bas` | 1100 |
| `CalcCorrelation` | Function | Private | `CorelationMatrix.bas` | 274 |
| `CalcMomentumScore` | Function | Private | `Sanner.bas` | 332 |
| `CalcPortBeta` | Function | Private | `PortfolioDashboard_v3.bas` | 1161 |
| `CalcPositions` | Sub | Private | `PortfolioDashboard_v3.bas` | 912 |
| `CalculateBeta30D` | Function | Public | `Attach.bas` | 419 |
| `CalculateBias` | Function | Public | `Attach.bas` | 356 |
| `CalculateRealizedPnL` | Sub | Public | `Attach.bas` | 510 |
| `CleanSheetName` | Function | Private | `TW_Coverage_Parser.bas` | 742 |
| `CleanWikilinks` | Function | Public | `Attach.bas` | 967 |
| `CleanWikilinks` | Function | Private | `TW_Coverage_Parser.bas` | 1015 |
| `ClearAllData` | Sub | Public | `Attach.bas` | 675 |
| `ClearDataSheet` | Sub | Private | `TW_Coverage_Parser.bas` | 753 |
| `ClearPriceCache` | Sub | Public | `TaiwanPriceFetcher.bas` | 64 |
| `ClearRealizedFilter` | Sub | Public | `Attach.bas` | 783 |
| `ClearTickerData` | Sub | Private | `TickerInsight.bas` | 559 |
| `CollectionToArray` | Function | Public | `Attach.bas` | 988 |
| `ColorizeRange` | Sub | Public | `Sanner.bas` | 486 |
| `CorrGetCommonDates` | Sub | Private | `PortfolioDashboard_v3.bas` | 1638 |
| `CorrPearson` | Function | Private | `PortfolioDashboard_v3.bas` | 1700 |
| `CorrRenderSheet` | Sub | Private | `PortfolioDashboard_v3.bas` | 1721 |
| `DarkRenderMdTable` | Function | Private | `TW_Coverage_Parser.bas` | 562 |
| `DarkSectionHeader` | Sub | Private | `TW_Coverage_Parser.bas` | 543 |
| `DB_Header` | Sub | Private | `Attach.bas` | 1405 |
| `DB_Row` | Sub | Private | `Attach.bas` | 1420 |
| `DB_Section` | Sub | Private | `Attach.bas` | 1394 |
| `DebugPriceTest` | Sub | Public | `SharedFunctions_Price.bas` | 117 |
| `DiagnoseTicker` | Sub | Public | `TickerInsight.bas` | 144 |
| `DrawActivePosition` | Sub | Private | `TickerInsight.bas` | 768 |
| `DrawAllocTable` | Function | Private | `PortfolioDashboard_v3.bas` | 614 |
| `DrawBrokerHeader` | Sub | Private | `PortfolioDashboard_v3.bas` | 740 |
| `DrawBrokerSubtotal` | Sub | Private | `PortfolioDashboard_v3.bas` | 754 |
| `DrawColumnHeaders` | Sub | Private | `PortfolioDashboard_v3.bas` | 659 |
| `DrawDeepAnalysis` | Sub | Public | `PortfolioDashboard_v3.bas` | 369 |
| `DrawDisclaimer` | Sub | Private | `PortfolioDashboard_v3.bas` | 889 |
| `DrawHeader` | Sub | Private | `PortfolioDashboard_v3.bas` | 275 |
| `DrawLifetimeMetrics` | Sub | Private | `TickerInsight.bas` | 819 |
| `DrawProjection` | Sub | Private | `TickerInsight.bas` | 878 |
| `DrawScanSummary` | Sub | Private | `Sanner.bas` | 300 |
| `DrawShellHeaders` | Sub | Private | `TickerInsight.bas` | 464 |
| `DrawTradeHistory` | Sub | Private | `TickerInsight.bas` | 923 |
| `EncodeURL` | Function | Public | `Attach.bas` | 825 |
| `EncodeUTF8Char` | Function | Private | `TW_Coverage_Parser.bas` | 1092 |
| `EnsureSheet` | Sub | Private | `TW_Coverage_Parser.bas` | 726 |
| `EnsureTickerInsightSheet` | Function | Private | `TickerInsight.bas` | 404 |
| `ExportScanCSV` | Sub | Public | `Sanner.bas` | 460 |
| `ExportSnapshot` | Sub | Public | `CommodityTermStructure.bas` | 1376 |
| `ExportTopToWatchlist` | Sub | Public | `Sanner.bas` | 413 |
| `ExtractListItems` | Function | Private | `TW_Coverage_Parser.bas` | 1024 |
| `ExtractNum` | Function | Private | `Sanner.bas` | 251 |
| `ExtractSection` | Function | Private | `TW_Coverage_Parser.bas` | 986 |
| `ExtractSectionLoose` | Function | Private | `TW_Coverage_Parser.bas` | 625 |
| `ExtractSubSection` | Function | Private | `TW_Coverage_Parser.bas` | 997 |
| `ExtractSubSectionLoose` | Function | Private | `TW_Coverage_Parser.bas` | 635 |
| `FetchFundamentals` | Sub | Public | `Sanner.bas` | 190 |
| `FetchPrice` | Function | Private | `CommodityTermStructure.bas` | 652 |
| `FetchPriceArray` | Function | Private | `CorelationMatrix.bas` | 229 |
| `FetchPriceWithFallback` | Function | Private | `CommodityTermStructure.bas` | 676 |
| `FetchRaw` | Function | Private | `TW_Coverage_Parser.bas` | 817 |
| `FetchRawContent` | Function | Public | `Attach.bas` | 891 |
| `FetchSingleReport` | Sub | Public | `TW_Coverage_Parser.bas` | 35 |
| `FetchTWSE` | Function | Private | `SharedFunctions_Price.bas` | 10 |
| `FetchYahoo` | Function | Private | `CommodityTermStructure.bas` | 793 |
| `FetchYahooClose` | Function | Private | `SharedFunctions_Price.bas` | 66 |
| `FetchYahooData` | Function | Public | `Attach.bas` | 338 |
| `FilterRealizedData` | Sub | Public | `Attach.bas` | 758 |
| `FindNearestPrice` | Function | Private | `PortfolioDashboard_v3.bas` | 1181 |
| `FixAndRebuild` | Sub | Public | `CommodityTermStructure.bas` | 734 |
| `FixTickerMap` | Sub | Public | `CommodityTermStructure.bas` | 890 |
| `FormatReportSheet` | Sub | Private | `TW_Coverage_Parser.bas` | 293 |
| `FrontMonthYahooSymbol` | Function | Private | `CommodityTermStructure.bas` | 713 |
| `FuturesTicker` | Function | Public | `CommodityTermStructure.bas` | 966 |
| `FuturesTickerQuarterly` | Function | Public | `CommodityTermStructure.bas` | 1006 |
| `GetAllMDPaths` | Function | Private | `TW_Coverage_Parser.bas` | 791 |
| `GetBenchmarkTicker` | Function | Public | `Attach.bas` | 405 |
| `GetCellPrice` | Function | Private | `CommodityTermStructure.bas` | 1364 |
| `GetCompanyName` | Function | Public | `Attach.bas` | 33 |
| `GetCurrencyType` | Function | Public | `Attach.bas` | 78 |
| `GetExchange` | Function | Public | `CommodityTermStructure.bas` | 980 |
| `GetExRate` | Function | Private | `PortfolioDashboard_v3.bas` | 1222 |
| `GetFundamentalData` | Sub | Public | `Attach.bas` | 318 |
| `GetFXToTWD` | Function | Private | `TickerInsight.bas` | 1144 |
| `GetGitHubToken` | Function | Public | `Attach.bas` | 876 |
| `GetHistoricalData` | Function | Public | `modvolatility.bas` | 239 |
| `GetHistoricalPrice` | Function | Public | `PortfolioDashboard_v3.bas` | 130 |
| `GetHistoryPrices` | Function | Public | `Attach.bas` | 470 |
| `GetInceptionDate` | Function | Private | `PortfolioDashboard_v3.bas` | 1229 |
| `GetPriceSymbol` | Function | Private | `TickerInsight.bas` | 1137 |
| `GetSectorTickers` | Function | Public | `Sanner.bas` | 506 |
| `GetStartingCapital` | Function | Private | `PortfolioDashboard_v3.bas` | 1236 |
| `GetStockPrice` | Function | Public | `Attach.bas` | 201 |
| `GetStockPriceSafe` | Function | Private | `TickerInsight.bas` | 1148 |
| `GetTechData` | Function | Public | `Attach.bas` | 90 |
| `GetToken` | Function | Private | `TW_Coverage_Parser.bas` | 1123 |
| `GetTWStockPrice` | Function | Public | `TaiwanPriceFetcher.bas` | 50 |
| `HoldingsCorrRenderSheet` | Sub | Private | `PortfolioDashboard_v3.bas` | 1474 |
| `HttpGet` | Function | Private | `TaiwanPriceFetcher.bas` | 171 |
| `HttpGet` | Function | Private | `TW_Coverage_Parser.bas` | 847 |
| `InitFilterInterface` | Sub | Public | `Attach.bas` | 717 |
| `InitSheetLayout` | Sub | Private | `TickerInsight.bas` | 418 |
| `IsArrayInitialized` | Function | Private | `TW_Coverage_Parser.bas` | 27 |
| `IsIndexFutureBase` | Function | Private | `CommodityTermStructure.bas` | 699 |
| `IsQuarterlyOnly` | Function | Public | `CommodityTermStructure.bas` | 1016 |
| `IsTWTicker` | Function | Private | `TickerInsight.bas` | 1130 |
| `JsonExtract` | Function | Public | `Attach.bas` | 367 |
| `JsonExtractTrend` | Function | Public | `Attach.bas` | 381 |
| `LogHistory` | Sub | Private | `PortfolioDashboard_v3.bas` | 85 |
| `MarkError` | Sub | Private | `CommodityTermStructure.bas` | 1333 |
| `MoneyFmt` | Function | Private | `TickerInsight.bas` | 1061 |
| `MonthCode` | Function | Public | `CommodityTermStructure.bas` | 960 |
| `Name_Safe` | Sub | Private | `CorelationMatrix.bas` | 300 |
| `NormalizeTicker` | Function | Private | `TickerInsight.bas` | 1112 |
| `NormScore` | Function | Private | `Sanner.bas` | 343 |
| `OpenDeleteForm` | Sub | Public | `Attach.bas` | 710 |
| `OpenMarketScanner` | Sub | Public | `Sanner.bas` | 482 |
| `OpenTickerInsight` | Sub | Public | `TickerInsight.bas` | 110 |
| `OpenTransactionForm` | Sub | Public | `Attach.bas` | 706 |
| `PaintHeaderRow` | Sub | Private | `CommodityTermStructure.bas` | 631 |
| `ParseClosePrice` | Function | Private | `Attach.bas` | 282 |
| `ParseFinancials` | Sub | Private | `TW_Coverage_Parser.bas` | 925 |
| `ParsePilotReports` | Sub | Public | `TW_Coverage_Parser.bas` | 647 |
| `ParsePrice` | Function | Private | `TaiwanPriceFetcher.bas` | 117 |
| `ParseValuationField` | Function | Private | `TW_Coverage_Parser.bas` | 253 |
| `ParseYahooPrice` | Function | Private | `CommodityTermStructure.bas` | 850 |
| `PauseMS` | Sub | Private | `CommodityTermStructure.bas` | 1371 |
| `PauseMS` | Sub | Private | `TaiwanPriceFetcher.bas` | 192 |
| `PnLColor` | Function | Private | `PortfolioDashboard_v3.bas` | 1200 |
| `PnLColorMuted` | Function | Private | `PortfolioDashboard_v3.bas` | 1214 |
| `PnLFmt` | Function | Private | `TickerInsight.bas` | 1066 |
| `QuarterlyMonthCode` | Function | Public | `CommodityTermStructure.bas` | 996 |
| `QueryCoverageReport` | Sub | Public | `Attach.bas` | 796 |
| `RebuildActiveXButtons` | Sub | Public | `PortfolioDashboard_v3.bas` | 537 |
| `RebuildPortfolioDashboard` | Sub | Public | `PortfolioDashboard_v3.bas` | 19 |
| `RebuildTickerMapDynamic` | Sub | Public | `CommodityTermStructure.bas` | 1028 |
| `RebuildTickerMapV3` | Sub | Public | `CommodityTermStructure.bas` | 1105 |
| `RefreshCorrelation` | Sub | Public | `CorelationMatrix.bas` | 305 |
| `RefreshDashboard` | Sub | Public | `CommodityTermStructure.bas` | 33 |
| `RefreshDashboardV2` | Sub | Public | `CommodityTermStructure.bas` | 200 |
| `RefreshSpreadMatrix` | Sub | Private | `CommodityTermStructure.bas` | 204 |
| `RefreshTickerInsight` | Sub | Public | `TickerInsight.bas` | 298 |
| `RefreshTickerProjection` | Sub | Public | `TickerInsight.bas` | 370 |
| `RegexFirst` | Function | Private | `TW_Coverage_Parser.bas` | 1040 |
| `RemoveBoldMarkers` | Function | Public | `Attach.bas` | 981 |
| `RemoveSheetShapes` | Sub | Private | `TickerInsight.bas` | 590 |
| `RenderReportToSheet` | Sub | Public | `Attach.bas` | 908 |
| `ResetSheetStyle` | Sub | Private | `PortfolioDashboard_v3.bas` | 236 |
| `RowHasValidPrices` | Function | Private | `CommodityTermStructure.bas` | 724 |
| `RunCompanyResearch` | Sub | Public | `Sanner.bas` | 18 |
| `RunSystemDebug` | Sub | Public | `Attach.bas` | 1010 |
| `SafeCell` | Function | Private | `TW_Coverage_Parser.bas` | 1049 |
| `SafeDate` | Function | Private | `TickerInsight.bas` | 1163 |
| `SafeNum` | Function | Private | `TickerInsight.bas` | 1173 |
| `SearchBySector` | Sub | Public | `TW_Coverage_Parser.bas` | 92 |
| `SearchGitHubFile` | Function | Public | `Attach.bas` | 843 |
| `SetGitHubToken` | Sub | Public | `Attach.bas` | 882 |
| `SetToken` | Sub | Public | `TW_Coverage_Parser.bas` | 1132 |
| `SetupPortfolioConfig` | Sub | Public | `PortfolioDashboard_v3.bas` | 1246 |
| `SetupSheets` | Sub | Private | `TW_Coverage_Parser.bas` | 698 |
| `SortArray` | Sub | Public | `Attach.bas` | 996 |
| `SplitCells` | Function | Private | `TW_Coverage_Parser.bas` | 1053 |
| `StartAutoRefresh` | Sub | Public | `PortfolioDashboard_v3.bas` | 1275 |
| `StopAutoRefresh` | Sub | Public | `PortfolioDashboard_v3.bas` | 1290 |
| `StripMetaLines` | Function | Private | `TW_Coverage_Parser.bas` | 1007 |
| `StripSuffix` | Function | Private | `TaiwanPriceFetcher.bas` | 164 |
| `SyncLayoutFromTickerMap` | Sub | Public | `CommodityTermStructure.bas` | 1186 |
| `TickerInsight_OnChange` | Sub | Public | `TickerInsight.bas` | 284 |
| `TPExPrice` | Function | Private | `TaiwanPriceFetcher.bas` | 93 |
| `TWSEExtract` | Function | Private | `SharedFunctions_Price.bas` | 44 |
| `TWSEPrice` | Function | Private | `TaiwanPriceFetcher.bas` | 73 |
| `UnmergeBand` | Sub | Private | `CommodityTermStructure.bas` | 624 |
| `UpdateVolatilityAnalysis_Pro` | Sub | Public | `modvolatility.bas` | 14 |
| `UrlEncodePath` | Function | Private | `TW_Coverage_Parser.bas` | 1069 |
| `USPnLColor` | Function | Private | `TickerInsight.bas` | 1096 |
| `WriteAllSheets` | Sub | Private | `TW_Coverage_Parser.bas` | 864 |
| `WriteColoredPct` | Sub | Private | `Sanner.bas` | 353 |
| `WriteDash` | Sub | Private | `TickerInsight.bas` | 1083 |
| `WriteHeaderRow` | Sub | Private | `TW_Coverage_Parser.bas` | 768 |
| `WriteHeaders` | Sub | Private | `TW_Coverage_Parser.bas` | 761 |
| `WriteKV` | Sub | Private | `PortfolioDashboard_v3.bas` | 1298 |
| `WriteLabel` | Sub | Private | `TickerInsight.bas` | 1003 |
| `WriteOnePositionRow` | Sub | Private | `PortfolioDashboard_v3.bas` | 783 |
| `WritePositionRows` | Function | Private | `PortfolioDashboard_v3.bas` | 697 |
| `WritePrice` | Sub | Private | `CommodityTermStructure.bas` | 1263 |
| `WriteSpread` | Sub | Private | `CommodityTermStructure.bas` | 1284 |
| `WriteValueMoney` | Sub | Private | `TickerInsight.bas` | 1020 |
| `WriteValuePct` | Sub | Private | `TickerInsight.bas` | 1071 |
| `WriteValuePnL` | Sub | Private | `TickerInsight.bas` | 1045 |
| `WriteValuePrice` | Sub | Private | `TickerInsight.bas` | 1033 |

---

## 1. Attach.bas

`VB_Name = "Attach"` · 1447 行 · 55.3 KB · 原始編碼 UTF-8

<details><summary><b>成員索引</b></summary>

| 行 | 型別 | 可見性 | 宣告 |
|---:|---|---|---|
| 33 | Function | Public | `Function GetCompanyName(Ticker As String) As String` |
| 78 | Function | Public | `Function GetCurrencyType(Ticker As String) As String` |
| 90 | Function | Public | `Function GetTechData(Ticker As String) As TechIndicators` |
| 201 | Function | Public | `Public Function GetStockPrice(Ticker As String) As Double` |
| 282 | Function | Private | `Private Function ParseClosePrice(resp As String) As Double` |
| 318 | Sub | Public | `Sub GetFundamentalData(http As Object, Ticker As String, ByRef res As TechIndicators)` |
| 338 | Function | Public | `Public Function FetchYahooData(http As Object, Ticker As String, mode As String) As String` |
| 356 | Function | Public | `Function CalculateBias(priceArr() As Double, period As Long, currentPrice As Double) As Double` |
| 367 | Function | Public | `Function JsonExtract(json As String, key As String) As Double` |
| 381 | Function | Public | `Function JsonExtractTrend(json As String, blockIndex As Long) As Double` |
| 405 | Function | Public | `Function GetBenchmarkTicker(Sector As String, stockTicker As String) As String` |
| 419 | Function | Public | `Function CalculateBeta30D(stockTicker As String, marketTicker As String) As Double` |
| 470 | Function | Public | `Function GetHistoryPrices(Ticker As String) As Object` |
| 510 | Sub | Public | `Sub CalculateRealizedPnL()` |
| 675 | Sub | Public | `Sub ClearAllData()` |
| 706 | Sub | Public | `Sub OpenTransactionForm()` |
| 710 | Sub | Public | `Sub OpenDeleteForm()` |
| 717 | Sub | Public | `Sub InitFilterInterface()` |
| 758 | Sub | Public | `Sub FilterRealizedData()` |
| 783 | Sub | Public | `Sub ClearRealizedFilter()` |
| 796 | Sub | Public | `Sub QueryCoverageReport()` |
| 825 | Function | Public | `Function EncodeURL(text As String) As String` |
| 843 | Function | Public | `Function SearchGitHubFile(tickerNum As String) As String` |
| 876 | Function | Public | `Function GetGitHubToken() As String` |
| 882 | Sub | Public | `Sub SetGitHubToken()` |
| 891 | Function | Public | `Function FetchRawContent(filePath As String) As String` |
| 908 | Sub | Public | `Sub RenderReportToSheet(content As String, tickerNum As String)` |
| 967 | Function | Public | `Function CleanWikilinks(text As String) As String` |
| 981 | Function | Public | `Function RemoveBoldMarkers(text As String) As String` |
| 988 | Function | Public | `Function CollectionToArray(col As Collection) As Variant` |
| 996 | Sub | Public | `Sub SortArray(ByRef arr As Variant)` |
| 1010 | Sub | Public | `Sub RunSystemDebug()` |
| 1394 | Sub | Private | `Private Sub DB_Section(ws As Worksheet, r As Long, title As String)` |
| 1405 | Sub | Private | `Private Sub DB_Header(ws As Worksheet, r As Long, c1 As String, c2 As String, c3 As String, c4 As String)` |
| 1420 | Sub | Private | `Private Sub DB_Row(ws As Worksheet, r As Long, category As String, item As String, status As String, detail As String, statusColor As Long)` |

</details>

```vba
Option Explicit

' ================================================================
' SHARED FUNCTIONS MODULE
' Used by: PortfolioDashboard_v3, UserForms, Correlation sheet
' ================================================================

Public Type TechIndicators
    price As Double
    ChangePercent As Double
    Change180D As Double
    DistFromHigh As Double
    Bias5 As Double
    Bias20 As Double
    Bias60 As Double
    Bias120 As Double
    Bias240 As Double
    PERatio As Double
    ForwardPE As Double
    PEGRatio As Double
    EPSNextQ As Double
    EPSNextY As Double
    Success As Boolean
End Type

' ================================================================
' PRICE CACHE (in-memory, 5-minute TTL)
' ================================================================
Public g_PriceCache As Object
Public g_CacheTime As Date

Function GetCompanyName(Ticker As String) As String
    Dim url As String, http As Object, response As String
    Dim nameStr As String, startPos As Long, endPos As Long

    Ticker = UCase(Trim(Ticker))
    If InStr(Ticker, ".TW") = 0 And InStr(Ticker, ".TWO") = 0 And IsNumeric(Ticker) Then
        Ticker = Ticker & ".TW"
    End If

    url = "https://query1.finance.yahoo.com/v8/finance/chart/" & Ticker & "?interval=1d&range=1d"

    On Error Resume Next
    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
    If http Is Nothing Then Set http = CreateObject("MSXML2.XMLHTTP")
    On Error GoTo 0

    With http
        .Open "GET", url, False
        .setRequestHeader "User-Agent", "Mozilla/5.0"
        .send
        response = .responseText
    End With

    Dim searchTag As String
    searchTag = """longName"":"""
    startPos = InStr(response, searchTag)

    If startPos > 0 Then
        startPos = startPos + Len(searchTag)
        endPos = InStr(startPos, response, """")
        GetCompanyName = Mid(response, startPos, endPos - startPos)
    Else
        searchTag = """shortName"":"""
        startPos = InStr(response, searchTag)
        If startPos > 0 Then
            startPos = startPos + Len(searchTag)
            endPos = InStr(startPos, response, """")
            GetCompanyName = Mid(response, startPos, endPos - startPos)
        Else
            GetCompanyName = Ticker
        End If
    End If
    Set http = Nothing
End Function

Function GetCurrencyType(Ticker As String) As String
    Ticker = UCase(Trim(Ticker))
    If InStr(Ticker, ".TW") > 0 Or InStr(Ticker, ".TWO") > 0 Then
        GetCurrencyType = "TWD"
    Else
        GetCurrencyType = "USD"
    End If
End Function

' ================================================================
' Technical Analysis
' ================================================================
Function GetTechData(Ticker As String) As TechIndicators
    Dim http As Object, response As String
    Dim result As TechIndicators
    Dim closeArr() As Double
    Dim i As Long, count As Long
    Dim currentPrice As Double, maxHigh As Double
    Dim rawTicker As String

    rawTicker = UCase(Trim(Ticker))
    result.Success = False
    Ticker = rawTicker

    If InStr(Ticker, ".TW") = 0 And InStr(Ticker, ".TWO") = 0 And IsNumeric(Ticker) Then
        Ticker = Ticker & ".TW"
    End If

    On Error Resume Next
    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
    If http Is Nothing Then Set http = CreateObject("MSXML2.XMLHTTP")
    On Error GoTo 0

    response = FetchYahooData(http, Ticker, "chart")

    If InStr(response, """timestamp""") = 0 And IsNumeric(rawTicker) Then
        Ticker = rawTicker & ".TWO"
        response = FetchYahooData(http, Ticker, "chart")
    End If

    Dim closeStart As Long, closeEnd As Long, closeStr As String, rawArr() As String
    Dim qPos As Long
    qPos = InStr(response, """quote"":[{")
    If qPos = 0 Then qPos = 1

    closeStart = InStr(qPos, response, """close"":[") + 9
    closeEnd = InStr(closeStart, response, "]")

    If closeStart > 8 And closeEnd > closeStart Then
        closeStr = Mid(response, closeStart, closeEnd - closeStart)
        rawArr = Split(closeStr, ",")
        ReDim closeArr(0 To UBound(rawArr))
        count = 0: maxHigh = 0

        For i = 0 To UBound(rawArr)
            If IsNumeric(rawArr(i)) Then
                closeArr(count) = Val(rawArr(i))
                If closeArr(count) > maxHigh Then maxHigh = closeArr(count)
                count = count + 1
            End If
        Next i

        If count > 0 Then
            ReDim Preserve closeArr(0 To count - 1)
            currentPrice = closeArr(count - 1)

            Dim metaPos As Long: metaPos = InStr(response, """regularMarketPrice"":")
            If metaPos > 0 Then
                Dim vals As Long: vals = metaPos + 21
                Dim valE As Long: valE = InStr(vals, response, ",")
                If valE = 0 Then valE = InStr(vals, response, "}")
                If valE > vals Then
                    Dim metaVal As String: metaVal = Trim(Mid(response, vals, valE - vals))
                    If IsNumeric(metaVal) And Val(metaVal) > 0 Then
                        Dim metaPx As Double: metaPx = Val(metaVal)
                        If metaPx <> currentPrice Then
                            currentPrice = metaPx
                            closeArr(count - 1) = currentPrice
                            If currentPrice > maxHigh Then maxHigh = currentPrice
                        End If
                    End If
                End If
            End If

            result.price = currentPrice

            If count >= 2 Then
                Dim prev As Double: prev = closeArr(count - 2)
                If prev > 0 Then result.ChangePercent = (currentPrice - prev) / prev
            End If

            Dim idx180 As Long: idx180 = count - 126
            If idx180 >= LBound(closeArr) And idx180 <= UBound(closeArr) Then
                If closeArr(idx180) > 0 Then
                    result.Change180D = (currentPrice - closeArr(idx180)) / closeArr(idx180)
                End If
            End If

            If maxHigh > 0 Then result.DistFromHigh = (currentPrice - maxHigh) / maxHigh

            result.Bias5 = CalculateBias(closeArr, 5, currentPrice)
            result.Bias20 = CalculateBias(closeArr, 20, currentPrice)
            result.Bias60 = CalculateBias(closeArr, 60, currentPrice)
            result.Bias120 = CalculateBias(closeArr, 120, currentPrice)
            result.Bias240 = CalculateBias(closeArr, 240, currentPrice)
            result.Success = True

            Call GetFundamentalData(http, Ticker, result)
        End If
    End If
    GetTechData = result
End Function

' ================================================================
' GET STOCK PRICE (with 5-min cache)
'
' Routing:
'   .TW / .TWO  -> TWSE / TPEx OpenAPI  (TaiwanPriceFetcher module)
'   US / others -> Yahoo Finance         (original path, Yahoo fallback)
'
' Adding routing here fixes ALL callers automatically:
'   PortfolioDashboard_v3, TickerInsight, CampaignSheet, etc.
' ================================================================
Public Function GetStockPrice(Ticker As String) As Double
    Ticker = UCase(Trim(Ticker))

    If g_PriceCache Is Nothing Then
        Set g_PriceCache = CreateObject("Scripting.Dictionary")
        g_CacheTime = Now
    End If

    If Now - g_CacheTime > TimeSerial(0, 5, 0) Then
        g_PriceCache.RemoveAll
        g_CacheTime = Now
    End If

    If g_PriceCache.Exists(Ticker) Then
        GetStockPrice = g_PriceCache(Ticker)
        Exit Function
    End If

    ' ── STEP 1: Normalise ticker ─────────────────────────────────────────
    Dim fullTicker As String: fullTicker = Ticker
    If InStr(fullTicker, ".TW") = 0 And InStr(fullTicker, ".TWO") = 0 And IsNumeric(fullTicker) Then
        fullTicker = fullTicker & ".TW"
    End If

    ' ── STEP 2: Yahoo Finance (primary) ──────────────────────────────────
    Dim http As Object
    On Error Resume Next
    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
    If http Is Nothing Then Set http = CreateObject("MSXML2.XMLHTTP")
    On Error GoTo 0

    Dim safeTicker As String: safeTicker = Replace(fullTicker, "^", "%5E")
    Dim url As String
    Dim resp As String: resp = ""
    url = "https://query1.finance.yahoo.com/v8/finance/chart/" & safeTicker & "?interval=1d&range=5d"

    On Error Resume Next
    With http
        .Open "GET", url, False
        .setRequestHeader "User-Agent", "Mozilla/5.0"
        .setRequestHeader "Accept", "application/json"
        .send
        resp = .responseText
    End With
    On Error GoTo 0

    Dim px As Double: px = ParseClosePrice(resp)

    ' .TWO fallback for bare numeric tickers (still within Yahoo)
    If px = 0 And IsNumeric(Ticker) And InStr(fullTicker, ".TWO") = 0 Then
        url = "https://query1.finance.yahoo.com/v8/finance/chart/" & Ticker & ".TWO?interval=1d&range=5d"
        On Error Resume Next
        With http
            .Open "GET", url, False
            .setRequestHeader "User-Agent", "Mozilla/5.0"
            .send
            resp = .responseText
        End With
        On Error GoTo 0
        px = ParseClosePrice(resp)
        If px > 0 Then fullTicker = Ticker & ".TWO"
    End If

    ' ── STEP 3: TWSE / TPEx OpenAPI (backup when Yahoo returns 0) ────────
    If px = 0 And InStr(fullTicker, ".TW") > 0 Then
        Dim openApiPx As Double: openApiPx = GetTWStockPrice(fullTicker)
        If openApiPx > 0 Then
            g_PriceCache(Ticker) = openApiPx
            GetStockPrice = openApiPx
            Set http = Nothing
            Exit Function
        End If
    End If
    ' ─────────────────────────────────────────────────────────────────────

    If px > 0 Then g_PriceCache(Ticker) = px
    GetStockPrice = px
    Set http = Nothing
End Function

' ----------------------------------------------------------------
Private Function ParseClosePrice(resp As String) As Double
    ParseClosePrice = 0

    Dim metaPos As Long: metaPos = InStr(resp, """regularMarketPrice"":")
    If metaPos > 0 Then
        Dim vals As Long: vals = metaPos + 21
        Dim valE As Long: valE = InStr(vals, resp, ",")
        If valE = 0 Then valE = InStr(vals, resp, "}")
        If valE > vals Then
            Dim metaVal As String: metaVal = Trim(Mid(resp, vals, valE - vals))
            If IsNumeric(metaVal) And Val(metaVal) > 0 Then
                ParseClosePrice = Val(metaVal)
                Exit Function
            End If
        End If
    End If

    Dim csPos As Long: csPos = InStr(resp, """close"":[")
    If csPos = 0 Then Exit Function
    Dim csS As Long: csS = csPos + 9
    Dim csE As Long: csE = InStr(csS, resp, "]")
    If csE <= csS Then Exit Function

    Dim parts() As String: parts = Split(Mid(resp, csS, csE - csS), ",")
    Dim i As Long
    For i = UBound(parts) To 0 Step -1
        Dim pt As String: pt = Trim(parts(i))
        If IsNumeric(pt) Then
            If CDbl(pt) > 0 Then
                ParseClosePrice = CDbl(pt)
                Exit Function
            End If
        End If
    Next i
End Function

Sub GetFundamentalData(http As Object, Ticker As String, ByRef res As TechIndicators)
    Dim response As String
    response = FetchYahooData(http, Ticker, "quoteSummary")
    If Len(response) < 100 Then Exit Sub

    response = Replace(Replace(Replace(Replace(response, " ", ""), vbCr, ""), vbLf, ""), vbTab, "")

    res.PERatio = JsonExtract(response, "trailingPE")
    res.ForwardPE = JsonExtract(response, "forwardPE")
    res.PEGRatio = JsonExtract(response, "pegRatio")
    res.EPSNextQ = JsonExtractTrend(response, 1)
    res.EPSNextY = JsonExtractTrend(response, 3)

    Dim apiPrice As Double: apiPrice = JsonExtract(response, "regularMarketPrice")
    Dim apiPrev As Double: apiPrev = JsonExtract(response, "regularMarketPreviousClose")
    If apiPrice > 0 And apiPrev > 0 Then
        res.ChangePercent = (apiPrice - apiPrev) / apiPrev
    End If
End Sub

Public Function FetchYahooData(http As Object, Ticker As String, mode As String) As String
    Dim url As String
    If mode = "chart" Then
        url = "https://query1.finance.yahoo.com/v8/finance/chart/" & Ticker & "?interval=1d&range=2y"
    Else
        url = "https://query1.finance.yahoo.com/v10/finance/quoteSummary/" & Ticker & _
              "?modules=summaryDetail,defaultKeyStatistics,earningsTrend,price"
    End If
    On Error Resume Next
    With http
        .Open "GET", url, False
        .setRequestHeader "User-Agent", "Mozilla/5.0"
        .send
        FetchYahooData = .responseText
    End With
    On Error GoTo 0
End Function

Function CalculateBias(priceArr() As Double, period As Long, currentPrice As Double) As Double
    If UBound(priceArr) + 1 < period Then CalculateBias = 0: Exit Function
    Dim k As Double: k = 2 / (period + 1)
    Dim ema As Double: ema = priceArr(0)
    Dim i As Long
    For i = 1 To UBound(priceArr)
        ema = (priceArr(i) - ema) * k + ema
    Next i
    If ema > 0 Then CalculateBias = (currentPrice - ema) / ema Else CalculateBias = 0
End Function

Function JsonExtract(json As String, key As String) As Double
    Dim searchStr As String: searchStr = """" & key & """:{""raw"":"
    Dim pos As Long: pos = InStr(json, searchStr)
    If pos = 0 Then JsonExtract = 0: Exit Function

    Dim valStart As Long: valStart = pos + Len(searchStr)
    Dim valEnd As Long: valEnd = InStr(valStart, json, ",")
    If valEnd = 0 Then valEnd = InStr(valStart, json, "}")
    If valEnd > valStart Then
        Dim tmp As String: tmp = Mid(json, valStart, valEnd - valStart)
        If IsNumeric(tmp) Then JsonExtract = Val(tmp)
    End If
End Function

Function JsonExtractTrend(json As String, blockIndex As Long) As Double
    Dim pos As Long: pos = 1
    Dim i As Long
    For i = 0 To blockIndex
        pos = InStr(pos, json, """earningsEstimate"":")
        If pos = 0 Then JsonExtractTrend = 0: Exit Function
        pos = pos + Len("""earningsEstimate"":")
    Next i
    Dim searchStr As String: searchStr = """avg"":{""raw"":"
    pos = InStr(pos, json, searchStr)
    If pos > 0 Then
        Dim valStart As Long: valStart = pos + Len(searchStr)
        Dim valEnd As Long: valEnd = InStr(valStart, json, ",")
        If valEnd = 0 Then valEnd = InStr(valStart, json, "}")
        If valEnd > valStart Then
            Dim tmp As String: tmp = Mid(json, valStart, valEnd - valStart)
            If IsNumeric(tmp) Then JsonExtractTrend = Val(tmp)
        End If
    End If
End Function

' ================================================================
' Beta Calculation
' ================================================================
Function GetBenchmarkTicker(Sector As String, stockTicker As String) As String
    If InStr(UCase(stockTicker), ".TW") > 0 Or InStr(UCase(stockTicker), ".TWO") > 0 Then
        GetBenchmarkTicker = "^TWII": Exit Function
    End If
    Select Case Sector
    Case ChrW(&H8CC7) & ChrW(&H8A0A) & ChrW(&H79D1) & ChrW(&H6280) & " (Information Technology)": GetBenchmarkTicker = "^IXIC"
    Case ChrW(&H5DE5) & ChrW(&H696D) & " (Industrials)": GetBenchmarkTicker = "^DJI"
    Case ChrW(&H91D1) & ChrW(&H878D) & " (Financials)": GetBenchmarkTicker = "XLF"
    Case ChrW(&H91AB) & ChrW(&H7642) & ChrW(&H4FDD) & ChrW(&H5065) & " (Health Care)": GetBenchmarkTicker = "XLV"
    Case ChrW(&H80FD) & ChrW(&H6E90) & " (Energy)": GetBenchmarkTicker = "XLE"
    Case Else: GetBenchmarkTicker = "^GSPC"
    End Select
End Function

Function CalculateBeta30D(stockTicker As String, marketTicker As String) As Double
    Dim stockPrices As Object, marketPrices As Object
    Set stockPrices = GetHistoryPrices(stockTicker)
    Set marketPrices = GetHistoryPrices(marketTicker)

    If stockPrices.Count < 5 Or marketPrices.Count < 5 Then
        CalculateBeta30D = -999: Exit Function
    End If

    Dim keys As Variant: keys = stockPrices.Keys
    Dim totalKeys As Long: totalKeys = UBound(keys) + 1
    Dim startIdx As Long: startIdx = IIf(totalKeys > 30, totalKeys - 30, 0)

    Dim arrStock() As Double
    Dim arrMarket() As Double
    ReDim arrStock(1 To totalKeys)
    ReDim arrMarket(1 To totalKeys)

    Dim validPairs As Long
    Dim prevStock As Double, prevMarket As Double
    Dim i As Long, dateKey As Variant

    For i = startIdx To UBound(keys)
        dateKey = keys(i)
        If marketPrices.Exists(dateKey) Then
            Dim currStock As Double: currStock = stockPrices(dateKey)
            Dim currMarket As Double: currMarket = marketPrices(dateKey)
            If prevStock > 0 Then
                validPairs = validPairs + 1
                arrStock(validPairs) = (currStock - prevStock) / prevStock
                arrMarket(validPairs) = (currMarket - prevMarket) / prevMarket
            End If
            prevStock = currStock: prevMarket = currMarket
        End If
    Next i

    If validPairs >= 10 Then
        Dim cleanS() As Double, cleanM() As Double
        ReDim cleanS(1 To validPairs): ReDim cleanM(1 To validPairs)
        For i = 1 To validPairs
            cleanS(i) = arrStock(i): cleanM(i) = arrMarket(i)
        Next i
        On Error Resume Next
        CalculateBeta30D = Application.WorksheetFunction.Slope(cleanS, cleanM)
        If Err.Number <> 0 Then CalculateBeta30D = -999
        On Error GoTo 0
    Else
        CalculateBeta30D = -999
    End If
End Function

Function GetHistoryPrices(Ticker As String) As Object
    Dim dict As Object: Set dict = CreateObject("Scripting.Dictionary")
    Ticker = UCase(Trim(Ticker))
    If InStr(Ticker, ".TW") = 0 And InStr(Ticker, ".TWO") = 0 And IsNumeric(Ticker) Then
        Ticker = Ticker & ".TW"
    End If

    Dim http As Object, response As String
    On Error Resume Next
    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
    If http Is Nothing Then Set http = CreateObject("MSXML2.XMLHTTP")
    With http
        .Open "GET", "https://query1.finance.yahoo.com/v8/finance/chart/" & Ticker & "?range=1y&interval=1d", False
        .setRequestHeader "User-Agent", "Mozilla/5.0"
        .send
        response = .responseText
    End With
    On Error GoTo 0

    Dim tsStart As Long: tsStart = InStr(response, """timestamp"":[") + 13
    Dim tsEnd As Long: tsEnd = InStr(tsStart, response, "]")
    Dim csStart As Long: csStart = InStr(response, """adjclose"":[") + 12
    Dim csEnd As Long: csEnd = InStr(csStart, response, "]")

    If tsStart > 12 And tsEnd > tsStart And csStart > 11 And csEnd > csStart Then
        Dim tsArr() As String: tsArr = Split(Mid(response, tsStart, tsEnd - tsStart), ",")
        Dim closeArr() As String: closeArr = Split(Mid(response, csStart, csEnd - csStart), ",")
        If UBound(tsArr) = UBound(closeArr) Then
            Dim i As Long
            For i = 0 To UBound(tsArr)
                If IsNumeric(tsArr(i)) And IsNumeric(closeArr(i)) Then
                    Dim dv As Date: dv = Int(CLng(tsArr(i)) / 86400 + 25569)
                    If Not dict.Exists(dv) Then dict.Add dv, CDbl(closeArr(i))
                End If
            Next i
        End If
    End If
    Set GetHistoryPrices = dict
End Function

Sub CalculateRealizedPnL()
    Dim wsTrans As Worksheet, wsReal As Worksheet, wsPort As Worksheet
    Set wsTrans = ThisWorkbook.Sheets("Transactions")
    Set wsReal = ThisWorkbook.Sheets("Realized")
    Set wsPort = ThisWorkbook.Sheets("RR4")
    Application.ScreenUpdating = False
    wsReal.Range("A2:I10000").ClearContents

    Dim hdrs As Variant
    hdrs = Array("TICKER", "PNL", "SHARES", "AVG COST", "NET AMT", "STRATEGY", "PNL(TWD)", "DATE", "BROKER")
    Dim hc As Integer
    For hc = 0 To 8
        wsReal.Cells(1, hc + 1).Value = hdrs(hc)
        wsReal.Cells(1, hc + 1).Font.Bold = True
    Next hc

    Dim dictLots As Object: Set dictLots = CreateObject("Scripting.Dictionary")

    Dim exRate As Double
    On Error Resume Next: exRate = Val(wsPort.Range("B2").Value): On Error GoTo 0
    If exRate <= 0 Then exRate = 32

    Dim lastRow As Long: lastRow = wsTrans.Cells(wsTrans.Rows.Count, "A").End(xlUp).Row
    Dim writeRow As Long: writeRow = 2
    Dim rr As Long

    For rr = 2 To lastRow
        Dim tDate As Variant: tDate = wsTrans.Cells(rr, 2).Value
        Dim Ticker As String: Ticker = UCase(Trim(wsTrans.Cells(rr, 3).Value))
        Dim Action As String: Action = UCase(Trim(wsTrans.Cells(rr, 4).Value))
        Dim shares As Double: shares = Val(wsTrans.Cells(rr, 5).Value)
        Dim NetAmount As Double: NetAmount = Val(wsTrans.Cells(rr, 9).Value)
        Dim strat As String: strat = CStr(wsTrans.Cells(rr, 12).Value)
        Dim broker As String: broker = Trim(CStr(wsTrans.Cells(rr, 14).Value))
        If broker = "" Then broker = "Default"

        Dim currType As String
        If InStr(Ticker, ".TW") > 0 Or InStr(Ticker, ".TWO") > 0 Then
            currType = "TWD"
        Else
            currType = "USD"
        End If

        If Ticker <> "" And shares > 0 Then
            Dim posKey As String: posKey = Ticker & "|" & broker

            If Not dictLots.Exists(posKey) Then
                Dim newCol As Collection
                Set newCol = New Collection
                dictLots.Add posKey, newCol
            End If

            Select Case Action
            Case "BUY"
                Dim costPerShare As Double
                costPerShare = NetAmount / shares
                dictLots(posKey).Add Array(shares, costPerShare)

            Case "ADJUSTCOST"
                Dim totalSh As Double: totalSh = 0
                Dim ll As Long
                For ll = 1 To dictLots(posKey).Count
                    totalSh = totalSh + dictLots(posKey)(ll)(0)
                Next ll
                If totalSh > 0 Then
                    Dim adjPerSh As Double: adjPerSh = NetAmount / totalSh
                    For ll = 1 To dictLots(posKey).Count
                        Dim lotArr As Variant: lotArr = dictLots(posKey)(ll)
                        lotArr(1) = lotArr(1) + adjPerSh
                        If ll < dictLots(posKey).Count Then
                            dictLots(posKey).Add lotArr, Before:=ll + 1
                            dictLots(posKey).Remove ll
                        Else
                            dictLots(posKey).Add lotArr
                            dictLots(posKey).Remove ll
                        End If
                    Next ll
                End If

            Case "SELL"
                Dim remaining As Double: remaining = shares
                Dim totalCostBasis As Double: totalCostBasis = 0

                Do While remaining > 0.00001 And dictLots(posKey).Count > 0
                    Dim frontLot As Variant: frontLot = dictLots(posKey)(1)
                    Dim lotSh As Double: lotSh = frontLot(0)
                    Dim lotCps As Double: lotCps = frontLot(1)

                    If lotSh <= remaining + 0.00001 Then
                        totalCostBasis = totalCostBasis + lotSh * lotCps
                        remaining = remaining - lotSh
                        dictLots(posKey).Remove 1
                    Else
                        totalCostBasis = totalCostBasis + remaining * lotCps
                        Dim updatedLot As Variant
                        updatedLot = Array(lotSh - remaining, lotCps)
                        If dictLots(posKey).Count > 1 Then
                            dictLots(posKey).Add updatedLot, Before:=2
                        Else
                            dictLots(posKey).Add updatedLot
                        End If
                        dictLots(posKey).Remove 1
                        remaining = 0
                    End If
                Loop

                Dim fifoAvgC As Double
                If shares > 0 Then fifoAvgC = totalCostBasis / shares Else fifoAvgC = 0
                Dim rlPnl As Double: rlPnl = NetAmount - totalCostBasis
                Dim rlTWD As Double: rlTWD = IIf(currType = "USD", rlPnl * exRate, rlPnl)

                Dim finalDate As Date
                If IsDate(tDate) Then finalDate = CDate(tDate) Else finalDate = Date

                With wsReal
                    .Cells(writeRow, 1) = Ticker
                    .Cells(writeRow, 2) = rlPnl: .Cells(writeRow, 2).NumberFormat = "#,##0.00"
                    .Cells(writeRow, 3) = shares
                    .Cells(writeRow, 4) = fifoAvgC
                    .Cells(writeRow, 5) = NetAmount
                    .Cells(writeRow, 6) = strat
                    .Cells(writeRow, 7) = rlTWD: .Cells(writeRow, 7).NumberFormat = "#,##0"
                    .Cells(writeRow, 8) = finalDate: .Cells(writeRow, 8).NumberFormat = "yyyy/m/d"
                    .Cells(writeRow, 9) = broker
                    If rlTWD > 0 Then
                        .Cells(writeRow, 2).Font.Color = RGB(255, 100, 100)
                        .Cells(writeRow, 7).Font.Color = RGB(255, 100, 100)
                    Else
                        .Cells(writeRow, 2).Font.Color = RGB(0, 255, 100)
                        .Cells(writeRow, 7).Font.Color = RGB(0, 255, 100)
                    End If
                End With
                writeRow = writeRow + 1

            End Select
        End If
    Next rr

    ' Update RR4 Entry PX
    Dim portLastRow As Long: portLastRow = wsPort.Cells(wsPort.Rows.Count, "A").End(xlUp).Row
    Dim pr As Long
    For pr = 5 To portLastRow
        Dim pTicker As String: pTicker = UCase(Trim(wsPort.Cells(pr, 1).Value))
        Dim pBroker As String: pBroker = Trim(CStr(wsPort.Cells(pr, 13).Value))
        If pBroker = "" Then pBroker = "Default"

        Dim pKey As String: pKey = pTicker & "|" & pBroker

        If dictLots.Exists(pKey) And pTicker <> "" Then
            Dim totSh As Double: totSh = 0
            Dim totCost As Double: totCost = 0
            For ll = 1 To dictLots(pKey).Count
                totSh = totSh + dictLots(pKey)(ll)(0)
                totCost = totCost + dictLots(pKey)(ll)(0) * dictLots(pKey)(ll)(1)
            Next ll
            If totSh > 0.00001 Then
                wsPort.Cells(pr, 9).Value = totCost / totSh
            End If
        End If
    Next pr

    wsReal.Columns("A:I").AutoFit
    Application.ScreenUpdating = True
End Sub

Sub ClearAllData()
    If MsgBox("Clear ALL data? This cannot be undone!", vbYesNo + vbExclamation + vbDefaultButton2) = vbNo Then Exit Sub

    Dim wsTrans As Worksheet, wsPort As Worksheet, wsReal As Worksheet, wsHist As Worksheet
    On Error Resume Next
    Set wsTrans = ThisWorkbook.Sheets("Transactions")
    Set wsPort = ThisWorkbook.Sheets("RR4")
    Set wsReal = ThisWorkbook.Sheets("Realized")
    Set wsHist = ThisWorkbook.Sheets("HistoryLog")
    On Error GoTo 0

    Dim LR As Long
    If Not wsTrans Is Nothing Then
        LR = wsTrans.Cells(wsTrans.Rows.Count, "A").End(xlUp).Row
        If LR > 1 Then wsTrans.Range("A2:N" & LR).ClearContents
    End If
    If Not wsReal Is Nothing Then
        LR = wsReal.Cells(wsReal.Rows.Count, "A").End(xlUp).Row
        If LR > 1 Then wsReal.Range("A2:I" & LR).ClearContents
    End If
    If Not wsHist Is Nothing Then
        LR = wsHist.Cells(wsHist.Rows.Count, "A").End(xlUp).Row
        If LR > 1 Then wsHist.Range("A2:G" & LR).ClearContents
    End If

    MsgBox "All data cleared.", vbInformation
End Sub

' ================================================================
' Form Launchers
' ================================================================
Sub OpenTransactionForm()
    frmTransaction.Show
End Sub

Sub OpenDeleteForm()
    frmDeleteTransaction.Show
End Sub

' ================================================================
' Realized Sheet Filter UI
' ================================================================
Sub InitFilterInterface()
    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets("Realized")
    On Error Resume Next: ws.Unprotect: On Error GoTo 0
    ws.AutoFilterMode = False
    With ws.Cells
        .Interior.Color = RGB(0, 0, 0)
        .Font.Color = RGB(255, 255, 255)
        .Font.Name = "Calibri"
        .Font.Size = 12
    End With
    With ws.Range("J1:L20")
        .UnMerge: .ClearContents: .ColumnWidth = 12
    End With
    With ws
        .Range("J1:L6").Interior.Color = RGB(30, 30, 30)
        .Range("J1:L1").Merge
        .Range("J1").HorizontalAlignment = xlCenter
        .Range("J1").Font.Bold = True
        .Range("J1").Font.Color = RGB(255, 215, 0)
        .Range("J2").Value = "Start Date:"
        .Range("J3").Value = "End Date:"
        .Range("K2").Value = DateSerial(Year(Date), 1, 1)
        .Range("K3").Value = Date
        With .Range("K2:K3")
            .NumberFormat = "yyyy/m/d"
            .Interior.Color = RGB(60, 60, 60)
            .Borders.LineStyle = xlContinuous
            .Font.Color = RGB(255, 255, 255)
            .Font.Bold = True
        End With
        .Range("J5").Value = "Period P&L:"
        .Range("J5").Font.Bold = True
        .Range("K5").NumberFormat = "$#,##0"
        .Range("K5").Font.Size = 14
        .Range("K5").Font.Bold = True
        .Range("L2").Value = "(e.g. 2025/1/1)"
        .Range("L2").Font.Color = RGB(150, 150, 150)
        .Range("L2").Font.Size = 10
    End With
End Sub

Sub FilterRealizedData()
    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets("Realized")
    Dim LR As Long: LR = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    If LR < 2 Then MsgBox "No data to filter", vbInformation: Exit Sub

    Dim sd As Date, ed As Date
    On Error Resume Next
    If IsDate(ws.Range("K2").Value) Then sd = CDate(ws.Range("K2").Value)
    If IsDate(ws.Range("K3").Value) Then ed = CDate(ws.Range("K3").Value)
    On Error GoTo 0
    If sd = 0 Or ed = 0 Then MsgBox "Enter valid dates in K2 and K3", vbExclamation: Exit Sub

    ws.AutoFilterMode = False
    ws.Range("A1:I" & LR).AutoFilter Field:=8, Criteria1:=">=" & sd, Operator:=xlAnd, Criteria2:="<=" & ed

    Dim PnL As Double
    On Error Resume Next
    PnL = Application.WorksheetFunction.Subtotal(109, ws.Range("G2:G" & LR))
    On Error GoTo 0

    ws.Range("K5").Value = PnL
    ws.Range("K5").Font.Color = IIf(PnL >= 0, RGB(255, 100, 100), RGB(0, 255, 100))
    MsgBox "Filter applied!", vbInformation
End Sub

Sub ClearRealizedFilter()
    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets("Realized")
    ws.AutoFilterMode = False
    Dim LR As Long: LR = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    Dim total As Double
    If LR >= 2 Then total = Application.WorksheetFunction.Sum(ws.Range("G2:G" & LR))
    ws.Range("K5").Value = total
    ws.Range("K5").Font.Color = IIf(total >= 0, RGB(255, 100, 100), RGB(0, 255, 100))
End Sub

' ================================================================
' GitHub Coverage Report
' ================================================================
Sub QueryCoverageReport()
    Dim Ticker As String
    Ticker = InputBox("Enter TW stock code (e.g. 2330):", "Coverage Report")
    Ticker = Trim(Ticker)
    If Ticker = "" Then Exit Sub
    Dim pureNum As String
    pureNum = Replace(Replace(UCase(Ticker), ".TWO", ""), ".TW", "")
    If Not IsNumeric(pureNum) Then MsgBox "Numeric code only (e.g. 2330)", vbExclamation: Exit Sub

    Application.StatusBar = "Searching " & pureNum & "..."
    Application.ScreenUpdating = False

    Dim filePath As String: filePath = SearchGitHubFile(pureNum)
    If filePath = "" Then
        Application.StatusBar = False: Application.ScreenUpdating = True
        MsgBox "Report not found for " & pureNum, vbExclamation: Exit Sub
    End If

    Dim content As String: content = FetchRawContent(filePath)
    If content = "" Then
        Application.StatusBar = False: Application.ScreenUpdating = True
        MsgBox "Failed to load report.", vbExclamation: Exit Sub
    End If

    Call RenderReportToSheet(content, pureNum)
    Application.StatusBar = pureNum & " loaded."
    Application.ScreenUpdating = True
End Sub

Function EncodeURL(text As String) As String
    Dim i As Long, result As String, c As String
    For i = 1 To Len(text)
        c = Mid(text, i, 1)
        Select Case c
        Case "A" To "Z", "a" To "z", "0" To "9", "-", "_", ".", "/", "~"
            result = result & c
        Case Else
            Dim bytes() As Byte: bytes = StrConv(c, vbFromUnicode)
            Dim b As Long
            For b = 0 To UBound(bytes)
                result = result & "%" & UCase(Hex(bytes(b)))
            Next b
        End Select
    Next i
    EncodeURL = result
End Function

Function SearchGitHubFile(tickerNum As String) As String
    Dim http As Object, response As String
    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
    Dim ghToken As String: ghToken = GetGitHubToken()
    On Error Resume Next
    With http
        .Open "GET", "https://api.github.com/repos/Timeverse/My-TW-Coverage/git/trees/master?recursive=1", False
        .setRequestHeader "User-Agent", "VBA-Excel-Client"
        .setRequestHeader "Accept", "application/vnd.github.v3+json"
        If ghToken <> "" Then .setRequestHeader "Authorization", "token " & ghToken
        .send
        response = .responseText
    End With
    On Error GoTo 0
    If Len(response) < 100 Then SearchGitHubFile = "": Exit Function

    Dim pat As String: pat = """path"":""Pilot_Reports"
    Dim pos As Long: pos = 1
    Do
        pos = InStr(pos, response, pat)
        If pos = 0 Then SearchGitHubFile = "": Exit Function
        Dim vs As Long: vs = pos + 8
        Dim ve As Long: ve = InStr(vs, response, """")
        If ve > vs Then
            Dim pv As String: pv = Mid(response, vs, ve - vs)
            If InStr(pv, "/" & tickerNum & "_") > 0 Then
                SearchGitHubFile = pv: Exit Function
            End If
        End If
        pos = pos + Len(pat)
    Loop
End Function

Function GetGitHubToken() As String
    On Error Resume Next
    GetGitHubToken = ThisWorkbook.Sheets("HistoryLog").Range("Z3").Value
    On Error GoTo 0
End Function

Sub SetGitHubToken()
    Dim token As String
    token = InputBox("Paste GitHub Token (ghp_...):", "Set Token")
    If token = "" Then Exit Sub
    ThisWorkbook.Sheets("HistoryLog").Range("Z3").Value = token
    ThisWorkbook.Sheets("HistoryLog").Range("Z3").Font.Color = RGB(0, 0, 0)
    MsgBox "Token saved!", vbInformation
End Sub

Function FetchRawContent(filePath As String) As String
    Dim http As Object, ghToken As String
    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
    ghToken = GetGitHubToken()
    Dim rawUrl As String
    rawUrl = "https://raw.githubusercontent.com/Timeverse/My-TW-Coverage/master/" & EncodeURL(filePath)
    On Error Resume Next
    With http
        .Open "GET", rawUrl, False
        .setRequestHeader "User-Agent", "VBA-Excel-Client"
        If ghToken <> "" Then .setRequestHeader "Authorization", "token " & ghToken
        .send
        If .status = 200 Then FetchRawContent = .responseText Else FetchRawContent = ""
    End With
    On Error GoTo 0
End Function

Sub RenderReportToSheet(content As String, tickerNum As String)
    Dim ws As Worksheet
    On Error Resume Next: Set ws = ThisWorkbook.Sheets("Coverage"): On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        ws.Name = "Coverage"
    End If
    With ws
        .Cells.ClearContents: .Cells.ClearFormats
        .Cells.Interior.Color = RGB(10, 12, 18)
        .Cells.Font.Color = RGB(210, 210, 210)
        .Cells.Font.Name = "Calibri"
        .Cells.Font.Size = 11
        .Columns("A").ColumnWidth = 5
        .Columns("B").ColumnWidth = 100
        .Rows.RowHeight = 16
    End With

    Dim lines() As String: lines = Split(content, vbLf)
    Dim writeRow As Long: writeRow = 2
    Dim i As Long
    For i = 0 To UBound(lines)
        Dim lt As String: lt = Replace(lines(i), vbCr, "")
        If lt = "" Then writeRow = writeRow + 1: GoTo NL
        With ws.Cells(writeRow, "B")
            Dim dt As String: dt = lt
            If Left(lt, 2) = "# " Then
                dt = CleanWikilinks(Mid(lt, 3))
                .Value = dt: .Font.Size = 18: .Font.Bold = True
                .Font.Color = RGB(255, 215, 0): .Interior.Color = RGB(25, 25, 35)
                writeRow = writeRow + 1
            ElseIf Left(lt, 3) = "## " Then
                dt = CleanWikilinks(Mid(lt, 4))
                .Value = "## " & dt: .Font.Size = 13: .Font.Bold = True
                .Font.Color = RGB(80, 180, 255): .Interior.Color = RGB(20, 22, 32)
            ElseIf Left(lt, 4) = "### " Then
                dt = CleanWikilinks(Mid(lt, 5))
                .Value = " ### " & dt: .Font.Size = 12: .Font.Bold = True
                .Font.Color = RGB(120, 230, 160)
            ElseIf Left(lt, 2) = "- " Or Left(lt, 2) = "* " Then
                .Value = " * " & CleanWikilinks(Mid(lt, 3))
                .Font.Color = RGB(200, 200, 200)
            ElseIf Left(lt, 1) = "|" Then
                .Value = CleanWikilinks(lt): .Font.Name = "Courier New"
                .Font.Size = 10: .Font.Color = RGB(180, 220, 180)
                .Interior.Color = RGB(18, 22, 28)
            ElseIf Left(lt, 3) = "```" Then
                GoTo NL
            Else
                .Value = RemoveBoldMarkers(CleanWikilinks(lt))
                .Font.Color = RGB(200, 200, 200)
            End If
        End With
        writeRow = writeRow + 1
NL:
    Next i
    ws.Activate: ws.Cells(1, 1).Select: ActiveWindow.ScrollRow = 1
End Sub

Function CleanWikilinks(text As String) As String
    Dim r As String: r = text
    Do While InStr(r, "[[") > 0
        Dim s As Long: s = InStr(r, "[[")
        Dim e As Long: e = InStr(s, r, "]]")
        If e > s Then
            r = Left(r, s - 1) & Mid(r, s + 2, e - s - 2) & Mid(r, e + 2)
        Else
            Exit Do
        End If
    Loop
    CleanWikilinks = r
End Function

Function RemoveBoldMarkers(text As String) As String
    RemoveBoldMarkers = Replace(text, "**", "")
End Function

' ================================================================
' Array Helpers
' ================================================================
Function CollectionToArray(col As Collection) As Variant
    If col.Count = 0 Then CollectionToArray = Empty: Exit Function
    Dim arr() As Variant: ReDim arr(0 To col.Count - 1)
    Dim i As Long
    For i = 0 To col.Count - 1: arr(i) = col(i + 1): Next i
    CollectionToArray = arr
End Function

Sub SortArray(ByRef arr As Variant)
    If IsEmpty(arr) Or Not IsArray(arr) Then Exit Sub
    If UBound(arr) - LBound(arr) < 1 Then Exit Sub
    Dim i As Long, j As Long, tmp As Variant
    For i = LBound(arr) To UBound(arr) - 1
        For j = i + 1 To UBound(arr)
            If arr(i) > arr(j) Then tmp = arr(i): arr(i) = arr(j): arr(j) = tmp
        Next j
    Next i
End Sub

' ================================================================
' DEBUG SYSTEM
' ================================================================
Sub RunSystemDebug()
    Application.ScreenUpdating = False

    Dim wsDB As Worksheet
    On Error Resume Next: Set wsDB = ThisWorkbook.Sheets("DebugLog"): On Error GoTo 0
    If wsDB Is Nothing Then
        Set wsDB = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        wsDB.Name = "DebugLog"
    End If

    wsDB.Cells.Clear
    With wsDB.Cells
        .Interior.Color = RGB(0, 0, 0)
        .Font.Color = RGB(200, 200, 200)
        .Font.Name = "Consolas"
        .Font.Size = 9
    End With
    wsDB.Activate
    ActiveWindow.DisplayGridlines = False
    wsDB.Columns("A").ColumnWidth = 22
    wsDB.Columns("B").ColumnWidth = 45
    wsDB.Columns("C").ColumnWidth = 18
    wsDB.Columns("D").ColumnWidth = 30

    With wsDB.Cells(1, 1)
        .Value = "SYSTEM DEBUG LOG | " & Format(Now, "yyyy/m/d hh:mm:ss")
        .Font.Color = RGB(255, 192, 0)
        .Font.Bold = True
        .Font.Size = 13
    End With
    wsDB.Rows(1).RowHeight = 26
    With wsDB.Range(wsDB.Cells(2, 1), wsDB.Cells(2, 4))
        .Interior.Color = RGB(255, 192, 0)
        .RowHeight = 3
    End With

    Dim wr As Long: wr = 4

    Call DB_Header(wsDB, wr, "CHECK", "ITEM", "STATUS", "DETAIL")
    wr = wr + 1

    ' 1. Worksheet Existence
    Call DB_Section(wsDB, wr, "WORKSHEET EXISTENCE")
    wr = wr + 1

    Dim sheetNames As Variant
    sheetNames = Array("RR4", "Transactions", "Realized", "HistoryLog", _
                       "Analysis", "HoldingsCorr", "Correlation", _
                       "DrawdownChart", "DebugLog")
    Dim sn As Variant
    For Each sn In sheetNames
        Dim wsCheck As Worksheet
        On Error Resume Next: Set wsCheck = ThisWorkbook.Sheets(CStr(sn)): On Error GoTo 0
        If wsCheck Is Nothing Then
            Call DB_Row(wsDB, wr, "Sheet", CStr(sn), "MISSING", "Sheet not found", RGB(255, 80, 80))
        Else
            Call DB_Row(wsDB, wr, "Sheet", CStr(sn), "OK", "Exists", RGB(0, 210, 100))
        End If
        wr = wr + 1
        Set wsCheck = Nothing
    Next sn

    ' 2. Transactions data check
    Call DB_Section(wsDB, wr, "TRANSACTIONS DATA")
    wr = wr + 1

    Dim wsTr As Worksheet
    On Error Resume Next: Set wsTr = ThisWorkbook.Sheets("Transactions"): On Error GoTo 0
    If Not wsTr Is Nothing Then
        Dim trLR As Long: trLR = wsTr.Cells(wsTr.Rows.Count, "A").End(xlUp).Row
        Dim trCount As Long: trCount = trLR - 1

        Call DB_Row(wsDB, wr, "Transactions", "Total rows", "INFO", trCount & " records", RGB(255, 192, 0))
        wr = wr + 1

        Dim blankDate As Long, blankTicker As Long, blankAction As Long, blankAmt As Long
        Dim tr As Long
        For tr = 2 To trLR
            If Trim(CStr(wsTr.Cells(tr, 2).Value)) = "" Then blankDate = blankDate + 1
            If Trim(CStr(wsTr.Cells(tr, 3).Value)) = "" Then blankTicker = blankTicker + 1
            If Trim(CStr(wsTr.Cells(tr, 4).Value)) = "" Then blankAction = blankAction + 1
            If Not IsNumeric(wsTr.Cells(tr, 9).Value) Then blankAmt = blankAmt + 1
        Next tr

        Call DB_Row(wsDB, wr, "Transactions", "Blank dates (Col B)", _
            IIf(blankDate = 0, "OK", "WARN"), blankDate & " rows", _
            IIf(blankDate = 0, RGB(0, 210, 100), RGB(255, 140, 0)))
        wr = wr + 1
        Call DB_Row(wsDB, wr, "Transactions", "Blank Ticker (Col C)", _
            IIf(blankTicker = 0, "OK", "WARN"), blankTicker & " rows", _
            IIf(blankTicker = 0, RGB(0, 210, 100), RGB(255, 140, 0)))
        wr = wr + 1
        Call DB_Row(wsDB, wr, "Transactions", "Blank Action (Col D)", _
            IIf(blankAction = 0, "OK", "WARN"), blankAction & " rows", _
            IIf(blankAction = 0, RGB(0, 210, 100), RGB(255, 140, 0)))
        wr = wr + 1
        Call DB_Row(wsDB, wr, "Transactions", "Non-numeric NetAmt (Col I)", _
            IIf(blankAmt = 0, "OK", "WARN"), blankAmt & " rows", _
            IIf(blankAmt = 0, RGB(0, 210, 100), RGB(255, 140, 0)))
        wr = wr + 1

        Dim invalidAction As Long
        For tr = 2 To trLR
            Dim act As String: act = UCase(Trim(CStr(wsTr.Cells(tr, 4).Value)))
            If act <> "BUY" And act <> "SELL" And act <> "ADJUSTCOST" And act <> "" Then
                invalidAction = invalidAction + 1
            End If
        Next tr
        Call DB_Row(wsDB, wr, "Transactions", "Invalid Action", _
            IIf(invalidAction = 0, "OK", "ERROR"), invalidAction & " rows (must be BUY/SELL/ADJUSTCOST)", _
            IIf(invalidAction = 0, RGB(0, 210, 100), RGB(255, 80, 80)))
        wr = wr + 1
    End If

    ' 3. HistoryLog update check
    Call DB_Section(wsDB, wr, "HISTORYLOG UPDATE CHECK")
    wr = wr + 1

    Dim wsH As Worksheet
    On Error Resume Next: Set wsH = ThisWorkbook.Sheets("HistoryLog"): On Error GoTo 0
    If Not wsH Is Nothing Then
        Dim hlLR As Long: hlLR = wsH.Cells(wsH.Rows.Count, "A").End(xlUp).Row
        Dim hlCount As Long: hlCount = hlLR - 1
        Call DB_Row(wsDB, wr, "HistoryLog", "Total rows", "INFO", hlCount & " records", RGB(255, 192, 0))
        wr = wr + 1

        If hlLR >= 2 Then
            Dim lastLogTime As Variant: lastLogTime = wsH.Cells(hlLR, "A").Value
            Dim lastLogMkt As Variant: lastLogMkt = wsH.Cells(hlLR, "B").Value

            Dim timeDiff As Double
            If IsDate(lastLogTime) Then
                timeDiff = Now - CDate(lastLogTime)
                Dim diffMin As Long: diffMin = CLng(timeDiff * 1440)
                Dim timeStatus As String
                Dim timeColor As Long
                If diffMin < 5 Then
                    timeStatus = "FRESH": timeColor = RGB(0, 210, 100)
                ElseIf diffMin < 60 Then
                    timeStatus = "OK": timeColor = RGB(0, 210, 100)
                ElseIf diffMin < 480 Then
                    timeStatus = "STALE": timeColor = RGB(255, 140, 0)
                Else
                    timeStatus = "OLD": timeColor = RGB(255, 80, 80)
                End If
                Call DB_Row(wsDB, wr, "HistoryLog", "Last update time", timeStatus, _
                    Format(CDate(lastLogTime), "m/d hh:mm") & " (" & diffMin & " min ago)", timeColor)
            Else
                Call DB_Row(wsDB, wr, "HistoryLog", "Last update time", "ERROR", "Cannot parse date", RGB(255, 80, 80))
            End If
            wr = wr + 1

            If IsNumeric(lastLogMkt) And CDbl(lastLogMkt) > 0 Then
                Call DB_Row(wsDB, wr, "HistoryLog", "Last market value (Col B)", "OK", _
                    Format(CDbl(lastLogMkt), "#,##0") & " TWD", RGB(0, 210, 100))
            Else
                Call DB_Row(wsDB, wr, "HistoryLog", "Last market value (Col B)", "ERROR", "Zero or non-numeric", RGB(255, 80, 80))
            End If
            wr = wr + 1

            Dim mLR As Long: mLR = wsH.Cells(wsH.Rows.Count, "M").End(xlUp).Row
            Dim mCount As Long: mCount = mLR - 1
            Call DB_Row(wsDB, wr, "HistoryLog", "M-col data count (Port Ret)", _
                IIf(mCount > 0, "OK", "ERROR"), mCount & " rows", _
                IIf(mCount > 0, RGB(0, 210, 100), RGB(255, 80, 80)))
            wr = wr + 1

            Dim todayLogged As Boolean: todayLogged = False
            If IsDate(lastLogTime) Then
                If Int(CDate(lastLogTime)) = Date Then todayLogged = True
            End If
            Call DB_Row(wsDB, wr, "HistoryLog", "Logged today?", _
                IIf(todayLogged, "OK", "WARN"), _
                IIf(todayLogged, "Logged today", "Not yet logged today"), _
                IIf(todayLogged, RGB(0, 210, 100), RGB(255, 140, 0)))
            wr = wr + 1
        End If
    End If

    ' 4. RR4 Dashboard Check
    Call DB_Section(wsDB, wr, "RR4 DASHBOARD CHECK")
    wr = wr + 1

    Dim wsP As Worksheet
    On Error Resume Next: Set wsP = ThisWorkbook.Sheets("RR4"): On Error GoTo 0
    If Not wsP Is Nothing Then
        Dim exRate As Double: exRate = 0
        On Error Resume Next: exRate = CDbl(wsP.Range("C22").Value): On Error GoTo 0
        Call DB_Row(wsDB, wr, "RR4", "FX rate C22 (USD/TWD)", _
            IIf(exRate > 20 And exRate < 50, "OK", "WARN"), _
            IIf(exRate > 0, Format(exRate, "0.00"), "No value"), _
            IIf(exRate > 20 And exRate < 50, RGB(0, 210, 100), RGB(255, 140, 0)))
        wr = wr + 1

        Dim totalMkt As Double: totalMkt = 0
        On Error Resume Next: totalMkt = CDbl(wsP.Cells(1, 1).Value): On Error GoTo 0
        Call DB_Row(wsDB, wr, "RR4", "Total market value (Cell A1)", _
            IIf(totalMkt > 0, "OK", "ERROR"), _
            IIf(totalMkt > 0, Format(totalMkt, "#,##0") & " TWD", "No value or zero"), _
            IIf(totalMkt > 0, RGB(0, 210, 100), RGB(255, 80, 80)))
        wr = wr + 1

        Dim posLR As Long: posLR = wsP.Cells(wsP.Rows.Count, 1).End(xlUp).Row
        Dim posCount As Long: posCount = posLR - 4
        Call DB_Row(wsDB, wr, "RR4", "Position count (Row 5+)", _
            IIf(posCount > 0, "OK", "WARN"), posCount & " positions", _
            IIf(posCount > 0, RGB(0, 210, 100), RGB(255, 140, 0)))
        wr = wr + 1

        Dim incDate As Date
        Dim stCap As Double
        On Error Resume Next
        incDate = CDate(wsP.Range("C2").Value)
        stCap = CDbl(wsP.Range("E2").Value)
        On Error GoTo 0
        Call DB_Row(wsDB, wr, "RR4", "InceptionDate (named range)", _
            IIf(incDate > DateSerial(2000, 1, 1), "OK", "ERROR"), _
            IIf(incDate > DateSerial(2000, 1, 1), Format(incDate, "yyyy/m/d"), "Not set"), _
            IIf(incDate > DateSerial(2000, 1, 1), RGB(0, 210, 100), RGB(255, 80, 80)))
        wr = wr + 1
        Call DB_Row(wsDB, wr, "RR4", "StartingCapital (named range)", _
            IIf(stCap > 0, "OK", "ERROR"), _
            IIf(stCap > 0, Format(stCap, "#,##0") & " TWD", "Not set"), _
            IIf(stCap > 0, RGB(0, 210, 100), RGB(255, 80, 80)))
        wr = wr + 1
    End If

    ' 5. Yahoo Finance API Test
    Call DB_Section(wsDB, wr, "YAHOO FINANCE API TEST")
    wr = wr + 1

    Dim testTickers As Variant
    testTickers = Array("SPY", "QQQ", "2330.TW", "^TWII")
    Dim tt As Variant
    For Each tt In testTickers
        Dim httpT As Object
        On Error Resume Next
        Set httpT = CreateObject("WinHttp.WinHttpRequest.5.1")
        On Error GoTo 0

        Dim apiURL As String
        apiURL = "https://query1.finance.yahoo.com/v8/finance/chart/" & _
                 CStr(tt) & "?interval=1d&range=5d"

        Dim apiResp As String: apiResp = ""
        On Error Resume Next
        With httpT
            .Open "GET", apiURL, False
            .setRequestHeader "User-Agent", "Mozilla/5.0"
            .setRequestHeader "Accept", "application/json"
            .send
            apiResp = .responseText
        End With
        On Error GoTo 0

        Dim apiPx As Double: apiPx = 0
        Dim csPos As Long
        csPos = InStr(apiResp, """close"":[")
        If csPos > 0 Then
            Dim csS As Long: csS = csPos + 9
            Dim csE As Long: csE = InStr(csS, apiResp, "]")
            If csE > csS Then
                Dim csStr As String: csStr = Mid(apiResp, csS, csE - csS)
                Dim csParts() As String: csParts = Split(csStr, ",")
                Dim cpIdx As Long
                For cpIdx = UBound(csParts) To 0 Step -1
                    Dim cpTrim As String: cpTrim = Trim(csParts(cpIdx))
                    If IsNumeric(cpTrim) And CDbl(cpTrim) > 0 Then
                        apiPx = CDbl(cpTrim)
                        Exit For
                    End If
                Next cpIdx
            End If
        End If

        Dim httpStatus As Long: httpStatus = 0
        On Error Resume Next: httpStatus = httpT.status: On Error GoTo 0

        Dim detailMsg As String
        If apiPx > 0 Then
            detailMsg = "Price = " & Format(apiPx, "#,##0.00") & " (HTTP " & httpStatus & ")"
        ElseIf httpStatus = 429 Then
            detailMsg = "HTTP 429 - Too Many Requests (rate limited)"
        ElseIf httpStatus = 401 Or httpStatus = 403 Then
            detailMsg = "HTTP " & httpStatus & " - Auth required / forbidden"
        ElseIf httpStatus = 0 Then
            detailMsg = "No response (network or firewall)"
        Else
            detailMsg = "HTTP " & httpStatus & " - Cannot parse response"
        End If

        Call DB_Row(wsDB, wr, "API Test", CStr(tt), _
            IIf(apiPx > 0, "OK", "FAIL"), detailMsg, _
            IIf(apiPx > 0, RGB(0, 210, 100), RGB(255, 80, 80)))
        wr = wr + 1
    Next tt

    ' 6. Module function existence check
    Call DB_Section(wsDB, wr, "MODULE FUNCTION CHECK")
    wr = wr + 1

    Dim funcNames As Variant
    funcNames = Array("GetStockPrice", "GetCompanyName", "GetCurrencyType", _
                      "GetHistoryPrices", "GetTechData", "CorrPearson", _
                      "CorrGetCommonDates", "BuildPositions", "CalculateRealizedPnL")
    Dim fn As Variant
    For Each fn In funcNames
        Dim fnExists As Boolean: fnExists = False
        On Error Resume Next
        Dim testCall As Variant
        testCall = Application.Run(CStr(fn))
        If Err.Number = 0 Or Err.Number = 1004 Or Err.Number = 450 Or Err.Number = 13 Then
            fnExists = True
        ElseIf Err.Number = 1000 Or Err.Number = 35 Then
            fnExists = False
        Else
            fnExists = True
        End If
        Err.Clear
        On Error GoTo 0
        Call DB_Row(wsDB, wr, "Function", CStr(fn), _
            IIf(fnExists, "OK", "MISSING"), _
            IIf(fnExists, "Callable", "Function not found, check module"), _
            IIf(fnExists, RGB(0, 210, 100), RGB(255, 80, 80)))
        wr = wr + 1
    Next fn

    ' 7. Health Summary
    wr = wr + 1
    Call DB_Section(wsDB, wr, "HEALTH SUMMARY")
    wr = wr + 1

    Dim okCount As Long, warnCount As Long, errCount As Long
    Dim scanR As Long
    For scanR = 4 To wr
        Dim statusVal As String: statusVal = CStr(wsDB.Cells(scanR, 3).Value)
        Select Case statusVal
        Case "OK", "FRESH": okCount = okCount + 1
        Case "WARN", "STALE": warnCount = warnCount + 1
        Case "ERROR", "FAIL", "MISSING", "OLD": errCount = errCount + 1
        End Select
    Next scanR

    Dim healthScore As Long
    If okCount + warnCount + errCount > 0 Then
        healthScore = CLng(okCount / (okCount + warnCount + errCount) * 100)
    End If

    Dim scoreColor As Long
    If healthScore >= 90 Then
        scoreColor = RGB(0, 210, 100)
    ElseIf healthScore >= 70 Then
        scoreColor = RGB(255, 192, 0)
    Else
        scoreColor = RGB(255, 80, 80)
    End If

    With wsDB.Cells(wr, 1)
        .Value = "HEALTH SCORE"
        .Font.Color = RGB(255, 192, 0)
        .Font.Bold = True
    End With
    With wsDB.Cells(wr, 2)
        .Value = healthScore & " / 100"
        .Font.Color = scoreColor
        .Font.Bold = True
        .Font.Size = 14
    End With
    With wsDB.Cells(wr, 3)
        .Value = "OK:" & okCount & " WARN:" & warnCount & " ERR:" & errCount
        .Font.Color = RGB(200, 200, 200)
    End With

    wsDB.Columns("A:D").AutoFit
    wsDB.Activate
    wsDB.Cells(1, 1).Select

    Application.ScreenUpdating = True
    Application.StatusBar = "Debug complete - Health: " & healthScore & "/100"
    MsgBox "Debug complete. Health score: " & healthScore & "/100" & vbCr & _
           "OK: " & okCount & " WARN: " & warnCount & " ERROR: " & errCount, vbInformation
End Sub

' ----------------------------------------------------------------
Private Sub DB_Section(ws As Worksheet, r As Long, title As String)
    With ws.Range(ws.Cells(r, 1), ws.Cells(r, 4))
        .Interior.Color = RGB(20, 20, 40)
        .Font.Color = RGB(100, 160, 255)
        .Font.Bold = True
    End With
    ws.Cells(r, 1).Value = "## " & title
    ws.Rows(r).RowHeight = 18
End Sub

' ----------------------------------------------------------------
Private Sub DB_Header(ws As Worksheet, r As Long, _
                      c1 As String, c2 As String, c3 As String, c4 As String)
    With ws.Range(ws.Cells(r, 1), ws.Cells(r, 4))
        .Interior.Color = RGB(10, 10, 10)
        .Font.Color = RGB(255, 192, 0)
        .Font.Bold = True
    End With
    ws.Cells(r, 1).Value = c1
    ws.Cells(r, 2).Value = c2
    ws.Cells(r, 3).Value = c3
    ws.Cells(r, 4).Value = c4
    ws.Rows(r).RowHeight = 16
End Sub

' ----------------------------------------------------------------
Private Sub DB_Row(ws As Worksheet, r As Long, _
                   category As String, item As String, _
                   status As String, detail As String, statusColor As Long)
    ws.Rows(r).RowHeight = 16
    With ws.Cells(r, 1)
        .Value = category
        .Font.Color = RGB(150, 150, 150)
        .Interior.Color = IIf(r Mod 2 = 0, RGB(8, 8, 8), RGB(12, 12, 12))
    End With
    With ws.Cells(r, 2)
        .Value = item
        .Font.Color = RGB(200, 200, 200)
        .Interior.Color = IIf(r Mod 2 = 0, RGB(8, 8, 8), RGB(12, 12, 12))
    End With
    With ws.Cells(r, 3)
        .Value = status
        .Font.Color = statusColor
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .Interior.Color = IIf(r Mod 2 = 0, RGB(8, 8, 8), RGB(12, 12, 12))
    End With
    With ws.Cells(r, 4)
        .Value = detail
        .Font.Color = RGB(180, 180, 180)
        .Interior.Color = IIf(r Mod 2 = 0, RGB(8, 8, 8), RGB(12, 12, 12))
    End With
End Sub
```

## 2. CommodityTermStructure.bas

`VB_Name = "CommodityTermStructure"` · 1433 行 · 55.8 KB · 原始編碼 UTF-8

<details><summary><b>成員索引</b></summary>

| 行 | 型別 | 可見性 | 宣告 |
|---:|---|---|---|
| 33 | Sub | Public | `Sub RefreshDashboard()` |
| 200 | Sub | Public | `Sub RefreshDashboardV2()` |
| 204 | Sub | Private | `Private Sub RefreshSpreadMatrix(wsM As Worksheet, wsTick As Worksheet, wsCfg As Worksheet, sec1 As Long, sec2Data As Long)` |
| 391 | Sub | Private | `Private Sub ApplyMatrixCellTheme(cell As Range)` |
| 410 | Sub | Public | `Sub ApplyDarkTheme()` |
| 624 | Sub | Private | `Private Sub UnmergeBand(wsM As Worksheet, rTop As Long, rBot As Long)` |
| 631 | Sub | Private | `Private Sub PaintHeaderRow(wsM As Worksheet, rowIdx As Long, labels As Variant)` |
| 652 | Function | Private | `Private Function FetchPrice(symbol As String) As Double` |
| 676 | Function | Private | `Private Function FetchPriceWithFallback(symbol As String, baseSym As String) As Double` |
| 699 | Function | Private | `Private Function IsIndexFutureBase(baseSym As String) As Boolean` |
| 713 | Function | Private | `Private Function FrontMonthYahooSymbol(baseSym As String) As String` |
| 724 | Function | Private | `Private Function RowHasValidPrices(wsM As Worksheet, r As Long) As Boolean` |
| 734 | Sub | Public | `Sub FixAndRebuild()` |
| 793 | Function | Private | `Private Function FetchYahoo(symbol As String) As Double` |
| 850 | Function | Private | `Private Function ParseYahooPrice(response As String) As Double` |
| 890 | Sub | Public | `Sub FixTickerMap()` |
| 960 | Function | Public | `Function MonthCode(m As Integer) As String` |
| 966 | Function | Public | `Function FuturesTicker(baseSymbol As String, exchange As String, monthsAhead As Long) As String` |
| 980 | Function | Public | `Function GetExchange(commodity As String) As String` |
| 996 | Function | Public | `Function QuarterlyMonthCode(m As Integer) As String` |
| 1006 | Function | Public | `Function FuturesTickerQuarterly(baseSymbol As String, exchange As String, monthsAhead As Long) As String` |
| 1016 | Function | Public | `Function IsQuarterlyOnly(baseSymbol As String) As Boolean` |
| 1028 | Sub | Public | `Sub RebuildTickerMapDynamic()` |
| 1105 | Sub | Public | `Sub RebuildTickerMapV3()` |
| 1186 | Sub | Public | `Sub SyncLayoutFromTickerMap()` |
| 1263 | Sub | Private | `Private Sub WritePrice(cell As Range, price As Double)` |
| 1284 | Sub | Private | `Private Sub WriteSpread(cell As Range, spd As Double)` |
| 1307 | Sub | Private | `Private Sub ApplyStructure(cell As Range, pM1 As Double, pM12 As Double)` |
| 1333 | Sub | Private | `Private Sub MarkError(ws As Worksheet, r As Long)` |
| 1364 | Function | Private | `Private Function GetCellPrice(cell As Range) As Double` |
| 1371 | Sub | Private | `Private Sub PauseMS(ms As Long)` |
| 1376 | Sub | Public | `Sub ExportSnapshot()` |

</details>

```vba
Option Explicit

' ================================================================
'  COMMODITY TERM STRUCTURE — Bloomberg-Style Dashboard
'  v3.0  |  Dark Theme  |  Section 1: Curve Signal  |  Section 2: Spread Matrix
' ================================================================

' ── Sheet / layout constants ───────────────────────────────────
Private Const SH_MAIN    As String = "BBG_COMM"
Private Const SH_TICK    As String = "TickerMap"
Private Const SH_CFG     As String = "_Config"
Private Const MAX_RETRY  As Integer = 3

' ── Dashboard layout columns (BBG_COMM) ────────────────────────
Private Const COL_FIRST      As Long = 1
Private Const COL_LAST       As Long = 13  ' STRUCTURE column
Private Const COL_PREVCACHE  As Long = 14  ' hidden cache for prior-session M1
Private Const COL_UNUSED_GAP As Long = 12  ' visually reserved blank column

' ── Theme tokens (pure-black canvas, monospace) ────────────────
Private Const THEME_FONT      As String = "Consolas"
Private Const THEME_FONT_SIZE As Long = 10
Private Const C_BG            As Long = 0  ' RGB(0,0,0) - pure black canvas

' ================================================================
'  ENTRY POINT - refresh full dashboard
'  - TickerMap layout (V3): A=Instrument | B=Base | C=M1 | D=M2 | E=M3 | F=M6 | G=M12
'  - Computes % change vs previous close (cached in hidden column 14)
'  - Calls RefreshSpreadMatrix and writes timestamp
'  - Per-row error guard: a single failed fetch does not abort the run
' ================================================================
Sub RefreshDashboard()
    Dim wsM    As Worksheet
    Dim wsTick As Worksheet
    Dim wsCfg  As Worksheet
    Dim sec1   As Long, instCnt As Long, sec2Data As Long
    Dim i      As Long
    Const PREV_COL As Integer = 14   ' hidden cache for prior-session M1 close

    Set wsM = ThisWorkbook.Sheets(SH_MAIN)
    Set wsTick = ThisWorkbook.Sheets(SH_TICK)
    Set wsCfg = ThisWorkbook.Sheets(SH_CFG)

    sec1 = wsCfg.Range("B1").Value
    instCnt = wsCfg.Range("B2").Value
    sec2Data = wsCfg.Range("B3").Value

    Application.ScreenUpdating = False
    Application.StatusBar = "Refreshing Commodity Term Structure..."
    wsM.Columns(PREV_COL).Hidden = True

    Dim okCount As Long, errCount As Long
    okCount = 0: errCount = 0

    For i = 1 To instCnt
        Dim dataRow As Long
        dataRow = sec1 + i - 1

        Application.StatusBar = "Fetching [" & i & "/" & instCnt & "]  " & _
                                  wsM.Cells(dataRow, 2).Value

        ' Per-row error guard so a single failed fetch cannot abort the whole run.
        On Error GoTo RowFail

        ' V3 TickerMap layout: C..G = M1, M2, M3, M6, M12.
        Dim tM1  As String, tM2  As String, tM3  As String
        Dim tM6  As String, tM12 As String
        tM1 = Trim(wsTick.Cells(i + 1, 3).Value)
        tM2 = Trim(wsTick.Cells(i + 1, 4).Value)
        tM3 = Trim(wsTick.Cells(i + 1, 5).Value)
        tM6 = Trim(wsTick.Cells(i + 1, 6).Value)
        tM12 = Trim(wsTick.Cells(i + 1, 7).Value)

        Dim baseSym As String
        baseSym = Trim(wsTick.Cells(i + 1, 2).Value)

        Dim pM1 As Double, pM2 As Double, pM3 As Double
        Dim pM6 As Double, pM12 As Double
        pM1 = FetchPriceWithFallback(tM1, baseSym)
        pM2 = FetchPriceWithFallback(tM2, baseSym)
        pM3 = FetchPriceWithFallback(tM3, baseSym)
        pM6 = FetchPriceWithFallback(tM6, baseSym)
        pM12 = FetchPriceWithFallback(tM12, baseSym)

        If pM1 > 0 Then
            ' % change vs cached prior M1 close.
            Dim prevClose As Double: prevClose = 0
            If IsNumeric(wsM.Cells(dataRow, PREV_COL).Value) Then
                prevClose = CDbl(wsM.Cells(dataRow, PREV_COL).Value)
            End If
            With wsM.Cells(dataRow, 3)
                .Font.Name = THEME_FONT
                .Font.Size = THEME_FONT_SIZE
                .Interior.Color = C_BG
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlCenter
                If prevClose > 0 Then
                    Dim chg As Double: chg = (pM1 - prevClose) / prevClose
                    .Value = chg
                    .NumberFormat = "+0.00%;-0.00%;0.00%"
                    .Font.Bold = True
                    If chg > 0 Then
                        .Font.Color = RGB(0, 255, 136)
                    ElseIf chg < 0 Then
                        .Font.Color = RGB(255, 51, 51)
                    Else
                        .Font.Color = RGB(255, 255, 0)
                    End If
                Else
                    .Value = "new"
                    .Font.Color = RGB(136, 136, 136)
                End If
            End With
            wsM.Cells(dataRow, PREV_COL).Value = pM1

            ' Prices
            WritePrice wsM.Cells(dataRow, 4), pM1
            WritePrice wsM.Cells(dataRow, 5), pM2
            WritePrice wsM.Cells(dataRow, 6), pM3
            WritePrice wsM.Cells(dataRow, 7), pM6
            WritePrice wsM.Cells(dataRow, 8), pM12

            ' Spreads (guard against zero far-month price -> blank, not bogus spread)
            If pM2 > 0 Then WriteSpread wsM.Cells(dataRow, 9), pM1 - pM2 _
                       Else wsM.Cells(dataRow, 9).Value = "-"
            If pM6 > 0 Then WriteSpread wsM.Cells(dataRow, 10), pM1 - pM6 _
                       Else wsM.Cells(dataRow, 10).Value = "-"
            If pM12 > 0 Then WriteSpread wsM.Cells(dataRow, 11), pM1 - pM12 _
                        Else wsM.Cells(dataRow, 11).Value = "-"

            ' Structure: prefer M12, fall back to M6 then M2 when far months unavailable.
            Dim farPrice As Double
            If pM12 > 0 Then
                farPrice = pM12
            ElseIf pM6 > 0 Then
                farPrice = pM6
            Else
                farPrice = pM2
            End If
            If farPrice > 0 Then ApplyStructure wsM.Cells(dataRow, 13), pM1, farPrice

            okCount = okCount + 1
        Else
            MarkError wsM, dataRow
            errCount = errCount + 1
        End If

        On Error GoTo 0
        GoTo RowDone
RowFail:
        MarkError wsM, dataRow
        errCount = errCount + 1
        On Error GoTo 0
RowDone:
    Next i

    ' Wipe stray rows between the last Section-1 instrument and the
    ' Section-2 title band. Stop at sec2Data - 3 so we never touch:
    '   sec2Data - 2  (Section 2 title row, merged A:M)
    '   sec2Data - 1  (Section 2 header row)
    ' Unmerge each row first so the wipe cannot trip "Cannot change
    ' part of a merged cell" (Excel error 1004).
    Dim wipeRow As Long, cc As Integer
    For wipeRow = sec1 + instCnt To sec2Data - 3
        On Error Resume Next
        wsM.Range(wsM.Cells(wipeRow, 1), wsM.Cells(wipeRow, COL_LAST)).UnMerge
        On Error GoTo 0
        For cc = 1 To 13
            With wsM.Cells(wipeRow, cc)
                .ClearContents
                .Interior.Color = C_BG
                .Font.Color = C_BG
                .Font.Name = THEME_FONT
                .Font.Size = THEME_FONT_SIZE
            End With
        Next cc
    Next wipeRow

    ' Section 2: Calendar Spread Matrix (reads selected row from _Config!B5).
    Call RefreshSpreadMatrix(wsM, wsTick, wsCfg, sec1, sec2Data)

    ' Timestamp: spread matrix occupies 6 rows -> first free row is sec2Data + 6.
    Dim tsRow As Long
    tsRow = sec2Data + 6
    With wsM.Cells(tsRow, 2)
        .Value = "Last updated: " & Format(Now(), "yyyy/mm/dd  hh:mm:ss")
        .Font.Name = THEME_FONT
        .Font.Size = 8
        .Font.Color = RGB(100, 100, 100)
        .Interior.Color = C_BG
        .HorizontalAlignment = xlLeft
    End With

    Application.ScreenUpdating = True
    Application.StatusBar = "Dashboard refreshed - " & okCount & " OK / " & errCount & " errors"
End Sub

' Backward-compat alias for any toolbar button still pointing at V2.
Sub RefreshDashboardV2()
    Call RefreshDashboard
End Sub

Private Sub RefreshSpreadMatrix(wsM As Worksheet, wsTick As Worksheet, _
                                 wsCfg As Worksheet, sec1 As Long, sec2Data As Long)
    ' Resolve which Section-1 instrument feeds the spread matrix.
    ' _Config!B5 stores a 1-based offset within Section 1 (1 = first instrument).
    ' Defaults to 1 if missing / out of range / non-numeric.
    Dim instCnt As Long
    instCnt = CLng(wsCfg.Range("B2").Value)
    Dim selOffset As Long: selOffset = 1
    If IsNumeric(wsCfg.Range("B5").Value) Then
        selOffset = CLng(wsCfg.Range("B5").Value)
        If selOffset < 1 Or selOffset > instCnt Then selOffset = 1
    End If
    Dim selRow As Long: selRow = sec1 + selOffset - 1

    ' Robustness: if the user-selected row has no usable M1 price, scan
    ' forward from sec1 for the first row that does. Avoids the matrix
    ' silently displaying all-zero spreads when _Config is stale.
    If Not RowHasValidPrices(wsM, selRow) Then
        Dim probe As Long
        For probe = sec1 To sec1 + instCnt - 1
            If RowHasValidPrices(wsM, probe) Then
                selRow = probe
                Exit For
            End If
        Next probe
    End If

    ' Echo the selected instrument name into the section header so the user
    ' can see which curve is being decomposed.
    On Error Resume Next
    With wsM.Cells(sec2Data - 1, 7)
        .Value = "Selected: " & wsM.Cells(selRow, 2).Value
        .Font.Name = THEME_FONT
        .Font.Size = THEME_FONT_SIZE - 1
        .Font.Color = RGB(180, 180, 180)
        .Font.Italic = True
        .HorizontalAlignment = xlCenter
    End With
    On Error GoTo 0

    ' Read M1..M12 prices (BBG_COMM columns 4..8).
    Dim prices(1 To 5) As Double
    Dim c As Integer
    For c = 1 To 5
        Dim v As Variant: v = wsM.Cells(selRow, 3 + c).Value
        If IsNumeric(v) Then
            If CDbl(v) > 0 Then prices(c) = CDbl(v) Else prices(c) = 0
        Else
            prices(c) = 0
        End If
    Next c

    ' Populate the TICKERS column (col 3 in the matrix) from the
    ' TickerMap row that corresponds to selRow.
    Dim tickRow As Long: tickRow = (selRow - sec1) + 2   ' +2: header row at 1, data starts at 2
    Dim labels As Variant
    labels = Array("M1-M2", "M1-M3", "M1-M6", "M1-M12", "M2-M3", "M3-M6")
    Dim tickCols As Variant
    ' Column indexes in TickerMap (V3 layout): C=M1, D=M2, E=M3, F=M6, G=M12
    tickCols = Array(Array(3, 4), Array(3, 5), Array(3, 6), Array(3, 7), Array(4, 5), Array(5, 6))
    Dim jj As Integer
    For jj = 0 To 5
        Dim ta As String, tb As String
        On Error Resume Next
        ta = CStr(wsTick.Cells(tickRow, tickCols(jj)(0)).Value)
        tb = CStr(wsTick.Cells(tickRow, tickCols(jj)(1)).Value)
        On Error GoTo 0
        With wsM.Cells(sec2Data + jj, 3)
            .Value = ta & " / " & tb
            .Font.Name = THEME_FONT
            .Font.Size = 8
            .Font.Color = RGB(140, 140, 140)
            .Interior.Color = C_BG
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
        End With
    Next jj

    Dim spreads(1 To 6) As Double
    spreads(1) = prices(1) - prices(2)
    spreads(2) = prices(1) - prices(3)
    spreads(3) = prices(1) - prices(4)
    spreads(4) = prices(1) - prices(5)
    spreads(5) = prices(2) - prices(3)
    spreads(6) = prices(3) - prices(4)

    Dim j As Integer
    For j = 1 To 6
        Dim r As Long: r = sec2Data + j - 1

        ' �w�w ���Ѱ��ӦC�Ҧ��X�֡A�קK 1004 �w�w�w�w�w�w�w�w�w�w�w�w�w�w
        On Error Resume Next
        wsM.Rows(r).UnMerge
        On Error GoTo 0

        Dim spd As Double: spd = spreads(j)

        ' Value (col 4)
        WriteSpread wsM.Cells(r, 4), spd

        ' Direction (col 5)
        With wsM.Cells(r, 5)
            ApplyMatrixCellTheme wsM.Cells(r, 5)
            .Font.Bold = True
            If spd > 0.005 Then
                .Value = "[+] BKWD"
                .Font.Color = RGB(0, 255, 136)
            ElseIf spd < -0.005 Then
                .Value = "[-] CNTNG"
                .Font.Color = RGB(255, 51, 51)
            Else
                .Value = "[=] FLAT"
                .Font.Color = RGB(255, 255, 0)
            End If
        End With

        ' Strength (col 6)
        Dim pctSpd As Double
        If prices(1) > 0 Then pctSpd = Abs(spd) / prices(1) Else pctSpd = 0
        With wsM.Cells(r, 6)
            ApplyMatrixCellTheme wsM.Cells(r, 6)
            .Font.Bold = True
            If pctSpd > 0.05 Then
                .Value = "STRONG": .Font.Color = RGB(0, 255, 136)
            ElseIf pctSpd > 0.01 Then
                .Value = "MODERATE": .Font.Color = RGB(255, 192, 0)
            Else
                .Value = "WEAK": .Font.Color = RGB(136, 136, 136)
            End If
        End With

        ' Interpretation (col 7)
        With wsM.Cells(r, 7)
            ApplyMatrixCellTheme wsM.Cells(r, 7)
            .Font.Bold = True
            If spd > 0.005 Then
                .Value = "TIGHT SUPPLY / STRONG SPOT DEMAND"
                .Font.Color = RGB(0, 255, 136)
            ElseIf spd < -0.005 Then
                .Value = "STORAGE SAT / FULL CARRY"
                .Font.Color = RGB(255, 51, 51)
            Else
                .Value = "MARKET IN EQUILIBRIUM"
                .Font.Color = RGB(255, 255, 0)
            End If
        End With

        ' Roll Yield % (col 8)
        With wsM.Cells(r, 8)
            ApplyMatrixCellTheme wsM.Cells(r, 8)
            .NumberFormat = "0.00%"
            .Font.Bold = True
            If prices(1) > 0 Then
                .Value = spd / prices(1)
                If spd >= 0 Then
                    .Font.Color = RGB(0, 255, 136)
                Else
                    .Font.Color = RGB(255, 51, 51)
                End If
            Else
                .Value = "-"
                .Font.Color = RGB(136, 136, 136)
            End If
        End With

        ' Signal Flag (col 13)
        With wsM.Cells(r, 13)
            ApplyMatrixCellTheme wsM.Cells(r, 13)
            .Font.Bold = True
            If Abs(spd) > 5 Then
                .Value = ">> WATCH": .Font.Color = RGB(255, 51, 51)
            ElseIf Abs(spd) > 1 Then
                .Value = "!! SHARP MOVE": .Font.Color = RGB(255, 192, 0)
            Else
                .Value = "OK STABLE": .Font.Color = RGB(136, 136, 136)
            End If
        End With

        ' Reserved gap col 12 stays clean.
        With wsM.Cells(r, COL_UNUSED_GAP)
            .ClearContents
            .Interior.Color = C_BG
        End With
    Next j
End Sub

' Shared per-cell theme for Section-2 matrix cells (Consolas / centered / black bg).
Private Sub ApplyMatrixCellTheme(cell As Range)
    cell.Font.Name = THEME_FONT
    cell.Font.Size = THEME_FONT_SIZE
    cell.Interior.Color = C_BG
    cell.HorizontalAlignment = xlCenter
    cell.VerticalAlignment = xlCenter
End Sub

' ================================================================
'  ApplyDarkTheme
'  One-shot UI sweep for BBG_COMM:
'    - Paint the entire used range pure black
'    - Set every cell font to Consolas
'    - Re-render Section 1 / Section 2 title bands (merged across A:M)
'    - Re-render legend band at top (rows 2-3)
'    - Center the instrument column (B) left, headers center
'    - Hide the prev-close cache column (14)
'  Run this once after any layout change, or simply re-run; it is idempotent.
' ================================================================
Sub ApplyDarkTheme()
    Dim wsM   As Worksheet: Set wsM = ThisWorkbook.Sheets(SH_MAIN)
    Dim wsCfg As Worksheet: Set wsCfg = ThisWorkbook.Sheets(SH_CFG)

    Dim sec1     As Long: sec1 = CLng(wsCfg.Range("B1").Value)
    Dim instCnt  As Long: instCnt = CLng(wsCfg.Range("B2").Value)
    Dim sec2Data As Long: sec2Data = CLng(wsCfg.Range("B3").Value)

    Application.ScreenUpdating = False

    ' Kill default Excel gridlines on this sheet so the pure-black canvas
    ' has no light-grey artefact lines bleeding through.
    On Error Resume Next
    wsM.Activate
    ActiveWindow.DisplayGridlines = False
    ActiveWindow.DisplayHeadings = True
    On Error GoTo 0

    ' --- 1. Paint everything pure black + Consolas across the dashboard band.
    Dim lastRow As Long: lastRow = sec2Data + 8
    Dim canvas As Range
    Set canvas = wsM.Range(wsM.Cells(1, 1), wsM.Cells(lastRow, COL_LAST))
    With canvas
        .Interior.Color = C_BG
        .Font.Name = THEME_FONT
        .Font.Size = THEME_FONT_SIZE
        .Font.Color = RGB(220, 220, 220)
        .VerticalAlignment = xlCenter
        ' Strip any leftover borders so frame color is fully managed by us.
        .Borders.LineStyle = xlNone
    End With

    ' Whole-sheet sweep so neighbouring cells (units, helper text) also get Consolas.
    With wsM.Cells
        .Font.Name = THEME_FONT
    End With

    ' --- 2. Legend band (rows 2-3). Merge B:M (column A stays as gutter).
    UnmergeBand wsM, 2, 3
    With wsM.Range(wsM.Cells(2, 2), wsM.Cells(2, COL_LAST))
        .Merge
        .Value = "BACKWARDATION  ->  Supply tightness. No storage incentive. Physical scarcity drives spot premium over forward."
        .Font.Name = THEME_FONT
        .Font.Color = RGB(0, 255, 136)
        .Font.Italic = True
        .Font.Size = THEME_FONT_SIZE
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .IndentLevel = 1
        .Borders.LineStyle = xlNone
    End With
    With wsM.Range(wsM.Cells(3, 2), wsM.Cells(3, COL_LAST))
        .Merge
        .Value = "CONTANGO       ->  Oversupply. Storage economical - buy spot cheap, sell forward at premium. Physical demand weak."
        .Font.Name = THEME_FONT
        .Font.Color = RGB(255, 51, 51)
        .Font.Italic = True
        .Font.Size = THEME_FONT_SIZE
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .IndentLevel = 1
        .Borders.LineStyle = xlNone
    End With

    ' --- 3. Section 1 title band (row sec1 - 2) and header row (sec1 - 1).
    '        Title merges B:M (column A stays as gutter).
    Dim sec1Title As Long: sec1Title = sec1 - 2
    Dim sec1Head  As Long: sec1Head = sec1 - 1
    If sec1Title >= 4 Then
        UnmergeBand wsM, sec1Title, sec1Title
        With wsM.Range(wsM.Cells(sec1Title, 2), wsM.Cells(sec1Title, COL_LAST))
            .Merge
            .Value = "< SECTION 1 >   CURVE STRUCTURE SIGNAL   -   FRONT MONTH vs DEFERRED MONTHS"
            .Font.Name = THEME_FONT
            .Font.Color = RGB(220, 220, 220)
            .Font.Bold = True
            .Font.Size = THEME_FONT_SIZE + 1
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .Interior.Color = RGB(20, 20, 20)
            .Borders.LineStyle = xlNone
            .Borders(xlEdgeBottom).LineStyle = xlContinuous
            .Borders(xlEdgeBottom).Color = RGB(60, 60, 60)
            .Borders(xlEdgeBottom).Weight = xlThin
        End With
    End If
    PaintHeaderRow wsM, sec1Head, _
        Array("", "INSTRUMENT", "% Change", "M1", "M2", "M3", "M6", "M12", _
              "M1-M2 SPRD", "M1-M6 SPRD", "M1-M12 SPRD", "", "STRUCTURE")

    ' --- 4. Section 1 instrument column (B) left-align bold, dim units.
    Dim i As Long
    For i = 0 To instCnt - 1
        With wsM.Cells(sec1 + i, 2)
            .Font.Name = THEME_FONT
            .Font.Size = THEME_FONT_SIZE
            .Font.Bold = True
            .Font.Color = RGB(220, 220, 220)
            .HorizontalAlignment = xlLeft
            .VerticalAlignment = xlCenter
            .IndentLevel = 1
        End With
    Next i

    ' --- 5. Section 2 title + header. Title merges B:M (column A stays gutter).
    Dim sec2Title As Long: sec2Title = sec2Data - 2
    Dim sec2Head  As Long: sec2Head = sec2Data - 1
    If sec2Title >= sec1 + instCnt Then
        UnmergeBand wsM, sec2Title, sec2Title
        With wsM.Range(wsM.Cells(sec2Title, 2), wsM.Cells(sec2Title, COL_LAST))
            .Merge
            .Value = "< SECTION 2 >   CALENDAR SPREAD MATRIX   -   SELECT INSTRUMENT VIA _Config!B5"
            .Font.Name = THEME_FONT
            .Font.Color = RGB(220, 220, 220)
            .Font.Bold = True
            .Font.Size = THEME_FONT_SIZE + 1
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .Interior.Color = RGB(20, 20, 20)
            .Borders.LineStyle = xlNone
            .Borders(xlEdgeBottom).LineStyle = xlContinuous
            .Borders(xlEdgeBottom).Color = RGB(60, 60, 60)
            .Borders(xlEdgeBottom).Weight = xlThin
        End With
    End If
    PaintHeaderRow wsM, sec2Head, _
        Array("", "SPREAD PERIOD", "TICKERS", "VALUE", "DIRECTION", "STRENGTH", _
              "INTERPRETATION", "ROLL YIELD %", "PRIOR WEEK", "WoW CHANGE", "MOMENTUM", "", "SIGNAL FLAG")

    ' --- 6. Section 2 first column (SPREAD PERIOD labels).
    Dim sLabels As Variant
    sLabels = Array("M1 - M2  [ 1mo ]", "M1 - M3  [ 2mo ]", _
                    "M1 - M6  [ 5mo ]", "M1 - M12 [11mo ]", _
                    "M2 - M3  [ roll ]", "M3 - M6  [ roll ]")
    Dim k As Long
    For k = 0 To 5
        With wsM.Cells(sec2Data + k, 2)
            .Value = sLabels(k)
            .Font.Name = THEME_FONT
            .Font.Size = THEME_FONT_SIZE
            .Font.Bold = True
            .Font.Color = RGB(220, 220, 220)
            .HorizontalAlignment = xlLeft
            .IndentLevel = 1
        End With
    Next k

    ' --- 7. Unified column widths sized for the LONGEST content per column.
    '   A : 3   - left gutter (visible margin so row-2/3/5 merges read as "from col B")
    '   B : 26  - INSTRUMENT names  / SPREAD PERIOD labels
    '   C : 22  - TICKERS pair "CLM26.NYM / CLM26.NYM" (Section 2) / % Change (Section 1)
    '   D : 12  - M1  / VALUE
    '   E : 13  - M2  / DIRECTION  ("[+] BKWD")
    '   F : 11  - M3  / STRENGTH   ("MODERATE")
    '   G : 36  - M6  / INTERPRETATION ("TIGHT SUPPLY / STRONG SPOT DEMAND" = 33)
    '   H : 13  - M12 / ROLL YIELD %
    '   I : 13  - M1-M2 SPRD / PRIOR WEEK
    '   J : 13  - M1-M6 SPRD / WoW CHANGE
    '   K : 13  - M1-M12 SPRD / MOMENTUM
    '   L : 3   - reserved gap
    '   M : 22  - STRUCTURE ("BACKWARDATION") / SIGNAL FLAG ("!! SHARP MOVE")
    wsM.Columns(1).ColumnWidth = 3
    wsM.Columns(2).ColumnWidth = 26
    wsM.Columns(3).ColumnWidth = 22
    wsM.Columns(4).ColumnWidth = 12
    wsM.Columns(5).ColumnWidth = 13
    wsM.Columns(6).ColumnWidth = 11
    wsM.Columns(7).ColumnWidth = 36
    wsM.Columns(8).ColumnWidth = 13
    wsM.Columns(9).ColumnWidth = 13
    wsM.Columns(10).ColumnWidth = 13
    wsM.Columns(11).ColumnWidth = 13
    wsM.Columns(COL_UNUSED_GAP).ColumnWidth = 3
    wsM.Columns(COL_LAST).ColumnWidth = 22
    wsM.Columns(COL_PREVCACHE).Hidden = True

    ' --- 7b. Uniform row heights. Rows 21-23 came out short because
    ' SyncLayoutFromTickerMap appended new instrument rows without resetting
    ' the row height. Normalise every data row and the title bands.
    Dim rh As Long
    For rh = sec1 To sec1 + instCnt - 1
        wsM.Rows(rh).RowHeight = 20
    Next rh
    For rh = sec2Data To sec2Data + 5
        wsM.Rows(rh).RowHeight = 20
    Next rh
    wsM.Rows(2).RowHeight = 18
    wsM.Rows(3).RowHeight = 18
    If sec1Title >= 4 Then wsM.Rows(sec1Title).RowHeight = 24
    If sec2Title >= sec1 + instCnt Then wsM.Rows(sec2Title).RowHeight = 24
    wsM.Rows(sec1Head).RowHeight = 22
    wsM.Rows(sec2Head).RowHeight = 22

    ' --- 8. Freeze panes at first data row of Section 1 so headers stay visible.
    On Error Resume Next
    wsM.Activate
    wsM.Cells(sec1, 1).Select
    ActiveWindow.FreezePanes = False
    ActiveWindow.FreezePanes = True
    On Error GoTo 0

    ' --- 9. Sheet tab color: Bloomberg-style amber so BBG_COMM stands out
    ' in the workbook tab bar.
    On Error Resume Next
    wsM.Tab.Color = RGB(230, 170, 40)
    On Error GoTo 0

    Application.ScreenUpdating = True
    MsgBox "Dark theme applied. Run RefreshDashboard to repopulate prices.", vbInformation
End Sub

' Helper - unmerge a horizontal band before re-merging (idempotent).
' Sweeps a wide range so legacy merges that include column A or extend past
' COL_LAST are also broken before the new merge is applied.
Private Sub UnmergeBand(wsM As Worksheet, rTop As Long, rBot As Long)
    On Error Resume Next
    wsM.Range(wsM.Cells(rTop, 1), wsM.Cells(rBot, COL_LAST + 4)).UnMerge
    On Error GoTo 0
End Sub

' Helper - paint a header row from an array of labels (COL_LAST cells wide).
Private Sub PaintHeaderRow(wsM As Worksheet, rowIdx As Long, labels As Variant)
    Dim c As Long
    For c = 1 To COL_LAST
        With wsM.Cells(rowIdx, c)
            .Value = labels(c - 1)
            .Font.Name = THEME_FONT
            .Font.Size = THEME_FONT_SIZE
            .Font.Bold = True
            .Font.Color = RGB(180, 180, 180)
            .Interior.Color = RGB(25, 25, 25)
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .Borders(xlEdgeBottom).LineStyle = xlContinuous
            .Borders(xlEdgeBottom).Color = RGB(80, 80, 80)
            .Borders(xlEdgeBottom).Weight = xlThin
        End With
    Next c
End Sub
' ================================================================
'  Price fetcher with dual-endpoint fallback + retry
' ================================================================
Private Function FetchPrice(symbol As String) As Double
    On Error GoTo Bail
    If symbol = "" Then FetchPrice = 0: Exit Function
    Dim attempt As Integer, result As Double
    For attempt = 1 To MAX_RETRY
        result = FetchYahoo(symbol)
        If result > 0 Then FetchPrice = result: Exit Function
        Call PauseMS(600 * attempt)
    Next attempt
    FetchPrice = 0
    Exit Function
Bail:
    FetchPrice = 0
End Function

' Tries the per-month ticker first, then falls back to the front-month
' continuous contract for this base if the monthly symbol returns nothing.
' Yahoo's coverage of CME monthly tickers (e.g. CLM26.NYM, ESH26.CME) is
' inconsistent, so this fallback keeps the dashboard populated rather than
' showing ERR for every far-month contract.
'
' Equity index futures (ES/NQ/YM/RTY/VX) are short-circuited: Yahoo only
' carries the continuous symbol for these, so we skip the monthly lookup
' entirely to avoid wasting 3 retries per cell on a guaranteed 404.
Private Function FetchPriceWithFallback(symbol As String, baseSym As String) As Double
    If IsIndexFutureBase(baseSym) Then
        FetchPriceWithFallback = FetchPrice(FrontMonthYahooSymbol(baseSym))
        Exit Function
    End If

    Dim p As Double
    p = FetchPrice(symbol)
    If p > 0 Then FetchPriceWithFallback = p: Exit Function

    If LenB(baseSym) > 0 Then
        Dim alt As String: alt = FrontMonthYahooSymbol(baseSym)
        ' Avoid an infinite loop if caller already passed the same form.
        If StrComp(alt, symbol, vbTextCompare) <> 0 Then
            p = FetchPrice(alt)
            If p > 0 Then FetchPriceWithFallback = p: Exit Function
        End If
    End If
    FetchPriceWithFallback = 0
End Function

' Whitelist of equity index futures whose individual-month contracts are
' NOT carried by Yahoo Finance. Treat as front-month-only.
Private Function IsIndexFutureBase(baseSym As String) As Boolean
    Select Case UCase(Trim(baseSym))
        Case "ES", "NQ", "YM", "RTY", "VX"
            IsIndexFutureBase = True
        Case Else
            IsIndexFutureBase = False
    End Select
End Function

' Resolves the Yahoo symbol that actually serves data for this base's
' front-month / spot quote. Most futures use <base>=F, but a handful of
' instruments need a special route:
'   VX  -> ^VIX  (Yahoo has no VX=F; ^VIX is the VIX spot index, which is
'                 the closest publicly available proxy for VX1 front-month)
Private Function FrontMonthYahooSymbol(baseSym As String) As String
    Select Case UCase(Trim(baseSym))
        Case "VX"
            FrontMonthYahooSymbol = "^VIX"
        Case Else
            FrontMonthYahooSymbol = baseSym & "=F"
    End Select
End Function

' True when col 4 (M1 price) is a positive number on a Section-1 row.
' Used by RefreshSpreadMatrix to skip rows that failed to fetch.
Private Function RowHasValidPrices(wsM As Worksheet, r As Long) As Boolean
    On Error GoTo Bail
    Dim v As Variant: v = wsM.Cells(r, 4).Value
    If IsNumeric(v) Then
        If CDbl(v) > 0 Then RowHasValidPrices = True: Exit Function
    End If
Bail:
    RowHasValidPrices = False
End Function

Sub FixAndRebuild()
    ' �w�w �B�J1�G���� _Config �ƭ� �w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w
    Dim wsCfg As Worksheet
    On Error Resume Next
    Set wsCfg = ThisWorkbook.Sheets("_Config")
    On Error GoTo 0
    If wsCfg Is Nothing Then
        Set wsCfg = ThisWorkbook.Sheets.Add
        wsCfg.Name = "_Config"
        wsCfg.Visible = xlSheetHidden
    End If
    
    ' �۰ʰ��� BBG_COMM �u�@������ڵ��c
    Dim wsM As Worksheet
    Set wsM = ThisWorkbook.Sheets("BBG_COMM")
    
    ' ���y�� Section1 �_�l�C�]��uINSTRUMENT�v���D�^
    Dim sec1 As Long, i As Long
    For i = 1 To 50
        If InStr(wsM.Cells(i, 2).Value, "INSTRUMENT") > 0 Then
            sec1 = i + 1   ' ��Ʊq���D�U�@�C�}�l
            Exit For
        End If
    Next i
    If sec1 = 0 Then sec1 = 7   ' fallback �w�]��
    
    ' ���y�� Section2 �_�l�C�]��uSPREAD PERIOD�v���D�^
    Dim sec2Hdr As Long, sec2Data As Long
    For i = sec1 To 100
        If InStr(wsM.Cells(i, 2).Value, "SPREAD PERIOD") > 0 Then
            sec2Hdr = i
            sec2Data = i + 1
            Exit For
        End If
    Next i
    If sec2Data = 0 Then sec2Data = sec1 + 20   ' fallback
    
    ' �p��ӫ~�ƶq�]sec1 �� sec2Hdr ���������e���C�^
    Dim instCnt As Long
    For i = sec1 To sec2Hdr - 1
        If Trim(wsM.Cells(i, 2).Value) <> "" Then instCnt = instCnt + 1
    Next i
    If instCnt = 0 Then instCnt = 14   ' fallback
    
    ' �g�J _Config
    wsCfg.Range("A1") = "SEC1_START": wsCfg.Range("B1") = sec1
    wsCfg.Range("A2") = "INST_COUNT": wsCfg.Range("B2") = instCnt
    wsCfg.Range("A3") = "SEC2_DATA_START": wsCfg.Range("B3") = sec2Data
    ' Spread-matrix instrument selector (1-based offset within Section 1).
    If Not IsNumeric(wsCfg.Range("B5").Value) Then
        wsCfg.Range("A5") = "SpreadMatrix_Offset": wsCfg.Range("B5") = 1
    End If
    
    MsgBox "? _Config �w�״_�I" & vbCr & vbCr & _
           "sec1      = " & sec1 & vbCr & _
           "instCnt   = " & instCnt & vbCr & _
           "sec2Data  = " & sec2Data & vbCr & vbCr & _
           "�{�b�i�H���� RefreshDashboard", vbInformation
End Sub
Private Function FetchYahoo(symbol As String) As Double
    ' Full error trap: any HTTP / parse failure returns 0 instead of bubbling
    ' up to the caller. A fresh MSXML2.XMLHTTP is created per endpoint so a
    ' stuck state on attempt #1 cannot poison attempt #2.
    On Error GoTo Bail

    ' URL-encode the few characters that appear in Yahoo symbols but break
    ' raw URL parsing in MSXML. Most important: ^VIX -> %5EVIX.
    Dim safeSym As String: safeSym = symbol
    safeSym = Replace(safeSym, "^", "%5E")

    Dim eps(1) As String
    eps(0) = "https://query1.finance.yahoo.com/v8/finance/chart/" & safeSym & "?interval=1d&range=1d"
    eps(1) = "https://query2.finance.yahoo.com/v8/finance/chart/" & safeSym & "?interval=1d&range=1d"

    Dim ep As Integer
    For ep = 0 To 1
        Dim http As Object
        Set http = Nothing
        Set http = CreateObject("MSXML2.XMLHTTP")

        Dim sent As Boolean: sent = False
        On Error Resume Next
        http.Open "GET", eps(ep), False
        http.setRequestHeader "User-Agent", "Mozilla/5.0 (Windows NT 10.0)"
        http.send
        sent = (Err.Number = 0)
        On Error GoTo Bail

        If sent Then
            Dim status As Long: status = 0
            On Error Resume Next
            status = http.status
            On Error GoTo Bail

            If status = 200 Then
                Dim response As String: response = ""
                On Error Resume Next
                response = http.responseText
                On Error GoTo Bail

                Dim parsed As Double
                parsed = ParseYahooPrice(response)
                If parsed > 0 Then FetchYahoo = parsed: Exit Function
            End If
        End If
        Set http = Nothing
    Next ep

    FetchYahoo = 0
    Exit Function
Bail:
    FetchYahoo = 0
End Function

' Pulls regularMarketPrice from the Yahoo chart JSON, with a fallback to
' the last numeric value in the close[] array.
Private Function ParseYahooPrice(response As String) As Double
    On Error GoTo Bail
    If LenB(response) = 0 Then ParseYahooPrice = 0: Exit Function

    Dim pos As Long
    pos = InStr(response, """regularMarketPrice"":")
    If pos > 0 Then
        Dim tmp As String
        tmp = Mid(response, pos + 22, 20)
        tmp = Split(tmp, ",")(0)
        tmp = Split(tmp, "}")(0)
        tmp = Trim(tmp)
        If IsNumeric(tmp) Then
            If CDbl(tmp) > 0 Then ParseYahooPrice = CDbl(tmp): Exit Function
        End If
    End If

    pos = InStr(response, """close"":[")
    If pos > 0 Then
        Dim arr As String
        arr = Mid(response, pos + 9, 300)
        arr = Split(arr, "]")(0)
        Dim parts() As String: parts = Split(arr, ",")
        ' Walk from the tail backwards, skipping nulls that Yahoo sometimes
        ' emits for the most recent (still-open) bar.
        Dim k As Long
        For k = UBound(parts) To LBound(parts) Step -1
            Dim cand As String: cand = Trim(parts(k))
            If IsNumeric(cand) Then
                If CDbl(cand) > 0 Then ParseYahooPrice = CDbl(cand): Exit Function
            End If
        Next k
    End If

    ParseYahooPrice = 0
    Exit Function
Bail:
    ParseYahooPrice = 0
End Function

Sub FixTickerMap()
    Dim wsTick As Worksheet
    On Error Resume Next
    Set wsTick = ThisWorkbook.Sheets("TickerMap")
    On Error GoTo 0
    If wsTick Is Nothing Then
        Set wsTick = ThisWorkbook.Sheets.Add
        wsTick.Name = "TickerMap"
        wsTick.Visible = xlSheetHidden
    End If
    
    wsTick.Cells.Clear
    
    ' ���D
    Dim h As Variant
    Dim headers As Variant
    headers = Array("Instrument", "M1", "M2", "M3", "M6", "M12")
    Dim ci As Integer
    For ci = 0 To 5
        wsTick.Cells(1, ci + 1).Value = headers(ci)
    Next ci
    
    ' �w�w Yahoo Finance �s��X���u�� =F�]���^�A
    '    ����� +1/+2/+5/+11 �몺����N�X
    '    �榡�GCLK25.NYM, CLM25.NYM... ��ڤW Yahoo ���䴩
    '    �̥i�a���k�G�����ΦP�@�� =F �N�X�A�� VBA ������
    '    ����ȥ� Quandl/CME �榡�A�Χ�� M2 spread �Ϊ�����
    
    Dim data As Variant
    data = Array( _
        Array("WTI CRUDE OIL", "CL=F", "CL=F", "CL=F", "CL=F", "CL=F"), _
        Array("BRENT CRUDE", "BZ=F", "BZ=F", "BZ=F", "BZ=F", "BZ=F"), _
        Array("RBOB GASOLINE", "RB=F", "RB=F", "RB=F", "RB=F", "RB=F"), _
        Array("ULSD DIESEL", "HO=F", "HO=F", "HO=F", "HO=F", "HO=F"), _
        Array("NATURAL GAS", "NG=F", "NG=F", "NG=F", "NG=F", "NG=F"), _
        Array("GOLD", "GC=F", "GC=F", "GC=F", "GC=F", "GC=F"), _
        Array("COPPER", "HG=F", "HG=F", "HG=F", "HG=F", "HG=F"), _
        Array("SILVER", "SI=F", "SI=F", "SI=F", "SI=F", "SI=F"), _
        Array("CORN", "ZC=F", "ZC=F", "ZC=F", "ZC=F", "ZC=F"), _
        Array("SOYBEANS", "ZS=F", "ZS=F", "ZS=F", "ZS=F", "ZS=F"), _
        Array("WHEAT", "ZW=F", "ZW=F", "ZW=F", "ZW=F", "ZW=F"), _
        Array("ALUMINUM", "ALI=F", "ALI=F", "ALI=F", "ALI=F", "ALI=F"), _
        Array("LUMBER", "LBS=F", "LBS=F", "LBS=F", "LBS=F", "LBS=F"), _
        Array("STEEL HRC", "HR=F", "HR=F", "HR=F", "HR=F", "HR=F"), _
        Array("S&P 500 E-MINI",  "ES=F",  "ES=F",  "ES=F",  "ES=F",  "ES=F"), _
        Array("NASDAQ-100 EMINI", "NQ=F", "NQ=F",  "NQ=F",  "NQ=F",  "NQ=F"), _
        Array("DOW E-MINI",      "YM=F",  "YM=F",  "YM=F",  "YM=F",  "YM=F"), _
        Array("RUSSELL 2000",    "RTY=F", "RTY=F", "RTY=F", "RTY=F", "RTY=F"), _
        Array("VIX",             "VX=F",  "VX=F",  "VX=F",  "VX=F",  "VX=F") _
    )
    
    Dim r As Integer, col As Integer
    For r = 0 To UBound(data)
        For col = 0 To 5
            wsTick.Cells(r + 2, col + 1).Value = data(r)(col)
        Next col
    Next r
    
    MsgBox "? TickerMap �w��s�I�@ " & UBound(data) + 1 & " �Ӱӫ~�C" & vbCr & _
           "�{�b���� RefreshDashboard", vbInformation
End Sub

' ================================================================
'  �ʺA�p��U��X���N�X + % Change �ץ�
'  ���N�쥻�� FixTickerMap �M RefreshDashboard ��������q��
' ================================================================

' �w�w ����N�X��Ӫ� �w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w
' Yahoo Finance ���f����GF=1�� G=2 H=3 J=4 K=5 M=6
'                         N=7�� Q=8 U=9 V=10 X=11 Z=12
Function MonthCode(m As Integer) As String
    Dim codes As String: codes = "FGHJKMNQUVXZ"
    MonthCode = Mid(codes, m, 1)
End Function

' �w�w �p��� n �Ӥ�᪺�X���N�X �w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w
Function FuturesTicker(baseSymbol As String, exchange As String, _
                       monthsAhead As Long) As String
    Dim targetDate As Date
    targetDate = DateAdd("m", monthsAhead, Date)
    
    Dim y As Integer: y = Year(targetDate)
    Dim m As Integer: m = Month(targetDate)
    
    ' �榡�GCLK26.NYM
    FuturesTicker = baseSymbol & MonthCode(m) & Right(CStr(y), 2) & exchange
End Function

' �w�w �U�ӫ~������ҫ�� �w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w�w
' .NYM = NYMEX  .CMX = COMEX  .CBT = CBOT
Function GetExchange(commodity As String) As String
    Select Case commodity
        Case "CL", "RB", "HO", "NG": GetExchange = ".NYM"
        Case "BZ": GetExchange = ".NYM"                        ' Brent �P NYMEX
        Case "GC", "SI", "HG", "ALI": GetExchange = ".CMX"
        Case "HR": GetExchange = ".CMX"
        Case "ZC", "ZS", "ZW": GetExchange = ".CBT"
        ' --- Equity index futures ---
        Case "ES", "NQ", "RTY": GetExchange = ".CME"           ' E-mini S&P / Nasdaq / Russell
        Case "YM": GetExchange = ".CBT"                        ' E-mini Dow on CBOT
        Case "VX": GetExchange = ".CFE"                        ' VIX on CFE
        Case Else: GetExchange = ".CMX"
    End Select
End Function

' Index futures only roll on quarterly cycle (H/M/U/Z = Mar/Jun/Sep/Dec)
Function QuarterlyMonthCode(m As Integer) As String
    Select Case m
        Case 1, 2, 3:    QuarterlyMonthCode = "H"   ' March
        Case 4, 5, 6:    QuarterlyMonthCode = "M"   ' June
        Case 7, 8, 9:    QuarterlyMonthCode = "U"   ' September
        Case 10, 11, 12: QuarterlyMonthCode = "Z"   ' December
    End Select
End Function

' Snap requested month-ahead forward to the NEXT quarterly contract.
Function FuturesTickerQuarterly(baseSymbol As String, exchange As String, _
                                 monthsAhead As Long) As String
    Dim targetDate As Date
    targetDate = DateAdd("m", monthsAhead, Date)
    Dim y As Integer: y = Year(targetDate)
    Dim m As Integer: m = Month(targetDate)
    FuturesTickerQuarterly = baseSymbol & QuarterlyMonthCode(m) & Right(CStr(y), 2) & exchange
End Function

' Returns True for symbols that only trade on the H/M/U/Z quarterly cycle.
Function IsQuarterlyOnly(baseSymbol As String) As Boolean
    Select Case baseSymbol
        Case "ES", "NQ", "YM", "RTY"
            IsQuarterlyOnly = True
        Case Else
            IsQuarterlyOnly = False
    End Select
End Function

' ================================================================
'  ���� TickerMap �X �ϥΰʺA����N�X
' ================================================================
Sub RebuildTickerMapDynamic()
    Dim wsTick As Worksheet
    On Error Resume Next
    Set wsTick = ThisWorkbook.Sheets("TickerMap")
    On Error GoTo 0
    If wsTick Is Nothing Then
        Set wsTick = ThisWorkbook.Sheets.Add
        wsTick.Name = "TickerMap"
        wsTick.Visible = xlSheetHidden
    End If
    wsTick.Cells.Clear
    
    ' ���D
    wsTick.Cells(1, 1) = "Instrument"
    wsTick.Cells(1, 2) = "Base"
    wsTick.Cells(1, 3) = "M1"
    wsTick.Cells(1, 4) = "M2"
    wsTick.Cells(1, 5) = "M3"
    wsTick.Cells(1, 6) = "M6"
    wsTick.Cells(1, 7) = "M12"
    
    ' [�ӫ~�W��, ��¦�N�X]
    Dim items As Variant
    items = Array( _
        Array("WTI CRUDE OIL", "CL"), _
        Array("BRENT CRUDE", "BZ"), _
        Array("RBOB GASOLINE", "RB"), _
        Array("ULSD DIESEL", "HO"), _
        Array("NATURAL GAS", "NG"), _
        Array("GOLD", "GC"), _
        Array("COPPER", "HG"), _
        Array("SILVER", "SI"), _
        Array("CORN", "ZC"), _
        Array("SOYBEANS", "ZS"), _
        Array("WHEAT", "ZW"), _
        Array("ALUMINUM", "ALI"), _
        Array("STEEL HRC", "HR") _
    )
    
    Dim offsets As Variant
    offsets = Array(1, 2, 3, 6, 12)  ' M1, M2, M3, M6, M12
    
    Dim i As Integer, j As Integer
    For i = 0 To UBound(items)
        Dim row As Integer: row = i + 2
        Dim base As String: base = items(i)(1)
        Dim exch As String: exch = GetExchange(base)
        
        wsTick.Cells(row, 1) = items(i)(0)
        wsTick.Cells(row, 2) = base
        
        For j = 0 To 4
            wsTick.Cells(row, j + 3) = FuturesTicker(base, exch, CInt(offsets(j)))
        Next j
    Next i
    
    ' ���G�w��
    Dim preview As String
    preview = "�ʺA�ͦ����X���N�X�]�H WTI ���ҡ^�G" & vbCr & vbCr
    For j = 0 To 4
        Dim mo As Variant: mo = Array("M1", "M2", "M3", "M6", "M12")
        preview = preview & mo(j) & " = " & wsTick.Cells(2, j + 3).Value & vbCr
    Next j
    
    MsgBox preview & vbCr & "�@ " & UBound(items) + 1 & " �Ӱӫ~�w��s�C", vbInformation
End Sub

' ================================================================
'  V3 TickerMap builder
'  - Uses Yahoo-style per-month tickers (<base><MonthCode><YY><.EXCH>)
'  - Honors quarterly-only contracts for equity index futures (ES/NQ/YM/RTY)
'  - Adds equity index futures (S&P 500, Nasdaq-100, Dow, Russell 2000, VIX)
'    forward curve alongside the commodity universe.
'  Layout (same as RebuildTickerMapDynamic):
'    A=Instrument | B=Base | C=M1 | D=M2 | E=M3 | F=M6 | G=M12
'  Use with RefreshDashboardV2.
' ================================================================
Sub RebuildTickerMapV3()
    Dim wsTick As Worksheet
    On Error Resume Next
    Set wsTick = ThisWorkbook.Sheets("TickerMap")
    On Error GoTo 0
    If wsTick Is Nothing Then
        Set wsTick = ThisWorkbook.Sheets.Add
        wsTick.Name = "TickerMap"
        wsTick.Visible = xlSheetHidden
    End If
    wsTick.Cells.Clear

    wsTick.Cells(1, 1) = "Instrument"
    wsTick.Cells(1, 2) = "Base"
    wsTick.Cells(1, 3) = "M1"
    wsTick.Cells(1, 4) = "M2"
    wsTick.Cells(1, 5) = "M3"
    wsTick.Cells(1, 6) = "M6"
    wsTick.Cells(1, 7) = "M12"

    ' [Instrument display name, Base symbol]
    Dim items As Variant
    items = Array( _
        Array("WTI CRUDE OIL",   "CL"), _
        Array("BRENT CRUDE",     "BZ"), _
        Array("RBOB GASOLINE",   "RB"), _
        Array("ULSD DIESEL",     "HO"), _
        Array("NATURAL GAS",     "NG"), _
        Array("GOLD",            "GC"), _
        Array("COPPER",          "HG"), _
        Array("SILVER",          "SI"), _
        Array("CORN",            "ZC"), _
        Array("SOYBEANS",        "ZS"), _
        Array("WHEAT",           "ZW"), _
        Array("ALUMINUM",        "ALI"), _
        Array("STEEL HRC",       "HR"), _
        Array("S&P 500 E-MINI",  "ES"), _
        Array("NASDAQ-100 EMINI", "NQ"), _
        Array("DOW E-MINI",      "YM"), _
        Array("RUSSELL 2000",    "RTY"), _
        Array("VIX",             "VX") _
    )

    Dim offsets As Variant
    offsets = Array(1, 2, 3, 6, 12)  ' M1, M2, M3, M6, M12

    Dim i As Integer, j As Integer
    For i = 0 To UBound(items)
        Dim r As Integer: r = i + 2
        Dim base As String: base = items(i)(1)
        Dim exch As String: exch = GetExchange(base)

        wsTick.Cells(r, 1) = items(i)(0)
        wsTick.Cells(r, 2) = base

        For j = 0 To 4
            If IsQuarterlyOnly(base) Then
                wsTick.Cells(r, j + 3) = FuturesTickerQuarterly(base, exch, CLng(offsets(j)))
            Else
                wsTick.Cells(r, j + 3) = FuturesTicker(base, exch, CLng(offsets(j)))
            End If
        Next j
    Next i

    MsgBox "TickerMap V3 rebuilt - " & UBound(items) + 1 & " instruments " & _
           "(commodities + equity index futures forward curve)." & vbCr & vbCr & _
           "WTI M1   = " & wsTick.Cells(2, 3).Value & vbCr & _
           "S&P M1   = " & wsTick.Cells(15, 3).Value & vbCr & _
           "S&P M6   = " & wsTick.Cells(15, 6).Value & vbCr & vbCr & _
           "Run SyncLayoutFromTickerMap then RefreshDashboard.", vbInformation
End Sub

' ================================================================
'  SyncLayoutFromTickerMap
'  - Reads every Instrument row from TickerMap and pushes it into
'    BBG_COMM Section 1 column B, then updates _Config!B2 (INST_COUNT).
'  - This is what makes newly added index futures actually appear in
'    the dashboard. Run once after RebuildTickerMapV3.
'  - If Section 1 is about to overflow into Section 2, shifts the
'    Section 2 anchor (_Config!B3) downward so they do not collide.
' ================================================================
Sub SyncLayoutFromTickerMap()
    Dim wsM    As Worksheet: Set wsM = ThisWorkbook.Sheets(SH_MAIN)
    Dim wsTick As Worksheet: Set wsTick = ThisWorkbook.Sheets(SH_TICK)
    Dim wsCfg  As Worksheet: Set wsCfg = ThisWorkbook.Sheets(SH_CFG)

    Dim sec1     As Long: sec1 = CLng(wsCfg.Range("B1").Value)
    Dim sec2Data As Long: sec2Data = CLng(wsCfg.Range("B3").Value)
    Dim oldCnt   As Long: oldCnt = CLng(wsCfg.Range("B2").Value)

    ' Count TickerMap data rows (skip header at row 1, stop on blank Instrument).
    Dim newCnt As Long: newCnt = 0
    Dim i As Long
    For i = 2 To 200
        If Trim(CStr(wsTick.Cells(i, 1).Value)) = "" Then Exit For
        newCnt = newCnt + 1
    Next i
    If newCnt = 0 Then
        MsgBox "TickerMap is empty - run RebuildTickerMapV3 first.", vbExclamation
        Exit Sub
    End If

    ' If the new instrument list runs past the existing Section 2 anchor,
    ' push Section 2 down (plus 1 blank gap row).
    If sec1 + newCnt > sec2Data Then
        Dim shift As Long: shift = (sec1 + newCnt + 1) - sec2Data
        sec2Data = sec2Data + shift
        wsCfg.Range("B3").Value = sec2Data
    End If

    ' Clear any rows beyond the new count (was 18 -> now 14 case).
    ' Unmerge each row first to survive previously merged title bands.
    If newCnt < oldCnt Then
        Dim wr As Long, cc As Integer
        For wr = sec1 + newCnt To sec1 + oldCnt - 1
            On Error Resume Next
            wsM.Range(wsM.Cells(wr, 1), wsM.Cells(wr, COL_LAST)).UnMerge
            On Error GoTo 0
            For cc = 1 To 13
                With wsM.Cells(wr, cc)
                    .ClearContents
                    .Interior.Color = C_BG
                    .Font.Color = C_BG
                    .Font.Name = THEME_FONT
                    .Font.Size = THEME_FONT_SIZE
                End With
            Next cc
        Next wr
    End If

    ' Write instrument names into BBG_COMM column B.
    For i = 1 To newCnt
        Dim r As Long: r = sec1 + i - 1
        Dim instName As String: instName = CStr(wsTick.Cells(i + 1, 1).Value)
        Dim baseSym  As String: baseSym = CStr(wsTick.Cells(i + 1, 2).Value)
        With wsM.Cells(r, 2)
            .Value = instName & "  [" & baseSym & "]"
            .Font.Name = THEME_FONT
            .Font.Size = THEME_FONT_SIZE
            .Font.Bold = True
            .Font.Color = RGB(220, 220, 220)
            .Interior.Color = C_BG
            .HorizontalAlignment = xlLeft
            .VerticalAlignment = xlCenter
            .IndentLevel = 1
        End With
    Next i

    wsCfg.Range("B2").Value = newCnt

    MsgBox "Layout synced: " & newCnt & " instruments (was " & oldCnt & ")." & vbCr & _
           "Section 2 anchor: row " & sec2Data & vbCr & vbCr & _
           "Now run RefreshDashboard.", vbInformation
End Sub

' ================================================================
'  Cell-writing helpers
' ================================================================
Private Sub WritePrice(cell As Range, price As Double)
    cell.Interior.Color = C_BG
    cell.Font.Name = THEME_FONT
    cell.Font.Size = THEME_FONT_SIZE
    cell.Font.Bold = False
    cell.HorizontalAlignment = xlRight
    cell.VerticalAlignment = xlCenter
    If price > 0 Then
        cell.NumberFormat = "#,##0.00"
        cell.Value = price
        cell.Font.Color = RGB(220, 220, 220)
    Else
        ' No usable quote for this contract month - show a soft dash so
        ' the user can distinguish "not available" from a real zero price.
        cell.NumberFormat = "@"
        cell.Value = "-"
        cell.Font.Color = RGB(100, 100, 100)
        cell.HorizontalAlignment = xlCenter
    End If
End Sub

Private Sub WriteSpread(cell As Range, spd As Double)
    ' Show "0.00" for a genuine zero so e.g. equity-index rows (where M1..M12
    ' all equal the front-month spot) read as "calculated and equal to zero"
    ' rather than the dash format which implies missing data.
    cell.NumberFormat = "#,##0.00;(#,##0.00);0.00"
    cell.Value = spd
    cell.Font.Name = THEME_FONT
    cell.Font.Size = THEME_FONT_SIZE
    cell.Font.Bold = True
    cell.HorizontalAlignment = xlCenter
    cell.VerticalAlignment = xlCenter
    If spd > 0.005 Then
        cell.Font.Color = RGB(0, 255, 136)
        cell.Interior.Color = RGB(0, 40, 20)
    ElseIf spd < -0.005 Then
        cell.Font.Color = RGB(255, 51, 51)
        cell.Interior.Color = RGB(40, 0, 0)
    Else
        cell.Font.Color = RGB(255, 255, 0)
        cell.Interior.Color = RGB(40, 40, 0)
    End If
End Sub

Private Sub ApplyStructure(cell As Range, pM1 As Double, pM12 As Double)
    Dim spd As Double: spd = pM1 - pM12
    cell.Font.Name = THEME_FONT
    cell.Font.Size = THEME_FONT_SIZE
    cell.Font.Bold = True
    cell.HorizontalAlignment = xlCenter
    cell.VerticalAlignment = xlCenter
    If spd > 0.005 Then
        cell.Value = "BACKWARDATION"
        cell.Font.Color = RGB(0, 255, 136)
        cell.Interior.Color = RGB(0, 51, 25)
    ElseIf spd < -0.005 Then
        cell.Value = "CONTANGO"
        cell.Font.Color = RGB(255, 51, 51)
        cell.Interior.Color = RGB(51, 0, 0)
    Else
        cell.Value = "MIXED / FLAT"
        cell.Font.Color = RGB(255, 255, 0)
        cell.Interior.Color = RGB(40, 40, 0)
    End If
End Sub

' Only marks a row as ERR when *every* price column is empty/zero.
' If at least one month came back valid, the prior WritePrice calls already
' rendered the cells correctly and we should not overwrite them.
' Skips COL_UNUSED_GAP (col 12) which is reserved as a visual divider.
Private Sub MarkError(ws As Worksheet, r As Long)
    Dim c As Integer
    For c = 4 To COL_LAST
        If c = COL_UNUSED_GAP Then GoTo NextCol
        Dim v As Variant: v = ws.Cells(r, c).Value
        Dim hasData As Boolean: hasData = False
        If IsNumeric(v) Then
            If CDbl(v) > 0 Then hasData = True
        End If
        If Not hasData Then
            With ws.Cells(r, c)
                .NumberFormat = "@"
                .Value = "ERR"
                .Font.Name = THEME_FONT
                .Font.Size = THEME_FONT_SIZE
                .Font.Bold = True
                .Font.Color = RGB(255, 102, 0)
                .Interior.Color = C_BG
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlCenter
            End With
        End If
NextCol:
    Next c
    ' Make sure the gap column stays clean.
    With ws.Cells(r, COL_UNUSED_GAP)
        .ClearContents
        .Interior.Color = C_BG
    End With
End Sub

Private Function GetCellPrice(cell As Range) As Double
    If IsNumeric(cell.Value) Then GetCellPrice = CDbl(cell.Value) Else GetCellPrice = 0
End Function

' ================================================================
'  Utilities
' ================================================================
Private Sub PauseMS(ms As Long)
    Dim t As Single: t = Timer + ms / 1000
    Do While Timer < t: DoEvents: Loop
End Sub

Sub ExportSnapshot()
    Dim ws As Worksheet:  Set ws = ThisWorkbook.Sheets(SH_MAIN)
    Dim wsCfg As Worksheet: Set wsCfg = ThisWorkbook.Sheets(SH_CFG)

    ' Resolve a writable output directory.
    ' ThisWorkbook.Path is empty when the workbook has never been saved -
    ' that produced the Open-failure ("\TermStructure_*.csv" has no drive).
    Dim outDir As String
    If LenB(ThisWorkbook.Path) > 0 Then
        outDir = ThisWorkbook.Path
    Else
        outDir = Environ$("USERPROFILE") & "\Desktop"
        If Len(Dir(outDir, vbDirectory)) = 0 Then outDir = Environ$("TEMP")
    End If

    Dim outPath As String
    outPath = outDir & "\TermStructure_" & Format(Now(), "yyyymmdd_hhmmss") & ".csv"

    ' Derive the export row range from _Config instead of hard-coding 60.
    Dim sec1     As Long: sec1 = CLng(wsCfg.Range("B1").Value)
    Dim instCnt  As Long: instCnt = CLng(wsCfg.Range("B2").Value)
    Dim sec2Data As Long: sec2Data = CLng(wsCfg.Range("B3").Value)
    Dim lastRow  As Long: lastRow = sec2Data + 6   ' include timestamp row

    Dim fNum As Integer: fNum = FreeFile

    On Error GoTo BadOpen
    Open outPath For Output As #fNum
    On Error GoTo 0

    Dim r As Long, c As Integer, line As String
    For r = 1 To lastRow
        line = ""
        For c = 2 To COL_LAST
            Dim v As String: v = CStr(ws.Cells(r, c).Value)
            ' RFC 4180: wrap in quotes when the cell contains a comma,
            ' embedded quote, or newline; escape internal quotes by doubling.
            If InStr(v, ",") > 0 Or InStr(v, """") > 0 Or InStr(v, vbLf) > 0 Then
                v = """" & Replace(v, """", """""") & """"
            End If
            line = line & IIf(c > 2, ",", "") & v
        Next c
        Print #fNum, line
    Next r
    Close #fNum

    MsgBox "Exported:" & vbCr & outPath, vbInformation
    Exit Sub

BadOpen:
    Dim errDesc As String: errDesc = Err.Description
    On Error Resume Next
    Close #fNum
    On Error GoTo 0
    MsgBox "Cannot write CSV:" & vbCr & outPath & vbCr & vbCr & _
           "Error " & Err.Number & ": " & errDesc, vbCritical
End Sub
```

## 3. CorelationMatrix.bas

`VB_Name = "CorelationMatrix"` · 309 行 · 11.9 KB · 原始編碼 CP950 (Big5)

<details><summary><b>成員索引</b></summary>

| 行 | 型別 | 可見性 | 宣告 |
|---:|---|---|---|
| 2 | Sub | Public | `Sub BuildCorrelationMatrix()` |
| 229 | Function | Private | `Private Function FetchPriceArray(Ticker As String, nDays As Long) As Double()` |
| 274 | Function | Private | `Private Function CalcCorrelation(retData() As Double, r1 As Integer, r2 As Integer, n As Long) As Double` |
| 300 | Sub | Private | `Private Sub Name_Safe(ws As Worksheet)` |
| 305 | Sub | Public | `Sub RefreshCorrelation()` |

</details>

```vba
Sub BuildCorrelationMatrix()
    Dim wsC As Worksheet
    On Error Resume Next
    Set wsC = ThisWorkbook.Sheets("Correlation")
    On Error GoTo 0
    If wsC Is Nothing Then
        Set wsC = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.count))
        wsC.Name = "Correlation"
    End If

    ' ── Layout constants ────────────────────────────────────────
    Dim MATRIX_ROW As Long: MATRIX_ROW = 2
    Dim MATRIX_COL As Long: MATRIX_COL = 4

    ' ── Read days BEFORE clearing (preserve user input) ─────────
    Dim nDays As Long: nDays = 60
    If IsNumeric(wsC.cells(2, 2).Value) And wsC.cells(2, 2).Value > 0 Then
        nDays = CLng(wsC.cells(2, 2).Value)
    End If

    ' ── Base style ───────────────────────────────────────────────
    wsC.cells.Clear
    With wsC.cells
        .Interior.Color = RGB(0, 0, 0)
        .Font.Color = RGB(200, 200, 200)
        .Font.Name = "Consolas"
        .Font.Size = 8
    End With
    wsC.Activate
    ActiveWindow.DisplayGridlines = False

    ' ── Title ────────────────────────────────────────────────────
    With wsC.cells(1, 1)
        .Value = "US SECTOR CORRELATION MATRIX"
        .Font.Color = RGB(255, 192, 0)
        .Font.Bold = True
        .Font.Size = 13
    End With

    ' ── Days label + input ───────────────────────────────────────
    wsC.cells(2, 1).Value = "DAYS:"
    wsC.cells(2, 1).Font.Color = RGB(150, 150, 150)
    wsC.cells(2, 1).Font.Bold = True
    With wsC.cells(2, 2)
        .Value = nDays
        .Interior.Color = RGB(30, 30, 0)
        .Font.Color = RGB(255, 192, 0)
        .Font.Bold = True
        .NumberFormat = "0"
        .HorizontalAlignment = xlCenter
    End With

    ' ── Sector ETF list ──────────────────────────────────────────
    Dim tickers As Variant
    Dim labels  As Variant
    tickers = Array( _
        "XLK", "XLF", "XLV", "XLE", "XLY", "XLP", "XLI", "XLU", "XLRE", "XLB", "XLC", _
        "IGV", "KBE", "XBI", "JETS", "XHB", "XRT", "XME", "XOP", "KIE", _
        "SMH", "SOXX", "ARKK", "ARKG", "KWEB", "GDX", "SLX", "MOO", "PHO", "ICLN")
    labels = Array( _
        "Tech", "Finance", "Health", "Energy", "ConDisc", "ConStap", "Industrl", "Utility", "RealEst", "Material", "CommSvc", _
        "SoftETF", "Banks", "BioTech", "Airlines", "HomeBld", "Retail", "Metals", "OilGas", "Insure", _
        "SemiETF", "SOXX", "ARKK", "ARKG", "China", "Gold", "Steel", "Agri", "Water", "CleanE")

    Dim nSec As Integer: nSec = UBound(tickers) + 1

    ' ── Fetch return data ────────────────────────────────────────
    Application.StatusBar = "Fetching price history for " & nSec & " sectors..."
    Dim retData() As Double
    ReDim retData(0 To nSec - 1, 0 To nDays - 2)

    Dim si As Integer
    For si = 0 To nSec - 1
        Application.StatusBar = "Fetching [" & si + 1 & "/" & nSec & "] " & tickers(si)
        DoEvents
        Dim prices() As Double
        prices = FetchPriceArray(CStr(tickers(si)), nDays)
        Dim j As Long
        For j = 0 To nDays - 2
            If j + 1 <= UBound(prices) And prices(j) > 0 Then
                retData(si, j) = (prices(j + 1) - prices(j)) / prices(j)
            Else
                retData(si, j) = 0
            End If
        Next j
    Next si

    ' ── Column widths ────────────────────────────────────────────
    wsC.Columns(1).ColumnWidth = 36
    wsC.Columns(2).ColumnWidth = 12
    wsC.Columns(3).Hidden = True
    Dim ci As Integer
    For ci = 0 To nSec - 1
        wsC.Columns(MATRIX_COL + ci).ColumnWidth = 6
    Next ci

    ' ── Row heights ──────────────────────────────────────────────
    wsC.Rows(1).RowHeight = 22
    wsC.Rows(MATRIX_ROW).RowHeight = 18
    Dim ri As Integer
    For ri = 0 To nSec - 1
        wsC.Rows(MATRIX_ROW + 1 + ri).RowHeight = 15
    Next ri

    ' ── Column headers (rotated) ─────────────────────────────────
    For ci = 0 To nSec - 2
        With wsC.cells(MATRIX_ROW, MATRIX_COL + ci)
            .Value = labels(ci)
            .Font.Color = RGB(255, 192, 0)
            .Font.Bold = True
            .Font.Size = 8
            .Interior.Color = RGB(10, 10, 10)
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlBottom
            .Orientation = 0
            .WrapText = False
        End With
    Next ci

    ' ── Matrix ───────────────────────────────────────────────────
    For ri = 0 To nSec - 1
        Dim dataRow As Long: dataRow = MATRIX_ROW + 1 + ri

        With wsC.cells(dataRow, 2)
            .Value = labels(ri)
            .Font.Color = RGB(255, 192, 0)
            .Font.Bold = True
            .Font.Size = 8
            .HorizontalAlignment = xlRight
            .VerticalAlignment = xlCenter
        End With

        For ci = 0 To nSec - 1
            Dim corr As Double
            If ri = ci Then
                corr = 1
            Else
                corr = CalcCorrelation(retData, ri, ci, nDays - 1)
            End If

            With wsC.cells(dataRow, MATRIX_COL + ci)
                .Value = corr
                .NumberFormat = "0.00"
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlCenter
                .Font.Size = 8
                .Font.Bold = (ri = ci)

                Select Case True
                    Case corr >= 0.7
                        .Interior.Color = RGB(120, 0, 0)
                        .Font.Color = RGB(255, 180, 180)
                    Case corr >= 0.4
                        .Interior.Color = RGB(70, 15, 15)
                        .Font.Color = RGB(230, 130, 130)
                    Case corr >= 0.1
                        .Interior.Color = RGB(25, 25, 25)
                        .Font.Color = RGB(190, 190, 190)
                    Case corr >= -0.1
                        .Interior.Color = RGB(10, 10, 10)
                        .Font.Color = RGB(120, 120, 120)
                    Case corr >= -0.4
                        .Interior.Color = RGB(10, 35, 10)
                        .Font.Color = RGB(100, 200, 100)
                    Case Else
                        .Interior.Color = RGB(0, 60, 0)
                        .Font.Color = RGB(80, 220, 80)
                End Select

                If ri = ci Then
                    .Interior.Color = RGB(35, 25, 0)
                    .Font.Color = RGB(255, 192, 0)
                End If
                ' Black border
                With .Borders
                    .LineStyle = xlContinuous
                    .Color = RGB(0, 0, 0)
                    .Weight = xlThin
                End With
            End With
        Next ci
    Next ri

    ' ── Gold separator ───────────────────────────────────────────
    With wsC.Range(wsC.cells(MATRIX_ROW, 1), _
                   wsC.cells(MATRIX_ROW, MATRIX_COL + nSec - 1)).Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Color = RGB(255, 192, 0)
        .Weight = xlThin
    End With

    ' ── Legend ───────────────────────────────────────────────────
    Dim legRow As Long: legRow = MATRIX_ROW + nSec + 2
    wsC.cells(legRow, 1).Value = "COLOUR LEGEND:"
    wsC.cells(legRow, 1).Font.Color = RGB(150, 150, 150)
    wsC.cells(legRow, 1).Font.Bold = True

    Dim legItems As Variant
    legItems = Array( _
        Array(">= 0.7  High +corr", RGB(120, 0, 0), RGB(255, 180, 180)), _
        Array("0.4~0.7 Mid  +corr", RGB(70, 15, 15), RGB(230, 130, 130)), _
        Array("-0.1~0.4 Neutral", RGB(25, 25, 25), RGB(190, 190, 190)), _
        Array("-0.4~-0.1 Mild -corr", RGB(10, 35, 10), RGB(100, 200, 100)), _
        Array("<= -0.4 High -corr", RGB(0, 60, 0), RGB(80, 220, 80)))

    Dim li As Integer
    For li = 0 To UBound(legItems)
        With wsC.cells(legRow + 1 + li, 1)
            .Value = legItems(li)(0)
            .Interior.Color = legItems(li)(1)
            .Font.Color = legItems(li)(2)
            .Font.Size = 9
            .Font.Bold = True
        End With
    Next li

    ' ── Freeze panes ─────────────────────────────────────────────
    wsC.cells(MATRIX_ROW + 1, MATRIX_COL).Select
    ActiveWindow.FreezePanes = False
    ActiveWindow.FreezePanes = True
    wsC.cells(1, 1).Select

    Application.StatusBar = "Correlation matrix done — " & Format(Now, "hh:mm:ss")
    MsgBox "Done! (" & nSec & " sectors, " & nDays & " days)", vbInformation
End Sub

' ── Fetch closing price array (nDays entries) ───────────────────
Private Function FetchPriceArray(Ticker As String, nDays As Long) As Double()
    Dim result() As Double
    ReDim result(0 To nDays - 1)

    Dim http As Object
    Set http = CreateObject("MSXML2.XMLHTTP")
    Dim url As String
    url = "https://query1.finance.yahoo.com/v8/finance/chart/" & Ticker & _
          "?interval=1d&range=" & CStr(CLng(nDays * 1.5 / 365 * 12)) & "mo"
    If CLng(nDays * 1.5) > 365 Then url = Replace(url, "range=", "range=2y&fake=")
    url = "https://query1.finance.yahoo.com/v8/finance/chart/" & Ticker & "?interval=1d&range=1y"

    On Error Resume Next
    http.Open "GET", url, False
    http.setRequestHeader "User-Agent", "Mozilla/5.0"
    http.send
    On Error GoTo 0

    If http.status <> 200 Then FetchPriceArray = result: Exit Function

    Dim resp As String: resp = http.responseText
    Dim pos  As Long: pos = InStr(resp, """close"":[")
    If pos = 0 Then FetchPriceArray = result: Exit Function

    Dim arrStr As String
    arrStr = Mid(resp, pos + 9, 5000)
    arrStr = Split(arrStr, "]")(0)
    Dim parts() As String: parts = Split(arrStr, ",")

    ' Take last nDays values
    Dim total As Long: total = UBound(parts) + 1
    Dim startIdx As Long: startIdx = IIf(total > nDays, total - nDays, 0)
    Dim idx As Long: idx = 0
    Dim k As Long
    For k = startIdx To UBound(parts)
        If IsNumeric(Trim(parts(k))) And idx < nDays Then
            result(idx) = CDbl(Trim(parts(k)))
            idx = idx + 1
        End If
    Next k

    FetchPriceArray = result
End Function

' ── Pearson correlation between two return series ───────────────
Private Function CalcCorrelation(retData() As Double, r1 As Integer, r2 As Integer, n As Long) As Double
    Dim sumX As Double, sumY As Double, sumXY As Double
    Dim sumX2 As Double, sumY2 As Double
    Dim i As Long, cnt As Long

    For i = 0 To n - 1
        Dim x As Double: x = retData(r1, i)
        Dim y As Double: y = retData(r2, i)
        If x <> 0 Or y <> 0 Then
            sumX = sumX + x
            sumY = sumY + y
            sumXY = sumXY + x * y
            sumX2 = sumX2 + x * x
            sumY2 = sumY2 + y * y
            cnt = cnt + 1
        End If
    Next i

    If cnt < 5 Then CalcCorrelation = 0: Exit Function

    Dim num As Double: num = cnt * sumXY - sumX * sumY
    Dim den As Double: den = Sqr((cnt * sumX2 - sumX ^ 2) * (cnt * sumY2 - sumY ^ 2))
    If den = 0 Then CalcCorrelation = 0 Else CalcCorrelation = num / den
End Function

' ── Helper: safely name a cell (optional) ───────────────────────
Private Sub Name_Safe(ws As Worksheet)
    ' placeholder, not used
End Sub

' ── Refresh with new days value ─────────────────────────────────
Sub RefreshCorrelation()
    Call BuildCorrelationMatrix
End Sub
```

## 4. DrawDownShadow.bas

`VB_Name = "DrawDownShadow"` · 478 行 · 16.7 KB · 原始編碼 ASCII

<details><summary><b>成員索引</b></summary>

| 行 | 型別 | 可見性 | 宣告 |
|---:|---|---|---|
| 30 | Sub | Public | `Public Sub BuildDrawdownShadow()` |

</details>

```vba
Option Explicit

' ================================================================
' DRAWDOWN SHADOW v2.0
' Builds drawdown shadow + pain-adjusted chart from HistoryLog!M
'
' v2 changes vs v1:
'   1. Detection rule:
'        v1: >= 2 consecutive negative-return days
'        v2: any 5-day rolling window with >= 4 negative days marks
'            all 5 days in that window as "in drawdown"
'   2. cumRet now SUMS ONLY THE NEGATIVE returns within a segment.
'      Positive bounce days inside the segment are zeroed out.
'   3. BUG FIX: the pain-adjusted shadowVal formula was missing a
'      '+' between the two terms (the '_' line-continuation in v1
'      silently dropped it). v2 restores it.
' ================================================================

Private Const SH_HIST    As String = "HistoryLog"
Private Const SH_OUT     As String = "DrawdownChart"
Private Const DATA_ROW   As Long = 3
Private Const PAIN_CAP   As Double = 5
Private Const PAIN_OVERFLOW As Double = 9999

' Detection window parameters - change these to tune rule.
Private Const WIN_SIZE     As Long = 5    ' rolling window length
Private Const WIN_MIN_NEG  As Long = 4    ' min negative days within window

Public Sub BuildDrawdownShadow()
    Application.ScreenUpdating = False
    Application.StatusBar = "Building drawdown shadow v2..."
    On Error GoTo CleanFail

    ' --- Read HistoryLog!M into typed arrays ---
    Dim wsH As Worksheet
    On Error Resume Next: Set wsH = ThisWorkbook.Sheets(SH_HIST): On Error GoTo CleanFail
    If wsH Is Nothing Then MsgBox "HistoryLog sheet not found", vbExclamation: GoTo CleanExit

    Dim lastRow As Long
    lastRow = wsH.Cells(wsH.Rows.Count, "M").End(xlUp).Row
    If lastRow < 3 Then MsgBox "Not enough data in column M (need >= 5 rows for 5-day window)", vbInformation: GoTo CleanExit

    Dim n As Long: n = lastRow - 1
    Dim ret() As Double: ReDim ret(1 To n)
    Dim dateStr() As String: ReDim dateStr(1 To n)
    Dim valid() As Boolean: ReDim valid(1 To n)

    Dim r As Long, i As Long
    For r = 2 To lastRow
        i = r - 1
        Dim v As Variant: v = wsH.Cells(r, "M").Value
        If IsNumeric(v) Then
            ret(i) = CDbl(v)
            valid(i) = True
        Else
            ret(i) = 0
            valid(i) = False
        End If
        dateStr(i) = "R" & r
        On Error Resume Next
        Dim d0 As Date: d0 = CDate(wsH.Cells(r, "A").Value)
        If d0 > 0 Then dateStr(i) = Format(d0, "m/d/yy")
        On Error GoTo CleanFail
    Next r

    ' --- Mark DD days using 5-day rolling window ---
    Dim isDD() As Boolean: ReDim isDD(1 To n)
    Dim w As Long, negCount As Long, k As Long
    For w = 1 To n - WIN_SIZE + 1
        negCount = 0
        For k = w To w + WIN_SIZE - 1
            If valid(k) And ret(k) < 0 Then negCount = negCount + 1
        Next k
        If negCount >= WIN_MIN_NEG Then
            For k = w To w + WIN_SIZE - 1
                isDD(k) = True
            Next k
        End If
    Next w

    ' --- Build segments from consecutive isDD-marked days ---
    Dim segments As Collection: Set segments = New Collection
    Dim segDates As Collection: Set segDates = New Collection

    Dim s As Long, segStart As Long, segEnd As Long
    s = 1
    Do While s <= n
        If isDD(s) Then
            segStart = s
            Do While s <= n
                If Not isDD(s) Then Exit Do
                s = s + 1
            Loop
            segEnd = s - 1

            Dim segLen As Long: segLen = segEnd - segStart + 1
            Dim segArr() As Double: ReDim segArr(0 To segLen - 1)
            Dim p As Long
            For p = 0 To segLen - 1
                ' Ignore positive returns - only negatives accumulate
                If ret(segStart + p) < 0 Then
                    segArr(p) = ret(segStart + p)
                Else
                    segArr(p) = 0
                End If
            Next p

            segments.Add segArr
            segDates.Add dateStr(segStart)
        Else
            s = s + 1
        End If
    Loop

    If segments.Count = 0 Then
        MsgBox "No drawdown segments found (no 5-day window has >= 4 negative days)", vbInformation
        GoTo CleanExit
    End If

    ' --- Find max segment length ---
    Dim maxDays As Long: maxDays = 0
    For s = 1 To segments.Count
        Dim a As Variant: a = segments(s)
        If UBound(a) + 1 > maxDays Then maxDays = UBound(a) + 1
    Next s

    ' --- Create / clear DrawdownChart sheet ---
    Dim wsD As Worksheet
    On Error Resume Next: Set wsD = ThisWorkbook.Sheets(SH_OUT): On Error GoTo CleanFail
    If wsD Is Nothing Then
        Set wsD = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        wsD.Name = SH_OUT
    End If
    wsD.Cells.Clear

    Dim shp As Shape
    For Each shp In wsD.Shapes
        shp.Delete
    Next shp

    With wsD.Cells
        .Interior.Color = RGB(0, 0, 0)
        .Font.Color = RGB(200, 200, 200)
        .Font.Name = "Consolas"
        .Font.Size = 9
    End With
    wsD.Activate
    ActiveWindow.DisplayGridlines = False

    With wsD.Cells(1, 1)
        .Value = "DRAWDOWN SHADOW v2 (rule: " & WIN_MIN_NEG & "/" & WIN_SIZE & ") | " & _
                 segments.Count & " episodes | " & Format(Now, "hh:mm:ss")
        .Font.Color = RGB(255, 192, 0)
        .Font.Bold = True
        .Font.Size = 12
    End With
    wsD.Rows(1).RowHeight = 24

    wsD.Cells(DATA_ROW - 1, 1).Value = "Day"
    wsD.Cells(DATA_ROW - 1, 1).Font.Color = RGB(150, 150, 150)

    Dim d As Long
    For d = 1 To maxDays
        wsD.Cells(DATA_ROW + d - 1, 1).Value = d
    Next d

    ' --- Compute shadow f(x) per segment per day ---
    Dim minCum As Double: minCum = 0
    Dim cumRet As Double, runMin As Double
    Dim xVal As Double, fxVal As Double, cellVal As Double
    Dim depth As Double, rVal As Long, colIdx As Long, segLenInner As Long

    For s = 1 To segments.Count
        Dim arr As Variant: arr = segments(s)
        segLenInner = UBound(arr) + 1
        colIdx = s + 1
        wsD.Cells(DATA_ROW - 1, colIdx).Value = segDates(s)
        wsD.Cells(DATA_ROW - 1, colIdx).Font.Color = RGB(180, 180, 180)
        wsD.Cells(DATA_ROW - 1, colIdx).Font.Size = 8

        cumRet = 0
        runMin = 0
        For d = 0 To segLenInner - 1
            xVal = d + 1
            cumRet = cumRet + arr(d)   ' arr(d) is 0 for positive days (ignored)
            fxVal = (Exp(xVal) / (xVal ^ 3 + WorksheetFunction.Pi() ^ 3) + Exp(Sqr(xVal))) * cumRet
            cellVal = fxVal
            wsD.Cells(DATA_ROW + d, colIdx).Value = Round(cellVal, 4)
            wsD.Cells(DATA_ROW + d, colIdx).NumberFormat = "0.0000"

            depth = Abs(cellVal)
            rVal = IIf(depth > 1000, 255, IIf(depth > 100, 230, IIf(depth > 10, 200, 170)))
            wsD.Cells(DATA_ROW + d, colIdx).Font.Color = RGB(rVal, 60, 60)

            If cellVal < runMin Then runMin = cellVal
        Next d
        If runMin < minCum Then minCum = runMin
    Next s

    ' --- Column widths / row heights ---
    wsD.Columns(1).ColumnWidth = 6
    Dim ci As Long
    For ci = 2 To segments.Count + 1
        wsD.Columns(ci).ColumnWidth = 9
    Next ci
    For d = 1 To maxDays
        wsD.Rows(DATA_ROW + d - 1).RowHeight = 15
    Next d

    ' --- Build chart ---
    Dim co As ChartObject
    Set co = wsD.ChartObjects.Add(wsD.Columns("M").Left, wsD.Rows(1).Top, 700, 400)
    With co.Chart
        .ChartType = xlLine
        .HasTitle = True
        .ChartTitle.Text = "Drawdown Shadow v2 (" & WIN_MIN_NEG & "/" & WIN_SIZE & " rule)"
        .ChartTitle.Font.Color = RGB(255, 192, 0)
        .ChartTitle.Font.Name = "Consolas"
        .ChartTitle.Font.Bold = True

        Do While .SeriesCollection.Count > 0
            .SeriesCollection(1).Delete
        Loop

        Dim xVals() As Long: ReDim xVals(0 To maxDays - 1)
        For d = 0 To maxDays - 1: xVals(d) = d + 1: Next d

        Dim yVals() As Variant
        Dim cum2 As Double, xv As Double
        Dim ser As Series
        Dim brightness As Long, lineClr As Long

        For s = 1 To segments.Count
            arr = segments(s)
            segLenInner = UBound(arr) + 1
            ReDim yVals(0 To maxDays - 1)
            cum2 = 0
            For d = 0 To maxDays - 1
                If d < segLenInner Then
                    xv = d + 1
                    cum2 = cum2 + arr(d)
                    yVals(d) = Round((Exp(xv) / (xv ^ 3 + WorksheetFunction.Pi() ^ 3) + Exp(Sqr(xv))) * cum2, 4)
                Else
                    yVals(d) = CVErr(xlErrNA)
                End If
            Next d
            Set ser = .SeriesCollection.NewSeries
            ser.Values = yVals
            ser.XValues = xVals
            ser.Name = CStr(segDates(s))

            brightness = 60 + CInt((s / segments.Count) * 160)
            lineClr = RGB(brightness, 20, 20)
            With ser.Format.Line
                .ForeColor.RGB = lineClr
                .Weight = IIf(s = segments.Count, 2, 1)
                .Transparency = IIf(s = segments.Count, 0, 0.4 - (s / segments.Count) * 0.3)
            End With
            ser.MarkerStyle = xlMarkerStyleNone
        Next s

        .PlotArea.Interior.Color = RGB(8, 8, 8)
        .PlotArea.Border.Color = RGB(40, 40, 40)
        .ChartArea.Interior.Color = RGB(0, 0, 0)
        .ChartArea.Border.Color = RGB(30, 30, 30)

        With .Axes(xlCategory)
            .HasTitle = True
            .AxisTitle.Text = "Days in Drawdown"
            .AxisTitle.Font.Color = RGB(255, 255, 255)
            .AxisTitle.Font.Name = "Consolas"
            .TickLabels.Font.Color = RGB(255, 255, 255)
            .TickLabels.Font.Name = "Consolas"
            .TickLabels.Font.Size = 8
            .MajorGridlines.Border.Color = RGB(35, 35, 35)
        End With
        With .Axes(xlValue)
            .HasTitle = True
            .AxisTitle.Text = "f(x) = sum(neg) * (e^x/(x^3+pi^3) + e^sqrt(x))"
            .AxisTitle.Font.Color = RGB(255, 255, 255)
            .AxisTitle.Font.Name = "Consolas"
            .TickLabels.Font.Color = RGB(255, 255, 255)
            .TickLabels.Font.Name = "Consolas"
            .TickLabels.Font.Size = 8
            .TickLabels.NumberFormat = "0.00"
            .MajorGridlines.Border.Color = RGB(35, 35, 35)
            .MaximumScale = 0
            .CrossesAt = 0
        End With

        .HasLegend = True
        With .Legend
            .Interior.Color = RGB(10, 10, 10)
            .Border.Color = RGB(30, 30, 30)
            .Font.Color = RGB(255, 255, 255)
            .Font.Name = "Consolas"
            .Font.Size = 7
        End With
    End With

    ' ============================================================
    ' Pain-Adjusted Section  (BUG FIX: missing '+' restored)
    ' ============================================================
    Dim painStartRow As Long: painStartRow = DATA_ROW + maxDays + 2
    Dim painHeaderRow As Long: painHeaderRow = painStartRow + 1
    Dim painDataRow As Long
    Dim painVal As Double, shadowVal As Double
    Dim minPain As Double: minPain = 0

    With wsD.Cells(painStartRow, 1)
        .Value = "PAIN-ADJUSTED DRAWDOWN | f(x) = TAN(ASIN(shadow/cap)) * cap"
        .Font.Color = RGB(255, 99, 71)
        .Font.Bold = True
        .Font.Size = 10
    End With

    With wsD.Cells(painStartRow, segments.Count + 3)
        .Value = "PAIN CAP%"
        .Font.Color = RGB(255, 192, 0)
        .Font.Bold = True
    End With
    With wsD.Cells(painStartRow, segments.Count + 4)
        .Value = PAIN_CAP
        .Font.Color = RGB(0, 191, 255)
        .Font.Bold = True
    End With

    With wsD.Cells(painHeaderRow, 1)
        .Value = "Day"
        .Font.Color = RGB(150, 150, 150)
    End With
    For d = 1 To maxDays
        With wsD.Cells(painHeaderRow + d, 1)
            .Value = d
            .Font.Color = RGB(200, 200, 200)
        End With
    Next d

    For s = 1 To segments.Count
        colIdx = s + 1
        With wsD.Cells(painHeaderRow, colIdx)
            .Value = segDates(s)
            .Font.Color = RGB(180, 180, 180)
            .Font.Size = 8
        End With
    Next s

    ' --- Compute pain values ---
    For s = 1 To segments.Count
        arr = segments(s)
        segLenInner = UBound(arr) + 1
        colIdx = s + 1
        cumRet = 0

        For d = 0 To segLenInner - 1
            xVal = d + 1
            cumRet = cumRet + arr(d)
            ' BUG FIX: previous version had '_' line continuation but missing '+'.
            ' Now the '+' is explicit so e^sqrt(x) is properly added.
            shadowVal = (Exp(xVal) / (xVal ^ 3 + WorksheetFunction.Pi() ^ 3) + _
                         Exp(Sqr(xVal))) * cumRet

            If Abs(shadowVal) >= PAIN_CAP Then
                painVal = Sgn(shadowVal) * PAIN_OVERFLOW
            Else
                painVal = Tan(WorksheetFunction.Asin(shadowVal / PAIN_CAP)) * PAIN_CAP
            End If

            painDataRow = painHeaderRow + xVal
            With wsD.Cells(painDataRow, colIdx)
                .Value = Round(painVal, 4)
                .NumberFormat = "0.0000"
                .Font.Size = 9
                depth = Abs(painVal)
                If depth >= PAIN_OVERFLOW Then
                    .Font.Color = RGB(255, 0, 0): .Font.Bold = True
                ElseIf depth > 5 Then
                    .Font.Color = RGB(255, 40, 20)
                ElseIf depth > 2 Then
                    .Font.Color = RGB(240, 60, 30)
                ElseIf depth > 0.5 Then
                    .Font.Color = RGB(210, 80, 50)
                Else
                    .Font.Color = RGB(180, 100, 80)
                End If
            End With

            If painVal < minPain Then minPain = painVal
        Next d
    Next s

    ' --- Pain stats ---
    Dim painStatCol As Long: painStatCol = segments.Count + 3
    With wsD.Cells(painHeaderRow + 1, painStatCol)
        .Value = "PAIN STATS"
        .Font.Color = RGB(255, 99, 71)
        .Font.Bold = True
    End With
    With wsD.Cells(painHeaderRow + 2, painStatCol)
        .Value = "Max Pain"
        .Font.Color = RGB(150, 150, 150)
    End With
    With wsD.Cells(painHeaderRow + 2, painStatCol + 1)
        .Value = Round(minPain, 4)
        .NumberFormat = "0.0000"
        .Font.Color = RGB(255, 69, 0)
    End With

    If minCum <> 0 Then
        With wsD.Cells(painHeaderRow + 3, painStatCol)
            .Value = "Pain/DD"
            .Font.Color = RGB(150, 150, 150)
        End With
        With wsD.Cells(painHeaderRow + 3, painStatCol + 1)
            .Value = Format(minPain / minCum, "0.00") & "x"
            .Font.Color = RGB(255, 69, 0)
        End With
    End If

    ' --- Summary panel ---
    Dim sumRow As Long: sumRow = DATA_ROW - 1
    Dim sumCol As Long: sumCol = segments.Count + 3
    With wsD.Cells(sumRow, sumCol)
        .Value = "DRAWDOWN SUMMARY v2"
        .Font.Color = RGB(255, 192, 0)
        .Font.Bold = True
    End With
    wsD.Cells(sumRow + 1, sumCol).Value = "Episodes"
    wsD.Cells(sumRow + 1, sumCol + 1).Value = segments.Count
    wsD.Cells(sumRow + 2, sumCol).Value = "Max Days"
    wsD.Cells(sumRow + 2, sumCol + 1).Value = maxDays
    wsD.Cells(sumRow + 3, sumCol).Value = "Min f(x)"
    wsD.Cells(sumRow + 3, sumCol + 1).Value = Format(minCum, "0.0000")
    wsD.Cells(sumRow + 4, sumCol).Value = "Rule"
    wsD.Cells(sumRow + 4, sumCol + 1).Value = WIN_MIN_NEG & "/" & WIN_SIZE & " neg"

    wsD.Cells(sumRow + 6, sumCol).Value = "Start"
    wsD.Cells(sumRow + 6, sumCol + 1).Value = "Days"
    wsD.Cells(sumRow + 6, sumCol + 2).Value = "Sum Neg%"
    wsD.Cells(sumRow + 6, sumCol).Font.Color = RGB(255, 192, 0)
    wsD.Cells(sumRow + 6, sumCol + 1).Font.Color = RGB(255, 192, 0)
    wsD.Cells(sumRow + 6, sumCol + 2).Font.Color = RGB(255, 192, 0)

    Dim sumNeg As Double
    For s = 1 To segments.Count
        arr = segments(s)
        sumNeg = 0
        For k = 0 To UBound(arr)
            sumNeg = sumNeg + arr(k)   ' arr already zero-filled for positives
        Next k
        wsD.Cells(sumRow + 6 + s, sumCol).Value = segDates(s)
        wsD.Cells(sumRow + 6 + s, sumCol + 1).Value = UBound(arr) + 1
        wsD.Cells(sumRow + 6 + s, sumCol + 2).Value = Format(sumNeg * 100, "0.00") & "%"
        wsD.Cells(sumRow + 6 + s, sumCol + 2).Font.Color = RGB(200, 80, 80)
    Next s

    For ci = sumCol To sumCol + 2
        wsD.Columns(ci).ColumnWidth = 14
    Next ci

CleanExit:
    Application.ScreenUpdating = True
    Application.StatusBar = False
    If segments.Count > 0 Then
        Application.StatusBar = "Drawdown Shadow v2 done: " & segments.Count & " episodes"
        MsgBox "Done! Found " & segments.Count & " drawdown segments." & vbCrLf & _
               "Rule: 5-day window with >= 4 negative days" & vbCrLf & _
               "Positive returns within segment are ignored in cumRet.", vbInformation
    End If
    Exit Sub

CleanFail:
    Application.ScreenUpdating = True
    Application.StatusBar = False
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical
End Sub
```

## 5. modvolatility.bas

`VB_Name = "modvolatility"` · 317 行 · 11.2 KB · 原始編碼 CP950 (Big5)

<details><summary><b>成員索引</b></summary>

| 行 | 型別 | 可見性 | 宣告 |
|---:|---|---|---|
| 14 | Sub | Public | `Sub UpdateVolatilityAnalysis_Pro()` |
| 61 | Sub | Public | `Sub AnalyzeAndOutput(ws As Worksheet, ByRef allPrices() As Double, lookbackDays As Long, startRow As Long, titleSuffix As String)` |
| 239 | Function | Public | `Function GetHistoricalData(Ticker As String, ByRef outDates() As Date, ByRef outPrices() As Double) As Boolean` |

</details>

```vba
Option Explicit




'=============================================================================
'  功能：個股波動率分佈與進階風險分析 (支援多週期：100天, 200天, 1000天)
'  工作表：Tickers volatility
'  輸入：B2 (Ticker)
'  輸出：多週期報表 (起始列分別為 4, 29, 54)
'=============================================================================

Sub UpdateVolatilityAnalysis_Pro()
    Dim ws As Worksheet
    Dim Ticker As String
    Dim allPrices() As Double, allDates() As Date
    Dim totalDays As Long
    
    ' 1. 設定工作表與輸入
    Set ws = ThisWorkbook.Sheets("Tickers volatility")
    Ticker = Trim(ws.Range("B2").Value)
    If Ticker = "" Then MsgBox "請在 B2 輸入股票代碼", vbExclamation: Exit Sub
    
    Application.ScreenUpdating = False
    
    ' 2. 清除舊數據 (範圍加大以容納三組分析表)
    ws.Range("A4:F100").ClearContents
    
    ' 3. 抓取歷史數據 (改為抓取 10 年，確保有 1000 個交易日)
    If Not GetHistoricalData(Ticker, allDates, allPrices) Then
        MsgBox "無法抓取數據，請檢查代碼或網絡。", vbCritical
        Application.ScreenUpdating = True
        Exit Sub
    End If
    
    totalDays = UBound(allPrices) + 1
    If totalDays < 20 Then
        MsgBox "數據筆數不足以進行進階分析。", vbExclamation
        Application.ScreenUpdating = True
        Exit Sub
    End If
    
    ' 4. 執行三個不同週期的分析 (傳入: 工作表, 全部價格, 欲分析天數, 輸出起始列, 標題)
    ' 1-a. 過去 100 天 (輸出在 A4:F26)
    Call AnalyzeAndOutput(ws, allPrices, 100, 4, "Past 100 Days")
    
    ' 1-b. 過去 200 天 (輸出在 A29:F51)
    Call AnalyzeAndOutput(ws, allPrices, 200, 29, "Past 200 Days")
    
    ' 1-c. 過去 1000 天 (輸出在 A54:F76)
    Call AnalyzeAndOutput(ws, allPrices, 1000, 54, "Past 1000 Days")
    
    Application.ScreenUpdating = True
    MsgBox "100天、200天、1000天進階波動率分析更新完成！", vbInformation
End Sub

'=============================================================================
'  副程式：執行特定天數的計算與報表輸出
'=============================================================================
Sub AnalyzeAndOutput(ws As Worksheet, ByRef allPrices() As Double, lookbackDays As Long, startRow As Long, titleSuffix As String)
    Dim prices() As Double, returns() As Double
    Dim i As Long, j As Long, n As Long
    Dim totalData As Long, useDays As Long
    
    totalData = UBound(allPrices) + 1
    useDays = lookbackDays
    If totalData < useDays Then useDays = totalData ' 如果歷史數據不夠，則用最大可用的天數
    
    ' 1. 截取最近 useDays 的股價數據
    ReDim prices(0 To useDays - 1)
    For i = 0 To useDays - 1
        prices(i) = allPrices(UBound(allPrices) - useDays + 1 + i)
    Next i
    
    n = UBound(prices)
    If n < 1 Then Exit Sub
    ReDim returns(0 To n - 1)
    
    ' --- 統計變數宣告 ---
    Dim mean As Double, stdDev As Double, maxRet As Double, minRet As Double
    Dim annVol As Double, skew As Double, kurt As Double
    Dim downDev As Double, mdd As Double
    Dim currRollVol As Double, maxRollVol As Double
    Dim upStreak As Long, downStreak As Long, maxUpStreak As Long, maxDownStreak As Long
    
    ' 2. 計算日漲跌幅 & 連續漲跌天數
    maxUpStreak = 0: maxDownStreak = 0
    upStreak = 0: downStreak = 0
    
    For i = 1 To n
        If prices(i - 1) <> 0 Then
            returns(i - 1) = (prices(i) - prices(i - 1)) / prices(i - 1)
        Else
            returns(i - 1) = 0
        End If
        
        If returns(i - 1) > 0 Then
            upStreak = upStreak + 1
            downStreak = 0
            If upStreak > maxUpStreak Then maxUpStreak = upStreak
        ElseIf returns(i - 1) < 0 Then
            downStreak = downStreak + 1
            upStreak = 0
            If downStreak > maxDownStreak Then maxDownStreak = downStreak
        Else
            upStreak = 0
            downStreak = 0
        End If
    Next i
    
    ' 3. 計算最大回撤 (MDD)
    Dim peak As Double, DrawDown As Double
    peak = prices(0)
    mdd = 0
    For i = 1 To n
        If prices(i) > peak Then peak = prices(i)
        DrawDown = (peak - prices(i)) / peak
        If DrawDown > mdd Then mdd = DrawDown
    Next i
    
    ' 4. 計算下行標準差 (Downside Deviation)
    Dim downReturns() As Double
    Dim downCount As Long
    downCount = 0
    For i = LBound(returns) To UBound(returns)
        If returns(i) < 0 Then
            ReDim Preserve downReturns(0 To downCount)
            downReturns(downCount) = returns(i)
            downCount = downCount + 1
        End If
    Next i
    
    ' 5. 計算 20日滾動波動率
    Dim rollWindow(1 To 20) As Double
    Dim rollVol As Double
    maxRollVol = 0
    If UBound(returns) >= 19 Then
        For i = 19 To UBound(returns)
            For j = 1 To 20
                rollWindow(j) = returns(i - 20 + j)
            Next j
            rollVol = Application.WorksheetFunction.StDev_S(rollWindow) * Sqr(252)
            If rollVol > maxRollVol Then maxRollVol = rollVol
            If i = UBound(returns) Then currRollVol = rollVol
        Next i
    End If
    
    ' 6. 基礎與進階統計指標
    On Error Resume Next
    mean = Application.WorksheetFunction.Average(returns)
    stdDev = Application.WorksheetFunction.StDev_S(returns)
    maxRet = Application.WorksheetFunction.Max(returns)
    minRet = Application.WorksheetFunction.Min(returns)
    annVol = stdDev * Sqr(252)
    skew = Application.WorksheetFunction.skew(returns)
    kurt = Application.WorksheetFunction.kurt(returns)
    If downCount > 1 Then
        downDev = Application.WorksheetFunction.StDev_S(downReturns)
    Else
        downDev = 0
    End If
    On Error GoTo 0
    
    ' 7. 填寫標題與數值 (結構化輸出)
    Dim labelsArr As Variant, valsArr As Variant
    labelsArr = Array("Mean (Daily)", "Std Dev (Daily)", "Annualized Volatility", _
                      "Skewness (不對稱性)", "Kurtosis (肥尾效應)", _
                      "Downside Deviation", "Max Drawdown (MDD)", _
                      "Current 20D Volatility", "Max 20D Volatility", _
                      "Max Up Streak (Days)", "Max Down Streak (Days)", _
                      "Max Daily Chg%", "Min Daily Chg%")
                      
    valsArr = Array(mean, stdDev, annVol, skew, kurt, downDev, -mdd, _
                    currRollVol, maxRollVol, maxUpStreak, maxDownStreak, maxRet, minRet)
    
    ' 寫入表頭
    ws.cells(startRow, 1).Value = "Statistic (" & titleSuffix & ")"
    ws.cells(startRow, 2).Value = "Value"
    
    For i = LBound(labelsArr) To UBound(labelsArr)
        ws.cells(startRow + 1 + i, 1).Value = labelsArr(i)
        ws.cells(startRow + 1 + i, 2).Value = valsArr(i)
    Next i
    
    ' 動態套用格式化
    ws.Range("B" & startRow + 1 & ":B" & startRow + 3 & ", B" & startRow + 6 & ":B" & startRow + 9 & ", B" & startRow + 12 & ":B" & startRow + 13).NumberFormat = "0.00%"
    ws.Range("B" & startRow + 4 & ":B" & startRow + 5).NumberFormat = "0.00"
    ws.Range("B" & startRow + 10 & ":B" & startRow + 11).NumberFormat = "0"
    
    ' 8. 建立分佈表
    Dim bins(0 To 21) As Long, labels(0 To 21) As String
    Dim retVal As Double, binIdx As Integer
    Dim totalReturnsCount As Long
    
    totalReturnsCount = UBound(returns) + 1
    labels(0) = "< -10%"
    labels(21) = "> 10%"
    
    For i = 1 To 20
        Dim lowerLimit As Double
        lowerLimit = -0.1 + (i - 1) * 0.01
        labels(i) = Format(lowerLimit, "0%") & " ~ " & Format(lowerLimit + 0.01, "0%")
    Next i
    
    For i = LBound(returns) To UBound(returns)
        retVal = returns(i)
        If retVal < -0.1 Then
            bins(0) = bins(0) + 1
        ElseIf retVal >= 0.1 Then
            bins(21) = bins(21) + 1
        Else
            binIdx = Int((retVal + 0.1) / 0.01) + 1
            If binIdx >= 1 And binIdx <= 20 Then bins(binIdx) = bins(binIdx) + 1
        End If
    Next i
    
    ' 填入分佈表格
    ws.cells(startRow, 4).Value = "Range"
    ws.cells(startRow, 5).Value = "Count"
    ws.cells(startRow, 6).Value = "Freq %"
    
    For i = 0 To 21
        ws.cells(startRow + 1 + i, 4).Value = "'" & labels(i)
        ws.cells(startRow + 1 + i, 5).Value = bins(i)
        If totalReturnsCount > 0 Then
            ws.cells(startRow + 1 + i, 6).Value = bins(i) / totalReturnsCount
        Else
            ws.cells(startRow + 1 + i, 6).Value = 0
        End If
    Next i
    ws.Range("F" & startRow + 1 & ":F" & startRow + 22).NumberFormat = "0.00%"

End Sub

'=============================================================================
'  核心函數：抓取 10 年數據 (供切割週期使用)
'=============================================================================
Function GetHistoricalData(Ticker As String, ByRef outDates() As Date, ByRef outPrices() As Double) As Boolean
    Dim http As Object
    Dim url As String, response As String
    Dim tsStr As String, closeStr As String
    Dim tsArr() As String, closeArr() As String
    Dim i As Long
    
    ' 1. 處理 Ticker
    If InStr(Ticker, ".TW") = 0 And InStr(Ticker, ".TWO") = 0 And IsNumeric(Ticker) Then
        Ticker = Ticker & ".TW"
    End If
    
    ' 2. 設定 URL (Range 改為 10y 確保涵蓋 1000 交易日)
    url = "https://query1.finance.yahoo.com/v8/finance/chart/" & Ticker & "?range=10y&interval=1d"
    
    On Error Resume Next
    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
    If http Is Nothing Then Set http = CreateObject("MSXML2.XMLHTTP")
    With http
        .Open "GET", url, False
        .setRequestHeader "User-Agent", "Mozilla/5.0"
        .send
        response = .responseText
    End With
    On Error GoTo 0
    
    ' 3. 解析 JSON
    Dim tsStart As Long, tsEnd As Long
    tsStart = InStr(response, """timestamp"":[") + 12
    tsEnd = InStr(tsStart, response, "]")
    If tsStart > 12 And tsEnd > tsStart Then
        tsStr = Mid(response, tsStart, tsEnd - tsStart)
        tsArr = Split(tsStr, ",")
    Else
        GetHistoricalData = False
        Exit Function
    End If
    
    Dim closeStart As Long, closeEnd As Long
    closeStart = InStr(response, """adjclose"":[{""adjclose"":[") + 23
    If closeStart < 24 Then closeStart = InStr(response, """close"":[") + 8
    
    closeEnd = InStr(closeStart, response, "]")
    If closeStart > 20 And closeEnd > closeStart Then
        closeStr = Mid(response, closeStart, closeEnd - closeStart)
        closeArr = Split(closeStr, ",")
    Else
        GetHistoricalData = False
        Exit Function
    End If
    
    Dim count As Long
    count = UBound(tsArr)
    If UBound(closeArr) < count Then count = UBound(closeArr)
    
    ReDim outDates(0 To count)
    ReDim outPrices(0 To count)
    
    Dim validCount As Long
    validCount = 0
    
    For i = 0 To count
        If IsNumeric(tsArr(i)) And IsNumeric(closeArr(i)) Then
            outDates(validCount) = (CDbl(tsArr(i)) / 86400) + 25569
            outPrices(validCount) = CDbl(closeArr(i))
            validCount = validCount + 1
        End If
    Next i
    
    If validCount > 0 Then
        ReDim Preserve outDates(0 To validCount - 1)
        ReDim Preserve outPrices(0 To validCount - 1)
        GetHistoricalData = True
    Else
        GetHistoricalData = False
    End If
End Function
```

## 6. PortfolioDashboard_v3.bas

`VB_Name = "PortfolioDashboard_v3"` · 1860 行 · 70 KB · 原始編碼 UTF-8

<details><summary><b>成員索引</b></summary>

| 行 | 型別 | 可見性 | 宣告 |
|---:|---|---|---|
| 19 | Sub | Public | `Sub RebuildPortfolioDashboard()` |
| 85 | Sub | Private | `Private Sub LogHistory(totalMkt As Double, totalPnL As Double, realPnL As Double)` |
| 130 | Function | Public | `Function GetHistoricalPrice(Ticker As String, targetDate As Date) As Double` |
| 169 | Sub | Public | `Sub BackfillHistory()` |
| 236 | Sub | Private | `Private Sub ResetSheetStyle(ws As Worksheet)` |
| 275 | Sub | Private | `Private Sub DrawHeader(ws As Worksheet, totalMkt As Double, totalCost As Double, totalUnrl As Double, realPnL As Double, portBeta As Double, posCount As Long, exRate As Double)` |
| 369 | Sub | Public | `Sub DrawDeepAnalysis(wsP As Worksheet, posData() As Variant, totalMkt As Double, totalCost As Double, totalUnrl As Double, realPnL As Double)` |
| 537 | Sub | Public | `Sub RebuildActiveXButtons()` |
| 614 | Function | Private | `Private Function DrawAllocTable(ws As Worksheet, startRow As Long, title As String, dict As Object, totalMkt As Double) As Long` |
| 659 | Sub | Private | `Private Sub DrawColumnHeaders(ws As Worksheet)` |
| 697 | Function | Private | `Private Function WritePositionRows(ws As Worksheet, posData() As Variant, totalMkt As Double, ByVal swingRiskMap As Object) As Long` |
| 740 | Sub | Private | `Private Sub DrawBrokerHeader(ws As Worksheet, r As Long, brokerName As String)` |
| 754 | Sub | Private | `Private Sub DrawBrokerSubtotal(ws As Worksheet, r As Long, brokerName As String, brokerMkt As Double, brokerUnrl As Double)` |
| 783 | Sub | Private | `Private Sub WriteOnePositionRow(ws As Worksheet, r As Long, i As Long, posData() As Variant, totalMkt As Double, swingRiskMap As Object)` |
| 889 | Sub | Private | `Private Sub DrawDisclaimer(ws As Worksheet, startRow As Long)` |
| 912 | Sub | Private | `Private Sub CalcPositions(positions As Object, exRate As Double, ByRef posData() As Variant, ByRef totalMkt As Double, ByRef totalCost As Double, ByRef totalUnrl As Double, ByRef posCount As Long)` |
| 1100 | Function | Private | `Private Function BuildPositions(wsTr As Worksheet) As Object` |
| 1161 | Function | Private | `Private Function CalcPortBeta(posData() As Variant) As Double` |
| 1181 | Function | Private | `Private Function FindNearestPrice(ByRef dict As Object, ByVal targetDate As Date) As Double` |
| 1200 | Function | Private | `Private Function PnLColor(v As Double) As Long` |
| 1214 | Function | Private | `Private Function PnLColorMuted(v As Double) As Long` |
| 1222 | Function | Private | `Private Function GetExRate(ws As Worksheet) As Double` |
| 1229 | Function | Private | `Private Function GetInceptionDate(ws As Worksheet) As Date` |
| 1236 | Function | Private | `Private Function GetStartingCapital(ws As Worksheet) As Double` |
| 1246 | Sub | Public | `Sub SetupPortfolioConfig()` |
| 1275 | Sub | Public | `Sub StartAutoRefresh()` |
| 1283 | Sub | Public | `Sub AutoRefreshLoop()` |
| 1290 | Sub | Public | `Sub StopAutoRefresh()` |
| 1298 | Sub | Private | `Private Sub WriteKV(ws As Worksheet, r As Long, label As String, val As Double, fmt As String, colorPnL As Boolean)` |
| 1313 | Sub | Public | `Sub BuildHoldingsCorrelation()` |
| 1474 | Sub | Private | `Private Sub HoldingsCorrRenderSheet(tickers() As String, tickerCount As Long, corrMatrix() As Double, usedCount As Long, commonDates() As Date)` |
| 1638 | Sub | Private | `Private Sub CorrGetCommonDates(priceData() As Object, tickerCount As Long, ByRef commonDates() As Date, ByRef usedCount As Long)` |
| 1700 | Function | Private | `Private Function CorrPearson(returns() As Double, r1 As Long, r2 As Long, n As Long) As Double` |
| 1721 | Sub | Private | `Private Sub CorrRenderSheet(tickers() As String, tickerCount As Long, corrMatrix() As Double, usedCount As Long, commonDates() As Date)` |

</details>

```vba
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

    ws.cells(1, 9).Value = "PORT.BETA"
    ws.cells(1, 9).Font.Color = RGB(150, 150, 150)
    ws.cells(1, 10).Value = portBeta
    ws.cells(1, 10).NumberFormat = "0.00"
    ws.cells(1, 10).Font.Color = RGB(255, 192, 0)
    ws.cells(1, 10).Font.Bold = True

    ws.cells(1, 11).Value = "POSITIONS"
    ws.cells(1, 11).Font.Color = RGB(150, 150, 150)
    ws.cells(1, 12).Value = posCount
    ws.cells(1, 12).Font.Color = RGB(221, 221, 221)

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

    ws.cells(2, 9).Value = "UNRL PNL%"
    ws.cells(2, 9).Font.Color = RGB(150, 150, 150)
    ws.cells(2, 10).Value = unrlPct
    ws.cells(2, 10).NumberFormat = "0.00%"
    ws.cells(2, 10).Font.Color = PnLColor(unrlPct)
    ws.cells(2, 10).Font.Bold = True

    ws.cells(2, 11).Value = "UNRL PNL"
    ws.cells(2, 11).Font.Color = RGB(150, 150, 150)
    ws.cells(2, 12).Value = totalUnrl
    ws.cells(2, 12).NumberFormat = "$#,##0"
    ws.cells(2, 12).Font.Color = PnLColor(totalUnrl)
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
```

## 7. Sanner.bas

`VB_Name = "Sanner"` · 644 行 · 32 KB · 原始編碼 CP950 (Big5)

<details><summary><b>成員索引</b></summary>

| 行 | 型別 | 可見性 | 宣告 |
|---:|---|---|---|
| 18 | Sub | Public | `Sub RunCompanyResearch(market As String, Sector As String)` |
| 190 | Sub | Public | `Sub FetchFundamentals(Ticker As String, ByRef pe As Double, ByRef fpe As Double, ByRef peg As Double)` |
| 251 | Function | Private | `Private Function ExtractNum(json As String, key As String) As Double` |
| 300 | Sub | Private | `Private Sub DrawScanSummary(ws As Worksheet, sumRow As Long, market As String, Sector As String, tickers As Variant, count As Long)` |
| 332 | Function | Private | `Private Function CalcMomentumScore(tech As TechIndicators) As Double` |
| 343 | Function | Private | `Private Function NormScore(val As Double, minV As Double, maxV As Double) As Double` |
| 353 | Sub | Private | `Private Sub WriteColoredPct(cell As Range, val As Double, fmt As String, positiveIsGood As Boolean)` |
| 381 | Sub | Private | `Private Sub ApplyScanTrend(cell As Range, tech As TechIndicators)` |
| 400 | Sub | Private | `Private Sub ApplyScoreColor(cell As Range, score As Double)` |
| 413 | Sub | Public | `Sub ExportTopToWatchlist()` |
| 460 | Sub | Public | `Sub ExportScanCSV()` |
| 482 | Sub | Public | `Sub OpenMarketScanner()` |
| 486 | Sub | Public | `Sub ColorizeRange(rng As Range)` |
| 506 | Function | Public | `Function GetSectorTickers(market As String, Sector As String) As Variant` |

</details>

```vba
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
```

## 8. SharedFunctions_Price.bas

`VB_Name = "SharedFunctions_Price"` · 186 行 · 6.4 KB · 原始編碼 CP950 (Big5)

<details><summary><b>成員索引</b></summary>

| 行 | 型別 | 可見性 | 宣告 |
|---:|---|---|---|
| 10 | Function | Private | `Private Function FetchTWSE(http As Object, exch As String) As Double` |
| 44 | Function | Private | `Private Function TWSEExtract(json As String, tag As String) As Double` |
| 66 | Function | Private | `Private Function FetchYahooClose(http As Object, Ticker As String) As Double` |
| 117 | Sub | Public | `Sub DebugPriceTest()` |

</details>

```vba
Option Explicit

' ================================================================
'  FetchTWSE
'  呼叫 TWSE/TPEX 即時 API，取得最新成交價（盤中）或收盤價
'  ex_ch 範例: "tse_2337.tw" 或 "otc_6510.tw"
' ================================================================
Private Function FetchTWSE(http As Object, exch As String) As Double

    Dim url As String
    url = "https://mis.twse.com.tw/stock/api/getStockInfo.jsp?ex_ch=" & _
          exch & "&json=1&delay=0&_=" & CStr(CLng((Now - DateSerial(1970, 1, 1)) * 86400))

    Dim response As String
    On Error Resume Next
    With http
        .Open "GET", url, False
        .setRequestHeader "User-Agent", "Mozilla/5.0"
        .setRequestHeader "Referer", "https://mis.twse.com.tw/"
        .setRequestHeader "Cache-Control", "no-cache"
        .send
        If .status = 200 Then response = .responseText
    End With
    On Error GoTo 0

    If Len(response) < 20 Then FetchTWSE = 0: Exit Function

    ' ── 優先取 "z"（最新成交價，盤中/收盤）──────────────────
    Dim price As Double
    price = TWSEExtract(response, """z"":""")

    ' ── 若 z 為 "-"（休市/未開盤），改取昨日收盤 "y" ─────────
    If price = 0 Then price = TWSEExtract(response, """y"":""")

    FetchTWSE = price
End Function

' ================================================================
'  TWSEExtract  從 TWSE JSON 取數字欄位
'  格式：  "z":"120.50"  或  "y":"121.00"
' ================================================================
Private Function TWSEExtract(json As String, tag As String) As Double
    Dim pos As Long: pos = InStr(json, tag)
    If pos = 0 Then TWSEExtract = 0: Exit Function

    pos = pos + Len(tag)
    Dim ep As Long: ep = pos
    Do While ep <= Len(json)
        Select Case Mid(json, ep, 1)
            Case """", ",", "}", "]": Exit Do
        End Select
        ep = ep + 1
    Loop

    Dim s As String: s = Trim(Mid(json, pos, ep - pos))
    If s = "-" Or s = "" Then TWSEExtract = 0: Exit Function
    If IsNumeric(s) Then TWSEExtract = CDbl(s)
End Function

' ================================================================
'  FetchYahooClose  使用 interval=1d&range=5d
'  從 close[] 陣列取最後一個有效正數
' ================================================================
Private Function FetchYahooClose(http As Object, Ticker As String) As Double

    Dim url As String
    url = "https://query2.finance.yahoo.com/v8/finance/chart/" & Ticker & "?interval=1d&range=5d"

    Dim response As String
    On Error Resume Next
    With http
        .Open "GET", url, False
        .setRequestHeader "User-Agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64)"
        .setRequestHeader "Referer", "https://finance.yahoo.com/"
        .setRequestHeader "Cache-Control", "no-cache"
        .send
        If .status = 200 Then response = .responseText
    End With
    On Error GoTo 0

    If Len(response) < 200 Then FetchYahooClose = 0: Exit Function

    ' 定位 quote 區塊內的 close[]
    Dim qPos As Long: qPos = InStr(response, """quote"":{")
    If qPos = 0 Then qPos = 1

    Dim csStart As Long: csStart = InStr(qPos, response, """close"":[")
    If csStart = 0 Then FetchYahooClose = 0: Exit Function
    csStart = csStart + 9

    Dim csEnd As Long: csEnd = InStr(csStart, response, "]")
    If csEnd <= csStart Then FetchYahooClose = 0: Exit Function

    Dim closeArr() As String
    closeArr = Split(Mid(response, csStart, csEnd - csStart), ",")

    Dim i As Long, price As Double
    For i = UBound(closeArr) To LBound(closeArr) Step -1
        Dim cv As String: cv = Trim(closeArr(i))
        ' 正確寫法：拆成兩層 If
    If IsNumeric(cv) Then
        If CDbl(cv) > 0 Then
            price = CDbl(cv)
            Exit For
        End If
    End If
    Next i

    FetchYahooClose = price
End Function

' ================================================================
'  DebugPriceTest  v6.0
' ================================================================
Sub DebugPriceTest()

    Dim testTicker As String
    testTicker = InputBox("輸入股票代號（如 2337 或 AAPL）:", "Price Debug v6", "2337")
    If testTicker = "" Then Exit Sub

    testTicker = UCase(Trim(testTicker))
    If IsNumeric(testTicker) Then testTicker = testTicker & ".TW"
    If Not g_PriceCache Is Nothing Then g_PriceCache.RemoveAll

    Dim http As Object
    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")

    Dim isTW As Boolean
    isTW = (InStr(testTicker, ".TW") > 0 Or InStr(testTicker, ".TWO") > 0)

    Dim twseResp As String, yahooResp As String
    Dim TWSEPrice As Double, yahooPrice As Double, finalPrice As Double
    Dim code As String

    If isTW Then
        code = Replace(Replace(testTicker, ".TWO", ""), ".TW", "")

        ' TWSE 請求
        Dim twseUrl As String
        twseUrl = "https://mis.twse.com.tw/stock/api/getStockInfo.jsp?ex_ch=tse_" & _
                  code & ".tw&json=1&delay=0"
        With http
            .Open "GET", twseUrl, False
            .setRequestHeader "User-Agent", "Mozilla/5.0"
            .setRequestHeader "Referer", "https://mis.twse.com.tw/"
            .send
            twseResp = .responseText
        End With
        TWSEPrice = FetchTWSE(http, "tse_" & code & ".tw")
    End If

    ' Yahoo 請求
    Dim yahooUrl As String
    yahooUrl = "https://query2.finance.yahoo.com/v8/finance/chart/" & testTicker & "?interval=1d&range=5d"
    With http
        .Open "GET", yahooUrl, False
        .setRequestHeader "User-Agent", "Mozilla/5.0"
        .setRequestHeader "Cache-Control", "no-cache"
        .send
        yahooResp = .responseText
    End With
    yahooPrice = FetchYahooClose(http, testTicker)

    finalPrice = GetStockPrice(testTicker)

    Dim info As String
    If isTW Then
        info = "【TWSE API】" & vbCr & _
               "URL    : tse_" & code & ".tw" & vbCr & _
               "價格   : " & TWSEPrice & vbCr & _
               "原始   : " & Left(twseResp, 300) & vbCr & vbCr
    End If
    info = info & "【Yahoo Finance】" & vbCr & _
           "URL    : " & yahooUrl & vbCr & _
           "Len    : " & Len(yahooResp) & " chars" & vbCr & _
           "價格   : " & yahooPrice & vbCr & vbCr & _
           "★ GetStockPrice 最終結果: " & finalPrice

    MsgBox info, vbInformation, "Price Debug v6: " & testTicker
    Set http = Nothing
End Sub
```

## 9. TaiwanPriceFetcher.bas

`VB_Name = "TaiwanPriceFetcher"` · 196 行 · 7.5 KB · 原始編碼 UTF-8

<details><summary><b>成員索引</b></summary>

| 行 | 型別 | 可見性 | 宣告 |
|---:|---|---|---|
| 50 | Function | Public | `Public Function GetTWStockPrice(ticker As String) As Double` |
| 64 | Sub | Public | `Public Sub ClearPriceCache()` |
| 73 | Function | Private | `Private Function TWSEPrice(code As String) As Double` |
| 93 | Function | Private | `Private Function TPExPrice(code As String) As Double` |
| 117 | Function | Private | `Private Function ParsePrice(json As String, keyField As String, code As String, valField As String) As Double` |
| 164 | Function | Private | `Private Function StripSuffix(ticker As String) As String` |
| 171 | Function | Private | `Private Function HttpGet(url As String) As String` |
| 192 | Sub | Private | `Private Sub PauseMS(ms As Long)` |

</details>

```vba
Option Explicit

' ================================================================
'  TAIWAN PRICE FETCHER  v1.1
'
'  Public Function GetTWStockPrice(ticker As String) As Double
'  Public Sub     ClearPriceCache()
'
'  專門處理台股（.TW / .TWO），不碰美股。
'  美股仍走 Attach 模組的 GetStockPrice（Yahoo Finance）。
'  兩模組 public 名稱不重疊，無衝突。
'
'  TickerInsight.bas 的 GetStockPriceSafe 負責路由：
'    .TW / .TWO suffix -> GetTWStockPrice (此模組)
'    其餘              -> GetStockPrice   (Attach 模組)
'
'  Session 快取:
'    STOCK_DAY_ALL 和 tpex_mainboard_daily_close_quotes 每次工作簿
'    開啟後只抓一次，同一個 session 內所有台股 ticker 共用快取。
'    Call ClearPriceCache() 可強制重新抓取（例如跨日後第一次查詢）。
' ================================================================

' ----- API 端點 -----
Private Const TWSE_URL As String = _
    "https://openapi.twse.com.tw/v1/exchangeReport/STOCK_DAY_ALL"

Private Const TPEX_URL As String = _
    "https://www.tpex.org.tw/openapi/v1/tpex_mainboard_daily_close_quotes"

' JSON 欄位名稱
' TWSE STOCK_DAY_ALL          -> Code / ClosingPrice
' TPEx mainboard_close_quotes -> SecuritiesCode / ClosingPrice
Private Const TWSE_CODE_KEY As String = "Code"
Private Const TPEX_CODE_KEY As String = "SecuritiesCode"
Private Const PRICE_KEY     As String = "ClosingPrice"

Private Const HTTP_RETRY   As Integer = 3
Private Const HTTP_WAIT_MS As Long = 800

' ----- Session 快取 (模組層級，工作簿關閉前持續有效) -----
Private g_twseJson  As String
Private g_tpexJson  As String
Private g_twseReady As Boolean
Private g_tpexReady As Boolean

' ================================================================
'  PUBLIC ENTRY POINT — 台股專用，非台股回傳 0
' ================================================================
Public Function GetTWStockPrice(ticker As String) As Double
    Dim t As String: t = UCase(Trim(CStr(ticker)))
    If Len(t) = 0 Then GetTWStockPrice = 0: Exit Function

    Dim code As String: code = StripSuffix(t)

    Select Case True
        Case InStr(t, ".TWO") > 0:  GetTWStockPrice = TPExPrice(code)
        Case InStr(t, ".TW") > 0:   GetTWStockPrice = TWSEPrice(code)
        Case Else:                   GetTWStockPrice = 0  ' 非台股請走 Attach 的 GetStockPrice
    End Select
End Function

' 強制清除快取，下次查詢時重新抓取
Public Sub ClearPriceCache()
    g_twseJson = "":  g_twseReady = False
    g_tpexJson = "":  g_tpexReady = False
End Sub

' ================================================================
'  TWSE (上市) — STOCK_DAY_ALL
'  JSON 格式: [{"Code":"2330","ClosingPrice":"900.00",...}, ...]
' ================================================================
Private Function TWSEPrice(code As String) As Double
    On Error GoTo Bail
    If Not g_twseReady Then
        Application.StatusBar = "TWSE: 抓取上市收盤行情..."
        g_twseJson = HttpGet(TWSE_URL)
        g_twseReady = (LenB(g_twseJson) > 0)   ' 只有成功取到資料才標 ready，失敗時允許下次重試
        Application.StatusBar = False
    End If
    If LenB(g_twseJson) > 0 Then
        TWSEPrice = ParsePrice(g_twseJson, TWSE_CODE_KEY, code, PRICE_KEY)
    End If
    Exit Function
Bail:
    TWSEPrice = 0
End Function

' ================================================================
'  TPEx (上櫃) — tpex_mainboard_daily_close_quotes
'  JSON 格式: [{"SecuritiesCode":"6547","ClosingPrice":"xx.xx",...}, ...]
' ================================================================
Private Function TPExPrice(code As String) As Double
    On Error GoTo Bail
    If Not g_tpexReady Then
        Application.StatusBar = "TPEx: 抓取上櫃收盤行情..."
        g_tpexJson = HttpGet(TPEX_URL)
        g_tpexReady = (LenB(g_tpexJson) > 0)   ' 只有成功取到資料才標 ready，失敗時允許下次重試
        Application.StatusBar = False
    End If
    If LenB(g_tpexJson) > 0 Then
        TPExPrice = ParsePrice(g_tpexJson, TPEX_CODE_KEY, code, PRICE_KEY)
    End If
    Exit Function
Bail:
    TPExPrice = 0
End Function

' ================================================================
'  JSON 解析器（適用於扁平陣列物件）
'
'  邏輯:
'  1. 在 JSON 中找到 "keyField":"code"
'  2. 從命中位置向前搜尋到當筆記錄結尾 "}"（避免抓到上一筆的欄位）
'  3. 在視窗內找到 "valField" 並解析數值（支援字串或數字欄位）
' ================================================================
Private Function ParsePrice(json As String, _
                             keyField As String, code As String, _
                             valField As String) As Double
    On Error GoTo Bail

    Dim needle As String: needle = """" & keyField & """:""" & code & """"
    Dim keyPos As Long:   keyPos = InStr(json, needle)
    If keyPos = 0 Then GoTo Bail

    ' 從 key 命中位置向後取到本筆記錄結尾 }（TWSE/TPEx 為簡單扁平物件，無巢狀）
    Dim afterKey As Long: afterKey = keyPos + Len(needle)
    Dim recEnd As Long:   recEnd = InStr(afterKey, json, "}")
    If recEnd = 0 Or recEnd > afterKey + 500 Then recEnd = afterKey + 500
    Dim win As String:    win = Mid(json, afterKey, recEnd - afterKey + 1)

    ' 嘗試字串值欄位: "valField":"<number>"
    Dim vfStr As String: vfStr = """" & valField & """:"""
    Dim vp As Long:      vp = InStr(win, vfStr)
    Dim raw As String

    If vp > 0 Then
        raw = Mid(win, vp + Len(vfStr), 30)
        raw = Split(raw, """")(0)                    ' 取到結尾 "
    Else
        ' 嘗試數字值欄位: "valField":<number>
        vfStr = """" & valField & """:"
        vp = InStr(win, vfStr)
        If vp = 0 Then GoTo Bail
        raw = Mid(win, vp + Len(vfStr), 30)
        raw = Split(Split(raw, ",")(0), "}")(0)      ' 取到 , 或 }
    End If

    raw = Trim(Replace(raw, ",", ""))                ' 移除千分位符號
    If IsNumeric(raw) And CDbl(raw) > 0 Then
        ParsePrice = CDbl(raw)
        Exit Function                                ' ← 必須在 Bail 前離開
    End If

Bail:
    ParsePrice = 0
End Function

' ================================================================
'  工具函式
' ================================================================

' 剝除 .TW / .TWO 後綴，取得純股票代號（輸入已 UCase）
Private Function StripSuffix(ticker As String) As String
    Dim n As String: n = ticker
    If Len(n) >= 4 And Right(n, 4) = ".TWO" Then n = Left(n, Len(n) - 4)
    If Len(n) >= 3 And Right(n, 3) = ".TW"  Then n = Left(n, Len(n) - 3)
    StripSuffix = n
End Function

Private Function HttpGet(url As String) As String
    On Error GoTo Bail
    Dim attempt As Integer
    For attempt = 1 To HTTP_RETRY
        Dim http As Object: Set http = CreateObject("MSXML2.XMLHTTP")
        On Error Resume Next
        http.Open "GET", url, False
        http.setRequestHeader "User-Agent", "Mozilla/5.0 (Windows NT 10.0)"
        http.setRequestHeader "Accept",     "application/json"
        http.send
        Dim ok As Boolean: ok = (Err.Number = 0)
        On Error GoTo Bail
        If ok And http.status = 200 Then HttpGet = http.responseText: Exit Function
        Set http = Nothing
        On Error GoTo Bail
        If attempt < HTTP_RETRY Then PauseMS HTTP_WAIT_MS * attempt
    Next attempt
Bail:
    HttpGet = ""
End Function

Private Sub PauseMS(ms As Long)
    Dim t As Single: t = Timer + ms / 1000
    Do While Timer < t: DoEvents: Loop
End Sub
```

## 10. TickerInsight.bas

`VB_Name = "TickerInsight"` · 1179 行 · 42.7 KB · 原始編碼 ASCII

<details><summary><b>成員索引</b></summary>

| 行 | 型別 | 可見性 | 宣告 |
|---:|---|---|---|
| 110 | Sub | Public | `Public Sub OpenTickerInsight()` |
| 144 | Sub | Public | `Public Sub DiagnoseTicker(Optional ByVal forcedTicker As String = "")` |
| 284 | Sub | Public | `Public Sub TickerInsight_OnChange(ByVal Target As Range)` |
| 298 | Sub | Public | `Public Sub RefreshTickerInsight(ByVal ticker As String)` |
| 370 | Sub | Public | `Public Sub RefreshTickerProjection()` |
| 404 | Function | Private | `Private Function EnsureTickerInsightSheet() As Worksheet` |
| 418 | Sub | Private | `Private Sub InitSheetLayout(ws As Worksheet)` |
| 464 | Sub | Private | `Private Sub DrawShellHeaders(ws As Worksheet, ticker As String)` |
| 559 | Sub | Private | `Private Sub ClearTickerData(ws As Worksheet)` |
| 590 | Sub | Private | `Private Sub RemoveSheetShapes(ws As Worksheet)` |
| 606 | Sub | Private | `Private Sub BuildFIFOHistory(ByVal ticker As String, ByRef records() As TIRecord, ByRef recCount As Long, ByRef lots() As TILot, ByRef lotCount As Long)` |
| 768 | Sub | Private | `Private Sub DrawActivePosition(ws As Worksheet, ticker As String, lots() As TILot, lotCount As Long)` |
| 819 | Sub | Private | `Private Sub DrawLifetimeMetrics(ws As Worksheet, ticker As String, records() As TIRecord, recCount As Long)` |
| 878 | Sub | Private | `Private Sub DrawProjection(ws As Worksheet, ticker As String, lots() As TILot, lotCount As Long)` |
| 923 | Sub | Private | `Private Sub DrawTradeHistory(ws As Worksheet, ticker As String, records() As TIRecord, recCount As Long)` |
| 1003 | Sub | Private | `Private Sub WriteLabel(ws As Worksheet, r As Long, c As Long, txt As String)` |
| 1020 | Sub | Private | `Private Sub WriteValueMoney(ws As Worksheet, r As Long, c As Long, v As Double, sym As String)` |
| 1033 | Sub | Private | `Private Sub WriteValuePrice(ws As Worksheet, r As Long, c As Long, v As Double)` |
| 1045 | Sub | Private | `Private Sub WriteValuePnL(ws As Worksheet, r As Long, c As Long, v As Double, sym As String)` |
| 1061 | Function | Private | `Private Function MoneyFmt(sym As String) As String` |
| 1066 | Function | Private | `Private Function PnLFmt(sym As String) As String` |
| 1071 | Sub | Private | `Private Sub WriteValuePct(ws As Worksheet, r As Long, c As Long, v As Double)` |
| 1083 | Sub | Private | `Private Sub WriteDash(ws As Worksheet, r As Long, c As Long)` |
| 1096 | Function | Private | `Private Function USPnLColor(v As Double) As Long` |
| 1112 | Function | Private | `Private Function NormalizeTicker(t As String) As String` |
| 1130 | Function | Private | `Private Function IsTWTicker(ticker As String) As Boolean` |
| 1137 | Function | Private | `Private Function GetPriceSymbol(ticker As String) As String` |
| 1144 | Function | Private | `Private Function GetFXToTWD(ticker As String) As Double` |
| 1148 | Function | Private | `Private Function GetStockPriceSafe(ticker As String) As Double` |
| 1163 | Function | Private | `Private Function SafeDate(v As Variant) As Date` |
| 1173 | Function | Private | `Private Function SafeNum(v As Variant) As Double` |

</details>

```vba
Option Explicit

' ================================================================
' TICKER INSIGHT v1.0
' Per-ticker FIFO trade history & performance dashboard
' Bloomberg terminal style, US color convention (green +, red -)
'
' ENTRY POINTS:
'   Sub OpenTickerInsight()          - call from Alt+F8 or button
'   Sub RefreshTickerInsight(ticker) - full rebuild for a ticker
'   Sub RefreshTickerProjection      - recompute PROJECTION block only
'   Sub TickerInsight_OnChange(t)    - dispatcher for sheet events
'
' WORKSHEET LAYOUT (auto-created on first run):
'   Row  1: TICKER <GO>: [B1 input]
'   Row  3: 3 section titles  (A=ACTIVE | D=LIFETIME | G=PROJECTION)
'   Row  4: NET EXPOSURE   | REALIZED PNL  | LAST PRICE
'   Row  5: ENTRY (W.AP)   | LIFETIME EFF  | PRICE TARGET [H5 input]
'   Row  6: LAST PRICE     | VELOCITY/DAY  | PROJECTED PNL
'   Row  7: UNREALIZED PNL |               | PROJECTED RETURN %
'   Row  8: RETURN %       |               |
'   Row 10: 4) RECENT TRADE HISTORY
'   Row 11: ENTRY | EXIT | DAYS | SHARES | PNL | RETURN % | BROKER
'   Row 12+: ... newest-first ...
'
' INSTALLATION:
'   1. VBE > File > Import File > TickerInsight.bas (this file)
'   2. Run OpenTickerInsight (Alt+F8) - "Ticker Insight" sheet created
'   3. In VBE left panel, double-click the "Ticker Insight" sheet
'      under "Microsoft Excel Objects". Paste this Worksheet_Change:
'
'        Private Sub Worksheet_Change(ByVal Target As Range)
'            Dim hitB1 As Boolean
'            Dim hitH5 As Boolean
'            hitB1 = Not Intersect(Target, Me.Range("B1")) Is Nothing
'            hitH5 = Not Intersect(Target, Me.Range("H5")) Is Nothing
'            If Not (hitB1 Or hitH5) Then Exit Sub
'            If Target.Count > 1 Then Exit Sub
'            If hitB1 And Trim(Me.Range("B1").Value) = "" Then Exit Sub
'
'            Application.EnableEvents = False
'            On Error GoTo ErrorHandler
'
'            If hitB1 Then
'                Call RefreshTickerInsight(CStr(Me.Range("B1").Value))
'            ElseIf hitH5 Then
'                Call RefreshTickerProjection
'            End If
'
'        ErrorHandler:
'            Application.EnableEvents = True
'            If Err.Number <> 0 Then
'                MsgBox "Ticker Insight error: " & Err.Description, _
'                       vbCritical, "Err " & Err.Number
'            End If
'        End Sub
'
'   4. Type a ticker in B1 (e.g. CRDO, 2330) and optionally H5 target.
' ================================================================

Private Const SH_TI   As String = "Ticker Insight"
Private Const SH_TR   As String = "Transactions"
Private Const SH_REAL As String = "Realized"
Private Const SH_RR4  As String = "RR4"

' Transaction column letters (from existing schema)
Private Const COL_DATE   As String = "B"
Private Const COL_TICKER As String = "C"
Private Const COL_ACTION As String = "D"
Private Const COL_SHARES As String = "E"
Private Const COL_NETAMT As String = "I"
Private Const COL_BROKER As String = "N"

' Exchange rate: USD -> TWD. All aggregate PnL / exposure values for
' US tickers are multiplied by this constant before display. Per-share
' prices (Entry, Last, Price Target) stay in NATIVE currency.
Private Const FX_USD_TWD As Double = 31.6

' ----------------------------------------------------------------
' One FIFO-paired trade record (one BUY lot matched to one SELL)
' ----------------------------------------------------------------
Private Type TIRecord
    EntryDate As Date
    ExitDate  As Date
    DaysHeld  As Long
    Shares    As Double
    Cost      As Double
    Proceeds  As Double
    PnL       As Double
    ReturnPct As Double
    Broker    As String
End Type

' ----------------------------------------------------------------
' One residual lot still open after all transactions processed
' ----------------------------------------------------------------
Private Type TILot
    LotDate      As Date
    Shares       As Double
    CostPerShare As Double
    Broker       As String
End Type

' ================================================================
' PUBLIC ENTRY: open the dashboard (auto-creates sheet)
' Event hooking is done via Worksheet_Change pasted on the sheet's
' code page (see install snippet in module header).
' ================================================================
Public Sub OpenTickerInsight()
    Dim ws As Worksheet
    Set ws = EnsureTickerInsightSheet()

    ' One-shot defensive cleanup: nuke any stray shapes/conditional
    ' formats inherited from previous versions of the sheet.
    Application.EnableEvents = False
    Call RemoveSheetShapes(ws)
    On Error Resume Next
    ws.Cells.FormatConditions.Delete
    ws.AutoFilterMode = False
    On Error GoTo 0
    Application.EnableEvents = True

    ws.Activate
    ws.Range("B1").Select

    ' If sheet already has a ticker, do a full refresh.
    ' Otherwise just redraw the empty shell so the user always sees structure.
    Dim existing As String: existing = Trim(CStr(ws.Range("B1").Value))
    If existing <> "" Then
        Call RefreshTickerInsight(existing)
    Else
        Application.EnableEvents = False
        Call DrawShellHeaders(ws, vbNullString)
        Application.EnableEvents = True
    End If
End Sub

' ================================================================
' DIAGNOSTIC: dumps the state of the Ticker Insight pipeline for the
' ticker currently in B1 (or a passed-in one). Useful when no values
' appear in the dashboard. Run via Alt+F8 -> DiagnoseTicker.
' ================================================================
Public Sub DiagnoseTicker(Optional ByVal forcedTicker As String = "")
    Dim msg As String
    Dim lf As String: lf = vbCrLf

    On Error GoTo ErrTrap

    ' --- 1. Sheets ---
    msg = msg & "=== SHEET CHECK ===" & lf
    Dim wsTI As Worksheet, wsTr As Worksheet
    On Error Resume Next
    Set wsTI = ThisWorkbook.Sheets(SH_TI)
    Set wsTr = ThisWorkbook.Sheets(SH_TR)
    On Error GoTo ErrTrap

    msg = msg & "Ticker Insight: " & IIf(wsTI Is Nothing, "MISSING", "OK") & lf
    msg = msg & "Transactions  : " & IIf(wsTr Is Nothing, "MISSING", "OK") & lf

    If wsTr Is Nothing Then GoTo ShowMsg

    ' --- 2. Ticker resolution ---
    msg = msg & lf & "=== TICKER ===" & lf
    Dim ticker As String
    If forcedTicker <> "" Then
        ticker = forcedTicker
    ElseIf Not (wsTI Is Nothing) Then
        ticker = CStr(wsTI.Range("B1").Value)
    End If
    ticker = UCase(Trim(ticker))

    If ticker = "" Then
        msg = msg & "(B1 empty - type a ticker then re-run, or DiagnoseTicker " & _
              Chr(34) & "CRDO" & Chr(34) & ")" & lf
        GoTo ShowMsg
    End If

    msg = msg & "Input         : [" & ticker & "]" & lf
    msg = msg & "Normalized    : [" & NormalizeTicker(ticker) & "]" & lf
    msg = msg & "Is TW ticker  : " & IsTWTicker(ticker) & lf
    msg = msg & "FX multiplier : " & GetFXToTWD(ticker) & lf
    msg = msg & "Live price    : " & GetStockPriceSafe(ticker) & lf

    ' --- 3. Transactions scan ---
    msg = msg & lf & "=== TRANSACTIONS SCAN ===" & lf
    Dim lastRow As Long
    lastRow = wsTr.Cells(wsTr.Rows.Count, "A").End(xlUp).Row
    msg = msg & "Total rows    : " & (lastRow - 1) & lf

    Dim normT As String: normT = NormalizeTicker(ticker)
    Dim matchCount As Long, sample As String, samples As Long
    Dim r As Long
    For r = 2 To lastRow
        Dim cellT As String: cellT = CStr(wsTr.Cells(r, COL_TICKER).Value)
        If NormalizeTicker(cellT) = normT Then
            matchCount = matchCount + 1
            If samples < 5 Then
                sample = sample & "  R" & r & ":  " & _
                    CStr(wsTr.Cells(r, COL_DATE).Value) & "  " & _
                    "[" & cellT & "]  " & _
                    CStr(wsTr.Cells(r, COL_ACTION).Value) & "  " & _
                    "sh=" & CStr(wsTr.Cells(r, COL_SHARES).Value) & "  " & _
                    "net=" & CStr(wsTr.Cells(r, COL_NETAMT).Value) & "  " & _
                    "brk=" & CStr(wsTr.Cells(r, COL_BROKER).Value) & lf
                samples = samples + 1
            End If
        End If
    Next r

    msg = msg & "Matching rows : " & matchCount & lf
    If matchCount > 0 Then
        msg = msg & lf & "Sample rows (up to 5):" & lf & sample
    Else
        msg = msg & "(none -- check ticker spelling and column C of Transactions)" & lf
        msg = msg & "Hint: Transactions C column normalised forms of first 10 tickers:" & lf
        Dim cnt As Long: cnt = 0
        For r = 2 To lastRow
            If cnt < 10 Then
                Dim t2 As String: t2 = CStr(wsTr.Cells(r, COL_TICKER).Value)
                If t2 <> "" Then
                    msg = msg & "  [" & t2 & "] -> [" & NormalizeTicker(t2) & "]" & lf
                    cnt = cnt + 1
                End If
            End If
        Next r
    End If

    ' --- 4. FIFO build ---
    msg = msg & lf & "=== FIFO RESULT ===" & lf
    Dim records() As TIRecord, recCount As Long
    Dim lots() As TILot, lotCount As Long
    Call BuildFIFOHistory(ticker, records, recCount, lots, lotCount)
    msg = msg & "recCount      : " & recCount & "  (closed FIFO pairs)" & lf
    msg = msg & "lotCount      : " & lotCount & "  (open lots = active pos)" & lf

    If recCount > 0 Then
        msg = msg & lf & "First record:" & lf
        msg = msg & "  Entry=" & records(1).EntryDate & "  Exit=" & records(1).ExitDate & lf
        msg = msg & "  Days=" & records(1).DaysHeld & "  Shares=" & records(1).Shares & lf
        msg = msg & "  Cost=" & records(1).Cost & "  Proceeds=" & records(1).Proceeds & lf
        msg = msg & "  PnL=" & records(1).PnL & "  Ret%=" & records(1).ReturnPct & lf
        msg = msg & "  Broker=" & records(1).Broker & lf
    End If

    If lotCount > 0 Then
        msg = msg & lf & "First open lot:" & lf
        msg = msg & "  Date=" & lots(1).LotDate & "  Shares=" & lots(1).Shares & lf
        msg = msg & "  CPS=" & lots(1).CostPerShare & "  Broker=" & lots(1).Broker & lf
    End If

    ' --- 5. Format probe ---
    msg = msg & lf & "=== FORMAT PROBE ===" & lf
    msg = msg & "MoneyFmt(NT$) = " & MoneyFmt("NT$") & lf
    msg = msg & "PnLFmt(NT$)   = " & PnLFmt("NT$") & lf

    ' --- 6. DrawLifetimeMetrics dry-run ---
    If Not (wsTI Is Nothing) Then
        msg = msg & lf & "=== DRAW PROBE ===" & lf
        On Error Resume Next
        Err.Clear
        Call DrawLifetimeMetrics(wsTI, ticker, records, recCount)
        If Err.Number <> 0 Then
            msg = msg & "DrawLifetimeMetrics: ERR " & Err.Number & " - " & Err.Description & lf
        Else
            msg = msg & "DrawLifetimeMetrics: OK (no error)" & lf
        End If
        On Error GoTo ErrTrap
    End If

ShowMsg:
    MsgBox msg, vbInformation, "Ticker Insight Diagnostic"
    Exit Sub

ErrTrap:
    msg = msg & lf & lf & "*** DIAGNOSTIC EXCEPTION ***" & lf & _
          "Err " & Err.Number & ": " & Err.Description
    MsgBox msg, vbCritical, "Ticker Insight Diagnostic - error"
End Sub

' ================================================================
' DISPATCHER: called from ThisWorkbook.Workbook_SheetChange
' ================================================================
Public Sub TickerInsight_OnChange(ByVal Target As Range)
    On Error Resume Next
    If Target.CountLarge > 1 Then Exit Sub
    Select Case Target.Address
        Case "$B$1"
            Call RefreshTickerInsight(CStr(Target.Value))
        Case "$H$5"
            Call RefreshTickerProjection
    End Select
End Sub

' ================================================================
' FULL REFRESH: rebuild every section for a ticker
' ================================================================
Public Sub RefreshTickerInsight(ByVal ticker As String)
    ticker = UCase(Trim(ticker))
    If ticker = "" Then Exit Sub

    Dim ws As Worksheet: Set ws = EnsureTickerInsightSheet()

    Application.EnableEvents = False
    Application.ScreenUpdating = False
    On Error GoTo ErrTrap

    ' --- Detect ticker change via hidden A2 tracker ------------
    ' A2 stores the PREVIOUS ticker (invisible: black-on-black).
    ' B1 has already been updated by the user, so we can't use it.
    Dim oldTicker As String
    oldTicker = UCase(Trim(CStr(ws.Cells(2, 1).Value)))
    Dim savedTarget As Variant: savedTarget = ws.Cells(5, 8).Value

    Call ClearTickerData(ws)
    Call DrawShellHeaders(ws, ticker)

    ' If user is refreshing the SAME ticker, preserve PRICE TARGET.
    ' If they switched tickers, leave H5 empty (set by shell).
    If oldTicker = ticker And IsNumeric(savedTarget) Then
        If CDbl(savedTarget) > 0 Then ws.Cells(5, 8).Value = CDbl(savedTarget)
    End If

    Dim records() As TIRecord, recCount As Long
    Dim lots() As TILot, lotCount As Long
    Call BuildFIFOHistory(ticker, records, recCount, lots, lotCount)

    If recCount = 0 And lotCount = 0 Then
        With ws.Cells(4, 1)
            .Value = "No transactions found for [" & ticker & "]"
            .Font.Color = RGB(255, 80, 80)
            .Font.Bold = True
            .Font.Italic = True
        End With
        GoTo StampTicker
    End If

    Call DrawActivePosition(ws, ticker, lots, lotCount)
    Call DrawLifetimeMetrics(ws, ticker, records, recCount)
    Call DrawProjection(ws, ticker, lots, lotCount)
    Call DrawTradeHistory(ws, ticker, records, recCount)

StampTicker:
    ' Stamp current ticker into invisible A2 for next refresh's compare
    With ws.Cells(2, 1)
        .Value = ticker
        .Font.Color = RGB(0, 0, 0)
        .Interior.Color = RGB(0, 0, 0)
    End With

Done:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Exit Sub

ErrTrap:
    ' Surface errors instead of silently swallowing - vital for debug
    Dim errN As Long: errN = Err.Number
    Dim errD As String: errD = Err.Description
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    MsgBox "RefreshTickerInsight error" & vbCrLf & vbCrLf & _
           "Err " & errN & ": " & errD & vbCrLf & _
           "Ticker: " & ticker, vbCritical, "Ticker Insight"
End Sub

' ================================================================
' LIGHT REFRESH: recompute PROJECTION only (price target changed)
' ================================================================
Public Sub RefreshTickerProjection()
    Dim ws As Worksheet
    On Error Resume Next: Set ws = ThisWorkbook.Sheets(SH_TI): On Error GoTo 0
    If ws Is Nothing Then Exit Sub

    Dim ticker As String: ticker = UCase(Trim(CStr(ws.Range("B1").Value)))
    If ticker = "" Then Exit Sub

    Application.EnableEvents = False
    Application.ScreenUpdating = False
    On Error GoTo ProjErr

    Dim records() As TIRecord, recCount As Long
    Dim lots() As TILot, lotCount As Long
    Call BuildFIFOHistory(ticker, records, recCount, lots, lotCount)
    Call DrawProjection(ws, ticker, lots, lotCount)

ProjDone:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Exit Sub

ProjErr:
    Dim errN As Long: errN = Err.Number
    Dim errD As String: errD = Err.Description
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    MsgBox "RefreshTickerProjection error" & vbCrLf & vbCrLf & _
           "Err " & errN & ": " & errD, vbCritical, "Ticker Insight"
End Sub

' ================================================================
' SHEET MANAGEMENT
' ================================================================
Private Function EnsureTickerInsightSheet() As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(SH_TI)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add( _
            After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        ws.Name = SH_TI
        Call InitSheetLayout(ws)
    End If
    Set EnsureTickerInsightSheet = ws
End Function

Private Sub InitSheetLayout(ws As Worksheet)
    ' Suppress events around shell setup so writing B1 doesn't recursively
    ' trigger Workbook_SheetChange while the sheet is half-built.
    Dim prevEvents As Boolean
    prevEvents = Application.EnableEvents
    Application.EnableEvents = False

    With ws.Cells
        .Interior.Color = RGB(0, 0, 0)
        .Font.Color = RGB(221, 221, 221)
        .Font.Name = "Consolas"
        .Font.Size = 10
    End With

    ws.Activate
    On Error Resume Next
    ActiveWindow.DisplayGridlines = False
    On Error GoTo 0

    ' Column widths: 3 metric groups + history area
    ws.Columns("A").ColumnWidth = 22   ' Label / Entry date
    ws.Columns("B").ColumnWidth = 16   ' Value / Exit date
    ws.Columns("C").ColumnWidth = 8    ' Days
    ws.Columns("D").ColumnWidth = 22   ' Label / Shares
    ws.Columns("E").ColumnWidth = 16   ' Value / PnL
    ws.Columns("F").ColumnWidth = 10   ' Return %
    ws.Columns("G").ColumnWidth = 22   ' Label / Broker
    ws.Columns("H").ColumnWidth = 16   ' Value

    On Error Resume Next
    ws.Tab.Color = RGB(255, 192, 0)
    On Error GoTo 0

    ' Draw the full static shell immediately so the user sees structure
    ' (even before they type a ticker). RefreshTickerInsight will repaint
    ' this same shell + fill in values.
    Call DrawShellHeaders(ws, vbNullString)

    Application.EnableEvents = prevEvents
End Sub

' ----------------------------------------------------------------
' Draws every STATIC element of the dashboard (labels, section
' titles, history table header). Called from InitSheetLayout (empty
' ticker) and from RefreshTickerInsight (with current ticker).
' ----------------------------------------------------------------
Private Sub DrawShellHeaders(ws As Worksheet, ticker As String)
    ' --- Row 1: TICKER <GO>: input ---
    With ws.Cells(1, 1)
        .Value = "TICKER <GO>:"
        .Font.Color = RGB(255, 192, 0)
        .Font.Bold = True
        .Font.Size = 11
    End With
    With ws.Cells(1, 2)
        .Value = ticker
        .Interior.Color = RGB(255, 192, 0)
        .Font.Color = RGB(0, 0, 0)
        .Font.Bold = True
        .Font.Size = 11
        .HorizontalAlignment = xlCenter
    End With
    ws.Rows(1).RowHeight = 22

    ' --- Row 3: section titles ---
    With ws.Cells(3, 1)
        .Value = "1) ACTIVE POSITION"
        .Font.Color = RGB(255, 192, 0)
        .Font.Bold = True
    End With
    With ws.Cells(3, 4)
        .Value = "2) LIFETIME METRICS"
        .Font.Color = RGB(255, 192, 0)
        .Font.Bold = True
    End With
    With ws.Cells(3, 7)
        .Value = "3) PROJECTION"
        .Font.Color = RGB(255, 192, 0)
        .Font.Bold = True
    End With
    ws.Rows(3).RowHeight = 20

    ' --- Row 4-8: metric labels (cyan) ---
    Call WriteLabel(ws, 4, 1, "NET EXPOSURE")
    Call WriteLabel(ws, 5, 1, "ENTRY PRICE (W.AP)")
    Call WriteLabel(ws, 6, 1, "LAST PRICE")
    Call WriteLabel(ws, 7, 1, "UNREALIZED PNL")
    Call WriteLabel(ws, 8, 1, "RETURN %")

    Call WriteLabel(ws, 4, 4, "REALIZED PNL")
    Call WriteLabel(ws, 5, 4, "LIFETIME EFF")
    Call WriteLabel(ws, 6, 4, "VELOCITY/DAY")

    Call WriteLabel(ws, 4, 7, "LAST PRICE")
    Call WriteLabel(ws, 5, 7, "PRICE TARGET")
    Call WriteLabel(ws, 6, 7, "PROJECTED PNL")
    Call WriteLabel(ws, 7, 7, "PROJECTED RETURN %")

    ' Row heights for metric block
    Dim rh As Long
    For rh = 4 To 8: ws.Rows(rh).RowHeight = 18: Next rh

    ' --- Row 10: history title ---
    With ws.Cells(10, 1)
        .Value = "4) RECENT TRADE HISTORY"
        .Font.Color = RGB(255, 192, 0)
        .Font.Bold = True
    End With
    ws.Rows(10).RowHeight = 20

    ' --- Row 11: history table header ---
    Dim hdrs As Variant
    hdrs = Array("ENTRY DATE", "EXIT DATE", "DAYS", "SHARES", _
                 "REALIZED PNL", "RETURN %", "BROKER")
    Dim c As Long
    For c = 0 To UBound(hdrs)
        With ws.Cells(11, c + 1)
            .Value = hdrs(c)
            .Font.Color = RGB(150, 150, 150)
            .Font.Bold = True
            .HorizontalAlignment = xlLeft
        End With
    Next c
    With ws.Range(ws.Cells(11, 1), ws.Cells(11, UBound(hdrs) + 1)).Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Color = RGB(255, 192, 0)
        .Weight = xlThin
    End With
    ws.Rows(11).RowHeight = 18

    ' --- Yellow background for H5 (PRICE TARGET input) ---
    With ws.Cells(5, 8)
        .Interior.Color = RGB(255, 192, 0)
        .Font.Color = RGB(0, 0, 0)
        .Font.Bold = True
        .Font.Size = 11
        .NumberFormat = "0.00"
        .HorizontalAlignment = xlCenter
    End With
End Sub

Private Sub ClearTickerData(ws As Worksheet)
    ' Wipe everything except row 1 (header) and row 2 (gutter)
    With ws.Range(ws.Cells(3, 1), ws.Cells(500, 8))
        .ClearContents
        .Interior.Color = RGB(0, 0, 0)
        .Font.Color = RGB(221, 221, 221)
        .Font.Bold = False
        .Font.Italic = False
        .Font.Size = 10
        .NumberFormat = "General"
        .HorizontalAlignment = xlGeneral
        .Borders.LineStyle = xlNone
        ' Strip any leftover conditional formatting (icon sets, data bars)
        .FormatConditions.Delete
    End With

    ' Remove ALL shapes (charts, autoshapes, sparklines, smart tags etc.)
    ' that may have been left on the sheet. We own this sheet entirely.
    Call RemoveSheetShapes(ws)

    ' Turn off any AutoFilter that may have been applied
    On Error Resume Next
    ws.AutoFilterMode = False
    On Error GoTo 0
End Sub

' ----------------------------------------------------------------
' Deletes every Shape on the worksheet (charts, autoshapes, comment
' callouts, smart tags, sparklines). Loops backwards to avoid index
' shifting as items are removed.
' ----------------------------------------------------------------
Private Sub RemoveSheetShapes(ws As Worksheet)
    On Error Resume Next
    Dim i As Long
    For i = ws.Shapes.Count To 1 Step -1
        ws.Shapes(i).Delete
    Next i
    On Error GoTo 0
End Sub

' ================================================================
' FIFO HISTORY BUILDER
' Walks Transactions chronologically per (ticker, broker).
' BUY pushes a lot. SELL pops lots from the front and emits one
' TIRecord per lot consumed (so one SELL that eats 3 lots = 3 rows).
' ADJUSTCOST distributes per share across all current lots.
' ================================================================
Private Sub BuildFIFOHistory(ByVal ticker As String, _
                             ByRef records() As TIRecord, _
                             ByRef recCount As Long, _
                             ByRef lots() As TILot, _
                             ByRef lotCount As Long)
    recCount = 0
    lotCount = 0
    ReDim records(1 To 256)
    ReDim lots(1 To 64)

    Dim wsTr As Worksheet
    On Error Resume Next: Set wsTr = ThisWorkbook.Sheets(SH_TR): On Error GoTo 0
    If wsTr Is Nothing Then Exit Sub

    Dim lastRow As Long
    lastRow = wsTr.Cells(wsTr.Rows.Count, "A").End(xlUp).Row
    If lastRow < 2 Then Exit Sub

    ' Collect rows for this ticker (normalised so 2330 / 2330.TW / 2330.TWO match)
    Dim tickerNorm As String: tickerNorm = NormalizeTicker(ticker)
    Dim tRows() As Long, tCount As Long
    ReDim tRows(1 To lastRow)
    Dim r As Long
    For r = 2 To lastRow
        If NormalizeTicker(CStr(wsTr.Cells(r, COL_TICKER).Value)) = tickerNorm Then
            tCount = tCount + 1
            tRows(tCount) = r
        End If
    Next r
    If tCount = 0 Then Exit Sub

    ' Sort by date ascending (insertion sort - tCount usually small)
    Dim i As Long, j As Long, tmp As Long
    For i = 1 To tCount - 1
        For j = i + 1 To tCount
            If SafeDate(wsTr.Cells(tRows(i), COL_DATE).Value) > _
               SafeDate(wsTr.Cells(tRows(j), COL_DATE).Value) Then
                tmp = tRows(i): tRows(i) = tRows(j): tRows(j) = tmp
            End If
        Next j
    Next i

    ' Per-broker FIFO lot queue
    Dim queues As Object: Set queues = CreateObject("Scripting.Dictionary")

    Dim k As Long
    For k = 1 To tCount
        r = tRows(k)
        Dim tDate As Date: tDate = SafeDate(wsTr.Cells(r, COL_DATE).Value)
        Dim action As String: action = UCase(Trim(CStr(wsTr.Cells(r, COL_ACTION).Value)))
        Dim shares As Double: shares = SafeNum(wsTr.Cells(r, COL_SHARES).Value)
        Dim netAmt As Double: netAmt = SafeNum(wsTr.Cells(r, COL_NETAMT).Value)
        Dim broker As String: broker = Trim(CStr(wsTr.Cells(r, COL_BROKER).Value))
        If broker = "" Then broker = "Default"

        If shares <= 0 And action <> "ADJUSTCOST" Then GoTo NextTx

        If Not queues.Exists(broker) Then queues.Add broker, New Collection
        Dim queue As Collection: Set queue = queues(broker)

        Select Case action
            Case "BUY"
                Dim cps As Double
                If shares > 0 Then cps = netAmt / shares Else cps = 0
                queue.Add Array(tDate, shares, cps)

            Case "SELL"
                Dim sellPx As Double
                If shares > 0 Then sellPx = netAmt / shares Else sellPx = 0
                Dim remaining As Double: remaining = shares
                Do While remaining > 0.0000001 And queue.Count > 0
                    Dim frontLot As Variant: frontLot = queue(1)
                    Dim lotDate As Date: lotDate = CDate(frontLot(0))
                    Dim lotShares As Double: lotShares = CDbl(frontLot(1))
                    Dim lotCps As Double: lotCps = CDbl(frontLot(2))

                    Dim consumed As Double
                    If lotShares <= remaining + 0.0000001 Then
                        consumed = lotShares
                    Else
                        consumed = remaining
                    End If

                    ' Emit a FIFO trade record
                    recCount = recCount + 1
                    If recCount > UBound(records) Then _
                        ReDim Preserve records(1 To UBound(records) * 2)
                    With records(recCount)
                        .EntryDate = lotDate
                        .ExitDate = tDate
                        .DaysHeld = CLng(tDate - lotDate)
                        .Shares = consumed
                        .Cost = consumed * lotCps
                        .Proceeds = consumed * sellPx
                        .PnL = .Proceeds - .Cost
                        If .Cost <> 0 Then
                            .ReturnPct = .PnL / .Cost
                        Else
                            .ReturnPct = 0
                        End If
                        .Broker = broker
                    End With

                    ' Update / pop the lot
                    If consumed >= lotShares - 0.0000001 Then
                        queue.Remove 1
                    Else
                        queue.Add Array(lotDate, lotShares - consumed, lotCps), Before:=1
                        queue.Remove 2
                    End If
                    remaining = remaining - consumed
                Loop

            Case "ADJUSTCOST"
                Dim totSh As Double: totSh = 0
                Dim li As Long
                For li = 1 To queue.Count
                    totSh = totSh + CDbl(queue(li)(1))
                Next li
                If totSh > 0 Then
                    Dim adjPerSh As Double: adjPerSh = netAmt / totSh
                    For li = 1 To queue.Count
                        Dim adjLot As Variant: adjLot = queue(li)
                        Dim newCps As Double: newCps = CDbl(adjLot(2)) + adjPerSh
                        queue.Remove li
                        If li > queue.Count Then
                            queue.Add Array(adjLot(0), adjLot(1), newCps)
                        Else
                            queue.Add Array(adjLot(0), adjLot(1), newCps), Before:=li
                        End If
                    Next li
                End If
        End Select
NextTx:
    Next k

    If recCount > 0 Then ReDim Preserve records(1 To recCount)

    ' Flatten residual lots (active position)
    Dim brokerKey As Variant
    For Each brokerKey In queues.Keys
        Dim bq As Collection: Set bq = queues(brokerKey)
        Dim li2 As Long
        For li2 = 1 To bq.Count
            lotCount = lotCount + 1
            If lotCount > UBound(lots) Then _
                ReDim Preserve lots(1 To UBound(lots) * 2)
            With lots(lotCount)
                .LotDate = CDate(bq(li2)(0))
                .Shares = CDbl(bq(li2)(1))
                .CostPerShare = CDbl(bq(li2)(2))
                .Broker = CStr(brokerKey)
            End With
        Next li2
    Next brokerKey

    If lotCount > 0 Then ReDim Preserve lots(1 To lotCount)
End Sub

' ================================================================
' SECTION 1: ACTIVE POSITION (columns A-B)
' ================================================================
Private Sub DrawActivePosition(ws As Worksheet, ticker As String, _
                                lots() As TILot, lotCount As Long)
    Dim totalShares As Double, totalCost As Double
    Dim i As Long
    For i = 1 To lotCount
        totalShares = totalShares + lots(i).Shares
        totalCost = totalCost + lots(i).Shares * lots(i).CostPerShare
    Next i

    Dim fx As Double: fx = GetFXToTWD(ticker)                   ' 31.6 for US, 1 for TW
    Dim lastPx As Double: lastPx = GetStockPriceSafe(ticker)
    Dim wap As Double, unrl As Double, retPct As Double, netExp As Double

    If totalShares > 0 Then
        wap = totalCost / totalShares
        netExp = lastPx * totalShares * fx              ' aggregate -> TWD
        unrl = (lastPx - wap) * totalShares * fx        ' aggregate -> TWD
        If totalCost > 0 Then retPct = (lastPx - wap) * totalShares / totalCost
    End If

    ' NOTE: labels are drawn by DrawShellHeaders. We only fill values here.
    If totalShares > 0 Then
        Call WriteValueMoney(ws, 4, 2, netExp, "NT$")
        Call WriteValuePrice(ws, 5, 2, wap)           ' per-share native (no fx)
    Else
        With ws.Cells(4, 2)
            .Value = "NO ACTIVE POSITION"
            .Font.Color = RGB(150, 150, 150)
            .Font.Italic = True
            .Interior.Color = RGB(0, 0, 0)
            .Borders.LineStyle = xlNone
        End With
        Call WriteDash(ws, 5, 2)
    End If

    If totalShares > 0 And lastPx > 0 Then
        Call WriteValuePrice(ws, 6, 2, lastPx)        ' per-share native (no fx)
        Call WriteValuePnL(ws, 7, 2, unrl, "NT$")
        Call WriteValuePct(ws, 8, 2, retPct)
    Else
        Call WriteDash(ws, 6, 2)
        Call WriteDash(ws, 7, 2)
        Call WriteDash(ws, 8, 2)
    End If
End Sub

' ================================================================
' SECTION 2: LIFETIME METRICS (columns D-E)
' LIFETIME EFF = Profit Factor (gross win / gross loss)
' VELOCITY/DAY = Realized PnL / sum(holding days), min 1 per trade
' ================================================================
Private Sub DrawLifetimeMetrics(ws As Worksheet, ticker As String, _
                                 records() As TIRecord, recCount As Long)
    Dim totalPnL As Double, grossWin As Double, grossLoss As Double
    Dim totalDays As Long
    Dim i As Long
    For i = 1 To recCount
        totalPnL = totalPnL + records(i).PnL
        If records(i).PnL > 0 Then grossWin = grossWin + records(i).PnL
        If records(i).PnL < 0 Then grossLoss = grossLoss + Abs(records(i).PnL)
        Dim d As Long: d = records(i).DaysHeld
        If d < 1 Then d = 1
        totalDays = totalDays + d
    Next i

    Dim fx As Double: fx = GetFXToTWD(ticker)   ' 31.6 for US, 1 for TW

    ' --- Row 4: REALIZED PNL (TWD-converted) ---
    Call WriteValuePnL(ws, 4, 5, totalPnL * fx, "NT$")

    ' --- Row 5: LIFETIME EFF (Profit Factor) ---
    ' Defensive: clear borders/background, set NumberFormat BEFORE Value,
    ' and Round the value so even if format application fails the cell
    ' still displays 2 decimals max.
    With ws.Cells(5, 5)
        .Borders.LineStyle = xlNone
        .Interior.Color = RGB(0, 0, 0)
        .HorizontalAlignment = xlLeft
        If grossLoss > 0 Then
            Dim pf As Double: pf = grossWin / grossLoss
            .NumberFormat = "0.00"
            .Value = Round(pf, 2)
            If pf >= 1 Then
                .Font.Color = RGB(0, 210, 100)
            Else
                .Font.Color = RGB(255, 80, 80)
            End If
            .Font.Bold = True
        ElseIf grossWin > 0 Then
            .NumberFormat = "@"
            .Value = "infinite"     ' only wins, no losses
            .Font.Color = RGB(0, 210, 100)
            .Font.Bold = True
        Else
            .NumberFormat = "@"
            .Value = ChrW(&H2014)   ' em dash (consistent with WriteDash)
            .Font.Color = RGB(150, 150, 150)
        End If
    End With

    ' --- Row 6: VELOCITY/DAY (TWD-converted, PnL per holding day) ---
    Dim velocity As Double
    If totalDays > 0 Then velocity = totalPnL / totalDays
    Call WriteValuePnL(ws, 6, 5, velocity * fx, "NT$")
End Sub

' ================================================================
' SECTION 3: PROJECTION (columns G-H)
' H5 is the user-editable PRICE TARGET (yellow background)
' ================================================================
Private Sub DrawProjection(ws As Worksheet, ticker As String, _
                            lots() As TILot, lotCount As Long)
    Dim totalShares As Double
    Dim i As Long
    For i = 1 To lotCount: totalShares = totalShares + lots(i).Shares: Next i

    Dim lastPx As Double: lastPx = GetStockPriceSafe(ticker)
    Dim fx As Double: fx = GetFXToTWD(ticker)

    ' NOTE: labels drawn by DrawShellHeaders. Values only here.
    ' Row 4: LAST PRICE (per-share native)
    If lastPx > 0 Then
        Call WriteValuePrice(ws, 4, 8, lastPx)
    Else
        Call WriteDash(ws, 4, 8)
    End If

    ' Row 5: PRICE TARGET (per-share native input) - preserve existing value
    With ws.Cells(5, 8)
        If Not IsNumeric(.Value) Then .Value = vbNullString
    End With

    Dim target As Double
    If IsNumeric(ws.Cells(5, 8).Value) Then target = CDbl(ws.Cells(5, 8).Value)

    ' Row 6: PROJECTED PNL (TWD-converted)
    If totalShares > 0 And target > 0 And lastPx > 0 Then
        Dim projPnL As Double: projPnL = (target - lastPx) * totalShares * fx
        Call WriteValuePnL(ws, 6, 8, projPnL, "NT$")
    Else
        Call WriteDash(ws, 6, 8)
    End If

    ' Row 7: PROJECTED RETURN %
    If target > 0 And lastPx > 0 Then
        Dim projRet As Double: projRet = (target / lastPx) - 1
        Call WriteValuePct(ws, 7, 8, projRet)
    Else
        Call WriteDash(ws, 7, 8)
    End If
End Sub

' ================================================================
' SECTION 4: RECENT TRADE HISTORY (table, newest first)
' ================================================================
Private Sub DrawTradeHistory(ws As Worksheet, ticker As String, _
                              records() As TIRecord, recCount As Long)
    ' NOTE: Title (Row 10) and headers (Row 11) are drawn by DrawShellHeaders.
    ' We only fill data rows starting Row 12.
    If recCount = 0 Then Exit Sub

    ' Sort by ExitDate descending (newest first)
    Dim i As Long, j As Long, tmp As TIRecord
    For i = 1 To recCount - 1
        For j = i + 1 To recCount
            If records(j).ExitDate > records(i).ExitDate Then
                tmp = records(i): records(i) = records(j): records(j) = tmp
            End If
        Next j
    Next i

    Dim fx As Double: fx = GetFXToTWD(ticker)
    Dim pnlFmtStr As String: pnlFmtStr = PnLFmt("NT$")   ' quoted-literal NT$ format

    Dim r As Long: r = 12
    For i = 1 To recCount
        With ws.Cells(r, 1)
            .NumberFormat = "yyyy-mm-dd"
            .Value = records(i).EntryDate
            .Font.Color = RGB(221, 221, 221)
            .HorizontalAlignment = xlLeft
            .Borders.LineStyle = xlNone
        End With
        With ws.Cells(r, 2)
            .NumberFormat = "yyyy-mm-dd"
            .Value = records(i).ExitDate
            .Font.Color = RGB(221, 221, 221)
            .HorizontalAlignment = xlLeft
            .Borders.LineStyle = xlNone
        End With
        With ws.Cells(r, 3)
            .Value = records(i).DaysHeld
            .Font.Color = RGB(221, 221, 221)
            .HorizontalAlignment = xlCenter
            .Borders.LineStyle = xlNone
        End With
        With ws.Cells(r, 4)
            .NumberFormat = "#,##0.##"
            .Value = records(i).Shares
            .Font.Color = RGB(221, 221, 221)
            .HorizontalAlignment = xlRight
            .Borders.LineStyle = xlNone
        End With
        ' --- PnL converted to TWD ---
        Dim pnlTWD As Double: pnlTWD = records(i).PnL * fx
        With ws.Cells(r, 5)
            .NumberFormat = pnlFmtStr
            .Value = pnlTWD
            .Font.Color = USPnLColor(pnlTWD)
            .Font.Bold = True
            .HorizontalAlignment = xlRight
            .Borders.LineStyle = xlNone
        End With
        With ws.Cells(r, 6)
            .NumberFormat = "0.00%"
            .Value = Round(records(i).ReturnPct, 4)
            .Font.Color = USPnLColor(records(i).PnL)
            .HorizontalAlignment = xlRight
            .Borders.LineStyle = xlNone
        End With
        With ws.Cells(r, 7)
            .Value = records(i).Broker
            .Font.Color = RGB(150, 150, 150)
            .Font.Size = 9
            .HorizontalAlignment = xlLeft
            .Borders.LineStyle = xlNone
        End With
        ws.Rows(r).RowHeight = 16
        r = r + 1
    Next i
End Sub

' ================================================================
' CELL HELPERS
' ================================================================
Private Sub WriteLabel(ws As Worksheet, r As Long, c As Long, txt As String)
    With ws.Cells(r, c)
        .Value = txt
        .Font.Color = RGB(0, 200, 255)
        .Font.Bold = False
        .HorizontalAlignment = xlLeft
    End With
End Sub

' All value-writers below set NumberFormat BEFORE Value and explicitly
' clear borders/background to prevent stale formatting from prior renders.
'
' Currency symbol "NT$" must be quoted in the format string. Bare
' "$" in Excel format codes is treated as a locale-aware currency
' specifier (in zh-TW it expands to "NT$"), so an unquoted "NT$..."
' was being mis-parsed and silently rejected on some Excel versions.
' Using Chr(34) (double quote) around the symbol forces literal text.
Private Sub WriteValueMoney(ws As Worksheet, r As Long, c As Long, _
                             v As Double, sym As String)
    With ws.Cells(r, c)
        .Borders.LineStyle = xlNone
        .Interior.Color = RGB(0, 0, 0)
        .NumberFormat = MoneyFmt(sym)
        .Value = Round(v, 0)
        .Font.Color = RGB(221, 221, 221)
        .Font.Bold = True
        .HorizontalAlignment = xlLeft
    End With
End Sub

Private Sub WriteValuePrice(ws As Worksheet, r As Long, c As Long, v As Double)
    With ws.Cells(r, c)
        .Borders.LineStyle = xlNone
        .Interior.Color = RGB(0, 0, 0)
        .NumberFormat = "#,##0.00"
        .Value = Round(v, 2)
        .Font.Color = RGB(221, 221, 221)
        .Font.Bold = True
        .HorizontalAlignment = xlLeft
    End With
End Sub

Private Sub WriteValuePnL(ws As Worksheet, r As Long, c As Long, _
                           v As Double, sym As String)
    With ws.Cells(r, c)
        .Borders.LineStyle = xlNone
        .Interior.Color = RGB(0, 0, 0)
        .NumberFormat = PnLFmt(sym)
        .Value = Round(v, 0)
        .Font.Color = USPnLColor(v)
        .Font.Bold = True
        .HorizontalAlignment = xlLeft
    End With
End Sub

' Quoted-literal currency format builders.
' For sym="NT$", returns:  "NT$"#,##0
' That outer quote pair is part of the Excel format code itself.
Private Function MoneyFmt(sym As String) As String
    Dim q As String: q = Chr(34) & sym & Chr(34)
    MoneyFmt = q & "#,##0"
End Function

Private Function PnLFmt(sym As String) As String
    Dim q As String: q = Chr(34) & sym & Chr(34)
    PnLFmt = q & "#,##0;-" & q & "#,##0"
End Function

Private Sub WriteValuePct(ws As Worksheet, r As Long, c As Long, v As Double)
    With ws.Cells(r, c)
        .Borders.LineStyle = xlNone
        .Interior.Color = RGB(0, 0, 0)
        .NumberFormat = "0.00%"
        .Value = Round(v, 4)
        .Font.Color = USPnLColor(v)
        .Font.Bold = True
        .HorizontalAlignment = xlLeft
    End With
End Sub

Private Sub WriteDash(ws As Worksheet, r As Long, c As Long)
    With ws.Cells(r, c)
        .Borders.LineStyle = xlNone
        .Interior.Color = RGB(0, 0, 0)
        .NumberFormat = "@"
        .Value = ChrW(&H2014)   ' em dash
        .Font.Color = RGB(150, 150, 150)
        .Font.Bold = False
        .HorizontalAlignment = xlLeft
    End With
End Sub

' US/Bloomberg color convention: green positive, red negative
Private Function USPnLColor(v As Double) As Long
    If v > 0 Then
        USPnLColor = RGB(0, 210, 100)
    ElseIf v < 0 Then
        USPnLColor = RGB(255, 80, 80)
    Else
        USPnLColor = RGB(200, 200, 200)
    End If
End Function

' ================================================================
' TICKER / CURRENCY HELPERS
' ================================================================

' Strip .TW / .TWO suffix and uppercase so "2330", "2330.TW", "2330.TWO"
' all normalise to the same key. Used by FIFO matching and ticker classification.
Private Function NormalizeTicker(t As String) As String
    Dim n As String: n = UCase(Trim(CStr(t)))
    If Len(n) >= 4 Then
        If Right(n, 4) = ".TWO" Then n = Left(n, Len(n) - 4)
    End If
    If Len(n) >= 3 Then
        If Right(n, 3) = ".TW" Then n = Left(n, Len(n) - 3)
    End If
    NormalizeTicker = n
End Function

' Suffix-based classification:
'   .TW or .TWO suffix -> TW (no FX conversion)
'   Everything else    -> US (FX x 31.6 applied to aggregates)
'
' This is a STRICT rule. Special codes like "00981A" without a suffix
' are treated as US. To classify a TW security correctly, the ticker
' MUST be typed with its full suffix (e.g. "00981A.TWO", "2330.TW").
Private Function IsTWTicker(ticker As String) As Boolean
    Dim raw As String: raw = UCase(Trim(CStr(ticker)))
    IsTWTicker = (InStr(raw, ".TW") > 0)
End Function

' Per-share PRICE currency symbol (native; never FX-converted)
'   TW -> "NT$"     US -> "$"
Private Function GetPriceSymbol(ticker As String) As String
    If IsTWTicker(ticker) Then GetPriceSymbol = "NT$" Else GetPriceSymbol = "$"
End Function

' Multiplier for aggregate PnL / exposure -> TWD
'   TW ticker -> 1
'   US ticker -> FX_USD_TWD (31.6)
Private Function GetFXToTWD(ticker As String) As Double
    If IsTWTicker(ticker) Then GetFXToTWD = 1 Else GetFXToTWD = FX_USD_TWD
End Function

Private Function GetStockPriceSafe(ticker As String) As Double
    GetStockPriceSafe = 0
    On Error Resume Next
    Dim raw As String: raw = UCase(Trim(CStr(ticker)))
    If InStr(raw, ".TW") > 0 Then
        ' .TW / .TWO -> TWSE / TPEx OpenAPI (TaiwanPriceFetcher module)
        GetStockPriceSafe = GetTWStockPrice(ticker)
    Else
        ' US stocks / ETFs / indices -> Yahoo Finance (Attach module)
        GetStockPriceSafe = GetStockPrice(ticker)
    End If
    On Error GoTo 0
End Function

' Safer date / number conversion (returns 0 if not parseable)
Private Function SafeDate(v As Variant) As Date
    On Error Resume Next
    If IsDate(v) Then
        SafeDate = CDate(v)
    ElseIf IsNumeric(v) Then
        SafeDate = CDate(CDbl(v))
    End If
    On Error GoTo 0
End Function

Private Function SafeNum(v As Variant) As Double
    On Error Resume Next
    If IsNumeric(v) Then SafeNum = CDbl(v)
    On Error GoTo 0
End Function
```

## 11. TW_Coverage_Parser.bas

`VB_Name = "TW_Coverage_Parser"` · 1148 行 · 45.4 KB · 原始編碼 CP950 (Big5)

<details><summary><b>成員索引</b></summary>

| 行 | 型別 | 可見性 | 宣告 |
|---:|---|---|---|
| 27 | Function | Private | `Private Function IsArrayInitialized(ByRef arr() As String) As Boolean` |
| 35 | Sub | Public | `Sub FetchSingleReport()` |
| 92 | Sub | Public | `Sub SearchBySector()` |
| 253 | Function | Private | `Private Function ParseValuationField(vSec As String, colIdx As Long) As String` |
| 274 | Sub | Private | `Private Sub ApplySheetDarkBg(ws As Worksheet, lastRow As Long)` |
| 293 | Sub | Private | `Private Sub FormatReportSheet(content As String, mdPath As String)` |
| 543 | Sub | Private | `Private Sub DarkSectionHeader(ws As Worksheet, r As Long, title As String, bgColor As Long, accentColor As Long)` |
| 562 | Function | Private | `Private Function DarkRenderMdTable(ws As Worksheet, mdBlock As String, startRow As Long, headerBg As Long, accentColor As Long) As Long` |
| 625 | Function | Private | `Private Function ExtractSectionLoose(content As String, heading As String) As String` |
| 635 | Function | Private | `Private Function ExtractSubSectionLoose(content As String, heading As String) As String` |
| 647 | Sub | Public | `Sub ParsePilotReports()` |
| 698 | Sub | Private | `Private Sub SetupSheets()` |
| 726 | Sub | Private | `Private Sub EnsureSheet(ByVal shName As String)` |
| 742 | Function | Private | `Private Function CleanSheetName(ByVal txt As String) As String` |
| 753 | Sub | Private | `Private Sub ClearDataSheet(ByVal shName As String)` |
| 761 | Sub | Private | `Private Sub WriteHeaders()` |
| 768 | Sub | Private | `Private Sub WriteHeaderRow(shName As String, headers As Variant, bgColor As Long)` |
| 791 | Function | Private | `Private Function GetAllMDPaths(token As String) As String()` |
| 817 | Function | Private | `Private Function FetchRaw(mdPath As String, token As String) As String` |
| 847 | Function | Private | `Private Function HttpGet(url As String, token As String) As String` |
| 864 | Sub | Private | `Private Sub WriteAllSheets(content As String, mdPath As String, nextRows() As Long)` |
| 925 | Sub | Private | `Private Sub ParseFinancials(finSection As String, ann() As String, qtr() As String)` |
| 986 | Function | Private | `Private Function ExtractSection(content As String, heading As String) As String` |
| 997 | Function | Private | `Private Function ExtractSubSection(content As String, heading As String) As String` |
| 1007 | Function | Private | `Private Function StripMetaLines(text As String) As String` |
| 1015 | Function | Private | `Private Function CleanWikilinks(text As String) As String` |
| 1024 | Function | Private | `Private Function ExtractListItems(text As String) As String` |
| 1040 | Function | Private | `Private Function RegexFirst(text As String, pattern As String) As String` |
| 1049 | Function | Private | `Private Function SafeCell(cells() As String, idx As Long) As String` |
| 1053 | Function | Private | `Private Function SplitCells(line As String) As String()` |
| 1069 | Function | Private | `Private Function UrlEncodePath(path As String) As String` |
| 1092 | Function | Private | `Private Function EncodeUTF8Char(c As String) As String` |
| 1107 | Sub | Private | `Private Sub AutoFitSheets()` |
| 1123 | Function | Private | `Private Function GetToken() As String` |
| 1132 | Sub | Public | `Sub SetToken()` |

</details>

```vba
Option Explicit

' ═══════════════════════════════════════════════════════════════════
'  My-TW-Coverage GitHub Parser  (完整合併版 + 全 Bug 修正)
'  Repo : Timeverse/My-TW-Coverage
'
'  功能 A (單筆) : 在「設定」B3 輸入代號 → 執行 FetchSingleReport
'  功能 B (批次) : 執行 ParsePilotReports → 寫入四張工作表
' ═══════════════════════════════════════════════════════════════════

Private Const RAW_BASE      As String = "https://raw.githubusercontent.com/Timeverse/My-TW-Coverage/master/"
Private Const API_BASE      As String = "https://api.github.com/repos/Timeverse/My-TW-Coverage/"

' ── 工作表名稱 ────────────────────────────────────────────────────
Private Const SETTINGS_SHEET As String = "Settings"
Private Const SH_REPORT     As String = "Company Report"   ' 單筆報告頁（不含非法字元）
Private Const SH_OVERVIEW   As String = "公司總覽"
Private Const SH_BIZ        As String = "業務與供應鏈"
Private Const SH_ANNUAL     As String = "年度財務"
Private Const SH_QTR        As String = "季度財務"

' ═══════════════════════════════════════════════════════════════════
'  共用輔助：安全偵測 String() 陣列是否已初始化
'  原始寫法 (Not paths) = -1 對陣列做 bitwise NOT → Type Mismatch 崩潰
' ═══════════════════════════════════════════════════════════════════
Private Function IsArrayInitialized(ByRef arr() As String) As Boolean
    On Error Resume Next
    Dim lb As Long
    lb = LBound(arr)
    IsArrayInitialized = (Err.Number = 0)
    On Error GoTo 0
End Function

Sub FetchSingleReport()
    Dim token As String: token = GetToken()
    Dim Ticker As String
    Dim wsReport As Worksheet
    
    ' 優先從「公司報告」工作表的 B1 取得代號
    EnsureSheet SH_REPORT
    Set wsReport = ThisWorkbook.Sheets(SH_REPORT)
    Ticker = Trim(wsReport.Range("B1").Value)
    
    ' 如果報告頁 B1 是空的，則回頭去找「設定」工作表的 B3
    If Ticker = "" Then
        Ticker = Trim(ThisWorkbook.Sheets(SETTINGS_SHEET).Range("B3").Value)
    End If

    If Ticker = "" Then
        MsgBox "錯誤：請在「公司報告」B1 或「設定」B3 輸入股票代號。", vbCritical
        Exit Sub
    End If

    Application.StatusBar = "讀取 GitHub 檔案清單..."
    Dim paths() As String
    paths = GetAllMDPaths(token)

    If Not IsArrayInitialized(paths) Then
        MsgBox "錯誤：無法取得目錄，請檢查 Token。", vbCritical
        Application.StatusBar = False
        Exit Sub
    End If

    ' 搜尋路徑... (其餘邏輯不變)
    Dim targetPath As String: targetPath = ""
    Dim i As Long
    For i = 0 To UBound(paths)
        If InStr(paths(i), Ticker) > 0 Then
            targetPath = paths(i)
            Exit For
        End If
    Next i

    If targetPath = "" Then
        MsgBox "搜尋結束：找不到包含「" & Ticker & "」的檔案。", vbInformation
        Exit Sub
    End If

    Dim mdContent As String
    mdContent = FetchRaw(targetPath, token)
    If Len(mdContent) > 0 Then
        FormatReportSheet mdContent, targetPath
    End If
    Application.StatusBar = False
End Sub
' ═══════════════════════════════════════════════════════════════════
'  功能 C：產業搜尋
'  在「設定」B4 輸入產業關鍵字（如 Semiconductors）→ 執行 SearchBySector
'  結果寫入「產業搜尋」工作表，深色主題
' ═══════════════════════════════════════════════════════════════════
Sub SearchBySector()
    Const SH_SECTOR As String = "產業搜尋"
    Dim token As String: token = GetToken()
    Dim keyword As String
    keyword = Trim(ThisWorkbook.Sheets(SETTINGS_SHEET).Range("B4").Value)

    If keyword = "" Then
        MsgBox "請在「設定」B4 輸入產業關鍵字（如 Semiconductors）", vbExclamation
        Exit Sub
    End If

    Application.ScreenUpdating = False
    Application.StatusBar = "讀取清單..."

    Dim paths() As String: paths = GetAllMDPaths(token)
    If Not IsArrayInitialized(paths) Then
        MsgBox "無法取得清單，請確認 Token。", vbCritical
        GoTo CleanupSector
    End If

    ' 收集符合路徑（不分大小寫比對產業資料夾名稱）
    Dim hits() As String: ReDim hits(0 To UBound(paths))
    Dim hitCount As Long
    Dim i As Long
    For i = 0 To UBound(paths)
        If InStr(LCase(paths(i)), LCase(keyword)) > 0 Then
            hits(hitCount) = paths(i): hitCount = hitCount + 1
        End If
    Next i

    If hitCount = 0 Then
        MsgBox "找不到包含「" & keyword & "」的資料夾。", vbInformation
        GoTo CleanupSector
    End If

    ' 建立 / 清空工作表
    EnsureSheet SH_SECTOR
    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets(SH_SECTOR)
    ws.cells.Clear
    ApplySheetDarkBg ws, hitCount + 10

    ' ── 頁首統計列 ───────────────────────────────────────────────
    With ws.cells(1, 1)
        .Value = "SECTOR SCAN  |  " & UCase(keyword) & "  |  " & hitCount & " COMPANIES  |  " & Now()
        .Font.Bold = True: .Font.Size = 11
        .Font.Color = RGB(13, 13, 13)
        .Interior.Color = RGB(212, 175, 55)
    End With
    ws.Range(ws.cells(1, 1), ws.cells(1, 10)).Merge
    ws.Rows(1).RowHeight = 24

    ' ── 欄位標題 ──────────────────────────────────────────────────
    Dim hd As Variant
    hd = Array("TICKER", "COMPANY", "SECTOR", "INDUSTRY", "MKT CAP", "EV", "P/E TTM", "P/B", "EV/EBITDA", "PATH")
    Dim ci As Long
    For ci = 0 To UBound(hd)
        With ws.cells(2, ci + 1)
            .Value = hd(ci)
            .Font.Bold = True
            .Font.Color = RGB(212, 175, 55)
            .Interior.Color = RGB(30, 30, 30)
            .HorizontalAlignment = xlCenter
        End With
    Next ci
    ws.Rows(2).RowHeight = 20

    ' ── 逐一抓取並解析 ───────────────────────────────────────────
    Dim r As Long: r = 3
    Dim j As Long
    For j = 0 To hitCount - 1
        Application.StatusBar = "抓取 " & (j + 1) & " / " & hitCount & ":  " & hits(j)
        Dim mc As String: mc = FetchRaw(hits(j), token)
        If Len(mc) = 0 Then GoTo NextHit

        Dim fp() As String: fp = Split(hits(j), "/")
        Dim fnb As String: fnb = Replace(fp(UBound(fp)), ".md", "")
        Dim tk As String, co As String
        If InStr(fnb, "_") > 0 Then
            tk = Split(fnb, "_")(0): co = Split(fnb, "_")(1)
        Else
            tk = fnb: co = fnb
        End If

        ' 估值指標
        Dim fSec As String: fSec = ExtractSectionLoose(mc, "財務概況")
        Dim vSec As String: vSec = ExtractSubSectionLoose(fSec, "估值指標")
        Dim pe As String, pb As String, evebitda As String
        pe = ParseValuationField(vSec, 0)
        pb = ParseValuationField(vSec, 3)
        evebitda = ParseValuationField(vSec, 4)

        ' 寫入每一欄
        Dim rowVals(0 To 9) As Variant
        rowVals(0) = tk
        rowVals(1) = co
        rowVals(2) = RegexFirst(mc, "\*\*板塊[:：]\*\*\s*(.+)")
        rowVals(3) = RegexFirst(mc, "\*\*產業[:：]\*\*\s*(.+)")
        rowVals(4) = RegexFirst(mc, "\*\*市值[:：]\*\*\s*(.+)")
        rowVals(5) = RegexFirst(mc, "\*\*企業價值[:：]\*\*\s*(.+)")
        rowVals(6) = pe
        rowVals(7) = pb
        rowVals(8) = evebitda
        rowVals(9) = hits(j)

        Dim k As Long
        For k = 0 To 9
            With ws.cells(r, k + 1)
                .Font.Color = RGB(210, 210, 210)
                Select Case k
                    Case 0  ' Ticker
                        .Value = rowVals(k)
                        .Font.Bold = True
                        .Font.Color = RGB(212, 175, 55)
                    Case 6, 7, 8  ' 估值數字
                        If IsNumeric(rowVals(k)) Then
                            .Value = CDbl(rowVals(k))
                            .NumberFormat = "0.00"
                            .Font.Color = RGB(24, 144, 255)
                        Else
                            .Value = rowVals(k)
                            .Font.Color = RGB(130, 130, 130)
                        End If
                    Case Else
                        .Value = rowVals(k)
                End Select
                ' 偶數列微亮底色
                If (r Mod 2) = 0 Then
                    .Interior.Color = RGB(22, 22, 22)
                End If
            End With
        Next k
        r = r + 1
NextHit:
        If j Mod 10 = 0 Then DoEvents
    Next j

    ' 欄寬
    ws.Columns("A").ColumnWidth = 10
    ws.Columns("B").ColumnWidth = 20
    ws.Columns("C").ColumnWidth = 22
    ws.Columns("D").ColumnWidth = 22
    ws.Columns("E").ColumnWidth = 16
    ws.Columns("F").ColumnWidth = 16
    ws.Columns("G:I").ColumnWidth = 12
    ws.Columns("J").ColumnWidth = 50

    ws.Activate
    ws.cells(1, 1).Select
    ActiveWindow.FreezePanes = False
    ws.cells(3, 1).Select
    ActiveWindow.FreezePanes = True

CleanupSector:
    Application.ScreenUpdating = True
    Application.StatusBar = False
    If hitCount > 0 Then
        MsgBox "完成！共找到 " & hitCount & " 家 " & keyword & " 公司。", vbInformation
    End If
End Sub

' ── 從估值表格取出指定欄（資料列第一列 = 標題列下面那列）────────
Private Function ParseValuationField(vSec As String, colIdx As Long) As String
    If vSec = "" Then Exit Function
    Dim normalized As String
    normalized = Replace(Replace(vSec, vbCrLf, vbLf), vbCr, vbLf)
    Dim lines() As String: lines = Split(normalized, vbLf)
    Dim dataRowCount As Long
    Dim ln As Variant
    For Each ln In lines
        Dim ls As String: ls = Trim(CStr(ln))
        If Left(ls, 1) = "|" And InStr(ls, "---") = 0 And Len(ls) > 1 Then
            dataRowCount = dataRowCount + 1
            If dataRowCount = 2 Then  ' 第 2 個有效列 = 數值列
                Dim cA() As String: cA = SplitCells(ls)
                If colIdx <= UBound(cA) Then ParseValuationField = Trim(cA(colIdx))
                Exit Function
            End If
        End If
    Next ln
End Function

' ── 基礎背景渲染：全域純黑 (無死角版) ──
Private Sub ApplySheetDarkBg(ws As Worksheet, lastRow As Long)
    ' 1. 將整張工作表的所有儲存格均設為純黑底色與無邊框
    With ws.cells
        .Interior.Color = RGB(0, 0, 0)
        .Borders.LineStyle = xlNone
    End With
    
    ' 2. 僅針對會有資料出現的範圍設定預設字型，避免消耗多餘記憶體
    With ws.Range(ws.cells(1, 1), ws.cells(lastRow + 30, 22))
        .Font.Color = RGB(210, 210, 210)
        .Font.Name = "Consolas"
        .Font.Size = 10
    End With
    
    ws.Tab.Color = RGB(21, 96, 130)
    ws.Activate
    ActiveWindow.DisplayGridlines = False
End Sub
' ── 單筆報告排版主引擎 ──
Private Sub FormatReportSheet(content As String, mdPath As String)
    EnsureSheet SH_REPORT
    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets(SH_REPORT)
    
    ' 從路徑解析 Ticker 與公司名稱
    Dim parts() As String: parts = Split(mdPath, "/")
    Dim filename As String: filename = Replace(parts(UBound(parts)), ".md", "")
    Dim Ticker As String, company As String
    If InStr(filename, "_") > 0 Then
        Ticker = Split(filename, "_")(0)
        company = Split(filename, "_")(1)
    Else
        Ticker = filename: company = ""
    End If

    ' 1. 清空並初始化純黑背景
    ws.cells.Clear
    ApplySheetDarkBg ws, 120

    ' 2. 建立頂部輸入控制區
    With ws.cells(1, 1)
        .Value = "查詢代號:"
        .Font.Color = RGB(130, 130, 130)
        .Font.Bold = True: .Font.Size = 9
    End With
    
    With ws.cells(1, 2)
        .Value = Ticker
        .HorizontalAlignment = xlCenter
        .Font.Color = RGB(255, 255, 255)
        .Font.Bold = True
        .Interior.Color = RGB(0, 0, 0) ' 純黑背景
        .Borders.LineStyle = xlContinuous
        .Borders.Color = RGB(212, 175, 55)
    End With
    
    With ws.cells(1, 3)
        .Value = " ← 輸入後直接執行"
        .Font.Color = RGB(100, 100, 100): .Font.Size = 8
    End With
    ws.Rows(1).RowHeight = 25

    Dim nr As Long
    ' Row 2: 公司名稱大標題
    With ws.cells(2, 1)
        .Value = company & " " & Ticker
        .Font.Size = 16: .Font.Bold = True
        .Font.Color = RGB(212, 175, 55)
        .Font.Name = "Consolas"
        .Interior.Color = RGB(0, 0, 0)
    End With
    ws.Rows(2).RowHeight = 35
    
    nr = 3
    
    ' ── Row 3: 基本資訊 stat bar ──
    Dim statLabels As Variant, statVals As Variant
    statLabels = Array("板塊", "產業", "市值", "企業價值")
    statVals = Array( _
        RegexFirst(content, "\*\*板塊[:：]\*\*\s*(.+)"), _
        RegexFirst(content, "\*\*產業[:：]\*\*\s*(.+)"), _
        RegexFirst(content, "\*\*市值[:：]\*\*\s*(.+)"), _
        RegexFirst(content, "\*\*企業價值[:：]\*\*\s*(.+)") _
    )
    Dim si As Long
    For si = 0 To 3
        Dim col As Long: col = si * 3 + 1
        With ws.cells(nr, col)
            .Value = UCase(statLabels(si))
            .Font.Color = RGB(130, 130, 130)
            .Font.Size = 8: .Font.Bold = True
        End With
        With ws.cells(nr, col + 1)
            .Value = statVals(si)
            .Font.Color = RGB(212, 175, 55)
            .Font.Bold = True: .Font.Size = 10
        End With
    Next si
    ws.Rows(nr).RowHeight = 18
    nr = nr + 1

    Dim sepRange As Range: Set sepRange = ws.Range(ws.cells(nr, 1), ws.cells(nr, 18))
    sepRange.Interior.Color = RGB(212, 175, 55)
    ws.Rows(nr).RowHeight = 2
    nr = nr + 1

    ' ── 業務簡介 ──
    DarkSectionHeader ws, nr, "  BUSINESS OVERVIEW", RGB(0, 0, 0), RGB(212, 175, 55)
    nr = nr + 1
    Dim biz As String
    biz = CleanWikilinks(StripMetaLines(ExtractSectionLoose(content, "業務")))
    With ws.cells(nr, 1)
        .Value = biz
        .WrapText = True
        .Font.Color = RGB(190, 190, 190)
        .Font.Size = 10
    End With
    ws.Rows(nr).RowHeight = 80
    ws.Range(ws.cells(nr, 1), ws.cells(nr, 18)).Merge
    nr = nr + 1

    ' ── 供應鏈 (已優化 Regex 邏輯) ──
    DarkSectionHeader ws, nr, "  SUPPLY CHAIN", RGB(0, 0, 0), RGB(24, 144, 255)
    nr = nr + 1
    Dim scSection As String: scSection = ExtractSectionLoose(content, "供應鏈")
    Dim scData(0 To 2) As Variant
    ' 使用高度寬容的正則匹配：無視星號位置，抓取冒號後所有文字
    scData(0) = Array("UPSTREAM", CleanWikilinks(RegexFirst(scSection, "上游.*?[:：][\*\s]*([^\r\n]+)")))
    scData(1) = Array("MIDSTREAM", CleanWikilinks(RegexFirst(scSection, "中游.*?[:：][\*\s]*([^\r\n]+)")))
    scData(2) = Array("DOWNSTREAM", CleanWikilinks(RegexFirst(scSection, "下游.*?[:：][\*\s]*([^\r\n]+)")))
    Dim sci As Long
    For sci = 0 To 2
        With ws.cells(nr, 1)
            .Value = scData(sci)(0)
            .Font.Color = RGB(130, 130, 130): .Font.Bold = True: .Font.Size = 8
        End With
        With ws.cells(nr, 2)
            .Value = scData(sci)(1)
            .Font.Color = RGB(210, 210, 210)
            .WrapText = True
        End With
        ws.Range(ws.cells(nr, 2), ws.cells(nr, 12)).Merge
        ws.Rows(nr).RowHeight = 18
        nr = nr + 1
    Next sci

    ' ── 主要客戶 / 供應商 ──
    DarkSectionHeader ws, nr, "  CUSTOMERS  &  SUPPLIERS", RGB(0, 0, 0), RGB(24, 144, 255)
    nr = nr + 1
    Dim csSection As String: csSection = ExtractSectionLoose(content, "主要客戶")
    Dim csData(0 To 1) As Variant
    csData(0) = Array("CUSTOMERS", CleanWikilinks(ExtractListItems(ExtractSubSectionLoose(csSection, "主要客戶"))))
    csData(1) = Array("SUPPLIERS", CleanWikilinks(ExtractListItems(ExtractSubSectionLoose(csSection, "主要供應商"))))
    Dim csi As Long
    For csi = 0 To 1
        With ws.cells(nr, 1)
            .Value = csData(csi)(0)
            .Font.Color = RGB(130, 130, 130): .Font.Bold = True: .Font.Size = 8
        End With
        With ws.cells(nr, 2)
            .Value = csData(csi)(1)
            .Font.Color = RGB(210, 210, 210)
            .WrapText = True
        End With
        ws.Range(ws.cells(nr, 2), ws.cells(nr, 18)).Merge
        ws.Rows(nr).RowHeight = 18
        nr = nr + 1
    Next csi

    ' ── 財務概況 ──
    Dim finSection As String: finSection = ExtractSectionLoose(content, "財務")

    If Len(finSection) > 0 Then
        ' ── 估值指標 (矩陣轉置直式排列) ──
        Dim vSec As String: vSec = ExtractSubSectionLoose(finSection, "估值指標")
        If Len(vSec) > 0 Then
            
            ' 【新增邏輯】提取括號內的股價基準日期
            Dim priceDateStr As String
            ' 尋找 "as of 2026-03-26" 這種格式
            priceDateStr = RegexFirst(content, "估值指標\s*\([^)]*as of\s*([0-9\-]+)[^)]*\)")
            
            Dim valHeader As String
            If priceDateStr <> "" Then
                valHeader = "  VALUATION METRICS (As of " & priceDateStr & ")"
            Else
                valHeader = "  VALUATION METRICS"
            End If

            nr = nr + 1
            ' 將抓取到的日期融合進紅色區塊標題
            DarkSectionHeader ws, nr, valHeader, RGB(0, 0, 0), RGB(255, 100, 100)
            nr = nr + 1
            
            Dim vLines() As String: vLines = Split(Replace(vSec, vbCr, ""), vbLf)
            Dim vHead() As String, vData() As String
            Dim validRowCount As Long: validRowCount = 0
            Dim lineIt As Variant
            
            ' 拆解 Markdown 表格為兩組陣列
            For Each lineIt In vLines
                Dim lineStr As String: lineStr = Trim(CStr(lineIt))
                If Left(lineStr, 1) = "|" And InStr(lineStr, "---") = 0 Then
                    validRowCount = validRowCount + 1
                    If validRowCount = 1 Then
                        vHead = SplitCells(lineStr)
                    ElseIf validRowCount = 2 Then
                        vData = SplitCells(lineStr)
                        Exit For
                    End If
                End If
            Next lineIt
            
            ' 垂直印出陣列
            If validRowCount = 2 Then
                Dim vi As Long
                For vi = 0 To UBound(vHead)
                    If Trim(vHead(vi)) <> "" Then ' 過濾掉潛在的空欄位
                        With ws.cells(nr, 1) ' 欄位名稱放在A欄
                            .Value = Trim(vHead(vi))
                            .Font.Color = RGB(130, 130, 130)
                            .Font.Bold = True
                            .HorizontalAlignment = xlRight
                        End With
                        With ws.cells(nr, 2) ' 數值放在B欄
                            .Value = Trim(SafeCell(vData, vi))
                            .Font.Color = RGB(255, 100, 100)
                            .Font.Bold = True
                            .HorizontalAlignment = xlLeft
                        End With
                        ws.Range(ws.cells(nr, 2), ws.cells(nr, 4)).Merge
                        nr = nr + 1
                    End If
                Next vi
            End If
        End If

        ' 年度財務
        Dim aSec As String: aSec = ExtractSubSectionLoose(finSection, "年度")
        If Len(aSec) > 0 Then
            nr = nr + 1
            DarkSectionHeader ws, nr, "  ANNUAL FINANCIALS", RGB(0, 0, 0), RGB(82, 196, 26)
            nr = nr + 1
            nr = DarkRenderMdTable(ws, aSec, nr, RGB(0, 0, 0), RGB(82, 196, 26))
        End If

        ' 季度財務
        Dim qSec As String: qSec = ExtractSubSectionLoose(finSection, "季度")
        If Len(qSec) > 0 Then
            nr = nr + 1
            DarkSectionHeader ws, nr, "  QUARTERLY FINANCIALS", RGB(0, 0, 0), RGB(250, 173, 20)
            nr = nr + 1
            nr = DarkRenderMdTable(ws, qSec, nr, RGB(0, 0, 0), RGB(250, 173, 20))
        End If
    End If

    ' 欄寬重設
    ws.Columns("A").ColumnWidth = 26
    Dim c As Long
    For c = 2 To 8
        ws.Columns(c).ColumnWidth = 14
    Next c

    ws.Activate
    ActiveWindow.FreezePanes = False
    ws.cells(4, 1).Select
    ActiveWindow.FreezePanes = True
    ws.Range("B1").Select
End Sub
' ── 深色區塊標題 ─────────────────────────────────────────────────
Private Sub DarkSectionHeader(ws As Worksheet, r As Long, _
                               title As String, bgColor As Long, accentColor As Long)
    Dim rng As Range: Set rng = ws.Range(ws.cells(r, 1), ws.cells(r, 18))
    rng.Interior.Color = bgColor
    ws.cells(r, 1).Interior.Color = accentColor
    ws.cells(r, 1).Value = ""  ' accent strip
    With ws.cells(r, 2)
        .Value = title
        .Font.Bold = True
        .Font.Color = accentColor
        .Font.Size = 10
        .Font.Name = "Consolas"
        .Interior.Color = bgColor
    End With
    ws.Range(ws.cells(r, 2), ws.cells(r, 18)).Merge
    ws.Rows(r).RowHeight = 22
End Sub

' ── Markdown 表格渲染器 (純黑版) ──
Private Function DarkRenderMdTable(ws As Worksheet, mdBlock As String, _
                                    startRow As Long, _
                                    headerBg As Long, accentColor As Long) As Long
    Dim normalized As String
    normalized = Replace(Replace(mdBlock, vbCrLf, vbLf), vbCr, vbLf)
    Dim lines() As String: lines = Split(normalized, vbLf)

    Dim r As Long: r = startRow
    Dim isHeader As Boolean: isHeader = True
    Dim rowNum As Long: rowNum = 0
    Dim ln As Variant

    For Each ln In lines
        Dim ls As String: ls = Trim(CStr(ln))
        If Left(ls, 1) = "|" And InStr(ls, "---") = 0 And Len(ls) > 1 Then
            Dim cA() As String: cA = SplitCells(ls)
            Dim c As Long
            For c = 0 To UBound(cA)
                Dim cv As String: cv = CleanWikilinks(Trim(cA(c)))
                With ws.cells(r, c + 1)
                    ' 移除交錯底色，全部設定為純黑
                    .Interior.Color = IIf(isHeader, headerBg, RGB(0, 0, 0))
                    
                    If isHeader Then
                        .Value = cv
                        .Font.Bold = True
                        .Font.Color = accentColor
                        .HorizontalAlignment = xlCenter
                    ElseIf c = 0 Then
                        .Value = cv
                        .Font.Color = RGB(170, 170, 170)
                        .Font.Bold = True
                    ElseIf IsNumeric(cv) Then
                        .Value = CDbl(cv)
                        .NumberFormat = "0.00"
                        Dim labelCell As String
                        labelCell = LCase(Trim(cA(0)))
                        If InStr(labelCell, "margin") > 0 Or InStr(labelCell, "income") > 0 Or _
                           InStr(labelCell, "profit") > 0 Then
                            If CDbl(cv) >= 0 Then
                                .Font.Color = RGB(82, 196, 26)
                            Else
                                .Font.Color = RGB(220, 53, 69)
                            End If
                        Else
                            .Font.Color = RGB(210, 210, 210)
                        End If
                        .HorizontalAlignment = xlRight
                    Else
                        .Value = cv
                        .Font.Color = RGB(130, 130, 130)
                        .HorizontalAlignment = xlCenter
                    End If
                End With
            Next c
            If Not isHeader Then rowNum = rowNum + 1
            isHeader = False
            r = r + 1
        End If
    Next ln
    DarkRenderMdTable = r
End Function
' ── 寬鬆版 ExtractSection（heading 後可接任意字元）──────────────
Private Function ExtractSectionLoose(content As String, heading As String) As String
    Static re As Object
    If re Is Nothing Then Set re = CreateObject("VBScript.RegExp")
    re.Global = False: re.MultiLine = False
    re.pattern = "## " & heading & "[^\r\n]*\r?\n([\s\S]+?)(?=\r?\n## |$)"
    Dim m As Object: Set m = re.Execute(content)
    If m.count > 0 Then ExtractSectionLoose = m(0).SubMatches(0) Else ExtractSectionLoose = ""
End Function

' ── 寬鬆版 ExtractSubSection（### heading 後可接任意字元）────────
Private Function ExtractSubSectionLoose(content As String, heading As String) As String
    Static re As Object
    If re Is Nothing Then Set re = CreateObject("VBScript.RegExp")
    re.Global = False: re.MultiLine = False
    re.pattern = "### " & heading & "[^\r\n]*\r?\n([\s\S]+?)(?=\r?\n###|$)"
    Dim m As Object: Set m = re.Execute(content)
    If m.count > 0 Then ExtractSubSectionLoose = m(0).SubMatches(0) Else ExtractSubSectionLoose = ""
End Function

' ═══════════════════════════════════════════════════════════════════
'  功能 B：批次解析所有公司
' ═══════════════════════════════════════════════════════════════════
Sub ParsePilotReports()
    Dim token As String: token = GetToken()

    If token = "" Then
        MsgBox "請先在「" & SETTINGS_SHEET & "」工作表 B1 填入 GitHub Token。", vbExclamation
        Exit Sub
    End If

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    SetupSheets

    Application.StatusBar = "讀取檔案清單..."
    Dim paths() As String
    paths = GetAllMDPaths(token)

    If Not IsArrayInitialized(paths) Then
        MsgBox "無法取得檔案清單，請確認 Token 是否有效或網路狀態。", vbCritical
        GoTo Cleanup
    End If

    Dim total As Long: total = UBound(paths) + 1
    Dim i As Long, written As Long
    Dim nextRows(1 To 4) As Long
    Dim n As Long
    For n = 1 To 4: nextRows(n) = 2: Next n

    For i = 0 To UBound(paths)
        Application.StatusBar = "解析 " & (i + 1) & " / " & total & "  (" & paths(i) & ")"
        Dim mdContent As String
        mdContent = FetchRaw(paths(i), token)
        If Len(mdContent) > 0 Then
            WriteAllSheets mdContent, paths(i), nextRows
            written = written + 1
        End If
        If (i + 1) Mod 50 = 0 Then DoEvents
    Next i

    AutoFitSheets

Cleanup:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.StatusBar = False
    MsgBox "完成！共解析 " & written & " 家公司", vbInformation
End Sub

' ═══════════════════════════════════════════════════════════════════
'  工作表初始化（批次用）
' ═══════════════════════════════════════════════════════════════════
Private Sub SetupSheets()
    EnsureSheet SETTINGS_SHEET
    With ThisWorkbook.Sheets(SETTINGS_SHEET)
        If .Range("A1").Value = "" Then
            .Range("A1").Value = "GitHub Token"
            .Range("A1").Font.Bold = True
            .Range("B1").Value = "（請貼上 ghp_xxxx Token）"
            .Range("B1").Font.Color = RGB(150, 150, 150)
        End If
        If .Range("A3").Value = "" Then
            .Range("A3").Value = "單筆查詢代號"
            .Range("A3").Font.Bold = True
        End If
    End With

    EnsureSheet SH_OVERVIEW
    EnsureSheet SH_BIZ
    EnsureSheet SH_ANNUAL
    EnsureSheet SH_QTR

    ClearDataSheet SH_OVERVIEW
    ClearDataSheet SH_BIZ
    ClearDataSheet SH_ANNUAL
    ClearDataSheet SH_QTR

    WriteHeaders
End Sub

Private Sub EnsureSheet(ByVal shName As String)
    Dim safeName As String: safeName = CleanSheetName(shName)
    If Len(safeName) = 0 Then Exit Sub

    On Error Resume Next
    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets(safeName)
    On Error GoTo 0

    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.count))
        On Error Resume Next
        ws.Name = safeName
        On Error GoTo 0
    End If
End Sub

Private Function CleanSheetName(ByVal txt As String) As String
    Dim illegal As Variant: illegal = Array("\", "/", "?", "*", "[", "]", ":")
    Dim i As Integer
    txt = Replace(txt, vbCr, ""): txt = Replace(txt, vbLf, "")
    For i = LBound(illegal) To UBound(illegal)
        txt = Replace(txt, illegal(i), "_")
    Next i
    If Len(txt) > 31 Then txt = Left(txt, 31)
    CleanSheetName = Trim(txt)
End Function

Private Sub ClearDataSheet(ByVal shName As String)
    With ThisWorkbook.Sheets(shName)
        If .UsedRange.Rows.count > 1 Then
            .Rows("2:" & .UsedRange.Rows.count + 1).ClearContents
        End If
    End With
End Sub

Private Sub WriteHeaders()
    WriteHeaderRow SH_OVERVIEW, Array("Ticker", "公司名稱", "產業資料夾", "板塊", "產業", "市值（百萬台幣）", "企業價值（百萬台幣）"), RGB(31, 78, 121)
    WriteHeaderRow SH_BIZ, Array("Ticker", "公司名稱", "產業資料夾", "業務簡介", "上游", "中游", "下游", "主要客戶", "主要供應商"), RGB(55, 86, 35)
    WriteHeaderRow SH_ANNUAL, Array("Ticker", "公司名稱", "產業資料夾", "年度營收1", "年度營收2", "年度營收3", "毛利率1", "毛利率2", "毛利率3", "營業利益率1", "營業利益率2", "營業利益率3", "EPS1", "EPS2", "EPS3"), RGB(123, 44, 44)
    WriteHeaderRow SH_QTR, Array("Ticker", "公司名稱", "產業資料夾", "Q1營收", "Q2營收", "Q3營收", "Q4營收", "Q1 EPS", "Q2 EPS", "Q3 EPS", "Q4 EPS"), RGB(123, 82, 19)
End Sub

Private Sub WriteHeaderRow(shName As String, headers As Variant, bgColor As Long)
    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets(shName)
    Dim i As Long
    For i = 0 To UBound(headers)
        With ws.cells(1, i + 1)
            .Value = headers(i)
            .Font.Bold = True
            .Font.Color = RGB(255, 255, 255)
            .Font.Name = "Arial"
            .Interior.Color = bgColor
            .HorizontalAlignment = xlCenter
        End With
    Next i
    ws.Rows(1).RowHeight = 24
    ws.Activate
    ActiveWindow.FreezePanes = False
    ws.cells(2, 4).Select
    ActiveWindow.FreezePanes = True
End Sub

' ═══════════════════════════════════════════════════════════════════
'  API 請求
' ═══════════════════════════════════════════════════════════════════
Private Function GetAllMDPaths(token As String) As String()
    Dim url As String: url = API_BASE & "git/trees/master?recursive=1"
    Dim json As String: json = HttpGet(url, token)
    If json = "" Then Exit Function

    Static re As Object
    If re Is Nothing Then Set re = CreateObject("VBScript.RegExp")
    re.Global = True
    re.pattern = """path"":""(Pilot_Reports/[^/]+/[^""]+\.md)"""

    Dim matches As Object: Set matches = re.Execute(json)
    If matches.count = 0 Then Exit Function

    Dim result() As String
    ReDim result(0 To matches.count - 1)
    Dim i As Long
    For i = 0 To matches.count - 1
        result(i) = matches(i).SubMatches(0)
    Next i
    GetAllMDPaths = result
End Function

' ── FetchRaw：改用 Contents API + raw accept header ──────────────
'  原本走 raw.githubusercontent.com，對中文檔名 URL 編碼敏感，容易靜默失敗。
'  Contents API 讓 GitHub 伺服器端解析路徑，Accept: v3.raw 直接回傳純文字，
'  兩者合用可完全繞過中文編碼問題，且失敗時能看到真實 HTTP 狀態碼。
Private Function FetchRaw(mdPath As String, token As String) As String
    ' 使用 Contents API（不是 raw 域名）
    Dim url As String: url = API_BASE & "contents/" & UrlEncodePath(mdPath)

    On Error GoTo ErrHandler
    Dim http As Object: Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
    http.Open "GET", url, False
    http.setRequestHeader "User-Agent", "Excel-VBA-TW-Coverage"
    If token <> "" Then http.setRequestHeader "Authorization", "token " & token
    ' 關鍵：告訴 GitHub API 直接回傳原始文字，而非 JSON 包裝
    http.setRequestHeader "Accept", "application/vnd.github.v3.raw"
    http.send

    If http.status = 200 Then
        FetchRaw = http.responseText
    Else
        ' 顯示真實錯誤碼，方便診斷（404=找不到路徑, 401=Token 無效, 403=權限不足）
        MsgBox "下載失敗！" & vbCrLf & _
               "HTTP " & http.status & "：" & http.StatusText & vbCrLf & vbCrLf & _
               "路徑：" & mdPath & vbCrLf & _
               "請求網址：" & url, vbCritical, "FetchRaw 診斷"
        FetchRaw = ""
    End If
    Exit Function
ErrHandler:
    MsgBox "網路連線失敗：" & Err.Description, vbCritical, "FetchRaw 網路錯誤"
    FetchRaw = ""
End Function

' HttpGet 僅供 GetAllMDPaths（抓目錄清單）使用，維持 JSON accept header
Private Function HttpGet(url As String, token As String) As String
    On Error GoTo ErrHandler
    Dim http As Object: Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
    http.Open "GET", url, False
    http.setRequestHeader "User-Agent", "Excel-VBA-TW-Coverage"
    If token <> "" Then http.setRequestHeader "Authorization", "token " & token
    http.setRequestHeader "Accept", "application/vnd.github.v3+json"
    http.send
    If http.status = 200 Then HttpGet = http.responseText Else HttpGet = ""
    Exit Function
ErrHandler:
    HttpGet = ""
End Function

' ═══════════════════════════════════════════════════════════════════
'  批次資料解析與寫入
' ═══════════════════════════════════════════════════════════════════
Private Sub WriteAllSheets(content As String, mdPath As String, nextRows() As Long)
    Dim parts() As String: parts = Split(mdPath, "/")
    If UBound(parts) < 2 Then Exit Sub

    Dim sectorFolder As String: sectorFolder = parts(1)
    Dim filename As String:     filename = parts(2)
    Dim fnBase As String:       fnBase = Left(filename, Len(filename) - 3)

    Dim Ticker As String, company As String
    Dim underPos As Long: underPos = InStr(fnBase, "_")
    If underPos > 0 Then
        Ticker = Left(fnBase, underPos - 1)
        company = Mid(fnBase, underPos + 1)
    Else
        Ticker = fnBase: company = fnBase
    End If

    Dim 板塊 As String, 產業 As String, 市值 As String, EV As String, 業務 As String
    Dim 上游 As String, 中游 As String, 下游 As String, 客戶 As String, 供應商 As String
    Dim ann(1 To 12) As String
    Dim qtr(1 To 8) As String

    板塊 = RegexFirst(content, "\*\*板塊[:：]\*\*\s*(.+)")
    產業 = RegexFirst(content, "\*\*產業[:：]\*\*\s*(.+)")
    市值 = Replace(RegexFirst(content, "\*\*市值[:：]\*\*\s*([\d,，]+)"), ",", "")
    EV = Replace(RegexFirst(content, "\*\*企業價值[:：]\*\*\s*([\d,，]+)"), ",", "")

    業務 = CleanWikilinks(StripMetaLines(ExtractSection(content, "業務簡介")))
    If Len(業務) > 400 Then 業務 = Left(業務, 400) & "..."

    Dim scSection As String: scSection = ExtractSection(content, "供應鏈位置")
    上游 = CleanWikilinks(RegexFirst(scSection, "\*\*上游[:：]\*\*\s*(.+)"))
    中游 = CleanWikilinks(RegexFirst(scSection, "\*\*中游[:：]\*\*\s*(.+)"))
    下游 = CleanWikilinks(RegexFirst(scSection, "\*\*下游[:：]\*\*\s*(.+)"))

    Dim csSection As String: csSection = ExtractSection(content, "主要客戶及供應商")
    客戶 = CleanWikilinks(ExtractListItems(ExtractSubSection(csSection, "主要客戶")))
    供應商 = CleanWikilinks(ExtractListItems(ExtractSubSection(csSection, "主要供應商")))

    ParseFinancials ExtractSection(content, "財務概況"), ann, qtr

    Dim ws As Worksheet, c As Long

    Set ws = ThisWorkbook.Sheets(SH_OVERVIEW)
    ws.cells(nextRows(1), 1).Resize(1, 7).Value = Array(Ticker, company, sectorFolder, 板塊, 產業, 市值, EV)

    Set ws = ThisWorkbook.Sheets(SH_BIZ)
    ws.cells(nextRows(2), 1).Resize(1, 9).Value = Array(Ticker, company, sectorFolder, 業務, 上游, 中游, 下游, 客戶, 供應商)

    Set ws = ThisWorkbook.Sheets(SH_ANNUAL)
    ws.cells(nextRows(3), 1).Resize(1, 3).Value = Array(Ticker, company, sectorFolder)
    For c = 1 To 12: ws.cells(nextRows(3), 3 + c) = ann(c): Next c

    Set ws = ThisWorkbook.Sheets(SH_QTR)
    ws.cells(nextRows(4), 1).Resize(1, 3).Value = Array(Ticker, company, sectorFolder)
    For c = 1 To 8: ws.cells(nextRows(4), 3 + c) = qtr(c): Next c

    Dim n As Long
    For n = 1 To 4: nextRows(n) = nextRows(n) + 1: Next n
End Sub

Private Sub ParseFinancials(finSection As String, ann() As String, qtr() As String)
    If finSection = "" Then Exit Sub

    Static re As Object
    If re Is Nothing Then Set re = CreateObject("VBScript.RegExp")
    re.Global = True
    re.MultiLine = True
    re.pattern = "(\|.+\|(\r?\n|\r)(\| *[-:| ]+ *\|(\r?\n|\r))?(\|.+\|(\r?\n|\r))+)"

    Dim tables As Object: Set tables = re.Execute(finSection)
    Dim t As Object
    For Each t In tables
        Dim lines() As String: lines = Split(t.Value, vbLf)
        Dim rowLines As New Collection
        Dim ln As Variant
        For Each ln In lines
            Dim lineStr As String
            lineStr = Trim(Replace(CStr(ln), vbCr, ""))
            If Left(lineStr, 1) = "|" And InStr(lineStr, "---") = 0 Then
                rowLines.Add SplitCells(lineStr)
            End If
        Next ln

        Dim row As Variant
        For Each row In rowLines
            Dim cells() As String: cells = row
            If UBound(cells) < 1 Then GoTo NextRow
            Dim label As String: label = Trim(cells(0))
            Dim v1 As String: v1 = SafeCell(cells, 1)
            Dim v2 As String: v2 = SafeCell(cells, 2)
            Dim v3 As String: v3 = SafeCell(cells, 3)
            Dim v4 As String: v4 = SafeCell(cells, 4)

            If InStr(label, "營收") > 0 Or InStr(label, "Revenue") > 0 Then
                If v4 <> "" Then
                    qtr(1) = v1: qtr(2) = v2: qtr(3) = v3: qtr(4) = v4
                Else
                    ann(1) = v1: ann(2) = v2: ann(3) = v3
                End If
            ElseIf InStr(label, "毛利率") > 0 Or InStr(label, "Gross") > 0 Then
                ann(4) = v1: ann(5) = v2: ann(6) = v3
            ElseIf InStr(label, "營業利益率") > 0 Or InStr(label, "Operating") > 0 Then
                ann(7) = v1: ann(8) = v2: ann(9) = v3
            ElseIf InStr(label, "EPS") > 0 Then
                If v4 <> "" Then
                    qtr(5) = v1: qtr(6) = v2: qtr(7) = v3: qtr(8) = v4
                Else
                    ann(10) = v1: ann(11) = v2: ann(12) = v3
                End If
            End If
NextRow:
        Next row
    Next t
End Sub

' ═══════════════════════════════════════════════════════════════════
'  Markdown 解析工具
' ═══════════════════════════════════════════════════════════════════

' [BUG FIX] \Z 在 VBScript RegExp 不支援，最後一個段落永遠抓不到。
' MultiLine=False 時，$ 等同整個字串結尾，行為正確。
Private Function ExtractSection(content As String, heading As String) As String
    Static re As Object
    If re Is Nothing Then Set re = CreateObject("VBScript.RegExp")
    re.Global = False
    re.MultiLine = False
    re.pattern = "## " & heading & "\r?\n([\s\S]+?)(?=\r?\n## |$)"
    Dim m As Object: Set m = re.Execute(content)
    If m.count > 0 Then ExtractSection = m(0).SubMatches(0) Else ExtractSection = ""
End Function

' [BUG FIX] 同上
Private Function ExtractSubSection(content As String, heading As String) As String
    Static re As Object
    If re Is Nothing Then Set re = CreateObject("VBScript.RegExp")
    re.Global = False
    re.MultiLine = False
    re.pattern = "### " & heading & "\r?\n([\s\S]+?)(?=\r?\n###|$)"
    Dim m As Object: Set m = re.Execute(content)
    If m.count > 0 Then ExtractSubSection = m(0).SubMatches(0) Else ExtractSubSection = ""
End Function

Private Function StripMetaLines(text As String) As String
    Static re As Object
    If re Is Nothing Then Set re = CreateObject("VBScript.RegExp")
    re.Global = True: re.MultiLine = True
    re.pattern = "^\*\*.+?[:：]\*\*.+\r?\n?"
    StripMetaLines = Trim(re.Replace(text, ""))
End Function

Private Function CleanWikilinks(text As String) As String
    If text = "" Then Exit Function
    Static re As Object
    If re Is Nothing Then Set re = CreateObject("VBScript.RegExp")
    re.Global = True
    re.pattern = "\[\[(.+?)\]\]"
    CleanWikilinks = re.Replace(text, "$1")
End Function

Private Function ExtractListItems(text As String) As String
    If text = "" Then Exit Function
    Static re As Object
    If re Is Nothing Then Set re = CreateObject("VBScript.RegExp")
    re.Global = True: re.MultiLine = True
    re.pattern = "^[-*]\s*(.+)"
    Dim matches As Object: Set matches = re.Execute(text)
    If matches.count = 0 Then Exit Function
    Dim results() As String: ReDim results(0 To matches.count - 1)
    Dim i As Long
    For i = 0 To matches.count - 1
        results(i) = Trim(matches(i).SubMatches(0))
    Next i
    ExtractListItems = Join(results, "、")
End Function

Private Function RegexFirst(text As String, pattern As String) As String
    Static re As Object
    If re Is Nothing Then Set re = CreateObject("VBScript.RegExp")
    re.Global = False: re.MultiLine = True
    re.pattern = pattern
    Dim m As Object: Set m = re.Execute(text)
    If m.count > 0 Then RegexFirst = Trim(m(0).SubMatches(0)) Else RegexFirst = ""
End Function

Private Function SafeCell(cells() As String, idx As Long) As String
    If idx <= UBound(cells) Then SafeCell = Trim(Replace(cells(idx), ",", "")) Else SafeCell = ""
End Function

Private Function SplitCells(line As String) As String()
    Dim inner As String: inner = line
    If Left(inner, 1) = "|" Then inner = Mid(inner, 2)
    If Right(inner, 1) = "|" Then inner = Left(inner, Len(inner) - 1)
    Dim parts() As String: parts = Split(inner, "|")
    Dim i As Long
    For i = 0 To UBound(parts): parts(i) = Trim(parts(i)): Next i
    SplitCells = parts
End Function

' ═══════════════════════════════════════════════════════════════════
'  URL 編碼
' ═══════════════════════════════════════════════════════════════════

' [BUG FIX] Hex(code) 對 code<16 只回傳 1 位，輸出 %A 而非 %0A。
' Right("0" & Hex(code), 2) 強制補零至兩位。
Private Function UrlEncodePath(path As String) As String
    Dim result As String, c As String
    Dim i As Long, code As Long
    For i = 1 To Len(path)
        c = Mid(path, i, 1)
        code = AscW(c)
        ' [BUG FIX] AscW 回傳有號 16-bit，U+8000~U+FFFF 的字元（如「電」U+96FB=38651）
        ' 超過 32767 後 AscW 回傳負數（-26885），負數 < &H800 為 True，
        ' 誤用 2-byte 公式產生錯誤位元組序列 → GitHub 回 404。
        ' 加 65536 將負數還原為正確的 Unicode code point。
        If code < 0 Then code = code + 65536
        If (code >= 65 And code <= 90) Or (code >= 97 And code <= 122) Or _
           (code >= 48 And code <= 57) Or c = "/" Or c = "_" Or c = "-" Or c = "." Then
            result = result & c
        ElseIf code < 128 Then
            result = result & "%" & Right("0" & Hex(code), 2)
        Else
            result = result & EncodeUTF8Char(c)
        End If
    Next i
    UrlEncodePath = result
End Function

Private Function EncodeUTF8Char(c As String) As String
    Dim code As Long: code = AscW(c)
    If code < 0 Then code = code + 65536  '[BUG FIX] AscW 對 U+8000+ 回傳負數，加 65536 還原正確 code point
    Dim result As String
    If code < &H800 Then
        result = "%" & Right("0" & Hex(&HC0 Or (code \ 64)), 2) & _
                 "%" & Right("0" & Hex(&H80 Or (code Mod 64)), 2)
    ElseIf code < &H10000 Then
        result = "%" & Right("0" & Hex(&HE0 Or (code \ 4096)), 2) & _
                 "%" & Right("0" & Hex(&H80 Or ((code \ 64) Mod 64)), 2) & _
                 "%" & Right("0" & Hex(&H80 Or (code Mod 64)), 2)
    End If
    EncodeUTF8Char = UCase(result)
End Function

Private Sub AutoFitSheets()
    Dim shNames As Variant: shNames = Array(SH_OVERVIEW, SH_BIZ, SH_ANNUAL, SH_QTR)
    Dim s As Variant
    For Each s In shNames
        Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets(CStr(s))
        ws.Columns.AutoFit
        If CStr(s) = SH_BIZ Then
            If ws.Columns(4).ColumnWidth > 60 Then ws.Columns(4).ColumnWidth = 60
            ws.Columns(4).WrapText = True
        End If
    Next s
End Sub

' ═══════════════════════════════════════════════════════════════════
'  Token 管理
' ═══════════════════════════════════════════════════════════════════
Private Function GetToken() As String
    On Error Resume Next
    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets(SETTINGS_SHEET)
    If ws Is Nothing Then Exit Function
    Dim val As String: val = Trim(ws.Range("B1").Value)
    If val <> "" And InStr(val, "請貼上") = 0 Then GetToken = val
    On Error GoTo 0
End Function

Sub SetToken()
    Dim token As String
    token = InputBox("請貼上 GitHub Personal Access Token：", "設定 Token")
    If token = "" Then Exit Sub
    EnsureSheet SETTINGS_SHEET
    With ThisWorkbook.Sheets(SETTINGS_SHEET)
        .Range("A1").Value = "GitHub Token"
        .Range("A1").Font.Bold = True
        .Range("B1").Value = token
        .Range("B1").Font.Color = RGB(100, 100, 100)
        .Range("B1").NumberFormat = "@"
    End With
    MsgBox "Token 已儲存到「" & SETTINGS_SHEET & "」B1。", vbInformation
End Sub
```

