# 3532 台勝科 FY2026Q2 圖表清單

各圖由 `projects/chart-tools/` 對應產生器（靜態 SVG）與 `projects/sankey-diagram-demo`
（自包含 div+script）產出，規格來源為 `../chart_specs/*.json`（由 `make_specs.py` 產生）。

- **supply_chain_chain.html**（chain-diagram）⚑ 上游原料→台勝科一貫廠→下游客戶→終端應用
  四層結構；關係人（SUMCO 單晶晶棒／回銷、台塑）以紫色虛線區隔於一般供應/銷售實線。
  來源：2026Q2 合併財報附註二六關係人交易；114 年度股東年報「主要客戶及供應商」章節。

- **supply_chain_sankey.html**（sankey-diagram-demo）⚑ 僅畫已揭露之量化流量：上游 SUMCO
  單晶晶棒 17.24%＋多晶矽供應商（Tokuyama）10.64% 進貨占比；下游客戶A 37.58%、SUMCO 回銷
  10.04%、其他晶圓廠 52.38% 銷貨占比。終端應用層因無量化拆分，依規則不進 Sankey。節點命名
  與 chain-diagram 一致。來源：2026Q2 合併財報附註二六關係人交易。

- **is_trend_line.html**（line-chart）⚑ 12 季（2023Q3–2026Q2）營收（左軸，億元）與毛利率／
  營業利益率（右軸，%）。2026Q2 毛利率 6.90% 為 12 季最低。來源：quarterly.json（MOPS
  t164sb01 合併財報）。

- **monthly_rev_line.html**（line-chart）⚑ 26 個月（2024/07–2026/08）月營收（左軸，億元）與
  YoY%（右軸）。2026/07、08 為 Q2 財報期間之後最新月份，YoY 連續轉正擴大（21.95%→29.27%）。
  來源：MOPS 上市月營收彙總表 t21sc03_sii，逐月以累計欄反算交叉驗證。

- **bu_bar.html**（bar-chart）⚑ 114 年度 vs 113 年度內銷／外銷金額（億元）。財報未拆分
  8 吋／12 吋季度營收，僅有年度內外銷比例可量化。來源：114 年度股東年報「主要產品別內外銷
  比例及地區」。

- **fcf_waterfall.html**（waterfall-chart）⚑ 2026 H1 YTD：稅前淨利 −0.22 億 → 折舊攤銷
  +17.07 億與營運資金調整 → 營運現金流 9.68 億 → 資本支出 −5.80 億 → 自由現金流 3.88 億。
  來源：2026Q2 合併現金流量表（H1 累計）。

- **peer_scatter.html**（scatter-chart）⚑ 採 **Forward P/E**（Siltronic 因預估獲利為負改用
  EV/EBITDA）vs 最近一季營收 YoY%；氣泡＝市值（十億美元，2026-09-16 匯率換算）。台勝科
  trailing P/E 338.5x 因獲利趨近於零而失真，故不採用。來源：peers_yahoo.json（Yahoo Finance
  quoteSummary）。

- **interest_coverage_line.html**（line-chart）⚑ 12 季利息保障倍數（營業利益÷財務成本，
  >60 倍截尾）。2026Q1 營業利益為負顯示為 0，2026Q2 降至 1.62x，為 12 季最低。來源：
  quarterly.json（各季合併財報附註財務成本、損益表營業利益）。

- **dso_line.html**（line-chart）⚑ 12 季應收帳款天數（DSO＝AR÷營收×91）與存貨天數。
  2026Q2 DSO 升至 70.9 天，12 季新高，與客戶A集中度上升方向一致。來源：quarterly.json
  （各季合併資產負債表與損益表）。

- **contract_liab_line.html**（area-chart，堆疊）⚑ 12 季合約負債（流動＋非流動）餘額，台股
  IFRS 無「預收貨款」科目，此為對應項目。由 2023Q3 的 50.9 億持續消化至 2026Q2 的 26.0 億。
  來源：各季合併資產負債表附註十九合約負債。

- **risk_heatmap.html**（heatmap-chart）⚑ 風險項（毛利率持續低迷／客戶A集中／利息保障惡化／
  中國8吋產能／母公司權利金-定價／12吋需求遞延）× 時間窗口（近1季／1年／3年），強度
  1–5，數值為本報告主觀評分（Fallback 模式，無逐字稿）。

## 略過的圖表

- **capex_gantt**：財報僅揭露「已簽約未交驗之建造工程與設備款」逐季**期末餘額**
  （45.11／47.15／47.83 億元），未依到期年度分層揭露起訖日期，不具備甘特圖所需的時程區間
  資料，故略過；該三個數字已收錄於 `fragments/quant.html` 資產負債表 tab 的必備追蹤表格中。
