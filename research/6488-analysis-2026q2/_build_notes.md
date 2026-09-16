# 環球晶圓 (GlobalWafers, TPEx: 6488) — FY2026 Q2 財報完整分析 · 建置筆記

## 決策（2026-09-16 使用者確認，10 題澄清）
| 項目 | 選擇 |
|---|---|
| 產出 | 完整 10-tab Earnings Call 報告（單檔） |
| 期別 | FY2026 Q2（董事會 2026-08-04 通過） |
| 模式 | Fallback（無逐字稿）：Q&A tab 改「法說會重點 × 財報口徑檢核」 |
| 資料夾／檔名 | `research/6488-analysis-2026q2/6488_FY2026Q2_Analysis/6488_FY2026Q2_Analysis.html` |
| 同業 | 全球五強（信越 4063.T、SUMCO 3436.T、Siltronic WAF.DE、SK siltron 非上市）＋合晶 6182、台勝科 3532 |
| 事件槽 | 德州 Sherman 廠轉固 × CHIPS/AMIC × Micron 5 億美元預付 |
| 圖表輸出 | 靜態 SVG（chart-tools）；Sankey 用 sankey-diagram-demo |
| Works 分組 | 既有「矽晶圓 · Silicon Wafer」（SUMCO 之後） |
| 本地 port | 8190（`6488-analysis`，directory 指到研究資料夾根，報告在 `/6488_FY2026Q2_Analysis/`） |

## 資料管線（data/）
- `raw/fin_YYYYQn.html`：MOPS t164sb01 iXBRL 12 季（cp950）→ `parse_fin.py` → `financials_raw.json` → `build_table.py` → `quarterly.json`（Q4＝FY−9M；CF＝YTD 相減）
- `raw/6488_2026Q2_consolidated.pdf`：doc.twse 全文（step=9 POST → /pdf/ 連結）→ `q2_pdf.txt`（附註：ECB、公司債、合約負債、承諾、附表七子公司損益、部門）
- `raw/6488_AR2025.pdf`：114 年度年報 → `ar2025.txt`（供應商/客戶 10%、銷售地區、市占）
- `raw/mrev_<民國年>_<月>.html`：上櫃月營收 26 個月 → `monthly_revenue.json`（全部累計欄反算 OK）
- `raw/ins_<月>.html`：MOPS ajax_stapap1 內部人持股月報 115/5–8
- `peers_yahoo.json`：Yahoo quoteSummary（需 fc.yahoo.com cookie + getcrumb）
- `make_specs.py` → 13 張圖 spec → Browser pane 驅動 chart-tools/sankey（scratchpad recv.py 同時 GET 供 spec、POST 收 SVG）→ `charts/*.html`
- `report_template.html` ＋ `{{CHART:name}}` → 最終 HTML

## 坑
- MOPS 重大訊息／內部人頁面是 UTF-8，不是 cp950（財報 iXBRL 才是 cp950）。
- Browser pane 隱藏時 clientWidth=0：Sankey 不畫、overflow 檢查全 true、截圖空白；先 resize_window 1280 再驗。
- ECB 嵌入選擇權 Q2 再衡量 −18.08 億與 Siltronic 持股 +48.4 億同向相抵，淨 +30.35 億——不能只看「FVTPL 利益」一個數字。
