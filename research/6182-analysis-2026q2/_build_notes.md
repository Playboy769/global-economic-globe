# 合晶科技 (Wafer Works, TPEx: 6182) — FY2026 Q2 財報完整分析 · 建置筆記

## 決策（2026-09-16 使用者確認，10 題澄清）
| 項目 | 選擇 |
|---|---|
| 產出 | 完整 10-tab Earnings Call 報告（單檔） |
| 期別 | FY2026 Q2 |
| 模式 | Fallback（無逐字稿）：Q&A tab 改「法說會重點 × 財報口徑檢核」（法說簡報有公開就抓） |
| 資料夾／檔名 | `research/6182-analysis-2026q2/6182_FY2026Q2_Analysis/6182_FY2026Q2_Analysis.html` |
| 版型／管線 | 沿用 `research/6488-analysis-2026q2/`（parse_fin / build_table / make_specs / report_template） |
| 同業 | 台股矽晶圓（環球晶 6488、台勝科 3532、嘉晶 3016）＋日系（信越 4063.T、SUMCO 3436.T） |
| 趨勢長度 | 12 季（FY2023Q3–FY2026Q2）＋ 月營收 24 個月以上 |
| 事件槽 | 由財報附註挑本季最重大事件 |
| 收尾 | 鏡像 globe-invest、上架 Works 矽晶圓分組、EARN_DATES、兩個 repo commit + push |
| 分工 | sonnet 子代理依序：資料 → 圖表 → 報告；主 session 規劃與審查 |

## 資料管線（data/）

完成於 2026-09-16。沿用 `research/6488-analysis-2026q2/data/` 的 `parse_fin.py`／
`build_table.py`（複製後只改 CO_ID=6182，IS/BS/CF 科目標籤與 6488 高度重疊，20個IS、
36–38個BS、21–24個CF標籤全數命中，未需增補新標籤）。

### 產出檔案清單
- `raw/fin_2023Q3.html` … `raw/fin_2026Q2.html`（12季）：MOPS t164sb01 iXBRL，cp950
  → `parse_fin.py` → `financials_raw.json` → `build_table.py` → `quarterly.json`
  （Q4＝FY YTD−9M；CF一律YTD相減）
- `derive_metrics.py` → `derived_metrics.json`：12季 DSO（91天口徑）、存貨天數、
  GM/OPM/NPM、利息保障倍數（OpInc/|IntExp|）、FCF（CFO−|Capex|）
- `raw/6182_2026Q2_consolidated.pdf`（1.6MB，91頁）→ `q2_pdf.txt`（doc.twse t57sb01，
  step=1列檔名202602_6182_AI1.pdf → POST step=9拿中介頁 → GET /pdf/連結）：附註含
  轉投資/大陸投資、合約負債、質押資產、背書保證、重大或有負債及未認列合約承諾
  （長期原料採購合約、在建工程設備合約）、關係人交易、租賃、股本/可轉債/現金增資
- `raw/6182_AR2025.pdf`（3.7MB，122頁）→ `ar2025.txt`（doc.twse同流程，mtype=F，
  114年報，filename前綴2025_）：主要客戶/供應商10%揭露、產品用途與產製流程、
  市場佔有率（SEMI統計）、銷售地區別、同業競爭
- `monthly_raw.json` → `monthly_revenue.json`：上櫃月營收32個月（2024-01–2026-08），
  全部通過「本月累計−上月累計＝本月單月」交叉驗證（cross_check_ok全數True）
- `insiders_raw.json` → `insiders_summary.txt`：MOPS ajax_stapap1內部人持股月報
  115/4–8月，19–20名內部人
- `peers_yahoo.json` → `peers_summary.txt`：Yahoo quoteSummary，6182.TWO/6488.TWO/
  3532.TW/3016.TW/4063.T/3436.T市值/P-E/P-B/EV/營收成長/毛利率
- `price_weekly.json`：6182.TWO近2年週線（106點）
- `qualitative_notes.md`：質化素材彙整（業務產品、同業、供應鏈節點建議、特殊事件
  候選3項、管理層展望背景、月營收動能對照、風險清單、6項必備追蹤、資料缺口）

### 坑
1. **MOPS ajax_stapap1（內部人持股）編碼是 UTF-8，不是 cp950**——用cp950解碼會把
   合法UTF-8多位元組序列拆成亂碼且不報錯（曾誤判為壞資料）；本次先照舊有筆記
   `mops-xbrl-encoding-and-annual-report-fetch` 的說法試cp950失敗後才發現這支
   ajax端點（`ajax_stapap1`／`ajax_t05st01`）跟t164sb01財報iXBRL不同，是UTF-8。
   財報iXBRL（t164sb01）本身仍是cp950，兩者不可套同一套解碼邏輯。
2. **ajax_stapap1直接POST基本參數（co_id/year/month）會回空殼頁**，必須加上
   `encodeURIComponent=1&step=1&firstin=1&off=1` 四個參數才會回真正資料——這組
   參數沿用了`mops-xbrl-encoding-and-annual-report-fetch`筆記提到的台股股利查詢
   坑（`firstin`等），原來内部人查詢也共用同一套防呆機制。
3. **MOPS重大訊息（ajax_t05st01）正確參數未試出**：嘗試SDATE/EDATE、
   startdate/enddate、dy1/dm1/dd1等多組日期欄位命名，加上前述四個防呆參數，
   一律回「年月日未輸入,請檢查」。時間有限下已放棄，列為資料缺口，建議下次
   改用瀏覽器實際發request觀察正確欄位名。
4. **上櫃月營收檔名月份不補零**（`t21sc03_115_8.html`而非`_08.html`），2026年8月
   若補零成`_08`會404；2026-09（115年9月）截至抓取當下尚未公告（回916 bytes提示頁）。
   2024-03（113年3月）首次嘗試回18 bytes空檔，重抓一次即正常，與既有筆記記載一致。
5. **peers_yahoo.json寫檔忘記指定`encoding='utf-8'`導致UnicodeDecodeError**（Windows
   預設cp950無法承載日股信越/SUMCO的公司名稱等非ASCII內容），改為
   `open(...,'w',encoding='utf-8')`後才正常讀回。
6. **bash工具的console本身是cp950**，Python腳本若直接print中文到stdout會噴
   UnicodeEncodeError或顯示亂碼；改用「寫檔→Read工具讀」的方式繞過，本次所有
   中文摘要（insiders_summary.txt、qualitative_notes.md、excerpt*.txt）都採此模式。
7. **q2_pdf.txt「轉投資事業相關資訊」大陸投資表格中河南芯晶半導體的持股比例欄位
   在PDF文字抽取時被截斷到下一頁**，未能完整擷取數值，已在qualitative_notes.md
   第9節列為缺口。
