# 台勝科（Formosa SUMCO Technology, TWSE 3532）FY2026Q2 資料層建置筆記

本次僅建置**資料層**（`data/` 目錄），未產出完整 10-tab HTML 報告，依上層指示不 git commit。

## 資料管線（沿用 6488 pipeline，`research/6488-analysis-2026q2/data/parse_fin.py`、`build_table.py` 原樣複製過來，未改內容——script 本身不寫死代號）

1. `raw/fin_<Y>Q<S>.html`：MOPS t164sb01 iXBRL，`CO_ID=3532&REPORT_ID=C`（合併財報，12季全部命中，未需退回 REPORT_ID=A）→ cp950 解碼 → `parse_fin.py` → `financials_raw.json` → `build_table.py` → `quarterly.json`（12季：2023Q3~2026Q2，Q4=FY YTD−Q3 YTD，CF=YTD相減）
2. `raw/mrev_<ROC年>_<月>.html`：MOPS 上市（**/sii/**，非/otc/）月營收，26個月（113/7~115/8，與6488同範圍）→ `parse_monthly.py` → `monthly_revenue.json`
3. `raw/3532_2026Q2_consolidated.pdf`：doc.twse t57sb01，115年第二季 IFRSs合併財報（202602_3532_AI1.pdf）→ pypdf → `q2_pdf.txt`（51頁）
4. `raw/3532_AR2025.pdf`：114年度股東年報（2025_3532_20260611F04.pdf）→ pypdf → `ar2025.txt`（139頁）
5. `peers_yahoo.json`：Yahoo quoteSummary，6檔同業（3532.TW/6488.TWO/6182.TWO/3436.T/4063.T/WAF.DE）全部命中
6. `price_weekly.json`：3532.TW 近2年週K（106週）
7. `raw/ins_5~8.html`：MOPS ajax_stapap1 董監持股（4次查詢內容相同，見下方坑）
8. `qualitative_notes.md`：關係人交易、主要客戶供應商、資本支出、借款、合約負債、股利、互證對照表

## 跟 6488 的差異點（依題目指示調整）

- 3532 是**上市（sii）**非上櫃，月營收URL改用 `/nas/t21/sii/`（6488是`/otc/`）——已確認26個月全部正確抓到且累計欄反算全部（除NT$1千的四捨五入尾差外）吻合。
- t164sb01 用 `REPORT_ID=C` 直接命中12季合併財報，未出現「無子公司」的情況，不需退回 `REPORT_ID=A`（個別財報）。

## 坑與驗證

- **月營收表格解析大小寫陷阱（新發現，6488當時未踩到）**：MOPS 月營收HTML的欄位標籤在部分列混用 `<td>` 與 `<Td>`（大寫T），原本沿用6488經驗寫的正規表達式 `<t[dD][^>]*>` 只涵蓋小寫開頭的 `td`/`tD`，不含 `Td`/`TD`，導致欄位整排錯位、YoY%欄讀出離譜數字（如10278649%）。改用大小寫不敏感的 `<td[^>]*>...</td>`（`re.I`）後全部正確，累計欄反算全部通過（僅NT$1千尾差）。
- MOPS偶發「Unreachable Server」空回應（114/12、115/1兩個月各遇到一次），單純重試即成功，非資料不存在。
- doc.twse 的 curl 呼叫兩次出現 exit code 1 但無明確錯誤訊息（SSL/TLS renegotiation相關，見-v輸出），單純重試即成功，非永久性問題。
- Yahoo crumb 首次呼叫遇到「Edge: Too Many Requests」，等待15秒後重試即成功。
- t164sb01 REPORT_ID=C 直接命中，未需要C→A退回邏輯。

## 資料缺口（詳見 qualitative_notes.md 第12節）

- 無法說會逐字稿，本案屬 Fallback 模式（若日後產出完整報告，Q&A/風險矩陣需依此調整）。
- 產能稼動率僅年報文字敘述，無量化百分比。
- 內部人持股變動：ajax_stapap1端點對3532固定回傳「目前持股」靜態快照（4次查詢/月份參數內容完全相同），並非逐月異動表，故無法呈現月度增減——這點與6488當時的做法可能不同，需要留意此端點的行為並非總是回傳月度歷史。
- 材料重大訊息（ajax_t05st01）查詢回傳空白模板，表單參數組合未能命中資料，本次未深入除錯，留作缺口。
- RD（研發費用）在 XBRL parse 結果中全部為空（12季皆無值），可能未獨立標記或併入其他費用科目，需要之後如需精細拆分時再查其他XBRL concept name。
- 未開始租賃（lease not yet commenced）：全文搜尋無此揭露，判斷不適用。
- 資料中心專案融資條款：不適用（非資料中心資本支出型公司）。

## 關鍵數字速覽（未經圖表化，僅供快速核對，詳細見 quarterly.json）

- 近4季營收（仟元）：2025Q3 3,088,467 / 2025Q4 3,268,741 / 2026Q1 3,308,768 / 2026Q2 3,610,137
- 近4季毛利率：15.4% / 11.6% / 0.8% / 6.9%（2026Q1毛利率驟降至近乎打平，2026Q2回升但仍處低檔）
- 近4季營益率：7.8% / 9.1% / -3.2% / 2.6%（2026Q1出現單季營業虧損）
- 近4季母公司歸屬淨利（仟元）：218,809 / 303,932 / -68,215 / 47,941
- 近4季EPS：0.56 / 0.78 / -0.18 / 0.12
- 2026年8月單月營收：1,331,863仟元，YoY +29.27%（累計YoY +19.1%左右，見monthly_revenue.json 2026-08 cum_yoy_pct）
- 2025年度現金股利：每股1.0元（113年度為1.8元，年減44%）
