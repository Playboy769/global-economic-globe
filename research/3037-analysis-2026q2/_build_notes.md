# 欣興電子 (Unimicron, TWSE: 3037) — FY2026 Q2 財報完整分析 · 建置筆記

## 決策 (2026-08-30 使用者確認)

| 項目 | 選擇 |
|---|---|
| 期別 | FY2026 Q2（民國115年第二季，MOPS 公告截止 2026-08-14） |
| 模式 | **Fallback**（無逐字稿）— 略過 Q&A 話術偵測 tab，改「法說會簡報重點」；風險矩陣註明資料限制；加重 24 個月月營收反推 |
| 範圍 | 完整 10-tab 合併單檔 HTML |
| 收尾 | 鏡像 globe-invest/app/research/ ＋ 上架 OutsideFramework Works ＋ push 兩個 repo |
| 資料夾 | `research/3037-analysis-2026q2/` |
| 輸出檔 | `3037_FY2026Q2_Analysis/3037_FY2026Q2_Analysis.html` |
| 同業對照 | 只比 ABF 載板雙雄：南電 8046、景碩 3189 |
| 逐季深度 | 近 12 季（2023Q3–2026Q2） |
| 特殊事件 tab | 由我方從財報／新聞判斷（候選：ABF 載板擴產資本支出與融資結構、AI 載板拉貨、泰國/山鶯路廠區配置） |
| 語言 | 純繁體中文 |
| 本地 port | 8177（`3037-analysis`） |

## 結構參照
- 主要：`research/8046-analysis-2026q1/8046_FY2026Q2_Analysis/8046_FY2026Q2_Analysis.html`（南電，同業、同季、同 Fallback 路徑）— CSS 元件庫逐字沿用
- 範本源：`research/jnj-analysis-2026q2/JNJ_FY2026Q2_Analysis/JNJ_FY2026Q2_Analysis.html`
- 其他台股：taiyen-6274、delta、liteon、2059、1303、3653

## 10 個 tab（Fallback）
1. 法說會簡報重點（grey note + stat-row 4 + quad-grid 4 + note）
2. 本季特殊交易/事件深度拆解（quad-grid 規模/期限/財務承諾/尚待驗證 + 核心疑問 note；含資料中心融資條款追蹤，不適用寫明）
3. 技術與產品線（tech-grid：ABF 載板 / BT 載板 / HDI / 類載板SLP / 一般PCB…）
4. 供應鏈（chain-diagram + Sankey，**同一組節點**；上游 CCL/玻纖布/銅箔/ABF膜/雷射鑽孔設備、下游 IC設計/封測/系統廠）
5. 風險矩陣（risk-item 徽章 + 至少一條獨創前瞻風險 + 內部人申報訊號）
6. 損益表（kpi-row 4 + 損益表 12 季 + line-chart 近12季 + **台股 24 個月月營收 line-chart** + 地區別/產品別）
7. 業務部門（bar-chart + bu-grid：ABF/BT/HDI/軟板/其他，含產品層級技術拆解）
8. 資產負債表（3 期別 + DSO 折線 + 未開始租賃 + 採購承諾 + 客戶預付款/合約負債逐季 + 信評/股利/庫藏股）
9. 現金流量（FCF 橋接 waterfall + 四類拆解 + H1 累計 vs 全年計畫執行進度）
10. 估值觀察（隱含 P/E 表 + 利息保障倍數 + **互證對照表（強制交付物）** + 同業 scatter 南電/景碩，必備圖）

## 6+1 必備圖
chain-diagram、sankey、損益表 line-chart（近12季）、**台股 24 個月月營收 line-chart**、業務部門 bar-chart、FCF 橋接 waterfall、估值 scatter（南電8046 / 景碩3189）

## 資料來源（台股路徑）
- MOPS 合併財報：`https://mopsov.twse.com.tw/server-java/t164sb01?step=1&CO_ID=3037&SYEAR=<西元>&SSEASON=<1-4>&REPORT_ID=C`
- MOPS 月營收（上市）：`https://mopsov.twse.com.tw/nas/t21/sii/t21sc03_<民國年>_<月>.html`（民國=西元−1911）
- MOPS 年報 / 內部人持股轉讓事前申報
- TWSE Open API 交叉驗證：`https://openapi.twse.com.tw/v1/`（僅覆核，全表端點會截斷，個股走 MOPS）
- ⚠️ 月營收「當月增減%」欄實為 MoM 非 YoY；每個單月數字用累計欄反算交叉驗證（本月累計 − 上月累計 = 本月單月）
