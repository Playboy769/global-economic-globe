# 景碩科技 (Kinsus Interconnect Technology, TWSE: 3189) — FY2026 Q2 財報完整分析 · 建置筆記

## 決策 (2026-08-31 使用者確認，11 題澄清)

| 項目 | 選擇 |
|---|---|
| 分析框架 | 完整 10-tab 合併單檔 HTML（Earnings Call 分析框架） |
| 期別 | FY2026 Q2（民國115年第二季，MOPS 公告截止 2026-08-14） |
| 模式 | **Fallback**（無逐字稿）— 略過 Q&A 話術偵測 tab，改「法說會簡報重點」；風險矩陣註明資料限制；加重 24 個月月營收反推 |
| 資料來源 | 我方直接抓 MOPS t164sb01 合併財報＋月營收頁，TWSE OpenAPI 交叉覆核 |
| 同業對照 | 台廠載板三雄（南電 8046、欣興 3037）為主，另補國際 ABF 載板廠（Ibiden、Shinko、AT&S） |
| 產業角度 | ABF 載板循環為主軸（AI/伺服器 GPU 需求、2023-24 下行後復甦） |
| 語言 | 純繁體中文 |
| 收尾 | 鏡像 globe-invest/app/research/ ＋ 上架 OutsideFramework Works ＋ push 兩個 repo |
| 資料夾 | `research/3189-analysis-2026q2/` |
| 輸出檔 | `3189_FY2026Q2_Analysis/3189_FY2026Q2_Analysis.html` |
| 逐季深度 | 近 12 季（2023Q3–2026Q2） |
| 本地 port | 8179（`3189-analysis`）— 8178=macro-tracker、8180=nvda、8179 空號 |

## 結構參照
- **主要**：`research/3037-analysis-2026q2/3037_FY2026Q2_Analysis/3037_FY2026Q2_Analysis.html`（欣興，同業、同季、同 Fallback 路徑）— CSS 元件庫逐字沿用
- 次要：`research/8046-analysis-2026q1/8046_FY2026Q2_Analysis/8046_FY2026Q2_Analysis.html`（南電，同業）
- 範本源：`research/jnj-analysis-2026q2/JNJ_FY2026Q2_Analysis/JNJ_FY2026Q2_Analysis.html`
- 既有 3189 素材：`research/3037-analysis-2026q2/data/peer_valuation.md`（含景碩全套估值/EPS/營收 YoY 數字，2026-08-28 收盤基準）

## 10 個 tab（Fallback）
1. 法說會簡報重點（grey note + stat-row + quad-grid + note）
2. 本季特殊交易/事件深度拆解（私募引進台系晶圓代工廠＋美系科技大廠、3 年 235 億資本支出、切入美系 AI GPU 載板第三供應商）
3. 技術與產品線（tech-grid：ABF 載板 / BT 載板 / 覆晶封裝載板 / 記憶體載板…）
4. 供應鏈（chain-diagram + Sankey，**同一組節點**；上游 ABF膜/CCL/銅箔/玻纖布/雷射鑽孔設備、下游 IC設計/封測/AI GPU 客戶）
5. 風險矩陣（risk-item 徽章 + 至少一條獨創前瞻風險 + 內部人申報訊號）
6. 損益表（kpi-row + 損益表 12 季 + line-chart 近12季 + **台股 24 個月月營收 line-chart** + 地區別/產品別）
7. 業務部門（bar-chart + bu-grid，含產品層級技術拆解）
8. 資產負債表（3 期別 + DSO 折線 + 未開始租賃 + 採購承諾 + 客戶預付款/合約負債逐季 + 信評/股利/庫藏股）
9. 現金流量（FCF 橋接 waterfall + 四類拆解 + H1 累計 vs 全年計畫執行進度）
10. 估值觀察（隱含 P/E 表 + 利息保障倍數 + **互證對照表（強制交付物）** + 同業 scatter 南電/欣興，必備圖）

## 6+1 必備圖
chain-diagram、sankey、損益表 line-chart（近12季）、**台股 24 個月月營收 line-chart**、業務部門 bar-chart、FCF 橋接 waterfall、估值 scatter（南電8046 / 欣興3037）

## 資料來源（台股路徑）
- MOPS 合併財報：`https://mopsov.twse.com.tw/server-java/t164sb01?step=1&CO_ID=3189&SYEAR=<西元>&SSEASON=<1-4>&REPORT_ID=C`
- MOPS 月營收（上市）：`https://mopsov.twse.com.tw/nas/t21/sii/t21sc03_<民國年>_<月>.html`（民國=西元−1911）
- MOPS 年報 / 內部人持股轉讓事前申報
- TWSE Open API 交叉驗證：`https://openapi.twse.com.tw/v1/`（僅覆核，全表端點會截斷，個股走 MOPS）
- ⚠️ 月營收「當月增減%」欄實為 MoM 非 YoY；每個單月數字用累計欄反算交叉驗證（本月累計 − 上月累計 = 本月單月）

## 進度
- [ ] MOPS 合併財報 2026Q2 / Q1 / FY2025 全套
- [ ] MOPS 合併財報 2023Q3 – 2024Q4（逐季趨勢用）
- [ ] MOPS 24 個月月營收
- [ ] 年報供應鏈揭露
- [ ] 內部人申報
- [ ] 估值/籌碼補查（大多已在 peer_valuation.md）
- [ ] 建 HTML 10 tab
- [ ] 7 張必備圖
- [ ] 收尾：鏡像 + 上架 + push
