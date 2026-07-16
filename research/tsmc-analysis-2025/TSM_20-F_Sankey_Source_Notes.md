# TSM 20-F (FY2025) — Sankey 圖表資料來源筆記

- **申報公司**：Taiwan Semiconductor Manufacturing Co Ltd（CIK 0001046179）
- **文件**：20-F（外國發行人年報），申報日期 2026-04-16，涵蓋會計年度至 2025-12-31
- **Accession Number**：0001628280-26-025362
- **原始文件**：https://www.sec.gov/Archives/edgar/data/1046179/000162828026025362/tsm-20251231.htm

以下每組數字都標注在 20-F 內的出處（Item / 附註），供三張 Sankey 圖回溯查核。

## 1. 供應鏈與客戶集中度（供應鏈 tab）

**上游 — 關係人採購交易**（Item 7. Major Shareholders and Related Party Transactions）：
- VIS（Vanguard International Semiconductor，TSMC 持股 27.6%，manufacturing agreement）：2025年採購 NT$878M（US$28M），占總營業成本 0.1%
- SSMC（Systems on Silicon Manufacturing，TSMC 持股 38.8%，與 NXP 合資，SSMC 產能專供 TSMC/NXP）：2025年採購 NT$4,113M（US$131M），占總營業成本 0.3%
- Xintec（精材科技，晶圓級封裝服務，TSMC 持股 41.0%）：2025年製造費用 NT$5,448M（US$174M），占總營業成本 0.4%

**上游 — 設備與原物料供應商**（Item 4.B Business Overview – Equipment / Raw Materials）：
20-F 僅揭露類別與地區，**未點名個別供應商，亦未揭露採購金額**，因此未計入 Sankey 線寬，僅列為文字說明：
- 設備：微影掃描機、清洗機、蝕刻機、爐管、濕式站、離子植入機、CVD、CMP、測試機/探針台等，「部分設備供應商數量有限」
- 原物料：矽晶圓（主要供應商位於台灣、日本、德國、新加坡）、化學品、氣體、貴金屬，「部分原物料採單一來源採購」

**下游 — 客戶集中度**（Item 18 Financial Statements, Note: Geographic and major customers' information）：
2025年占淨營收 10% 以上的客戶：
- Customer A：NT$726,974.3M，19%（2024年為 NT$352,271.2M，12%；2023年未達10%）
- Customer B：NT$645,178.7M，17%（2024年 NT$624,345.5M，22%；2023年 NT$546,550.9M，25%）
- 其他客戶合計（個別均低於10%揭露門檻，本筆為推算值）：2025淨營收總額 NT$3,809,054M − Customer A − Customer B = NT$2,436,901.0M（約64%）
- 20-F 未揭露客戶真實身份（僅稱 Customer A/B/C），圖中維持原始匿名代號，不做外部臆測

**下游 — 關係人銷售交易**（Item 7）：
- GUC（創意電子 Global Unichip，TSMC 持股 34.8%，SoC設計服務）：2025年銷售 NT$31,094M（US$991M），占總營收 0.8%

## 2. 2025 營收分部（營收分部 tab）

來源：Item 4.B「net revenue by platform」表 + 同段落前段「geographic breakdown of our net revenue」表。兩表總額一致（NT$3,809,054M），故可用同一 TSMC 營收節點串接。

**依地區**（2025）：North America 75%／Asia Pacific(不含中日) 9%／China 9%／Japan 4%／EMEA 3%

**依平台**（2025）：High Performance Computing 58%／Smartphone 29%／IoT 5%／Automotive 5%／Digital Consumer Electronics 1%／Others 2%

## 3. 股權結構（股權結構 tab）

**誰持有 TSMC**（Item 7 Major Shareholders，資料基準日 2026-02-28）：
- 國發基金（National Development Fund, Executive Yuan）：1,653,709,980股，6.38%
- ADS（美國存託憑證）持有人，經 Citibank N.A.（存託機構）名義持有：5,313,450,863股（對應 1,062,690,167 ADS），約20.49%（=5,313,450,863 / 25,932,524,521）
- 董事及高階經理人合計：60,237,681股，0.23%
- 已發行普通股總數：25,932,524,521股
- 其他分散股東（推算值 = 總數 − 上列三者）：18,905,125,997股，約72.90%

**TSMC 對外主要權益投資**（Item 18, Note: Investments Accounted for Using Equity Method，帳面價值為 2025-12-31 資產負債表數字；持股%取 Item 7 較新的 2026-02-28 基準）：
- VIS：帳面價值 NT$18,068.7M，持股 28%（財報附註）／27.6%（Item 7，2026-02-28）
- SSMC：帳面價值 NT$12,419.2M，持股 39%／38.8%
- Xintec：帳面價值 NT$4,470.4M，持股 41%／41.0%
- GUC：帳面價值 NT$2,893.6M，持股 35%／34.8%

## 資料限制說明

- 設備與原物料供應商未具名、未揭露金額，Sankey 圖僅呈現有揭露金額的關係人交易與客戶集中度，設備/原物料供應商改列為文字附註，避免用虛構數字撐出線寬。
- Customer A/B 為 20-F 原文匿名代號，非本文推測其真實身份。
- 「其他客戶合計」「其他分散股東」為由已揭露數字反推的計算值，非申報文件直接揭露的單一數字。
