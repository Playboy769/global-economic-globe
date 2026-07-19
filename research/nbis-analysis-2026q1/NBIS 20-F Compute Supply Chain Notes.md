# NBIS FY2025 Form 20-F — 計算供應鏈揭露筆記

- **申報公司**：Nebius Group N.V.（CIK 0001513845；荷蘭 N.V.，前身 Yandex N.V.，2024-08-16 更名）
- **文件**：20-F，申報日期 2026-04-30，涵蓋會計年度至 2025-12-31
- **Accession Number**：0001104659-26-052948
- **原始文件**：https://www.sec.gov/Archives/edgar/data/1513845/000110465926052948/nbis-20251231x20f.htm
- 本文件為標準逐條 Item 編號格式（非 ASML 式 Annual Report 混合格式），章節對應清楚，出處頁碼見文末對照表。

> **框架調整說明**：Nebius 是雲端／AI 基礎設施服務商，不是硬體製造商，沒有傳統意義上的零組件供應鏈（BOM）。依使用者指示，本 tab 改為「計算供應鏈」框架：上游＝GPU/晶片依賴＋資料中心／電力／土地取得，下游＝客戶集中度（含 Meta/Microsoft），並納入與 NVIDIA 的股權／商業夥伴關係。

---

## 上游 — GPU／晶片供應

### NVIDIA 依賴（風險因子原文，Item 3.D「RISKS RELATED TO OUR OPERATIONS」第 a 項）

> "We currently rely on Nvidia for the GPU chips we use and on a limited number of other suppliers for other key components in our infrastructure."

風險因子條列的供應商集中度風險包括：關鍵零組件（尤其先進 GPU 晶片）取得受限、對生產成本/交期/條款/定價缺乏掌控、可能被迫簽署高於市價的綁定採購承諺、供應商技術路線圖若落後將直接限制 Nebius 服務能力、供應商優先供應其他客戶的風險、供應商違約風險、地緣政治/天災/公衛事件對供應鏈的衝擊、供應商自身又依賴複雜的第三方半導體製造網絡（Nebius 對此完全不可控）。文中特別點名關稅/反關稅/制裁/出口管制可能影響 GPU 晶片的成本與可得性。**出處**：Item 3.D，年報頁 14（本文件行 774-804）。

### 供應商集中度量化揭露（財報附註 "Certain Risks and Concentrations"）

- **2025 年**：資本支出的供應商集中度歸因於「兩家供應商」（2024 年僅「一家」），**未揭露供應商名稱，也未給百分比數字**（客戶集中度有給百分比，供應商集中度沒有——這是刻意的資訊不對稱揭露）。
- **出處**：財報附註（本文件行 11457-11459）。

### NVIDIA 夥伴關係定性描述（Item 4「Information on the Company」）

> "The performance of our hardware is supported by our long-standing partnerships and collaboration with leading chipmakers and OEMs, such as NVIDIA, and a consistent track record of being one of the first-to-deploy the latest generation of NVIDIA GPU chips."

London（UK）站點（2025-11 簽約）明確點名「featuring NVIDIA Blackwell Ultra GPUs」。**出處**：Item 4，年報頁 33-34（本文件行 1306、1320、1376）。

### NVIDIA 股權暨商業夥伴關係（財報附註 17 SUBSEQUENT EVENTS）

**不是單純股權投資，是預先出資認股權證（pre-funded warrant）**：

- 2026-03-11，Nebius 與 NVIDIA 簽署證券購買協議，私募發行認股權證，可認購 **21,065,396 股 Class A 普通股**（每股面值 €0.01），**行權價僅 $0.0001/股**（實質等同股份），總發行毛額約 **$2.0 billion**。
- 認股權證自 2026-03-11 起 NVIDIA 隨時可行使；行權後股份可由庫藏股撥發或新發行。
- 資金用途：支持全端 AI 雲平台開發＋綠地資料中心的開發與建設。
- 定性為「Nebius 與 NVIDIA 更廣泛策略合作的一部分，以支持 Group AI 雲基礎設施擴張」。
- 現金流量表列為融資活動「Proceeds from issuance of pre-funded warrants sale」$2,000.0M（與可轉債分開列示）。
- **出處**：財報附註 17（本文件行 20029-20031）。

法說會上 Andrey 提到的「5 GW capacity commitment by the end of 2030」「line of sight」為法說會揭露內容，20-F 財報附註未見對應的量化文字（可能在 MD&A 其他段落，本輪未逐頁查核，列為待驗證項）。

---

## 上游 — 資料中心／電力／土地

### 四種資料中心類型（Item 4，年報頁 34-35）

1. **Greenfield（綠地自建）**：自有土地＋自管電力＋自行設計。**Missouri、Finland 為綠地設施**；Finland 擁有業界領先的 PUE（power usage effectiveness）水準，採氣冷式免費冷卻（不依賴外部取水），並有熱回收系統——2025 年回收近 20 GWh 伺服器餘熱，估計為當地降低約 10% 住宅暖氣成本。同樣冷卻設計將複製到 Missouri、Alabama（2026 年初宣布）新綠地站點。
2. **Brownfield（棕地）**：利用既有建物/電力設施加速上線。
3. **Build-to-suit（客製化委建）**：與擁有土地及電力的開發商合作，Nebius 提供客製化規格。**New Jersey（Vineland）設施採此模式**（2025-02 簽約，「我們第一個 build-to-suit 設施」）。
4. **Co-location（共置）**：向第三方機房業者租用容量，快速部署但不擁有設施。目前共置站點包括 France、Iceland、UK、Israel、Kansas City（US）。

### 站點清單（20-F 明確列名，Item 4 年報頁 35）

| 地點 | 類型 | 簽約/公告時間 | 備註 |
|---|---|---|---|
| Mäntsälä, Finland | Greenfield（自有） | －（既有） | 領先 PUE、免取水氣冷、熱回收供暖 |
| Kansas City, Missouri, US | Co-location | 2024-11 | 美國第一個共置資料中心 |
| Paris, France | Co-location | 2024-07 | 第一個共置設施 |
| Keflavik, Iceland | Co-location | 2024-12 | 數千 GPU 叢集 |
| Vineland, New Jersey, US | Build-to-suit | 2025-02 | 第一個 build-to-suit 設施 |
| Israel（一） | Co-location | 2025-10 | 該國首批公開可用 AI 部署場域之一 |
| London, UK | Co-location | 2025-11 | 搭載 NVIDIA Blackwell Ultra GPU，呼應英國政府 AI Opportunities Action Plan |
| Missouri（州內新址） | Greenfield（自有） | 2026-02 | 取得土地與電力容量 |
| Alabama | Greenfield（自有） | 2026-02 | 取得土地與電力容量 |
| Israel（新增兩處） | Co-location | 2026-02 | 新增兩個共置站點 |
| Oklahoma、Minnesota（美）＋France、UK（歐）＋Israel（中東）| 未逐一標示類型 | 2026-02 公告 | 20-F 稱「九個新增站點、七個地點」的擴張公告，**Oklahoma 與 Minnesota 完全未在法說會提及** |

**Pennsylvania（法說會 2026-05-13 當天宣布，1.2GW、「第二個自有 GW 級站點」）不在本 20-F 名單中**——時間點合理，因為 20-F 申報日 2026-04-30 早於法說會兩週，屬於申報後才發生的事件，不代表遺漏或矛盾。

### 容量口徑（三種定義，Item 4）

- **Contracted power（已簽約電力）**：已取得土地與電力承諾的容量。2026 年 2 月時 >2 GW（法說會 5/13 更新為 >3.5GW，目標年底 ≥4GW——約 2.5 個月內從 2GW 衝到 3.5GW）。
- **Connected power（已接電）**：電力已接入資料中心的容量。
- **Active power（實際運轉）**：IT 設備正在消耗、可產生營收的容量。**截至 2025-12-31 僅約 170 MW**。

**⚑ 關鍵反推**：170MW active vs >2GW contracted（2026年2月口徑）——實際「可產生營收」的容量占已簽約容量約 **8.5% 或更低**（若用法說會 5/13 的 3.5GW 口徑計算則約 4.9%）。這量化了法說會「everything we build is sold」敘事背後的另一面：絕大多數已簽約容量目前完全不產生營收，CapEx 認列（$20-25B guidance）與實際可變現產能之間存在巨大時間落差，直接支撐財報 D&A/資本效率相關的風險論證。**出處**：Item 4，年報頁 35（本文件行 1346-1356）。

### 資料中心擴張風險因子（Item 3.D，年報頁 14-16）

完整風險清單包含：專案延遲、預算變動、建材與設備供應/價格延遲、勞工爭議、環境與地質問題、**「日益升高的地方社區對資料中心專案的反對」（increasing public opposition to data center projects in certain localities）**、政治與法規審查、政府單位/公用事業/其他組織核准延遲。另有專段描述能源市場波動、部分司法管轄區要求資料中心業者承擔更多電網升級/發電成本，以及地方社群對能源消耗、環境衝擊、土地使用、在地基礎設施壓力的疑慮可能導致政治壓力、許可延遲、公投倡議、更嚴格法規、訴訟或專案反對。**這與法說會 Tom Blackwell 對「美國資料中心政治反對」問題的公關式回答（強調社區參與、透明度、長期夥伴關係）形成對照——20-F 風險因子把同一議題列為正式、具體的營運風險項目，用詞遠比法說會的公關語言更直接。** **出處**：Item 3.D，年報頁 14-16（本文件行 806-846）。

---

## 下游 — 客戶集中度

### Meta 合約（財報附註 17 SUBSEQUENT EVENTS，正式財報揭露版本）

> "On March 16, 2026, the Group entered into a long-term AI infrastructure supply agreement with Meta. Under the five-year agreement, the Group will provide approximately $12 billion of dedicated compute capacity across multiple locations, based on large-scale deployments of the NVIDIA Vera Rubin platform, with delivery expected to commence in early 2027. In addition, in connection with access to these deployments, Meta has committed to purchase additional available compute capacity across certain upcoming Nebius clusters of up to $15 billion over a five-year period. **The Group currently intends to sell such capacity to third-party customers of its AI cloud business, with any remaining capacity to be purchased by Meta.**"

**⚑ 措辭差異值得注意**：財報附註的預設語氣是「Nebius 打算把這批容量賣給第三方客戶，剩餘的才由 Meta 買」（Meta = 兜底買家／價格支撐者）；法說會上 Marc Boroditsky 的說法則是「我們有裁量權，可以分配給 Meta 或以更高市價賣給 AI Cloud 客戶」（強調 Nebius 的主動選擇權，語氣更中性/正面）。兩者不衝突，但財報書面用語更清楚地把 Meta 定位為容量的最終承接者/保底方，而非單純「其中一個選項」。

- 兩份獨立合約皆已列為 20-F Exhibit：Exhibit 4.2「Cloud Infrastructure Services Agreement dated as of November 1, 2025」（原始合約）＋ Exhibit 4.3「Infrastructure Services Agreement dated as of March 13, 2026」（新擴大合約，注意簽約日 3/13 與附註敘述的 3/16「entered into」有 3 天差異，屬簽署日 vs 生效/宣布日的常見落差，非矛盾）。
- 法說會稱「formally, this is a $27 billion contract」＝ $12B（dedicated）＋$15B（option/backstop），與財報附註數字一致。

### Microsoft 合約

20-F Exhibit 4.4「Statement of Work effective as of September 7, 2025, between Nebius, Inc. and Microsoft」＋ Exhibit 4.5「Addendum Number 1... effective as of January 21, 2026」。財報本文（本輪查核範圍）未見 Microsoft 合約金額的具體揭露；法說會亦未給出總金額數字，僅稱「Microsoft contract」與 Meta 合約並列為資產擔保融資（asset-backed financing）的基礎。**金額未知，屬揭露缺口。**

### 客戶集中度量化揭露（財報附註 "Significant Customers"，年度數字，非季度）

| 客戶 | FY2023 | FY2024 | FY2025 |
|---|--:|--:|--:|
| Customer A | ＜10% | 27% | 25% |
| Customer B | ＜10% | ＜10% | 15% |
| Customer C | ＜10% | 11% | ＜10% |
| Customer D | ＜10% | ＜10% | ＜10% |

**應收帳款集中度（信用風險揭露，年底口徑）**：2024-12-31 為 Customer A，$6.6M／59%；**2025-12-31 為 Customer D，$597.0M／83%**（不具名）。20-F 未點名任何客戶身份。

**⚑ 獨創推論（非財報直接揭露，僅供參考，信心度中等）**：Customer D 在所有揭露年度的「營收」占比皆低於 10% 門檻，卻在 2025 年底占「應收帳款」高達 83%——顯示這是一筆帳款規模遠超過已認列營收的巨額新簽約案。時間點（2025 年底）與 Nebius 和 Meta 的原始合約（Exhibit 4.2，2025-11-01 簽署）高度吻合：新簽約、尚未大規模交付/認列營收，但帳款/開票已大幅入帳。這與 2026 Q1 資產負債表應收帳款進一步翻倍（$720.3M→$1,479.2M）、遞延收入單季暴增近 $3.2B 的走勢方向一致。**此為交叉比對後的合理推論，20-F 並未點名 Customer D 身份為 Meta，勿當作已證實的事實陳述。**

### 客戶產業/地理背景（Item 4，定性描述）

「客戶群目前主要集中在美國，AI 採用正加速全球化」；資料中心遍布 Finland、France、Iceland、UK、Israel，加上美國 GPU 叢集，總部設於阿姆斯特丹。競爭對手（Item 4）：運算層——CoreWeave、Crusoe、Lambda Labs；推論層——Fireworks AI、Together AI。**出處**：Item 4，年報頁 36、44（本文件行 1555、1679）。

---

## 股權結構／關聯方

- **Avride**（自駕車與外送機器人）：Nebius 合併子公司，主要營運於美國 Austin, TX，歐洲/以色列/南韓設有研發據點。2025-02-27 完成集團內部重組，Avride Holding Inc.（德拉瓦州公司）成為 Avride 集團的中間控股公司。2025-03-06 通過 Avride 員工認股計畫，最高可發行相當於 **Avride Group Inc 完全稀釋股本 20%** 的股權獎酬（隱含 Nebius 目前持有 Avride 約 80%+ 完全稀釋權益，未逐字確認精確比例）。法說會確認 Avride「目前仍合併入帳」，管理層意圖未來尋找策略/財務夥伴、解除合併。
- **TripleTen**（教育科技）：合併子公司，Q1 2026 營收僅年增 10%（遠低於核心業務），占集團營收比重由去年同期約 20.6% 稀釋至約 2.9%。
- **ClickHouse**：非合併，「重大少數股權」（significant minority equity stake），持股比例未具體揭露。2026 年 1 月完成 Series D 可轉換特別股融資，由 Nebius 以外投資人出資 $400M，投後估值約 **$15 billion**。這筆交易觸發 Q1 2026 財報認列 **$780.6M「投資證券重估利得」（gain from revaluation of investment in equity securities）**，MD&A 明確指出此利得「主要即為 ClickHouse 重估」，是 Q1 淨利 $621.2M 的主要來源（且此利得本身已超過淨利數字，代表若無此筆非現金利得，公司本業仍是虧損）。
- **Toloka**（資料標註）：2025 年第二季解除合併，此後採權益法認列；Q1 2026「權益法投資損益」-$7.6M 主要反映 Toloka 及少數創投基金部位。
- **Tavily**（併購，2026-02 完成）：現金對價 **$177.3M**（其中 $0.6M 存入託管帳戶），另有最高 $30.3M 或有留任獎金（綁定關鍵員工在職條件）＋最高 $0.4M 或有indemnification調整款。Eigen AI、Clarifai 兩筆併購（法說會提及但 20-F Subsequent Events 未揭露，很可能是在 20-F 申報日 2026-04-30 之後才簽約/完成，或金額未達重大性門檻須揭露）——**兩筆交易的財務條款目前完全沒有任何 SEC 申報文件可查證，屬揭露缺口，法說會上 Alex Dubov 問及併購邏輯時管理層也未提供金額**。
- **創辦人投票權集中**：Class A 股 1 票/股、Class B 股 10 票/股（複數表決權結構）；20-F 風險因子明列「創辦股東投票權集中，限制少數股東影響公司事務（含董事選舉）的能力」，惟具體創辦人持股/投票權百分比本輪未逐表擷取。已發行股數：Class A 220,406,311 股＋Class B 33,491,883 股（2026-03-31 基準）。

---

## 內部控制重大缺失（Item 15，年報頁 84-87 — 法說會完全未提及）

管理層評估截至 2025-12-31，**公司財務報導內部控制「無效」（not effective）**，認定兩項重大缺失（material weakness）：

1. **固定資產相關控制設計/執行無效**：未完整落實折舊起算日期、資產盤點對帳流程的相關控制，導致無法完全信賴伺服器與網路設備等固定資產的資料完整性與準確性。
2. **TripleTen 事業單位資訊科技一般控制（ITGC）未及時落實**，營收認列相關業務流程控制執行未充分留存紀錄（TripleTen 於揭露時點約占總營收 10%——此為年度口徑，非 Q1 2026 的 2.9%）。

管理層聲明上述缺失「認定未造成」歷史財報重大誤述，2025 年及 2026 年迄今已採取補償性控制措施與整體補救計畫（引入外部顧問），**預計於 2026 年底前完成補救**，但無法保證屆時能完全補救。

**⚑ 交叉關聯**：固定資產／折舊起算日控制缺失，直接對應 Q1 2026 財報 D&A 費用暴增至 $212.0M（年增 332%）——這正是造成「調整後 EBITDA 轉正（$130M/13%）」與「GAAP 營業虧損反而擴大（-$128.0M）」兩套敘事分歧的最主要科目。換言之，最容易被管理層/市場用來論證「獲利能力改善」的调整後指標，其排除項目（D&A）恰好落在公司自己向 SEC 承認控制尚未到位的科目上。**出處**：Item 15，年報頁 84-87（本文件行 5864-5928）。

---

## 資料限制說明

- 供應商集中度（CapEx 口徑）僅揭露「一家/兩家供應商」，無金額、無名稱——Sankey 圖上游 GPU/供應商節點的連結寬度**只能用客戶下游揭露的相對量級或法說會敘事做示意性呈現，不可比照下游客戶連結線寬**（下游有具體 % 數字，上游沒有）。
- Customer A/B/C/D 為 20-F 原文匿名代號，Customer D＝Meta 為本文交叉比對後的**推論**，非申報文件直接揭露，繪圖與行文都需明確標註信心等級。
- Microsoft 合約金額未見於本輪查核的任何文件（20-F 本文或 6-K），Sankey 圖中 Microsoft 節點的連結寬度須標記為「未揭露金額，示意性」。
- Eigen AI、Clarifai 兩筆併購的財務條款完全未揭露（無論金額或於哪份文件），僅知交易存在（法說會提及）。
- Form 4 內部人交易：確認申報存在（含 2026-05-13 法說會當天交易日的 Form 4／Form 4/A），但因本輪 SEC EDGAR 存取於研究後段間歇性中斷，未能取得申報人身份與交易明細，不納入本次報告的內部人訊號分析，僅註記存在。

---

## 出處對照表

| 揭露主題 | 20-F Item | 年報頁碼（本文件行號） |
|---|---|---|
| 風險因子總覽（供應商依賴、資料中心擴張、地緣政治） | Item 3.D | 頁 5-16（行 478-859） |
| NVIDIA GPU 依賴細節 | Item 3.D | 頁 14（行 774-804） |
| 業務概況、資料中心類型與站點清單 | Item 4 | 頁 33-36（行 1282-1382） |
| 容量口徑（contracted/connected/active） | Item 4 | 頁 35（行 1346-1356） |
| 內部控制重大缺失 | Item 15 | 頁 84-87（行 5864-5928） |
| Major Shareholders 表格（創辦人投票權） | Item 7 | 頁 65 起（行 4481+） |
| 客戶/供應商集中度、應收帳款集中度 | 財報附註（Certain Risks and Concentrations） | F 頁碼（行 11333-11460） |
| NVIDIA 認股權證、Meta 合約、Tavily 併購、可轉債、ClickHouse Series D | 財報附註 17 Subsequent Events | F-57～F-59（行 20013-20055） |
| Exhibit 清單（Meta/Microsoft 合約、NVIDIA warrant 等文件名） | Item 19 | 頁 92 起（行 20085+） |

---
*備註：本文件為讀取用途的萃取整理，非法律或投資建議；金額除特別標註外均為美元。文中「⚑」與「獨創推論」段落為分析性解讀，已明確與申報文件原文區隔。*
