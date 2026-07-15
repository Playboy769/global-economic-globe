# GS SEC Filing 交叉查詢筆記

**限制說明**：Q2 FY2026（截至2026/6/30財季）10-Q 尚未申報（大型加速申報者一般於季末後40天內申報，預計2026年8月初）。本次交叉驗證改用最近一期已申報的 **10-Q（截至2026年3月31日，2026/5/1申報，accession 0000886982-26-000118）**，涵蓋資本結構、法律訴訟、Platform Solutions、信用集中度等「結構性」科目；Q2 當季損益數字本身以法說會逐字稿/簡報揭露為準。

## 資本比率（10-Q，Standardized approach）

| 指標 | 2026/3（Q1 10-Q） | 2025/12（10-Q對照欄） | 2026/6（法說會口頭揭露） |
|---|---|---|---|
| CET1 ratio (Standardized) | 12.5% | 13.3% | 12.9% |
| Tier 1 capital ratio (Std) | 14.1% | 15.1% | — |
| Total capital ratio (Std) | 15.9% | 16.6% | — |
| CET1 ratio (Advanced) | 14.3% | 15.1% | — |
| CET1 capital ($ Dec25) | — | $104,297M | — |
| RWA (Advanced, Dec25) | — | $727,338M | — |

來源：10-Q「Capital Ratios」表格，Item 2 MD&A。

## SLR（Supplementary Leverage Ratio）與 eSLR 改革

- **10-Q 揭露的 SLR 要求**：2026/3 = 3.75%；2025/12 = 5.0%（要求本身下降，非僅GS自行優化使用率）。
- 原因：**2026/1/1 起 GS 提前採用修訂版 Enhanced SLR（eSLR）標準**，將 G-SIB 適用的緩衝從固定 2% 改為「G-SIB附加資本的50%（Method 1）」。
- 10-Q 揭露槓桿比率原始數據：Tier 1 capital $115,189M ／ Total leverage exposure $2,476,612M（2026/3）→ 反推實際 SLR ≈ 4.65%；Dec 2025：Tier 1 $118,943M。
- 法說會（Christian Bolu 提問）：Q2 2026 實際 SLR = **4.3%**，較前期下滑約40bp（與10-Q反推的4.65%→4.3%相符，落差約35bp，量級一致）。
- ⚑ **交叉驗證重點**：法說會上管理層將 SLR 下滑歸因於「業務成長（prime/financing擴張）」，但10-Q顯示同期SLR**要求本身**也因監管新規下降（5.0%→3.75%），兩者疊加使GS有更大空間擴張低RWA、槓桿密集的融資業務（如prime brokerage），而這部分未在法說會上被管理層主動提及或被分析師追問。

## 股票回購與股利

- 10-Q：2026年1-3月回購 5,416,274 股，均價約$906，截至2026/3月底剩餘授權回購額度 **$27.0B**。
- 法說會：Q2 2026 回購 $4B；季度股利上調至 $5/股（+25% YoY，五年+150%）。

## Platform Solutions／Apple Card 退出

- 10-Q：「Substantially all of the revenues in Platform Solutions are from activities related to issuing credit cards to and raising deposits from Apple Card customers and related to businesses that have been exited. In December 2025, we entered into an agreement to transition the Apple Card program to another issuer. The transition is expected to be completed in approximately 24 months from the date of the agreement.」
- 換算：轉出完成時間點約為 **2027年12月前後**。
- 10-Q Platform Solutions資產（2026/3）：貸款 $19,116M（2025/12: $19,823M，略降）。
- 法說會：Platform Solutions Q2營收僅 $221M，且管理層預期「後續季度大致持平」——與10-Q揭露的退出時程一致，屬確定中的萎縮業務，非成長引擎。

## Industry Ventures 收購結構（AWM無機成長）

- 10-Q Part II Item 2：GSAM Ignite Holdings LP 於2026/1/2發行約40萬股可交換單位（公允價值約$315M）作為收購 Industry Ventures 的部分對價；另約定最多再發行25萬股可交換單位，需視 Industry Ventures 至2030年的績效目標達成情況（部分以現金結算）。
- 對應法說會敘述：Solomon稱 Industry Ventures 與 Innovator 兩筆收購「整合初期動能良好」，但未提及對賭條款細節；10-Q揭露的2030年績效門檻代表這筆收購仍有相當比例對價尚未確定兌現。

## 信用集中度（10-Q Note 26，2026/3）

| 貸款類別 | 帳面價值 | Americas | EMEA | Asia |
|---|---|---|---|---|
| Corporate | $38,156M | 69% | 22% | 9% |
| Commercial real estate | $38,779M | 78% | 18% | 4% |
| Residential real estate | $32,885M | — | — | — |

## 法律訴訟（10-Q Item 1 / Note 27）

標準化揭露語言，無本季新增重大個案被特別點名；管理層維持「訴訟費用可能維持高檔」的一般性警語。截至Q1 10-Q無足以影響財務狀況的重大不利認定。

## Form 4 內部人交易訊號（2026年4月中—6月17日，法說會前約4週至11週）

透過 SEC EDGAR 逐筆下載並解析（非透過xslF345X06樣式頁，直接讀取原始XML）：

| 申報人 | 職稱 | 交易日 | 性質 | 備註 |
|---|---|---|---|---|
| David M. Solomon | Chairman & CEO | 2026/5/1 | 出售 3,470股，均價$930.43/$931.25 | 透過家族信託持有，footnote未載明是否為Rule 10b5-1計畫 |
| Denis P. Coleman | CFO | 2026/5/14 | 出售 6,857股，均價$972.29–$975.21 | 同上，未載明10b5-1 |
| Kathryn H. Ruemmler | Chief Legal Officer | 2026/5/1, 2026/5/6 | 分兩批出售，共約13,790股，$928–$941 | 同上 |
| Sheara J. Fredman | Chief Accounting Officer | 2026/5/1 | 出售 10,301股，$925.18/$930.21 | 同上 |

- 觀察：四位最高層具名高管（CEO/CFO/CLO/CAO）**均於同一窗口（5/1–5/14）出售股票**，且股價已處於當時歷史相對高檔（$925–$975，事後Q2財報公布後股價進一步走高）；**無任何一筆內部人買進**紀錄；6/17後至法說會（7/14）當日無新增Form 4。
- 規模判斷：以GS股價量級與高管一般持股基數衡量，屬**例行、小額**出售，非集中大額出脫；filing footnote未明載10b5-1計畫，故無法100%排除非預先排程賣出，但同一窗口四人同步賣出的模式，較符合年度分紅股票解禁後常規賣股節奏，而非單一突發訊號。
- （另有一筆2026/6/17申報人為「GOLDMAN SACHS GROUP INC」本身作為做市商就*其他發行人*股票的Form 4，經核為做市業務常規申報，與GS自身內部人動向無關，已排除。）

---
*本筆記僅供 GS_FY26Q2_Analysis / GS_FY26Q2_Financials 兩份報告交叉引用，非投資建議。*
