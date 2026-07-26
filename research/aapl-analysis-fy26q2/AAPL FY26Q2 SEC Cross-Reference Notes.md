# AAPL — SEC EDGAR 交叉查詢筆記（FY2026 Q2）

擷取日期：2026-07-26 · CIK 0000320193 · 全部經 `data.sec.gov` 官方 API／`www.sec.gov/Archives` 取得（帶 User-Agent）

| 文件 | Accession | 申報日 | 期別 | 主檔 |
|---|---|---|---|---|
| 10-Q（本季） | 0000320193-26-000013 | 2026-05-01 | 2026-03-28 | `aapl-20260328.htm` |
| 10-Q（前季，比較基期） | 0000320193-26-000006 | 2026-01-30 | 2025-12-27 | `aapl-20251227.htm` |
| 10-Q（去年同季） | 0000320193-25-000057 | 2025-05-02 | 2025-03-29 | `aapl-20250329.htm` |
| 10-K（FY2025） | 0000320193-25-000079 | 2025-10-31 | 2025-09-27 | `aapl-20250927.htm` |
| Form 4 × 11 檔 | 見下方內部人一節 | 2026-04-03 ～ 2026-06-17 | — | — |

---

## 一、損益表（10-Q Part I Item 1，單位：US$M）

| 科目 | FY26 Q2 | FY25 Q2 | YoY | FY26 Q1（推算） | QoQ |
|---|---|---|---|---|---|
| Products | 80,208 | 68,714 | +16.7% | 113,743 | −29.5% |
| Services | 30,976 | 26,645 | +16.3% | 30,013 | +3.2% |
| 總營收 | 111,184 | 95,359 | +16.6% | 143,756 | −22.7% |
| 毛利 | 54,781 | 44,867 | +22.1% | 69,231 | −20.9% |
| 毛利率 | 49.27% | 47.05% | +2.2pp | 48.16% | +1.11pp |
| R&D | 11,419 | 8,550 | +33.6% | 10,887 | +4.9% |
| SG&A | 7,477 | 6,728 | +11.1% | 7,492 | −0.2% |
| 營業利益 | 35,885 | 29,589 | +21.3% | 50,852 | −29.4% |
| 有效稅率 | 17.5% | 15.5% | +2.0pp | 17.5%(6M) | — |
| 淨利 | 29,578 | 24,780 | +19.4% | 42,097 | −29.7% |
| 稀釋 EPS | $2.01 | $1.65 | +21.8% | $2.84 | −29.2% |

Products / Services 毛利率：38.7% / 76.7%（去年同季 35.9% / 75.7%；前季推算 40.68% / 76.52%）。

**⚑ SG&A 一次性費用回推**：Kevan 說 opex 略高於財測上緣是「a one-time expense in SG&A」，但未給金額。FY25 同期 SG&A QoQ 為 −6.2%（7,175→6,728）；FY26 為 −0.2%（7,492→7,477）。若沿用去年季節性衰退幅度，本季 SG&A 應約 7,027，**差額約 4.5 億美元**即為該一次性費用的量級。同一份 10-Q 的 Legal Proceedings 寫「The Company settled certain matters during the second quarter of 2026」——時間點吻合。

---

## 二、業務別與地區別（10-Q Note 2 / Note 10）

**產品線（US$M）**：iPhone 56,994（+22%）／Mac 8,399（+6%）／iPad 6,914（+8%）／穿戴家居配件 7,901（+5%）／Services 30,976（+16%）

**地區（US$M，Q2）**：Americas 45,093（+12%）／Europe 28,055（+15%）／Greater China 20,497（+28%）／Japan 8,401（+15%）／Rest of Asia Pacific 9,138（+25%）

**地區營益率（自 Note 10 計算）**：Americas 43.0%（去年 41.6%）／Europe 46.5%（42.2%）／Greater China 44.8%（41.4%）／**Japan 45.7%（47.1%，唯一下滑）**／RoAP 45.2%（41.0%）。Corporate 費用 −13,695（去年 −10,547，+29.8%）。

> Note 2 附註：「except in Greater China, where iPhone revenue represented a moderately higher proportion of net sales」——大中華區的 iPhone 佔比高於集團平均，是財報唯一明示的區域產品組合差異。

---

## 三、資產負債表關鍵變動（10-Q Part I Item 1 + Note 5）

| 科目（US$M） | 2026-03-28 | 2025-09-27 | 變動 |
|---|---|---|---|
| 現金及約當現金 | 45,572 | 35,934 | +9,638 |
| 有價證券（流動＋非流動） | 101,023 | 96,486 | +4,537 |
| 應收帳款 | 30,339 | 39,777 | −9,438 |
| 供應商非貿易應收款 | 23,172 | 33,180 | −10,008 |
| 存貨 | 6,747 | 5,718 | +1,029 |
| **無形資產（非流動）** | **21,334** | **11,093** | **+10,241** |
| 其他非流動資產 | 77,430 | 72,634 | +4,796 |
| 應付帳款 | 57,349 | 69,860 | −12,511 |
| 商業本票 | 1,997 | 7,979 | −5,982 |
| 長期負債（流動＋非流動） | 82,714 | 90,678 | −7,964 |
| **其他非流動負債** | **55,546** | **41,549** | **+13,997** |
| 股東權益 | 106,491 | 73,733 | +32,758 |

**Note 5 無形資產明細**：毛額 24,950 → **37,767（+12,817）**；累計攤銷 11,649 → 11,970（僅 +321）；淨額 13,301 → 25,797，其中流動 2,208 → 4,463。

**⚑ 揭露層級改變**：FY26 Q1 的 10-Q 資產負債表**沒有**「Intangible assets, net」這一行（Other non-current assets 為 93,146）；Q2 才把它獨立拉出來，並把 2025-09-27 的比較數重編為 72,634 + 11,093 = 83,727（與 Q1 表列數一致）。獨立列示通常代表金額跨過重大性門檻。

**⚑ 資產與負債同步跳升**：其他非流動資產＋無形資產合計由 83,727 →（Q1）93,146 →（Q2）98,764，六個月 +15,037；其他非流動負債同期 41,549 → 52,055 → 55,546，+13,997。兩者幾乎等額同步，型態符合「認列一項長期資產、同時認列對應的長期給付義務」，且大部分發生在 **FY26 Q1（12 月季）**。法說會與 Q&A 全程未提。

---

## 四、承諾與或有事項（10-Q Note 9 + Item 2 Liquidity）

| 項目（US$bn） | FY25 10-K<br>(2025-09-27) | FY26 Q1<br>(2025-12-27) | FY26 Q2<br>(2026-03-28) | FY25 Q2<br>(2025-03-29) |
|---|---|---|---|---|
| 製造採購承諾 | 56.2 | 44.4 | 44.6 | 38.4 |
| **其他採購承諾** | **14.8** | **35.1** | **30.4** | 未單獨揭露 |
| ├ 12 個月內到期 | 7.0 | 9.3 | 9.3 | — |
| 無條件採購義務（>1年，Note 9） | — | — | 27.7 | — |

- 製造採購承諾 YoY +16.1%（38.4→44.6），與營收 +16.6% 同步，**看不到為了因應缺料而超前備料的跡象**。
- 其他採購承諾在 12 月季一口氣 **+20.3bn**，Q2 回落 4.7bn。10-Q 對其內容的描述為「supplier arrangements, licensed intellectual property and content, distribution rights, and the acquisition of capital assets related to product manufacturing」。
- Note 9 無條件採購義務到期表（US$M）：2026 剩餘 2,994／2027 7,343／2028 6,130／2029 5,394／2030 5,281／其後 549，合計 27,691 — **加權期限約 3.6 年，且「其後」只有 549**，代表這批承諾集中在未來四年、幾乎沒有超過五年的長尾。

---

## 五、現金流量（10-Q，六個月累計，US$M）

| 項目 | FY26 6M | FY25 6M |
|---|---|---|
| 營運現金流 | 82,627 | 53,887 |
| 折舊攤銷 | 6,653 | 5,741 |
| 股份基礎薪酬 | 7,122 | 6,512 |
| 其他流動及非流動資產變動 | −14,329 | −4,371 |
| 應付帳款變動 | −12,297 | −14,604 |
| 其他流動及非流動負債變動 | +7,301 | −15,579 |
| **資本支出** | **−4,344** | **−6,011** |
| 投資活動淨額 | −11,054 | +12,709 |
| 庫藏股買回 | −36,989 | −49,504 |
| 股利 | −7,743 | −7,614 |
| 償還長期債務 | −7,914 | −4,009 |
| 償還商業本票淨額 | −5,911 | −3,968 |
| 現金所得稅（補充揭露） | 20,397 | 31,683 |

Kevan 在電話會議說 Q2 單季營運現金流 28.7bn（六個月 82.6bn ⇒ Q1 約 53.9bn）。

**⚑ 資本支出反向**：R&D +33.6% 且 10-Q 明寫增幅「primarily due to higher infrastructure-related costs」，但同期 capex **減少 27.7%**（6,011→4,344）。AI 基礎設施走的是費用化／租賃而非自建。

**⚑ 遞延海外盈餘稅**：六個月內付清 TCJA 遞延匯回稅剩餘 **8.8bn**，是現金稅負由 31.7bn 降至 20.4bn 的反向項（去年同期還包含愛爾蘭 State Aid 15.4bn 解付）。

---

## 六、集中度（10-Q Note 4「Accounts Receivable」）

| 指標 | FY24 年底 | FY25 年底<br>(2025-09-27) | FY26 Q1<br>(2025-12-27) | FY26 Q2<br>(2026-03-28) |
|---|---|---|---|---|
| 單一客戶佔貿易應收 | — | 12% | 12% | **17%** |
| 電信商合計佔貿易應收 | — | 34% | 35% | **30%** |
| 供應商非貿易應收 — 第一大 | 44% | 46% | 47% | **51%** |
| 供應商非貿易應收 — 第二大 | 23% | 23% | 26% | **18%** |

第一大供應商佔比連四期走高並首度過半；第二大同期大幅回落。Apple 從不揭露其身分。

---

## 七、10-K（FY2025）供應鏈揭露

**Item 1 — Supply of Components**
- 多數關鍵零組件「generally available from multiple sources」，但**部分零組件目前取自單一或有限來源**（single or limited sources）。
- 使用**同業不用的客製零組件**；新產品採用的客製件常常「available from only one source」。
- 新技術導入初期會因供應商良率／產能未成熟而有產能限制。
- 貿易限制會推升成本或限制零組件、**稀土與其他原物料**的可得性。

**Item 1A — 委外製造**
- 「A significant majority of the Company's manufacturing is performed in whole or in part by outsourcing partners located primarily in **China mainland, India, Japan, South Korea, Taiwan and Vietnam**, in addition to sourcing from partners and facilities located in the U.S.」
- 「The Company relies on **single-source partners in the U.S., Asia and Europe** to supply and manufacture many components, and on partners **primarily located in Asia, for final assembly of substantially all of the Company's hardware products**.」
- 運輸與物流管理亦大量外包。

**Item 1A — 信用集中**
- 「a significant portion of the Company's trade receivables can be concentrated within **cellular network carriers** or other resellers」
- 「the Company has made **prepayments associated with long-term supply agreements to secure supply of inventory components**」
- 「As of September 27, 2025, the Company's vendor non-trade receivables were **concentrated among a few individual vendors located primarily in Asia**.」

**Item 1 — 通路**：FY2025 直營／間接通路佔總營收 **40% / 60%**。

**Item 1 — 人力**：FY2025 年底約 **166,000** 名全職人員。

**FY2025 承諾（10-K Liquidity）**：製造採購承諾 56.2bn（55.4bn 於 12 個月內）／其他採購承諾 14.8bn／固定租賃給付義務 16.8bn（含資料中心）／Notes 未來利息 37.0bn。

> **限制說明**：Apple 的 10-K／10-Q **不點名任何供應商或客戶**。本報告供應鏈圖中的具名節點（TSMC 等）僅來自法說會逐字稿明示者；其餘節點以 filing 揭露的「類別＋集中度」層級呈現，並在圖旁註明出處。

---

## 八、稅務、關稅與法律（10-Q Item 2 / Part II Item 1）

- **關稅時序**：2026-01-14 商務部 Section 232 半導體調查初步結果公布，**未對 Apple 產品加徵新關稅**；2026-02-20 **最高法院判決推翻部分 IEEPA 關稅**，Apple 正依 CBP 程序申請退稅；Section 122 下的關稅另行課徵。Tim Cook 表示退稅款將「reinvest into U.S. innovation and advanced manufacturing」，且屬既有承諾之外的新增投資。**退稅金額未量化、未入帳。**
- **EU DMA**：Article 5(4) 已於 2025-04-23 遭罰 €500M（上訴中）；Article 6(4) 若最終認定違規，罰則上限為**全球年度淨銷售額的 10%**。
- **Epic**：2025-12-11 第九巡迴法院部分維持、部分修改 2025 年禁制令，**准許 Apple 對 link-out 交易收取佣金**，發回地院修改命令。
- **Google 搜尋授權**：D.C. 地院 2025-09-02 命令中的救濟措施仍在後續程序與上訴中；若上訴翻案採 DOJ 原提案（禁止 Google 向 Apple 提供搜尋分潤條件），將直接衝擊該筆授權收入。
- **新增風險因子（相對 FY25 10-K）**：本季 10-Q 新增／改寫四項風險因子，其中「net sales and gross margins are subject to volatility and downward pressure」明列 **industry-wide supply constraints and increasing costs for components such as advanced semiconductors, storage (NAND) and memory (DRAM)**，並新增獨立的 AI 風險因子。
- **Item 2 MD&A 巨觀段**新增：「The Company **expects these trends to intensify**, which, together with actions that may be taken by the Company in response to such trends, **may materially adversely affect demand** for the Company's products…」——比法說會口徑強硬得多。

---

## 九、庫藏股與股利（10-Q Note 7 / Part II Item 2）

| 期間 | 買回股數（千股） | 均價 |
|---|---|---|
| 2025-12-28 ～ 2026-01-31 | 41,032 | $259.26 |
| 2026-02-01 ～ 2026-02-28 | 1,395 | $265.20 |
| **2026-03-01 ～ 2026-03-28** | **0** | **—** |
| Q2 合計 | 42,427 | — |

- Q2 買回金額 11.0bn（去年同季 25.2bn）；六個月 36.0bn（135M 股）。
- 2025-05-01 授權的 $100bn 計畫至 2026-03-28 已用 36.2bn，**剩餘額度 63.8bn**；2026-04-30 再授權 $100bn。
- 季配息由 $0.26 調升至 $0.27（+3.85%），自 FY26 Q3 起發放。

**⚑ 三月零買回**：整整四週沒有任何庫藏股執行，而同期 Apple 已擁有 63.8bn 未動用額度。Kevan 在電話會議只說「Our repurchase activity at any time can be affected by a number of factors that we take into account. As you're aware, we recently announced a CEO transition.」——把兩句話放在一起，但沒有把它們連起來說。

---

## 十、Form 4 內部人交易（法說會 2026-04-30 前後）

| 申報日 | 交易日 | 申報人 | 職務 | 交易 | 股數 | 價格 | 交易後持股 | 10b5-1 註記 |
|---|---|---|---|---|---|---|---|---|
| 2026-04-03 | 04-01 | Timothy D. Cook | CEO、董事 | M（RSU 交割） | 131,576 | — | 3,411,994 | — |
| | 04-01 | | | F（扣稅） | 66,627 | 255.63 | 3,345,367 | — |
| | 04-02 | | | **S ×6 筆** | **64,949** | 251.25–256.00 | **3,280,418** | **有**（2024-05-24 訂立） |
| 2026-04-03 | 04-01 | Sabih Khan | COO | M | 64,317 | — | 1,107,212 | — |
| | 04-01 | | | F（扣稅） | 33,317 | 255.63 | **1,073,895** | 附註明載「**No shares were sold**」 |
| 2026-04-03 | 04-01 | Deirdre O'Brien | SVP | M / F | 64,317 / 34,315 | 255.63 | 166,812 | — |
| | 04-02 | | | S ×2 | 30,002 | 255.12／255.82 | 136,810 | — |
| 2026-04-17 | 04-15 | Kevan Parekh | CFO | M / F | 10,928 / 4,793 | 266.43 | 14,900 | — |
| 2026-04-27 | 04-23 | Kevan Parekh | CFO | S | 1,534 | 275.00 | **13,366** | **有**（2025-11-21 訂立） |
| 2026-04-17 | 04-15 | Ben Borders | 主計長 | M / F | 1,717 / 892 | 266.43 | 39,987 | — |
| **2026-05-08** | **05-06** | **Arthur D. Levinson** | **董事長** | **S ×2** | **250,000** | **284.57／285.04** | **3,819,576** | **無** |
| | 05-06 | | | G（贈與） | 5,000 | — | 3,814,576 | — |
| 2026-05-12 | 05-08 | Ben Borders | 主計長 | S | 1,274 | 290.00 | 38,713 | —（另見 10-Q Item 5） |
| **2026-05-29** | **05-27** | **Arthur D. Levinson** | **董事長** | **S** | **50,000** | **311.02** | **3,764,576** | **無** |
| | 05-27 | | | G（贈與） | 65,000 | — | **3,699,576** | — |
| 2026-06-17 | 06-15 | Jennifer Newstead | SVP、總法律顧問 | M / F | 30,104 / 16,238 | 296.42 | 41,546 | — |
| 2026-06-17 | 06-15/16 | Ben Borders | 主計長 | M / F / S | 240 / 124 / 116 | 296.42／295.14 | 38,713 | — |

**⚑ 三點觀察**

1. **Levinson（董事長）**：法說會後 6 天賣出 250,000 股（約 **US$71.2M**），再 27 天後賣出 50,000 股（約 **US$15.6M**），另贈與 70,000 股。持股由 4,069,576 降至 3,699,576，三週內 **−9.1%**。兩份 Form 4 **均未附任何 Rule 10b5-1 註記**；相對地，同期 Cook 與 Parekh 的賣出都明確標註依 10b5-1 計畫執行。Levinson 正是主導 CEO 接班決策的董事會主席。
2. **Cook**：4/1 是 Apple 制式 RSU 交割日，隨後的賣出依 **2024-05-24**（近兩年前）訂立的 10b5-1 計畫執行，資訊優勢極低；且**接班宣布後未見任何額外賣出**。
3. **Khan（COO）**：同一批 RSU 交割，扣稅後 **一股未賣**，是本批申報中唯一完全留倉的高階主管。
4. **未申報者**：候任 CEO **John Ternus 不在 Section 16 申報人名單中**，其持股與交易無從查核 —— 這是本季內部人訊號最大的空白。
5. **10-Q Item 5**：主計長 Ben Borders 於 **2026-02-06** 訂立 10b5-1 計畫（至多 898 股＋2026-04-15 至 12-15 之間既得股份，2026-12-31 到期）。

---

## 十一、資料限制

- Apple 自 FY2019 起**不再揭露任何產品的出貨量**，因此「量價拆分」無法由 filing 直接取得，本報告以營收、管理層口述的市佔／升級／新客紀錄與第三方（IDC、Worldpanel、451 Research）引述進行**區間推估**，並在對應表格標明為推估。
- Apple 的損益表**不揭露按業務別的營業利益**，僅有地區別營益（Note 10），因此「業務部門」tab 的獲利分析只能到 Products / Services 兩層毛利率。
- 供應商與客戶身分全部匿名，供應鏈圖之具名節點僅限逐字稿明示者。
