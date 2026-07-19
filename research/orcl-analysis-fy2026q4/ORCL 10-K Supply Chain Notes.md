# Oracle FY2026 10-K — Supply Chain & Insider Trading Notes

來源：SEC EDGAR，CIK 0001341439。10-K（accession 0001193125-26-277521，filed 2026-06-22，period 2026-05-31）；Q4 FY26 8-K/press release（accession 0001193125-26-265848，filed 2026-06-10）；Form 4 filings（2026-06-02、2026-06-26 批次）。

## 供應鏈（10-K Item 1 "Manufacturing"、Item 1A Risk Factors）

- **代工模式**："We rely on third-party manufacturing partners to produce most of our hardware products
  that we market and sell to customers and utilize internally to deliver Oracle Cloud offerings, and we
  distribute most of our hardware products from these partners' facilities."（Item 1）——Oracle 自身不生產硬體，OCI 資料中心用的伺服器/機櫃也是委外代工組裝。
- **零組件來源**："For most of our hardware products, we have existing alternative sources of supply or
  such sources are readily available. However, we do rely on sole and/or single sources for certain
  hardware components."（Item 1）——未點名具體零組件或供應商。
- **GPU/記憶體供給是主要風險因子**（Item 1A）："industry supply capacity for AI accelerators, including
  graphics processing units, as well as memory devices, is competitive, and we at times have to accept
  less favorable terms with suppliers to minimize supply constraints. While this permits us to secure
  cloud infrastructure capacity, it has increased excess and obsolescence risk of such hardware..."
  ⚑ 這句是本份 10-K 少數承認「為了搶產能接受更差供應商條款」的段落，與法說會 Clay Magouyrk
  强調「we have a very robust set of mechanisms to ensure Oracle is not sitting there with reduced
  margins」的樂觀敘事形成對照。
- **資料中心擴張三大瓶頸**（Item 1A）：可建置場址、穩定電力、伺服器/GPU/記憶體供應——三者皆被列為
  風險因子；"we have faced, and may continue to face, challenges with securing reliable and
  cost-effective power sources for our data center energy demands, which are constrained globally due
  to the significant increase in demand for and limited availability of energy to power AI compute."
- **與第三方資料中心業者的長約承諾**："we have entered, and expect to continue to enter, into long-term
  lease commitments with third-party data center providers and other significant commitments with
  suppliers of chips and other data center infrastructure."
- **未點名供應商**：全篇 10-K 未出現「NVIDIA」「AMD」「TSMC」等具體供應商名稱，也未揭露任何客戶集中度
  （無「10% of revenue」等字樣）——與法說會 Q&A 中管理層拒絕拆解「哪些客戶屬於資料中心」的留白手法
  一致，此為系統性的資訊管理，而非單一場合迴避。
- **人力**：截至 2026/5/31 全職員工約 141,000 人（美國 49,000／海外 92,000）；按業務線：雲端與軟體
  26,000、服務 34,000、業務與行銷 25,000、研發 43,000、硬體 2,000、一般與行政 11,000。研發占比
  30.5%，反映軟體/雲端公司屬性。

## RPO（Remaining Performance Obligations）10-K 揭露

$638B（2026/5/31）vs $138B（2025/5/31），10-K MD&A 僅說明「primarily attributable to certain
significant cloud contracts entered into during the period」，未進一步拆解客戶別或到期結構（比法說會
揭露的「12%於12個月內確認、34%於13–36個月」更簡略）。

## Form 4 內部人交易（法說會前後兩週）

### 2026-06-02（法說會前 8 天）—— 全數為董事會例行年度配股，非公開市場買賣
8 位董事同日申報，全部為 Code M（履約既有 RSU/選擇權，非市場買賣）：Awo Ablo、Jeffrey Berg、
Michael J. Boskin、Bruce R. Chizen、Rona Alison Fairhead、Tomislav Mihaljevic、Charles W. Moorman、
Stephen H. Rusckowski。多數為「2,114股 @ Code M，交易日 05/31/2026」（會計年度末當日既定歸屬），
Mihaljevic 與 Rusckowski（新任董事）另各獲配 1,550股 $0成本RSU。⚑ 此批屬每年固定週期性配股，
非管理層對本季業績的主動判斷訊號。

### 2026-06-26（法說會後 16 天）—— Jeffrey Henley（董事會執行副主席，前CFO/董事長）
交易日 2026/6/24：
- Code M：履約深度價內選擇權，以 $40.93 履約價取得 400,000 股（此履約價顯示為多年前授予的舊選擇權）
- Code S：同日分 8 筆以上市價出售，成交價區間 $156.06–$162.996（加權均價約 $158.5），已確認出售
  股數合計 359,068 股（尚有餘額 40,932 股持續持有），總價金約 $5,700萬
- ⚑ 觀察：此交易發生於法說會（6/10）後兩週，且股價已從 6/24 反映法說會利多大漲至 $156–163
  區間（相對 Henley 履約成本 $40.93 有超過 280% 帳面利得）。時點落在盈餘公布後的正常交易窗口，
  性質上更像深度價內舊選擇權的例行變現/分散化，而非對本季業績的負面訊號——但股價區間本身
  （$156–163）是法說會後市場實際反應的獨立佐證，可與分析師目標價交叉比對。

## 財報數字來源
詳見同資料夾 `ORCL Q4 FY26 Key Financial Data.md`（源自 10-K 合併財報與 8-K/EX-99.1 新聞稿之
GAAP/non-GAAP 調節表、分季揭露、trailing four-quarter free cash flow 表）。
