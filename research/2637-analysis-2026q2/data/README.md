# 慧洋-KY (2637) FY2026Q2 資料包 README

抓取日期：2026-09-20。標的確認為 TWSE 上市（sii），已用 TWSE openapi `t187ap03_L` 交叉驗證
（isin.twse.com.tw 單一來源不可信，見記憶 [[isin-twse-market-type-unreliable]]）。

## 檔案清單

| 檔案 | 內容 | 來源 | 抓取日期 |
|---|---|---|---|
| `quarterly.json` | 12季（2023Q3–2026Q2）單季損益表/資產負債表/現金流量表結構化數字 | MOPS `t164sb01` iXBRL，12支 raw HTML 在 `raw/fin_*.html` | 2026-09-20 |
| `financial_table.md` | 上述數字的可讀 Markdown 表格 | 同上（本地 build_table.py 產生）| 2026-09-20 |
| `financials_raw.json` | 12季原始 XBRL fact 抽取結果（parse_fin.py 中間產物）| 同上 | 2026-09-20 |
| `raw/q2_2026_fin.pdf` / `fin_2026Q2.txt` | 2026Q2合併財報全文PDF（122頁）與抽出文字 | doc.twse.com.tw mtype=A，申報檔`202602_2637_AI1.pdf` | 2026-09-20 |
| `raw/ar2025.pdf` / `ar2025.txt` | 114年報/2025股東會年報（82頁）與抽出文字 | doc.twse.com.tw mtype=F，申報檔`2025_2637_20260522F04.pdf` | 2026-09-20 |
| `fleet_2026q2.json` | 截至2026-06-30完整船隊清單（130艘：船名/建造年/DWT/船型）| 2026Q2財報附註十二「船隊資訊」| 2026-09-20 |
| `monthly_revenue.json` | 2024-06至2026-08共27個月月營收（含YoY%、累計、累計反算驗證）| MOPS `t21sc03` sii 月營收彙總表 | 2026-09-20 |
| `insider_and_announcements.md` | 2026/03–2026/08董監大股東持股變化 + 2025/06起重大訊息（含2026Q2起） | MOPS `ajax_stapap1`／`ajax_t05st01` | 2026-09-20 |
| `dividend_history.md` | 2021–2026發放年度現金股利，2025/2026兩筆已與年報交叉驗證 | wantgoo.com（單一來源，信心度中等）+ ar2025.txt（官方，信心度高）| 2026-09-20 |
| `market_cycle.md` | BDI零星時點、公司自身市場評述、合約結構（指數連結型長約占比） | WebSearch新聞彙整 + ar2025.txt + 富果直送法說會摘要 | 2026-09-20 |
| `qualitative_notes.md` | 管理層展望、合約結構細節、新造船時程、財務體質、主要客戶/供應商、資料口徑落差註記 | 財報+年報+第三方法說會摘要（富果直送）| 2026-09-20 |
| `price_weekly.json` | 2637.TW 近2年週K（106筆）| Yahoo chart API | 2026-09-20 |
| `peers_yahoo.json` | 8檔同業估值指標（P/E、P/B、營收成長等）| Yahoo quoteSummary（fc.yahoo.com cookie + crumb）| 2026-09-20 |
| `raw/monthly/`、`raw/insider_*.html`、`raw/ann_*.html`、`raw/peers/` | 上述各項的原始 HTML/JSON | 見上 | 2026-09-20 |

## 表達貨幣與單位

- **XBRL主表（quarterly.json / financial_table.md）：新台幣（TWD），原始scale為仟元，
  build_table.py已換算為百萬元顯示**。已用 `ix:unit id="TWD"` 與頁首「Unit: NT$ thousands」
  標籤確認（見 fin_2026Q2.txt 頁首）。
- **財報附註（fin_2026Q2.txt原文）大多以「美元」為主要單位、新台幣為對照**——慧洋為KY
  （開曼）控股、子公司多為美元功能性貨幣的航運SPV，附註原文明載「金額除另予註明者外，
  均以美元為單位」。**使用附註數字時務必留意欄位標題是「美元」還是「新台幣(千)」，
  兩者並列時金額差三個數量級，容易誤讀**。
- EPS 單位為新台幣元/股，未換算。

## 監理與交叉驗證重點

1. **市場別**：TWSE上市（sii），已用 openapi `t187ap03_L` 確認（非isin.twse.com.tw單一來源）。
2. **月營收累計反算驗證**：24個月中有部分（約20+個月）implied single-month與官方回報值有
   小額落差（多數 <NT$1千萬，占營收比重 <1%），屬MOPS常見的當月數字事後小幅修正/前次
   累計調整，非解析錯誤——已交叉核對cum欄與XBRL財報YTD營收完全一致（如2026Q2累計
   9,679,888千元 = XBRL的IS_YTD Revenue 9,679,888,000元，精確吻合）。落差明細見
   `monthly_revenue.json`中的`implied_single_month_k`欄。
3. **MOPS股利分派查詢`ajax_t05st09_2`持續遭安全性阻擋**（`mops.twse.com.tw`與
   `mopsov.twse.com.tw`兩網域皆是），2021–2024年股利數字退而求其次採第三方來源
   （wantgoo.com），信心度中等；2025、2026兩筆已用年報官方數字交叉驗證一致。
4. **重大訊息`ajax_t05st01`需用ROC年份（year=115/114）單參數查詢**，非日期區間參數
   （date1/date2嘗試失敗）；已成功取得2025/06起共59則公告。
5. **內部人持股`ajax_stapap1`需用ROC年（year=115）而非西元年**，回傳UTF-8編碼（注意：
   與t164sb01財報頁的cp950編碼不同，此為本次任務新確認的細節，先前記憶只記載
   ajax_t05st01/ajax_stapap1回UTF-8，這裡進一步確認：t21月營收頁與t164sb01財報頁仍是
   cp950，兩類端點編碼不同，抓取時務必依端點分別處理，不可一律假設cp950）。
6. **MOPS月營收檔名月份不補零**：`t21sc03_115_8.html`（單位數月份無前導零）是正確路徑，
   `t21sc03_115_08.html`回404——這是本次任務新發現的坑，先前6488記憶未記載此細節
   （6488抓取期間剛好月份多為10–12月雙位數，未觸發此問題）。

## 資料缺口（誠實列出，未捏造數字）

- **Baltic Exchange官方BDI/BSI/BHSI逐季平均值未取得**（需付費訂閱），market_cycle.md
  僅有零星時點與新聞定性描述，信心度中等偏低。
- **法說會簡報PDF原始檔未取得**（公司IR官網未深度爬梳），qualitative_notes.md中管理層
  展望與合約結構細節均轉引自第三方部落格（富果直送）的法說會摘要，非一手簡報，
  信心度中等。
- **公司自結逐季平均日租金（TCE）未在財報或年報中直接揭露**，無法提供結構化TCE數列。
- **2021–2024發放年度現金股利僅單一第三方來源（wantgoo.com）**，未與MOPS官方股利公告
  逐筆核對（官方端點遭安全性阻擋）。
- **2027/2028年新造船交船時程**在財報附註九（截至2026-06-30尚未交船餘額：2026年3艘+
  2027年4艘=7艘）與法說會摘要（2026年8艘、2027年4艘、2028年1艘）之間存在口徑落差，
  未能向公司查證，已在qualitative_notes.md中並列並加註⚠️提醒。
- **應收帳款/合約負債等資產負債表科目在2637（散裝航運）幾乎不存在**——與6488半導體
  範本不同，此公司無`ContractLiabilities`類科目（XBRL標籤搜尋確認），船舶運送業以
  「預收款項」（CurrentAdvances）取代，金額很小，非核心觀察指標。
- **供應商集中度低**（無單一供應商達10%進貨門檻），與部分製造業標的的供應鏈集中度
  分析框架不完全適用，qualitative_notes.md已據實記錄「無主要供應商10%以上名單」。

## 可重用管線筆記（供下次台股航運/KY控股公司報告參考）

- 本次沿用6488的`parse_fin.py`/`build_table.py`架構，但**IS/BS/CF標籤清單需針對航運業
  重新盤點**（無ContractLiabilities、多了GainsOnDisposalsOfPropertyPlantAndEquipment
  處分資產利益、ShareOfProfitLossOfAssociates權益法損益等航運/控股特有科目）。建議下次
  抓KY控股公司或航運業標的時，先跑一次「列出全部ix:nonFraction name」再挑欄位，
  不要直接照搬6488半導體的標籤清單。
- doc.twse全文PDF抓取流程（GET清單→POST拿中介頁→GET/pdf/連結）與6488記憶一致，
  本次無新增坑。
