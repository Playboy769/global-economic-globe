# 台灣證券交易所 (TWSE) OpenAPI 端點總覽

**Base URL**: `https://openapi.twse.com.tw/v1`  
**Swagger UI**: `https://openapi.twse.com.tw/`  
**格式**: 所有端點均支援 `application/json`，部分亦支援 `text/csv`

---

## 公司治理

| 端點 | 說明 |
|------|------|
| `GET /opendata/t187ap03_L` | 上市公司基本資料 |
| `GET /opendata/t187ap03_P` | 公開發行公司基本資料 |
| `GET /opendata/t187ap04_L` | 上市公司每日重大訊息 |
| `GET /opendata/t187ap02_L` | 上市公司持股逾 10% 大股東名單 |
| `GET /opendata/t187ap14_L` | 上市公司各產業 EPS 統計資訊 |
| `GET /opendata/t187ap45_L` | 上市公司股利分派情形 |
| `GET /company/newlisting` | 最近上市公司 |
| `GET /company/suspendListingCsvAndHtml` | 終止上市公司 |
| `GET /company/applylistingLocal` | 申請上市之本國公司 |
| `GET /company/applylistingForeign` | 外國公司向證交所申請第一上市 |
| `GET /announcement/punish` | 集中市場公布處置股票 |

### 董監事 / 持股

| 端點 | 說明 |
|------|------|
| `GET /opendata/t187ap08_L` | 上市公司董事、監察人持股不足法定成數彙總表 |
| `GET /opendata/t187ap10_L` | 上市公司董事、監察人持股不足法定成數連續達 3 個月以上彙總表 |
| `GET /opendata/t187ap11_L` | 上市公司董監事持股餘額明細資料 |
| `GET /opendata/t187ap11_P` | 公發公司董監事持股餘額明細 |
| `GET /opendata/t187ap09_L` | 上市公司董事、監察人質權設定占持有股數彙總表 |
| `GET /opendata/t187ap12_L` | 每日內部人持股轉讓事前申報表－持股轉讓日報表 |
| `GET /opendata/t187ap13_L` | 每日內部人持股轉讓事前申報表－持股未轉讓日報表 |
| `GET /opendata/t187ap29_A_L` | 上市公司董事酬金相關資訊 |
| `GET /opendata/t187ap29_B_L` | 上市公司監察人酬金相關資訊 |
| `GET /opendata/t187ap29_C_L` | 上市公司合併報表董事酬金相關資訊 |
| `GET /opendata/t187ap29_D_L` | 上市公司合併報表監察人酬金相關資訊 |
| `GET /opendata/t187ap30_L` | 上市公司獨立董監事兼任情形彙總表 |
| `GET /opendata/t187ap33_L` | 上市公司董事長是否兼任總經理 |
| `GET /opendata/t187ap34_L` | 上市公司採累積投票制、候選人提名制選任董監事彙總表 |

### 股東會

| 端點 | 說明 |
|------|------|
| `GET /opendata/t187ap38_L` | 上市公司股東會公告彙總表（95 年度起） |
| `GET /opendata/t187ap41_L` | 上市公司召開股東常(臨時)會日期、地點及電子投票彙總表 |
| `GET /opendata/t187ap35_L` | 上市公司股東行使提案權情形彙總表 |

### 經營權變動

| 端點 | 說明 |
|------|------|
| `GET /opendata/t187ap24_L` | 上市公司經營權異動公司 |
| `GET /opendata/t187ap25_L` | 上市公司營業範圍重大變更公司 |
| `GET /opendata/t187ap26_L` | 上市公司經營權異動且營業範圍重大變更停止買賣公司 |
| `GET /opendata/t187ap27_L` | 上市公司經營權異動且列為變更交易公司 |

### 裁罰 / 規範

| 端點 | 說明 |
|------|------|
| `GET /opendata/t187ap22_L` | 上市公司金管會證券期貨局裁罰案件 |
| `GET /opendata/t187ap23_L` | 上市公司違反資訊申報、重大訊息及說明記者會規定 |
| `GET /opendata/t187ap32_L` | 上市公司公司治理相關規程規則 |

### ESG 資訊揭露

| 端點 | 說明 |
|------|------|
| `GET /opendata/t187ap46_L_1` | ESG－溫室氣體排放 |
| `GET /opendata/t187ap46_L_2` | ESG－能源管理 |
| `GET /opendata/t187ap46_L_3` | ESG－水資源管理 |
| `GET /opendata/t187ap46_L_4` | ESG－廢棄物管理 |
| `GET /opendata/t187ap46_L_5` | ESG－人力發展 |
| `GET /opendata/t187ap46_L_6` | ESG－董事會 |
| `GET /opendata/t187ap46_L_7` | ESG－投資人溝通 |
| `GET /opendata/t187ap46_L_8` | ESG－氣候相關議題管理 |
| `GET /opendata/t187ap46_L_9` | ESG－功能性委員會 |
| `GET /opendata/t187ap46_L_10` | ESG－燃料管理 |
| `GET /opendata/t187ap46_L_11` | ESG－產品生命週期 |
| `GET /opendata/t187ap46_L_12` | ESG－食品安全 |
| `GET /opendata/t187ap46_L_13` | ESG－供應鏈管理 |
| `GET /opendata/t187ap46_L_14` | ESG－產品品質與安全 |
| `GET /opendata/t187ap46_L_15` | ESG－社區關係 |
| `GET /opendata/t187ap46_L_16` | ESG－資訊安全 |
| `GET /opendata/t187ap46_L_17` | ESG－普惠金融 |
| `GET /opendata/t187ap46_L_18` | ESG－持股及控制力 |
| `GET /opendata/t187ap46_L_19` | ESG－風險管理政策 |
| `GET /opendata/t187ap46_L_20` | ESG－反競爭行為法律訴訟 |
| `GET /opendata/t187ap46_L_21` | ESG－職業安全衛生 |

---

## 證券交易

### 個股行情

| 端點 | 說明 |
|------|------|
| `GET /exchangeReport/STOCK_DAY_ALL` | 上市個股日成交資訊（全部） |
| `GET /exchangeReport/STOCK_DAY_AVG_ALL` | 上市個股日收盤價及月平均價 |
| `GET /exchangeReport/FMSRFK_ALL` | 上市個股月成交資訊 |
| `GET /exchangeReport/FMNPTK_ALL` | 上市個股年成交資訊 |
| `GET /exchangeReport/BWIBBU_ALL` | 上市個股日本益比、殖利率及股價淨值比（依代碼） |
| `GET /exchangeReport/BWIBBU_d` | 上市個股日本益比、殖利率及股價淨值比（依日期） |
| `GET /exchangeReport/TWT84U` | 上市個股股價升降幅度 |
| `GET /exchangeReport/TWT88U` | 上市個股首五日無漲跌幅 |
| `GET /exchangeReport/TWT48U_ALL` | 上市股票除權除息預告表 |

> **RR5 常用**: `STOCK_DAY_ALL`（搭配 `?date=YYYYMMDD&stockNo=CODE` 查個股日成交，`title` 欄含股名）

### 大盤 / 市場

| 端點 | 說明 |
|------|------|
| `GET /exchangeReport/MI_INDEX` | 每日收盤行情－大盤統計資訊 |
| `GET /exchangeReport/MI_INDEX20` | 集中市場每日成交量前二十名證券 |
| `GET /exchangeReport/MI_INDEX4` | 每日上市上櫃跨市場成交資訊 |
| `GET /exchangeReport/FMTQIK` | 集中市場每日市場成交資訊 |
| `GET /exchangeReport/MI_5MINS` | 每 5 秒委託成交統計 |
| `GET /opendata/twtazu_od` | 集中市場漲跌證券數統計表 |

### 融資融券 / 借券

| 端點 | 說明 |
|------|------|
| `GET /exchangeReport/MI_MARGN` | 集中市場融資融券餘額 |
| `GET /SBL/TWT96U` | 上市上櫃股票當日可借券賣出股數 |
| `GET /exchangeReport/BFI84U` | 集中市場停資停券預告表 |

### 當日沖銷

| 端點 | 說明 |
|------|------|
| `GET /exchangeReport/TWTB4U` | 上市股票每日當日沖銷交易標的及統計 |
| `GET /exchangeReport/TWTBAU1` | 集中市場暫停先賣後買當日沖銷交易標的預告表 |
| `GET /exchangeReport/TWTBAU2` | 集中市場暫停先賣後買當日沖銷交易歷史查詢 |

### 外資

| 端點 | 說明 |
|------|------|
| `GET /fund/MI_QFIIS_cat` | 集中市場外資及陸資投資類股持股比率表 |
| `GET /fund/MI_QFIIS_sort_20` | 集中市場外資及陸資持股前 20 名彙總表 |

### 其他交易

| 端點 | 說明 |
|------|------|
| `GET /exchangeReport/TWT53U` | 集中市場零股交易行情單 |
| `GET /exchangeReport/TWTAWU` | 集中市場暫停交易證券 |
| `GET /exchangeReport/BFT41U` | 集中市場盤後定價交易 |
| `GET /exchangeReport/TWT85U` | 集中市場證券變更交易 |
| `GET /exchangeReport/STOCK_FIRST` | 每日第一上市外國股票成交量值 |
| `GET /block/BFIAUU_d` | 集中市場鉅額交易日成交量值統計 |
| `GET /block/BFIAUU_m` | 集中市場鉅額交易月成交量值統計 |
| `GET /block/BFIAUU_y` | 集中市場鉅額交易年成交量值統計 |
| `GET /holidaySchedule/holidaySchedule` | 有價證券集中交易市場開（休）市日期 |
| `GET /Announcement/BFZFZU_T` | 投資理財節目異常推介個股 |
| `GET /announcement/notice` | 集中市場當日公布注意股票 |
| `GET /announcement/notetrans` | 集中市場公布注意累計次數異常資訊 |
| `GET /opendata/t187ap19` | 電子式交易統計資訊 |

---

## 財務報表

### 上市公司

| 端點 | 說明 |
|------|------|
| `GET /opendata/t187ap05_L` | 上市公司每月營業收入彙總表 |
| `GET /opendata/t187ap05_P` | 公開發行公司每月營業收入彙總表 |
| `GET /opendata/t187ap15_L` | 上市公司截至各季綜合損益財測達成情形（簡式） |
| `GET /opendata/t187ap16_L` | 上市公司當季損益經會計師查核與預測差異達 10% 以上（簡式） |
| `GET /opendata/t187ap17_L` | 上市公司營益分析查詢彙總表（全體公司） |
| `GET /opendata/t187ap31_L` | 上市公司財務報告經監察人承認情形 |

### 綜合損益表

| 端點 | 說明 |
|------|------|
| `GET /opendata/t187ap06_L_basi` | 上市公司綜合損益表（金融業） |
| `GET /opendata/t187ap06_L_bd` | 上市公司綜合損益表（證券期貨業） |
| `GET /opendata/t187ap06_L_ci` | 上市公司綜合損益表（一般業） |
| `GET /opendata/t187ap06_L_fh` | 上市公司綜合損益表（金控業） |
| `GET /opendata/t187ap06_L_ins` | 上市公司綜合損益表（保險業） |
| `GET /opendata/t187ap06_L_mim` | 上市公司綜合損益表（異業） |
| `GET /opendata/t187ap06_X_basi` | 公發公司綜合損益表（金融業） |
| `GET /opendata/t187ap06_X_bd` | 公發公司綜合損益表（證券期貨業） |
| `GET /opendata/t187ap06_X_ci` | 公發公司綜合損益表（一般業） |
| `GET /opendata/t187ap06_X_fh` | 公發公司綜合損益表（金控業） |
| `GET /opendata/t187ap06_X_ins` | 公發公司綜合損益表（保險業） |
| `GET /opendata/t187ap06_X_mim` | 公發公司綜合損益表（異業） |

### 資產負債表

| 端點 | 說明 |
|------|------|
| `GET /opendata/t187ap07_L_basi` | 上市公司資產負債表（金融業） |
| `GET /opendata/t187ap07_L_bd` | 上市公司資產負債表（證券期貨業） |
| `GET /opendata/t187ap07_L_ci` | 上市公司資產負債表（一般業） |
| `GET /opendata/t187ap07_L_fh` | 上市公司資產負債表（金控業） |
| `GET /opendata/t187ap07_L_ins` | 上市公司資產負債表（保險業） |
| `GET /opendata/t187ap07_L_mim` | 上市公司資產負債表（異業） |
| `GET /opendata/t187ap07_X_basi` | 公發公司資產負債表（金融業） |
| `GET /opendata/t187ap07_X_bd` | 公發公司資產負債表（證券期貨業） |
| `GET /opendata/t187ap07_X_ci` | 公發公司資產負債表（一般業） |
| `GET /opendata/t187ap07_X_fh` | 公發公司資產負債表（金控業） |
| `GET /opendata/t187ap07_X_ins` | 公發公司資產負債表（保險業） |
| `GET /opendata/t187ap07_X_mim` | 公發公司資產負債表（異業） |

---

## 指數

| 端點 | 說明 |
|------|------|
| `GET /indicesReport/MI_5MINS_HIST` | 發行量加權股價指數歷史資料 |
| `GET /indicesReport/MFI94U` | 發行量加權股價報酬指數 |
| `GET /indicesReport/TAI50I` | 臺灣 50 指數歷史資料 |
| `GET /indicesReport/FRMSA` | 寶島股價指數歷史資料 |

---

## 權證

| 端點 | 說明 |
|------|------|
| `GET /opendata/t187ap37_L` | 上市權證基本資料彙總表 |
| `GET /opendata/t187ap36_L` | 上市認購(售)權證年度發行量概況統計表 |
| `GET /opendata/t187ap42_L` | 上市認購(售)權證每日成交資料檔 |
| `GET /opendata/t187ap43_L` | 上市認購(售)權證交易人數檔 |

> **RR5 常用**:
> - `t187ap37_L` — 上市權證基本資料，含代號、名稱、標的、到期日（查權證名稱最快）
> - `t187ap42_L` — 每日成交資料，可補強收盤價來源

---

## 券商資料

| 端點 | 說明 |
|------|------|
| `GET /brokerService/brokerList` | 證券商總公司基本資料 |
| `GET /brokerService/secRegData` | 開辦定期定額業務證券商名單 |
| `GET /ETFReport/ETFRank` | 定期定額交易戶數統計排行月報表 |
| `GET /opendata/t187ap18` | 證券商基本資料 |
| `GET /opendata/t187ap01` | 券商業務別人員數 |
| `GET /opendata/t187ap20` | 各券商每月月計表 |
| `GET /opendata/t187ap21` | 各券商收支概況表資料 |
| `GET /opendata/OpenData_BRK01` | 證券商營業員男女人數統計資料 |
| `GET /opendata/OpenData_BRK02` | 證券商分公司基本資料 |

---

## 其他

| 端點 | 說明 |
|------|------|
| `GET /opendata/t187ap47_L` | 基金基本資料彙總表 |
| `GET /exchangeReport/BFI61U` | 中央登錄公債補息資料表 |
| `GET /news/newsList` | 證交所新聞 |
| `GET /news/eventList` | 證交所活動訊息 |

---

## RR5 專用速查

| 用途 | 端點 | 關鍵欄位 |
|------|------|----------|
| 查個股日收盤 + 股名 | `GET /rwd/zh/afterTrading/STOCK_DAY?response=json&date=YYYYMMDD&stockNo=CODE` | `title`（含股名）、`data[][6]`（收盤價） |
| 上市權證基本資料（名稱/到期日/標的） | `GET /opendata/t187ap37_L` | `SecuritiesCode`, `Name`, `UnderlyingSecurities` |
| 權證每日成交 | `GET /opendata/t187ap42_L` | `SecuritiesCode`, `ClosingPrice` |
| 除權除息預告 | `GET /exchangeReport/TWT48U_ALL` | — |
| 本益比 / 殖利率 / 淨值比 | `GET /exchangeReport/BWIBBU_ALL` | `PriceEarningRatio`, `DividendYield`, `PriceBookRatio` |
