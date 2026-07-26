# 證券櫃檯買賣中心 (TPEx) OpenAPI 端點總覽

**Base URL**: `https://www.tpex.org.tw/openapi/v1`  
**Swagger UI**: `https://www.tpex.org.tw/openapi/`  
**Spec JSON**: `https://www.tpex.org.tw/openapi/swagger.json`  
**版本**: 1.0.0 / OAS3  
**資料範疇**: 上櫃、興櫃、創櫃及債券等發行及交易資訊

---

## 上櫃 (OTC Main Board)

### 行情 / 報價

| 端點 | 說明 |
|------|------|
| `GET /tpex_mainborad_highlight` | 上櫃股票市場現況 |
| `GET /tpex_mainboard_daily_close_quotes` | 上櫃股票行情 |
| `GET /tpex_mainboard_quotes` | 上櫃股票收盤行情 |
| `GET /tpex_mainboard_peratio_analysis` | 上櫃個股本益比、殖利率、股價淨值比 |
| `GET /tpex_active_dollar_volume` | 上櫃盤中個股成交金額排行 |
| `GET /tpex_active_advanced` | 上櫃盤中個股漲幅排行 |
| `GET /tpex_active_declined` | 上櫃盤中個股跌幅排行 |
| `GET /tpex_ceil_non_trading` | 上櫃漲跌停未成交資訊 |
| `GET /tpex_prvol` | 上櫃股票等價系統成交分價表 |

### 歷史排行

| 端點 | 說明 |
|------|------|
| `GET /tpex_daily_market_value` | 上櫃歷史個股市值排行 |
| `GET /tpex_daily_turnover` | 上櫃歷史個股週轉率排行 |
| `GET /tpex_trading_volumes_avg` | 上櫃歷史個股日均量排行 |
| `GET /tpex_trading_amount_avg` | 上櫃歷史個股日均值排行 |
| `GET /tpex_volume_rank` | 上櫃歷史個股成交量排行 |
| `GET /tpex_amount_rank` | 上櫃歷史個股成交值排行 |
| `GET /tpex_pe_ratio_top10` | 上櫃歷史個股本益比排行 |
| `GET /tpex_trading_volume_ratio` | 上櫃歷史類股成交價量比重 |

### 日成交統計

| 端點 | 說明 |
|------|------|
| `GET /tpex_daily_trading_index` | 上櫃日成交量值指數 |
| `GET /tpex_active_broker_volume` | 上櫃股票熱門股證券商進出排行 |
| `GET /tpex_daily_broker1` | 上櫃各券商當日營業金額統計表 |

### 融資融券

| 端點 | 說明 |
|------|------|
| `GET /tpex_mainboard_margin_balance` | 上櫃股票融資融券餘額 |
| `GET /tpex_margin_sbl` | 上櫃股票融券借券賣出餘額 |
| `GET /tpex_margin_trading_term` | 上櫃融資融券暫停融券賣出預告表 |
| `GET /tpex_margin_trading_adjust` | 上櫃融資融券調整成數 |
| `GET /tpex_margin_trading_lend` | 上櫃融資融券標借 |
| `GET /tpex_margin_trading_marginspot` | 上櫃信用交易餘額概況表 |
| `GET /tpex_margin_trading_margin_mark` | 上櫃平盤下得融(借)券賣出之證券名單 |
| `GET /tpex_margin_trading_margin_used` | 上櫃融資融券使用率報表 |
| `GET /tpex_margin_trading_short_sell` | 上櫃融資融券增減排行表 |
| `GET /tpex_short_sell` | 上櫃當日融券賣出與借券賣出成交量值 |

### 當日沖銷

| 端點 | 說明 |
|------|------|
| `GET /tpex_securities` | 上櫃股票現股當沖交易標的資訊 |
| `GET /tpex_intraday_trading_statistics` | 上櫃股票現股當沖交易統計資訊 |
| `GET /tpex_intraday_fee` | 上櫃應付現股當日沖銷券差借券費率 |
| `GET /tpex_intraday_trading_pre` | 上櫃暫停先賣後買當日沖銷交易標的預告表 |
| `GET /tpex_intraday_trading_his` | 上櫃暫停先賣後買當日沖銷交易歷史查詢 |

### 三大法人 / 外資

| 端點 | 說明 |
|------|------|
| `GET /tpex_3insti_qfii` | 上櫃僑外資及陸資持股比例排行表 |
| `GET /tpex_3insti_qfii_industry` | 上櫃各類股僑外資及陸資持股比例表 |
| `GET /tpex_3insti_qfii_trading` | 上櫃股票外資及陸資買賣超彙總表 |
| `GET /tpex_3insti_daily_trading` | 上櫃股票三大法人買賣明細資訊 |
| `GET /tpex_3insti_dealer_trading` | 上櫃股票自營商買賣超彙總表 |
| `GET /tpex_3insti_summary` | 上櫃股票三大法人買賣金額彙總表 |
| `GET /tpex_3insti_trading` | 上櫃股票投信買賣超彙總表 |

### 鉅額交易

| 端點 | 說明 |
|------|------|
| `GET /tpex_daily_qutoes_block` | 上櫃鉅額交易日成交資訊 |
| `GET /tpex_daily_trading_block` | 上櫃個股單一證券鉅額交易日成交資訊 |
| `GET /tpex_daily_trading_summary_odd` | 上櫃鉅額交易日成交量值統計 |
| `GET /tpex_monthly_trading_summary_block` | 上櫃鉅額交易月成交量值統計 |
| `GET /tpex_yearly_trading_summary_block` | 上櫃鉅額交易年成交量值統計 |
| `GET /tpex_daily_trade_block_day` | 鉅額交易歷史成交資訊 |

### 其他交易資訊

| 端點 | 說明 |
|------|------|
| `GET /tpex_exright_daily` | 上櫃股票除權除息計算結果表 |
| `GET /tpex_exright_prepost` | 上櫃股票除權除息預告表 |
| `GET /tpex_cmode` | 上櫃股票變更交易、分盤交易、管理股票與停止交易資訊 |
| `GET /tpex_odd_stock` | 上櫃股票零股交易資訊 |
| `GET /tpex_off_market` | 上櫃股票盤後定價行情 |
| `GET /tpex_spendi_today` | 上櫃當日公布暫停/恢復交易股票 |
| `GET /tpex_spendi_history` | 上櫃歷史公布暫停/恢復交易股票 |
| `GET /tpex_delayed_stock_open` | 上櫃每日暫緩開盤股票 |
| `GET /tpex_delayed_stock_close` | 上櫃每日暫緩收盤股票 |
| `GET /tpex_ipo_no_limit` | 上櫃首五日無漲跌幅資訊 |
| `GET /tpex_esb_applicant_companies` | 申請上櫃公司 |
| `GET /tpex_trading_warning_information` | 上櫃公布注意股票資訊 |
| `GET /tpex_disposal_information` | 上櫃處置有價證券資訊 |
| `GET /tpex_trading_warning_note` | 上櫃公布注意累計次數異常資訊 |
| `GET /mopsfin_t187ap19_O` | 電子式交易統計資訊（上櫃） |

---

## 指數系列

| 端點 | 說明 |
|------|------|
| `GET /tpex_index` | 櫃買指數歷史資料 |
| `GET /tpex_index_consti` | 櫃買指數成分股 |
| `GET /tpex_reward_index` | 櫃買指數與報酬指數收市指數 |
| `GET /tpex50_index` | 富櫃 50 指數歷史收盤指數 |
| `GET /tpex50_constituents` | 富櫃 50 指數當日成分股 |
| `GET /tpex50_change` | 富櫃 50 指數當日收盤指數 |
| `GET /tpex200_constituents` | 富櫃 200 指數當日成分股 |
| `GET /tpex200_change` | 富櫃 200 指數當日收盤指數 |
| `GET /tpcgi_constituents` | 上櫃公司治理指數當日成分股 |
| `GET /tpcgi_change` | 上櫃公司治理指數當日收盤指數 |
| `GET /tpcgi_reward_index` | 上櫃公司治理指數歷史收盤指數 |
| `GET /tphd_constituents` | 高殖利率指數當日成分股 |
| `GET /tphd_change` | 高殖利率指數當日收盤指數 |
| `GET /tphd_index` | 高殖利率指數歷史收盤指數 |
| `GET /tpci_constituents` | 薪酬指數當日成分股 |
| `GET /tpci_change` | 薪酬指數當日收盤指數 |
| `GET /tpci_reward_index` | 薪酬指數歷史收盤指數 |
| `GET /tpex_emp88_constituents` | 勞工就業 88 指數當日成分股 |
| `GET /tpex_emp88_change` | 勞工就業 88 指數當日收盤指數 |
| `GET /tpex_emp88_reward_index` | 勞工就業 88 指數歷史收盤指數 |

---

## 公司治理

### 基本資料

| 端點 | 說明 |
|------|------|
| `GET /mopsfin_t187ap03_O` | 上櫃股票基本資料 |
| `GET /mopsfin_t187ap03_R` | 興櫃公司基本資料 |
| `GET /mopsfin_t187ap04_O` | 上櫃公司每日重大訊息 |
| `GET /mopsfin_t187ap02_O` | 上櫃公司持股逾 10% 大股東名單 |
| `GET /mopsfin_t187ap05_O` | 上櫃公司每月營業收入彙總表 |
| `GET /t187ap05_R` | 興櫃公司每月營業收入彙總表 |
| `GET /mopsfin_t187ap05_OA` | 二十九大類股營收變化統計表 |
| `GET /mopsfin_t187ap05_OB` | 發行公司營收創新高一覽表（上櫃） |
| `GET /mopsfin_t187ap14_O` | 上櫃公司各產業 EPS 統計資訊 |
| `GET /tpex_esb_eps_rank` | 本國興櫃公司 EPS 排名 |
| `GET /tpex_esb_capitals_rank` | 興櫃公司資本額排名 |
| `GET /mopsfin_t187ap39_O` | 上櫃股利分派情形－董事會通過 |

### 董監事 / 持股

| 端點 | 說明 |
|------|------|
| `GET /mopsfin_t187ap08_O` | 上櫃公司董事、監察人持股不足法定成數彙總表 |
| `GET /mopsfin_t187ap10_O` | 上櫃公司董事、監察人持股不足法定成數連續達 3 個月以上彙總表 |
| `GET /mopsfin_t187ap11_O` | 上櫃公司董監事持股餘額明細資料 |
| `GET /mopsfin_t187ap11_R` | 興櫃公司董監事持股餘額明細資料 |
| `GET /mopsfin_t187ap09_O` | 上櫃公司董事、監察人質權設定占持有股數彙總表 |
| `GET /mopsfin_t187ap12_O` | 上櫃公司每日內部人持股轉讓事前申報表－持股轉讓日報表 |
| `GET /mopsfin_t187ap13_O` | 上櫃公司每日內部人持股轉讓事前申報表－持股未轉讓日報表 |
| `GET /mopsfin_t187ap29_A_O` | 上櫃公司董事酬金相關資訊 |
| `GET /mopsfin_t187ap29_B_O` | 上櫃公司監察人酬金相關資訊 |
| `GET /mopsfin_t187ap29_C_O` | 上櫃公司合併報表董事酬金相關資訊 |
| `GET /mopsfin_t187ap29_D_O` | 上櫃公司合併報表監察人酬金相關資訊 |
| `GET /mopsfin_t187ap30_O` | 上櫃公司獨立董監事兼任情形彙總表 |
| `GET /mopsfin_t187ap33_O` | 上櫃公司董事長是否兼任總經理 |
| `GET /mopsfin_t187ap34_O` | 上櫃公司採累積投票制、候選人提名制選任董監事彙總表 |

### 股東會

| 端點 | 說明 |
|------|------|
| `GET /t187ap41_O` | 上櫃公司召開股東常(臨時)會日期、地點及電子投票彙總表 |
| `GET /mopsfin_t187ap35_O` | 上櫃公司股東行使提案權情形彙總表 |

### 經營權變動 / 裁罰

| 端點 | 說明 |
|------|------|
| `GET /mopsfin_t187ap24_O` | 上櫃公司經營權異動公司 |
| `GET /mopsfin_t187ap25_O` | 上櫃公司營業範圍重大變更公司 |
| `GET /mopsfin_t187ap26_O` | 上櫃公司經營權異動且營業範圍重大變更停止買賣公司 |
| `GET /mopsfin_t187ap27_O` | 上櫃公司經營權異動且列為變更交易公司 |
| `GET /mopsfin_t187ap22_O` | 上櫃公司金管會證券期貨局裁罰案件 |
| `GET /mopsfin_t187ap23_O` | 上櫃公司違反資訊申報、重大訊息及說明記者會規定 |
| `GET /mopsfin_t187ap32_O` | 上櫃公司公司治理相關規程規則 |
| `GET /mopsfin_t187ap31_O` | 上櫃公司財務報告經監察人承認情形 |

### ESG 資訊揭露（上櫃）

| 端點 | 說明 |
|------|------|
| `GET /t187ap46_O_1` | ESG－溫室氣體排放 |
| `GET /t187ap46_O_2` | ESG－能源管理 |
| `GET /t187ap46_O_3` | ESG－水資源管理 |
| `GET /t187ap46_O_4` | ESG－廢棄物管理 |
| `GET /t187ap46_O_5` | ESG－人力發展 |
| `GET /t187ap46_O_6` | ESG－董事會 |
| `GET /t187ap46_O_7` | ESG－投資人溝通 |
| `GET /t187ap46_O_8` | ESG－氣候相關議題管理 |
| `GET /t187ap46_O_9` | ESG－功能性委員會 |
| `GET /t187ap46_O_12` | ESG－食品安全 |
| `GET /t187ap46_O_13` | ESG－供應鏈管理 |
| `GET /t187ap46_O_14` | ESG－產品品質與安全 |
| `GET /t187ap46_O_15` | ESG－社區關係 |
| `GET /t187ap46_O_19` | ESG－風險管理政策 |
| `GET /t187ap46_O_20` | ESG－反競爭行為法律訴訟 |
| `GET /t187ap46_O_21` | ESG－職業安全衛生 |

---

## 權證

| 端點 | 說明 |
|------|------|
| `GET /tpex_warrant` | 上櫃股票權證資訊（含 Code、Name、標的、履約價、行使比例、到期日） |
| `GET /tpex_warrant_issue` | 上櫃權證發行基本資料（含 Code、Name、上櫃日期、到期日、標的） |
| `GET /tpex_warrant_quts` | 單筆權證成交資料 |
| `GET /tpex_warrant_daily_quts` | 上櫃權證收盤行情日報表 |
| `GET /tpex_warrant_monthly_quts` | 上櫃權證收盤行情月報表 |
| `GET /tpex_warrant_statistics` | 上櫃權證統計資訊 |
| `GET /tpex_warrant_suspend_today` | 上櫃今日暫停交易權證 |
| `GET /tpex_warrant_suspend_history` | 上櫃歷史暫停交易權證 |
| `GET /tpex_warrant_gold` | 上櫃黃金現貨權證資訊 |
| `GET /tpex_warrant_gold_quts` | 上櫃黃金現貨權證成交資料 |
| `GET /tpex_warrant_wcb_daily_quts` | 上櫃牛熊證收盤行情日報表 |
| `GET /tpex_warrant_wcb_monthly_quts` | 上櫃牛熊證收盤行情月報表 |
| `GET /tpex_warrant_wcb_issue` | 上櫃牛熊證發行基本資料 |
| `GET /tpex_warrant_wxy_daily_quts` | 上櫃外匯權證收盤行情日報表 |
| `GET /tpex_warrant_wxy_monthly_quts` | 上櫃外匯權證收盤行情月報表 |
| `GET /tpex_warrant_wxy_issue` | 上櫃外匯權證發行基本資料 |

> **RR5 常用**:
> - `/tpex_warrant` — 查上櫃權證名稱最快（回傳全部；搜尋 `"Code":"XXXXXX"` 取 `Name`）
> - `/tpex_warrant_issue` — 同上，含上市日期和到期日，適合補足 RR5 欄位

---

## 債券

### 國際債券 / 寶島債

| 端點 | 說明 |
|------|------|
| `GET /tpex_international_bond_quotes` | 國際債券當日盤中報價行情表（含寶島債） |
| `GET /tpex_international_bond_trade` | 國際債券當日盤中成交行情表（含寶島債） |
| `GET /tpex_international_bond_issue_investor` | 國際債券（一般投資人） |
| `GET /tpex_international_bond_issue_org` | 國際債券（僅售予專業投資人者） |

### 公司債 / 可轉債

| 端點 | 說明 |
|------|------|
| `GET /tpex_dpsp_monthly_CBmcs007` | 可轉債資產交換 ASO 及 ASW 銀行承作餘額 |
| `GET /bond_cb_daily` | 轉(交)換公司債買賣斷券商買賣日報表 |

### 債券發行資料

| 端點 | 說明 |
|------|------|
| `GET /bond_ISSBD1_data` | 公債發行資料 |
| `GET /bond_ISSBD2_data` | 外國金融債發行資料 |
| `GET /bond_ISSBD3_data` | 金融債發行資料 |
| `GET /bond_ISSBD4_data` | 普通債發行資料 |
| `GET /bond_ISSBD5_data` | 轉(交)換債發行資料 |
| `GET /bond_ISSBD6_data` | 海外轉換債發行資料 |
| `GET /bond_ISSBD7_data` | 附認股權公司債發行資料 |
| `GET /bond_ISSBD8_data` | 海外附認股權公司債發行資料 |
| `GET /bond_ISSBD9_data` | 海外普通債發行資料 |
| `GET /bond_ISSBD10_data` | 國際債券（寶島債）－本國發行人及第一、二上市(櫃)外國發行人資料 |
| `GET /bond_ISSBD11_data` | 國際債券（寶島債）－外國發行人（未公開發行股權商品者）資料 |

### 理論價格

| 端點 | 說明 |
|------|------|
| `GET /BDdos209UTF` | 美元零息可贖回國際債券理論價格 |
| `GET /BDdos215UTF` | 美元附息固定利率可贖回國際債券理論價格 |
| `GET /BDdos216UTF` | 美元固定利率不可贖回國際債券理論價格 |

---

## 興櫃 (Emerging Board)

| 端點 | 說明 |
|------|------|
| `GET /tpex_esb_highlight` | 興櫃股票市場現況 |
| `GET /tpex_esb_latest_statistics` | 興櫃股票當日行情表 |
| `GET /tpex_esb_recommended_dealer` | 興櫃推薦證券商與推薦之股票 |
| `GET /tpex_esb_warning_information` | 興櫃公布注意有價證券資訊 |
| `GET /tpex_esb_disposal_information` | 興櫃處置有價證券資訊 |

---

## 創櫃 (GISA)

| 端點 | 說明 |
|------|------|
| `GET /tpex_gisa_highlight` | 創櫃板公司市場現況 |
| `GET /tpex_gisa_company` | 創櫃板公司資訊 |
| `GET /tpex_gisa_financing_before` | 於登錄創櫃板前辦理籌資資訊 |
| `GET /tpex_gisa_financing_history` | 創櫃板公司透過籌資系統辦理籌資資訊 |
| `GET /tpex_gisa_financing_in_process` | 創櫃板辦理中籌資資訊 |

---

## 開放式基金

| 端點 | 說明 |
|------|------|
| `GET /tpex_opfund_latest` | 開放式基金當日行情表 |
| `GET /tpex_opfund_recommended_dealer` | 開放式基金受益憑證造市商與造市之基金 |
| `GET /tpex_opfund_market_highlight` | 開放式基金市場現況 |

---

## 黃金現貨

| 端點 | 說明 |
|------|------|
| `GET /tpex_gold_market_highlight` | 黃金現貨市場現況 |
| `GET /tpex_gold_latest` | 黃金現貨當日行情表 |
| `GET /tpex_gold_recommended_dealer` | 造市商與造市之黃金現貨 |

---

## 財務報表

### 上櫃公司

#### 綜合損益表

| 端點 | 說明 |
|------|------|
| `GET /mopsfin_t187ap06_O_basi` | 上櫃公司綜合損益表（金融業） |
| `GET /mopsfin_t187ap06_O_bd` | 上櫃公司綜合損益表（證券期貨業） |
| `GET /mopsfin_t187ap06_O_ci` | 上櫃公司綜合損益表（一般業） |
| `GET /mopsfin_t187ap06_O_fh` | 上櫃公司綜合損益表（金控業） |
| `GET /mopsfin_t187ap06_O_ins` | 上櫃公司綜合損益表（保險業） |
| `GET /mopsfin_t187ap06_O_mim` | 上櫃公司綜合損益表（異業） |
| `GET /mopsfin_t187ap06_O_basiA` | 上櫃公司財報資訊（金融業） |
| `GET /mopsfin_t187ap06_O_bdA` | 上櫃公司財報資訊（證券期貨業） |
| `GET /mopsfin_t187ap06_O_ciA` | 上櫃公司財報資訊（一般業） |
| `GET /mopsfin_t187ap06_O_fhA` | 上櫃公司財報資訊（金控業） |
| `GET /mopsfin_t187ap06_O_insA` | 上櫃公司財報資訊（保險業） |
| `GET /mopsfin_t187ap06_O_mimA` | 上櫃公司財報資訊（異業） |
| `GET /mopsfin_t187ap15_O` | 上櫃公司截至各季綜合損益財測達成情形（簡式） |
| `GET /mopsfin_t187ap16_O` | 上櫃公司當季損益與預測差異達 10% 以上（簡式） |
| `GET /mopsfin_187ap17_O` | 上櫃公司營益分析查詢彙總表（全體公司） |

#### 資產負債表

| 端點 | 說明 |
|------|------|
| `GET /mopsfin_t187ap07_O_basi` | 上櫃公司資產負債表（金融業） |
| `GET /mopsfin_t187ap07_O_bd` | 上櫃公司資產負債表（證券期貨業） |
| `GET /mopsfin_t187ap07_O_ci` | 上櫃公司資產負債表（一般業） |
| `GET /mopsfin_t187ap07_O_fh` | 上櫃公司資產負債表（金控業） |
| `GET /mopsfin_t187ap07_O_ins` | 上櫃公司資產負債表（保險業） |
| `GET /mopsfin_t187ap07_O_mim` | 上櫃公司資產負債表（異業） |

### 興櫃公司

#### 綜合損益表

| 端點 | 說明 |
|------|------|
| `GET /mopsfin_t187ap06_U_basi` | 興櫃公司綜合損益表（金融業） |
| `GET /mopsfin_t187ap06_U_bd` | 興櫃公司綜合損益表（證券期貨業） |
| `GET /mopsfin_t187ap06_U_ci` | 興櫃公司綜合損益表（一般業） |
| `GET /mopsfin_t187ap06_U_fh` | 興櫃公司綜合損益表（金控業） |
| `GET /mopsfin_t187ap06_U_ins` | 興櫃公司綜合損益表（保險業） |
| `GET /mopsfin_t187ap06_U_mim` | 興櫃公司綜合損益表（異業） |

#### 資產負債表

| 端點 | 說明 |
|------|------|
| `GET /mopsfin_t187ap07_U_basi` | 興櫃公司資產負債表（金融業） |
| `GET /mopsfin_t187ap07_U_bd` | 興櫃公司資產負債表（證券期貨業） |
| `GET /mopsfin_t187ap07_U_ci` | 興櫃公司資產負債表（一般業） |
| `GET /mopsfin_t187ap07_U_fh` | 興櫃公司資產負債表（金控業） |
| `GET /mopsfin_t187ap07_U_ins` | 興櫃公司資產負債表（保險業） |
| `GET /mopsfin_t187ap07_U_mim` | 興櫃公司資產負債表（異業） |

---

## 券商資料

| 端點 | 說明 |
|------|------|
| `GET /mopsfin_t187ap01` | 券商業務別人員數 |
| `GET /tpex_daily_broker2` | 上櫃股票各券商總公司當日營業金額統計表 |

---

## RR5 專用速查（TPEx 上櫃）

| 用途 | 端點 | 關鍵欄位 |
|------|------|----------|
| **查上櫃權證名稱**（現役） | `GET /tpex_warrant` | `Code`, `Name`, `UnderlyingCode`, `ExpirationDate` |
| **查上櫃權證發行資料**（含到期日） | `GET /tpex_warrant_issue` | `Code`, `Name`, `ListedDate`, `ExpiryDate`, `UnderlyingStockCode`, `Latest ExerciseRatio` |
| **上櫃權證日收盤行情** | `GET /tpex_warrant_daily_quts` | ClosingPrice 等 |
| **上櫃個股行情**（含本益比） | `GET /tpex_mainboard_peratio_analysis` | PE ratio, Dividend Yield |
| **上櫃除權除息預告** | `GET /tpex_exright_prepost` | — |
| **上櫃基本資料** | `GET /mopsfin_t187ap03_O` | 公司代號、名稱、產業別 |

> **注意**: `tpex_warrant` 和 `tpex_warrant_issue` 均為無參數端點，回傳**全部**上櫃權證。
> 使用方式：下載後在 JSON 中搜尋 `"Code":"XXXXXX"` 再取 `Name` 欄位。
> 已到期或下市的權證不會出現在這兩個端點中（改用 TWSE STOCK_DAY 補查）。
