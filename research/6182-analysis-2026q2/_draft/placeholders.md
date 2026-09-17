# 6182 report_template.html 佔位符清單

沿用 `research/6488-analysis-2026q2/data/report_template.html` 的視覺骨架（style／tab 結構／script），
但 6488 那份檔案其實已經是**填完文字、只剩圖表**的成品（見下方「發現的問題」），不是可重用的泛用範本。
本檔重新把全部公司專屬文字也換成佔位符，共 **102 個**（12 CHART + 18 TABLE + 72 TEXT），
供下一個撰寫子代理依 `research/6182-analysis-2026q2/data/` 的資料逐一填入。

資料來源請對應 `research/6182-analysis-2026q2/data/quarterly.json`、`financials_raw.json`、
`derived_metrics.json`（目前 data/ 尚缺 monthly_revenue.json、peers_yahoo.json、raw 法說會/年報文字檔——
待資料子代理補齊，命名可比照 6488 的 `data/` 同名檔案）。

## Tab 1 · 法說會重點檢核（Q&A Fallback）
| 佔位符 | 型態 | 內容 | 資料來源 |
|---|---|---|---|
| `qa_fallback_note` | TEXT | Fallback 說明：合晶有無法說會逐字稿/簡報，本 tab 改採口徑對照法 | 公開新聞、公司法說會公告 |
| `qa_stat_row` | TABLE(4×stat-box) | 本季 EPS 特殊性、關鍵一次性項目、毛利率相對位置、季後最新月營收 YoY | quarterly.json + monthly_revenue.json |
| `qa_stat1_val/lbl`…`qa_stat4_val/lbl` | TEXT | 若不用 TABLE 逐格填，個別 stat-box 數值與標籤（與上列二擇一填法） | 同上 |
| `qa_summary` | TEXT | 一句話結論，含矛盾點（如本業 vs 非本業） | quarterly.json |
| `qa_verification` | TABLE | 管理層口徑 vs 財報印證，6-10 列，含一列留白反推、一列月營收驗證 | 法說會轉述/簡報 + quarterly.json + monthly_revenue.json |
| `qa_verification_interp` | TEXT | ⚑ 解讀，總結印證/未印證/留白比例 | 上表推導 |
| `qa_orig_quotes` | TEXT | 逐字引用法說會/簡報原文（.orig-box），找不到則留空並在 header 註記 | 公開新聞轉述 |

## Tab 2 · 特殊交易/事件（由財報附註挑選本季最重大事件）
| 佔位符 | 型態 | 內容 | 資料來源 |
|---|---|---|---|
| `deal_title` | TEXT | 事件標題 | financials_raw.json 附註 |
| `deal_intro` | TEXT | 事件說明卡片內文 | 同上 |
| `deal_scale_items` | TEXT(<li>) | 規模象限項目 | 同上 |
| `deal_timeline_items` | TEXT(<li>) | 期限象限項目 | 同上 |
| `deal_commitment_items` | TEXT(<li>) | 財務承諾象限項目 | 同上 |
| `deal_unverified_items` | TEXT(<li>) | 尚待驗證象限項目 | 同上 |
| `datacenter_financing` | TABLE | 資料中心專案融資條款；合晶預期無此類資本支出，至少一列寫「不適用」 | financials_raw.json（Commitments/VIE 附註） |
| `deal_financing_interp` | TEXT | ⚑ 解讀 | 上表推導 |
| `deal_key_questions` | TEXT(<ol><li>) | 3-5 條盡職調查式追問 | 綜合推演 |
| `deal_timeline_table` | TABLE | 關鍵時程：時間｜事件｜對財報的預期影響 | 財報期後事項 + 公開新聞 |

## Tab 3 · 技術與產品線
| 佔位符 | 型態 | 內容 | 資料來源 |
|---|---|---|---|
| `tech_positioning` | TEXT | 合晶市場定位、產品組合、本季收入結構 | 年報、財報附註 |
| `tech_cards` | TABLE(.tech-card×4-6) | 產品線卡片，含產品是什麼/用在哪/客戶為何需要 | 年報產品描述 |
| `bu_mix`（CHART） | CHART | 產品線營收圖 | quarterly.json 附註六產品別 |
| `tech_bu_mix_interp` | TEXT | ⚑ 解讀 + 出處 | 上圖推導 |
| `tech_causal_chain` | TEXT | 終端需求→矽晶圓耗用因果鏈（正向＋反向） | 產業研究、年報 |

## Tab 4 · 供應鏈
| 佔位符 | 型態 | 內容 | 資料來源 |
|---|---|---|---|
| `chain_intro` | TEXT | 上下游集中度概述 | 年報主要客戶/供應商揭露 |
| `chain_supply`（CHART） | CHART | chain-diagram 結構圖 | 年報 |
| `chain_structure_interp` | TEXT | ⚑ 解讀 | 同上 |
| `sankey_supply`（CHART） | CHART | Sankey 流量圖，節點需與 chain_supply 一致 | 年報 |
| `chain_flow_interp` | TEXT | ⚑ 解讀 | 同上 |
| `chain_detail` | TABLE | 上下游關係明細（含 section-head 上游/下游） | 年報 + 財報附註 |
| `chain_crosslink_note` | TEXT | ⚑ 交叉連結（.note yellow），對應 BS tab 與估值觀察互證表 | 綜合 |

## Tab 5 · 風險矩陣
| 佔位符 | 型態 | 內容 | 資料來源 |
|---|---|---|---|
| `risk_fallback_note` | TEXT | 資料限制說明 | — |
| `risk_items` | TABLE(.risk-item×6-8) | 風險卡片，至少一條【獨創前瞻風險】 | 綜合財報+新聞 |
| `risk_heat`（CHART） | CHART | 風險矩陣熱力圖 | 上列風險主觀評分 |
| `risk_heat_interp` | TEXT | ⚑ 解讀 | 同上 |

## Tab 6 · 損益表
| 佔位符 | 型態 | 內容 | 資料來源 |
|---|---|---|---|
| `inc_kpi_row` | TABLE(.kpi×4) | 營收/毛利率/營業利益率/EPS + QoQ/YoY | quarterly.json |
| `inc_statement` | TABLE | 逐行損益表（含 q_curr/q_prev/q_yoy/ytd_curr/ytd_prev 表頭） | quarterly.json |
| `q_curr`/`q_prev`/`q_yoy`/`ytd_curr`/`ytd_prev` | TEXT | 表頭期別標籤 | quarterly.json |
| `inc_statement_interp` | TEXT | ⚑ 解讀，本業線 vs 非本業線拆解 | 上表推導 |
| `income_trend`（CHART） | CHART | 12 季營收與利潤率趨勢 | quarterly.json |
| `income_trend_interp` | TEXT | ⚑ 解讀 | 同上 |
| `monthly_revenue`（CHART） | CHART | 近 24+ 個月月營收與 YoY（台股必備） | monthly_revenue.json |
| `monthly_revenue_interp` | TEXT | ⚑ 解讀，回答月營收是否支持管理層展望 | 同上 |
| `guidance_vs_actual` | TABLE | 財測/口徑 vs 實際 | qa_verification 延伸 |

## Tab 7 · 業務部門
| 佔位符 | 型態 | 內容 | 資料來源 |
|---|---|---|---|
| `bu_disclosure_note` | TEXT | 部門揭露限制與替代做法說明 | quarterly.json 部門附註 |
| `bu_mix`（CHART，與 tab3 共用） | CHART | 各 BU/產品線營收利潤率對比 | quarterly.json |
| `bu_chart_interp` | TEXT | ⚑ 解讀 | 同上 |
| `bu_cards` | TABLE(.bu-card×2-4) | 各 BU/產品線卡片，含產品技術深度說明 | 年報 + 財報附註 |
| `bu_region_table` | TABLE | 地區別/子公司損益，無揭露寫「無此資料/不適用」 | 財報部門附註 |
| `bu_quantity_price_interp` | TEXT | ⚑ 解讀（量價拆分） | quarterly.json 反推 |

## Tab 8 · 資產負債表
| 佔位符 | 型態 | 內容 | 資料來源 |
|---|---|---|---|
| `bs_curr_date`/`bs_yearend_date`/`bs_yoy_date` | TEXT | 三期別表頭日期 | quarterly.json |
| `bs_statement` | TABLE | 三期別資產負債表（含資產/負債/權益 section-head） | quarterly.json |
| `bs_statement_interp` | TEXT | ⚑ 解讀，結構性變化 | 上表推導 |
| `dso_value`/`dso_source`/`dso_interp` | TEXT | DSO 追蹤（框架必備①） | derived_metrics.json |
| `unstarted_lease_value`/`_source`/`_interp` | TEXT | 未開始租賃（框架必備②，缺資料寫無此資料） | 財報租賃附註 |
| `purchase_commit_value`/`_source`/`_interp` | TEXT | 採購承諾（框架必備③） | 財報承諾附註 |
| `customer_prepay_value`/`_source`/`_interp` | TEXT | 客戶預付款/合約負債（框架必備④） | 財報合約負債附註 |
| `bs_tracking_four`（TABLE，實為上四組 TEXT 組成，保留占位名一致） | — | 見上四項 | — |
| `contract_liab`（CHART） | CHART | 合約負債 12 季走勢 | quarterly.json |
| `contract_liab_interp` | TEXT | ⚑ 解讀 | 同上 |
| `dso`（CHART） | CHART | DSO 與存貨天數趨勢 | derived_metrics.json |
| `dso_chart_interp` | TEXT | ⚑ 解讀 | 同上 |
| `bs_misc_notes` | TEXT(<li>) | 股利/庫藏股/員工酬勞/擔保質押/信評/母公司持股 | 財報附註 |

## Tab 9 · 現金流量
| 佔位符 | 型態 | 內容 | 資料來源 |
|---|---|---|---|
| `cf_curr_period_lbl`/`cf_ytd_period_lbl` | TEXT | 期別標籤 | quarterly.json |
| `cf_quarter_block` | TEXT(HTML 區塊) | 單季現金流三段（營業/投資/籌資）+ FCF | quarterly.json（YTD 相減） |
| `cf_ytd_block` | TEXT(HTML 區塊) | 累計 vs 去年同期現金流 | quarterly.json |
| `cf_interp` | TEXT | ⚑ 解讀 | 上兩塊推導 |
| `fcf_waterfall`（CHART） | CHART | FCF 橋接（必備） | quarterly.json |
| `fcf_waterfall_interp` | TEXT | ⚑ 解讀 | 同上 |
| `capex_dep`（CHART） | CHART | 資本支出 vs 折舊 vs CFO | quarterly.json |
| `capex_dep_interp` | TEXT | ⚑ 解讀 | 同上 |

## Tab 10 · 估值觀察
| 佔位符 | 型態 | 內容 | 資料來源 |
|---|---|---|---|
| `val_stat_row` | TABLE(4×stat-box) | 收盤價/市值、Trailing P/E、核心 P/E、P/B | peers_yahoo.json |
| `val_ratios` | TABLE | 估值與體質比率，必含利息保障倍數（框架必備） | derived_metrics.json |
| `val_ratios_interp` | TEXT | ⚑ 隱含 P/E 反推市場懷疑程度 | 上表推導 |
| `int_cov`（CHART） | CHART | 利息保障倍數趨勢 | derived_metrics.json |
| `int_cov_interp` | TEXT | ⚑ 解讀 | 同上 |
| `valuation_scatter`（CHART） | CHART | 同業 Forward P/E vs 營收成長（必備） | peers_yahoo.json |
| `valuation_scatter_interp` | TEXT | ⚑ 解讀，若比較意義有限仍需完整說明理由 | 同上 |
| `peer_valuation` | TABLE | 同業估值明細（6182 自身 + 6488/3532/3016/4063.T/3436.T 等） | peers_yahoo.json |
| `cross_verification` | TABLE | 互證對照表（強制交付物），10-13 列，含月營收實績對照一列 | 綜合全篇 |
| `cross_verification_summary` | TEXT | ⚑ 總結，待驗項與下次財報日 | 同上 |

## Footer / Header
| 佔位符 | 型態 | 內容 | 資料來源 |
|---|---|---|---|
| `header_subtitle` | TEXT | 主標題下方副標，列 10 個 tab 名稱、主軸期間、資料來源清單、Fallback 模式說明 | 綜合 |
| `footer_sources` | TEXT | 完整資料來源清單＋免責聲明，結尾「僅供資訊參考，非投資建議」 | 綜合 |

---

## 完成後合併檢查清單（提醒撰寫子代理）
1. 全部 102 個佔位符填完後，用 `grep -c "{{"` 確認為 0。
2. 全檔僅一個 `class="panel active"`（目前在 `panel-qa`）。
3. 全檔搜尋 `class="... panel"`，確認沒有非分頁元素誤用 `panel` class。
4. 6 張必備圖（chain_supply、sankey_supply、income_trend、bu_mix 對應長條圖、fcf_waterfall、
   valuation_scatter）＋台股第 7 張必備圖 monthly_revenue，共 7 張必須全部有內容，不可留空。
5. 資產負債表 tab 四項框架必備追蹤（DSO/未開始租賃/採購承諾/客戶預付款）與估值觀察 tab
   利息保障倍數、互證對照表，缺資料一律明確寫「無此資料/不適用」，不可略過不提。
6. 完成後複製到 `globe-invest/app/research/`（檔名一致、不建子資料夾），並在
   `app/OutsideFramework/index.html` Works → 矽晶圓分組上架、加 `data-published` 屬性與
   `EARN_DATES`，兩個 repo 分別 commit + push（見 CLAUDE.md「分析完成後強制同步與上架」）。
