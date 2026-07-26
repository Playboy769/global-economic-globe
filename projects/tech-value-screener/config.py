"""
設定檔：股票清單、篩選條件與基準數值
"""

# ─── 美股科技股觀察清單 ───────────────────────────────────────────────────────
US_STOCKS = [
    # 巨型科技
    "AAPL", "MSFT", "GOOGL", "META", "AMZN", "NVDA",
    # 雲端 / SaaS
    "CRM", "NOW", "ADBE", "WDAY", "SNOW", "DDOG", "MDB",
    "HUBS", "ZS", "CRWD", "NET", "PANW", "FTNT",
    # 半導體
    "AMD", "AVGO", "QCOM", "MRVL", "AMAT", "KLAC",
    # 其他高研發科技
    "ORCL", "INTU", "CDNS", "ANSS",
    # 平台 / 消費科技
    "UBER", "SHOP", "ABNB",
]

# ─── 台股科技股觀察清單（Yahoo Finance .TW 格式，免費不需 API Key）──────────────
TW_STOCKS = [
    "2330.TW",  # 台積電
    "2454.TW",  # 聯發科
    "2317.TW",  # 鴻海
    "2308.TW",  # 台達電
    "3711.TW",  # 日月光投控
    "2382.TW",  # 廣達
    "2379.TW",  # 瑞昱
    "3008.TW",  # 大立光
    "2303.TW",  # 聯電
    "2357.TW",  # 華碩
]

# ─── 篩選門檻 ─────────────────────────────────────────────────────────────────
SCREENING_CRITERIA = {
    "min_market_cap_b": 1.0,       # 最低市值 10億美元 (十億)
    "min_gross_margin": 40.0,      # 最低毛利率 40%
    "min_rd_pct": 5.0,             # 最低研發佔營收 5%
    "min_fcf_margin": 5.0,         # 最低 FCF 利潤率 5%
    "min_revenue_growth": -5.0,    # 允許輕微衰退（篩出倒閉公司）
}

# ─── 評分權重（合計 100）─────────────────────────────────────────────────────
SCORE_WEIGHTS = {
    "hidden_profit": 25,   # 隱藏盈利潛力
    "fcf_quality":   20,   # 現金流品質
    "growth":        20,   # 成長性 (Rule of 40)
    "valuation":     20,   # 估值合理性
    "rd_intensity":  15,   # 研發投入強度
}

# ─── 財務調整基準（傳統產業平均水準）──────────────────────────────────────────
BENCHMARKS = {
    "traditional_rd_pct":  3.0,   # 傳統企業研發率約 3%
    "traditional_sm_pct": 10.0,   # 傳統企業行銷率約 10%
    "us_tax_rate":        0.21,   # 美國企業稅率
    "tw_tax_rate":        0.20,   # 台灣企業稅率
}
