"""
adjuster.py — 財務調整計算器

核心邏輯（對應圖片概念）：
  科技公司為追求未來成長，在研發(R&D)和行銷(S&M)超額支出。
  這些費用在 GAAP 下立即認列為損失，壓低當期盈利。
  「調整」的目的：還原這些被隱藏的盈利潛力，找出被低估的科技股。

  調整方法 A：加回全部 R&D（最樂觀）
    ➜ 假設所有研發費用都是未來投資，如同製造業的資本支出
  調整方法 B：加回超額 R&D（較保守）
    ➜ 只加回超過「傳統企業基準（3%）」的部分
  方法 C：直接用 FCF（最客觀）
    ➜ 現金流不受會計準則扭曲，是最真實的獲利衡量
"""
import logging
from typing import Optional, Dict, Any

from config import BENCHMARKS
from fetcher import StockFetcher

logger = logging.getLogger(__name__)

_fetcher = StockFetcher()


def calculate(raw: Dict[str, Any]) -> Optional[Dict[str, Any]]:
    """輸入 fetcher.fetch() 的原始資料，輸出所有調整後指標"""
    fin = _fetcher.extract_financials(raw)
    if not fin:
        return None

    symbol = raw["symbol"]
    name   = raw["name"]
    mktcap = raw["market_cap"]
    currency = raw.get("currency", "USD")

    # 台股市值換算 (TWD → USD)
    fx = 31.0 if currency == "TWD" else 1.0
    mktcap_usd = mktcap / fx

    rev  = fin["revenue"]
    if rev <= 0:
        return None

    gross      = fin["gross"]
    rd         = fin["rd"]
    sm         = fin["sm"]
    op_income  = fin["op_income"]
    net_income = fin["net_income"]
    fcf        = fin["fcf"]
    prev_rev   = fin["prev_rev"]
    eff_tax    = fin["eff_tax"]

    trd = BENCHMARKS["traditional_rd_pct"]  # 傳統基準研發率 3%
    excess_rd = max(0.0, rd - rev * trd / 100)

    # ── GAAP 率值 ──────────────────────────────────────────────────────────
    gross_m   = gross / rev * 100
    op_m      = op_income / rev * 100
    net_m     = net_income / rev * 100
    rd_pct    = rd / rev * 100
    sm_pct    = sm / rev * 100
    fcf_m     = fcf / rev * 100

    # ── 成長率 ────────────────────────────────────────────────────────────
    rev_growth = (rev - prev_rev) / prev_rev * 100 if prev_rev > 0 else 0.0

    # ── 調整後指標 ────────────────────────────────────────────────────────
    # 方法 A：加回全部 R&D（資本化處理概念）
    adj_op_a    = op_income + rd
    adj_op_m_a  = adj_op_a / rev * 100
    adj_ni_a    = net_income + rd * (1 - eff_tax)

    # 方法 B：加回超額 R&D（保守版）
    adj_op_b    = op_income + excess_rd
    adj_op_m_b  = adj_op_b / rev * 100
    adj_ni_b    = net_income + excess_rd * (1 - eff_tax)

    # ── 估值指標 ──────────────────────────────────────────────────────────
    trailing_pe = raw.get("trailing_pe")

    adj_pe_a = (mktcap_usd / adj_ni_a) if adj_ni_a > 0 and mktcap_usd > 0 else None
    adj_pe_b = (mktcap_usd / adj_ni_b) if adj_ni_b > 0 and mktcap_usd > 0 else None
    p_to_fcf = (mktcap_usd / fcf)      if fcf > 0 and mktcap_usd > 0 else None
    fcf_yield = (fcf / mktcap_usd * 100) if mktcap_usd > 0 else 0.0

    # ── 複合指標 ──────────────────────────────────────────────────────────
    # Rule of 40 = 營收成長% + FCF利潤率%
    rule_of_40 = rev_growth + fcf_m

    # 隱藏盈利比 = 調整後淨利 / GAAP淨利（倍數越大代表隱藏越多）
    hidden_ratio = (adj_ni_a / net_income) if net_income > 0 and adj_ni_a > 0 else None

    return {
        # 基本資訊
        "symbol":        symbol,
        "name":          name[:28],
        "sector":        raw.get("sector", ""),
        "industry":      raw.get("industry", ""),
        "currency":      currency,
        "market_cap_b":  mktcap_usd / 1e9,
        "price":         raw.get("price", 0),

        # ── GAAP 指標 ──
        "revenue_b":         rev / 1e9,
        "revenue_growth":    rev_growth,
        "gross_margin":      gross_m,
        "op_margin":         op_m,
        "net_margin":        net_m,
        "rd_pct":            rd_pct,
        "sm_pct":            sm_pct,
        "fcf_margin":        fcf_m,
        "trailing_pe":       trailing_pe,

        # ── 調整後指標 ──
        "adj_op_margin_a":   adj_op_m_a,    # 全 R&D 加回
        "adj_op_margin_b":   adj_op_m_b,    # 超額 R&D 加回
        "adj_pe_a":          adj_pe_a,
        "adj_pe_b":          adj_pe_b,
        "p_to_fcf":          p_to_fcf,
        "fcf_yield":         fcf_yield,

        # ── 複合指標 ──
        "rule_of_40":        rule_of_40,
        "hidden_ratio":      hidden_ratio,
        "margin_gap":        adj_op_m_a - op_m,  # 隱藏盈利的利潤率差距

        # ── 原始金額（給 Excel 明細用）──
        "_rev":              rev,
        "_rd":               rd,
        "_excess_rd":        excess_rd,
        "_sm":               sm,
        "_op_income":        op_income,
        "_net_income":       net_income,
        "_fcf":              fcf,
        "_mktcap":           mktcap_usd,
    }
