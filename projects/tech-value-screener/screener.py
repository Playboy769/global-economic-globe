"""
screener.py — 評分篩選器

每支股票從 5 個維度評分（各項滿分 100，加權後合計 100 分）：
  1. 隱藏盈利潛力 (25)：調整後 vs GAAP 利潤率的差距
  2. 現金流品質   (20)：FCF 利潤率
  3. 成長性       (20)：Rule of 40
  4. 估值合理性   (20)：Price / FCF 倍數
  5. 研發投入強度 (15)：研發費用佔營收比
"""
from typing import Dict, Any, Tuple, List

from config import SCREENING_CRITERIA, SCORE_WEIGHTS


def passes_filter(m: Dict[str, Any]) -> Tuple[bool, List[str]]:
    """回傳 (通過?, 失敗原因清單)"""
    c = SCREENING_CRITERIA
    fails = []

    if m["market_cap_b"] < c["min_market_cap_b"]:
        fails.append(f"市值 {m['market_cap_b']:.1f}B < {c['min_market_cap_b']}B")
    if m["gross_margin"] < c["min_gross_margin"]:
        fails.append(f"毛利率 {m['gross_margin']:.1f}% < {c['min_gross_margin']}%")
    if m["rd_pct"] < c["min_rd_pct"]:
        fails.append(f"研發率 {m['rd_pct']:.1f}% < {c['min_rd_pct']}%")
    if m["fcf_margin"] < c["min_fcf_margin"]:
        fails.append(f"FCF率 {m['fcf_margin']:.1f}% < {c['min_fcf_margin']}%")
    if m["revenue_growth"] < c["min_revenue_growth"]:
        fails.append(f"成長率 {m['revenue_growth']:.1f}% < {c['min_revenue_growth']}%")

    return len(fails) == 0, fails


def _clamp(v: float, lo: float, hi: float) -> float:
    return max(lo, min(hi, v))


def score(m: Dict[str, Any]) -> float:
    """計算 0-100 綜合評分"""
    w = SCORE_WEIGHTS
    total = 0.0

    # 1. 隱藏盈利潛力：利潤率差距 (gap = adj_A - GAAP)
    #    差距 ≥30% → 100分；差距 ≤5% → 0分
    gap = m.get("margin_gap", 0)
    hp  = _clamp((gap - 5) / 25, 0, 1) * 100
    total += hp * w["hidden_profit"] / 100

    # 2. 現金流品質：FCF 利潤率
    #    FCF ≥25% → 100分；FCF ≤0% → 0分
    fq = _clamp(m["fcf_margin"] / 25, 0, 1) * 100
    total += fq * w["fcf_quality"] / 100

    # 3. 成長性：Rule of 40
    #    R40 ≥60 → 100分；R40 ≤15 → 0分
    r40 = _clamp((m.get("rule_of_40", 0) - 15) / 45, 0, 1) * 100
    total += r40 * w["growth"] / 100

    # 4. 估值合理性：P/FCF（越低越好）
    #    P/FCF ≤12 → 100分；P/FCF ≥50 → 0分
    ptf = m.get("p_to_fcf")
    if ptf and ptf > 0:
        vs = _clamp((50 - ptf) / 38, 0, 1) * 100
    elif m.get("adj_pe_b") and m["adj_pe_b"] > 0:
        vs = _clamp((60 - m["adj_pe_b"]) / 40, 0, 1) * 100
    else:
        vs = 0.0
    total += vs * w["valuation"] / 100

    # 5. 研發投入強度：R&D%
    #    R&D ≥20% → 100分；R&D ≤5% → 0分
    ri = _clamp((m["rd_pct"] - 5) / 15, 0, 1) * 100
    total += ri * w["rd_intensity"] / 100

    return round(total, 1)
