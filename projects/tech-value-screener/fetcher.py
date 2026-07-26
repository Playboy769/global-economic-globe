"""
fetcher.py — Yahoo Finance 資料抓取器
支援美股（MSFT）與台股（2330.TW），完全免費、不需 API Key
"""
import time
import logging
from typing import Optional, Dict, Any

import yfinance as yf
import pandas as pd

logger = logging.getLogger(__name__)

# ─── 欄位別名對照（不同公司/市場格式有差異）──────────────────────────────────
_KEYS = {
    "revenue":    ["Total Revenue", "Revenue", "Operating Revenue"],
    "gross":      ["Gross Profit"],
    "rd":         ["Research And Development", "Research Development",
                   "Research And Development Expense"],
    "sm":         ["Selling And Marketing Expense", "Sales And Marketing",
                   "Selling Expense"],
    "sga":        ["Selling General Administrative",
                   "Selling General And Administrative",
                   "General And Administrative Expense"],
    "op_income":  ["Operating Income", "Operating Income Loss",
                   "Total Operating Profit Loss", "EBIT"],
    "net_income": ["Net Income", "Net Income From Continuing Operations",
                   "Net Income Common Stockholders"],
    "pretax":     ["Pretax Income", "Income Before Tax"],
    "tax":        ["Income Tax Expense", "Tax Provision"],
    "op_cf":      ["Operating Cash Flow",
                   "Cash Flow From Continuing Operating Activities",
                   "Net Cash Provided By Used In Operating Activities"],
    "capex":      ["Capital Expenditure", "Capital Expenditures",
                   "Purchase Of Property Plant And Equipment"],
    "fcf":        ["Free Cash Flow"],
}


def _safe(df: pd.DataFrame, key: str, col) -> float:
    """嘗試多個欄位名稱，取第一個非零有效值"""
    for k in _KEYS.get(key, [key]):
        try:
            if k in df.index:
                v = df.loc[k, col]
                if pd.notna(v) and v != 0:
                    return float(v)
        except (KeyError, TypeError):
            continue
    return 0.0


class StockFetcher:
    """通用股票資料抓取器（Yahoo Finance）"""

    def __init__(self, delay: float = 0.8):
        self.delay = delay

    def fetch(self, symbol: str) -> Optional[Dict[str, Any]]:
        """抓取單支股票完整財務資料，失敗回傳 None"""
        try:
            tk = yf.Ticker(symbol)
            info = tk.info or {}

            # 市值必須存在才繼續
            market_cap = info.get("marketCap") or 0
            if market_cap <= 0:
                logger.debug(f"[{symbol}] 無市值資料，跳過")
                return None

            income = tk.income_stmt
            cashflow = tk.cashflow

            if income is None or income.empty:
                logger.debug(f"[{symbol}] 無損益表資料，跳過")
                return None

            time.sleep(self.delay)

            # 判斷幣別（台股是 TWD）
            currency = info.get("currency", "USD")

            return {
                "symbol":      symbol,
                "name":        info.get("longName", symbol),
                "sector":      info.get("sector", ""),
                "industry":    info.get("industry", ""),
                "currency":    currency,
                "market_cap":  market_cap,
                "price":       info.get("currentPrice") or info.get("regularMarketPrice", 0),
                "trailing_pe": info.get("trailingPE"),
                "forward_pe":  info.get("forwardPE"),
                "income":      income,
                "cashflow":    cashflow,
            }

        except Exception as e:
            logger.warning(f"[{symbol}] 抓取失敗: {e}")
            return None

    def extract_financials(self, data: Dict[str, Any]) -> Optional[Dict[str, float]]:
        """從原始資料提取財務數字，含成長率計算"""
        income = data["income"]
        cashflow = data.get("cashflow")
        currency = data.get("currency", "USD")

        if income.empty or len(income.columns) < 1:
            return None

        c0 = income.columns[0]  # 最近年度
        c1 = income.columns[1] if len(income.columns) > 1 else None  # 前一年度

        # ── 損益表 ────────────────────────────────────────────────────────────
        rev  = _safe(income, "revenue",    c0)
        if rev <= 0:
            return None

        prev_rev   = _safe(income, "revenue",    c1) if c1 else 0.0
        gross      = _safe(income, "gross",      c0)
        rd         = abs(_safe(income, "rd",     c0))
        sm         = abs(_safe(income, "sm",     c0))
        sga        = abs(_safe(income, "sga",    c0))
        op_income  = _safe(income, "op_income",  c0)
        net_income = _safe(income, "net_income", c0)
        pretax     = _safe(income, "pretax",     c0)
        tax        = _safe(income, "tax",        c0)

        # 有效稅率
        if pretax > 0 and tax > 0:
            eff_tax = min(0.40, max(0.05, tax / pretax))
        else:
            eff_tax = 0.21 if currency == "USD" else 0.20

        # ── 現金流量表 ────────────────────────────────────────────────────────
        op_cf = 0.0
        capex = 0.0
        fcf   = 0.0

        if cashflow is not None and not cashflow.empty:
            cf0 = cashflow.columns[0]
            op_cf = _safe(cashflow, "op_cf",  cf0)
            capex = abs(_safe(cashflow, "capex", cf0))
            fcf_direct = _safe(cashflow, "fcf", cf0)
            fcf = fcf_direct if fcf_direct != 0 else (op_cf - capex)

        # 台股幣別換算：新台幣→美元（約 1 USD = 31 TWD，僅用於規模比較）
        fx = 31.0 if currency == "TWD" else 1.0

        return {
            "revenue":    rev / fx,
            "prev_rev":   prev_rev / fx,
            "gross":      gross / fx,
            "rd":         rd / fx,
            "sm":         sm / fx,
            "sga":        sga / fx,
            "op_income":  op_income / fx,
            "net_income": net_income / fx,
            "op_cf":      op_cf / fx,
            "capex":      capex / fx,
            "fcf":        fcf / fx,
            "eff_tax":    eff_tax,
        }
