#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
RR4 Valuation (nav page VA / VA!) - ROIC-based relative valuation, no WACC.

Called by RR4/modValuationPage.bas (Excel VBA shells this script and reads
the result file back).  Everything numeric lives here so it can be tested
from the command line without Excel:

    python valuation.py --ticker 3017 --market TW --peers 3017,2308,2421 --out out.txt
    python valuation.py --req request.json --out out.txt      # what VBA does

Definitions follow SEC-Filing-Fetcher/modValuation.bas so the numbers line
up with the Filings / TW_Filings sheets:
    Invested capital IC = Total assets - Current liabilities
    NOPAT              = Operating income x (1 - effective tax rate)
    ROIC               = NOPAT / IC (period end)
Two extra IC readings are shown next to it (never instead of it):
    IC ex-cash = IC - cash - short-term investments   (operating ROIC)
    IC std     = IC + short-term borrowings           (textbook: only non-
                 interest-bearing current liabilities are netted)
Flows use TTM (last four quarters) for the peer table and fair-value block;
the target's quarterly history shows both the x4 annualised quarter (the
Excel column convention) and the rolling TTM.

Data:
    US  SEC companyfacts (data.sec.gov), one request per company, plus the
        ticker->CIK map from sec.gov.  Q4 = 10-K full year minus the three
        quarters inside that fiscal year.
    TW  MOPS t164sb01 inline XBRL (cp950).  Flows come as year-to-date, so
        TTM = current YTD + prior FY - prior-year YTD of the same quarter,
        and quarters use the 3-month contexts (Q4 = FY - Q3 YTD).  Shares
        = ordinary share capital / NT$10 par.
    Prices  Yahoo chart API (last close).  TW tries .TW then .TWO.
Peer group comes from the workbook (tblGroups), passed in the request.

Output is a plain text file of tab-separated blocks (VBA has no JSON
parser): "#NAME" starts a block, following lines are rows.
"""
import argparse
import json
import math
import os
import re
import statistics
import sys
import time
from datetime import date, datetime, timedelta

try:
    import requests
except ImportError:  # pragma: no cover
    sys.stderr.write("requests is required (pip install requests)\n")
    sys.exit(2)

UA = "Ryan Personal Research Tool ryan929929@gmail.com"   # same as shared-vba/modHttp.SEC_USER_AGENT
YAHOO_UA = ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
            "(KHTML, like Gecko) Chrome/124 Safari/537.36")
CACHE_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), ".valuation-cache")
TAX_FALLBACK = {"US": 0.21, "TW": 0.20}
ROIC_FLOOR = {"US": 0.10, "TW": 0.08}      # absolute hurdle replacing WACC
HIST_QUARTERS = 8
LOG = []


def log(msg):
    LOG.append(msg)
    sys.stderr.write(msg + "\n")


# ----------------------------------------------------------------------------
# small helpers
# ----------------------------------------------------------------------------
def cache_path(key):
    os.makedirs(CACHE_DIR, exist_ok=True)
    safe = re.sub(r"[^A-Za-z0-9_.-]", "_", key)
    return os.path.join(CACHE_DIR, safe)


def fetch(url, key, headers, ttl_days=1, binary=False, params=None):
    """GET with an on-disk cache (so re-running the same day does not hammer
    MOPS / SEC).  Returns text (or bytes)."""
    p = cache_path(key)
    if os.path.exists(p) and (time.time() - os.path.getmtime(p)) < ttl_days * 86400:
        with open(p, "rb") as f:
            data = f.read()
        return data if binary else data.decode("utf-8", "replace")
    for attempt in range(3):
        try:
            r = requests.get(url, headers=headers, params=params, timeout=40)
            if r.status_code == 429 or r.status_code >= 500:
                time.sleep(2 + attempt * 3)
                continue
            if r.status_code >= 400:
                raise RuntimeError("HTTP %d %s" % (r.status_code, url))
            r.raise_for_status()
            data = r.content
            with open(p, "wb") as f:
                f.write(data)
            return data if binary else data.decode("utf-8", "replace")
        except requests.RequestException as e:
            log("fetch retry %d %s: %s" % (attempt + 1, url, e))
            time.sleep(2 + attempt * 3)
    raise RuntimeError("fetch failed: " + url)


def num(v):
    return v is not None and isinstance(v, (int, float)) and not (isinstance(v, float) and math.isnan(v))


def div(a, b):
    if num(a) and num(b) and b != 0:
        return a / b
    return None


def fmt(v, kind="num"):
    if not num(v):
        return ""
    if kind == "pct":
        return "%.4f" % v
    return repr(float(v))


def parse_date(s):
    return datetime.strptime(s, "%Y-%m-%d").date()


# ----------------------------------------------------------------------------
# Yahoo price
# ----------------------------------------------------------------------------
def yahoo_last_close(sym):
    url = "https://query1.finance.yahoo.com/v8/finance/chart/%s" % sym
    try:
        txt = fetch(url, "yahoo_%s.json" % sym, {"User-Agent": YAHOO_UA}, ttl_days=0.5,
                    params={"range": "5d", "interval": "1d"})
        j = json.loads(txt)
        res = j["chart"]["result"][0]
        meta = res["meta"]
        px = meta.get("regularMarketPrice")
        if not num(px):
            closes = [c for c in res["indicators"]["quote"][0]["close"] if c is not None]
            px = closes[-1] if closes else None
        return px, meta.get("currency", "")
    except Exception as e:  # noqa
        log("yahoo %s: %s" % (sym, e))
        return None, ""


def price_for(ticker, market):
    if market == "TW":
        for suf in (".TW", ".TWO"):
            px, cur = yahoo_last_close(ticker + suf)
            if num(px):
                return px, ticker + suf
        return None, ""
    px, cur = yahoo_last_close(ticker.replace(".", "-"))
    return px, ticker


# ----------------------------------------------------------------------------
# US: SEC companyfacts
# ----------------------------------------------------------------------------
_CIK_MAP = None


def cik_for(ticker):
    global _CIK_MAP
    if _CIK_MAP is None:
        txt = fetch("https://www.sec.gov/files/company_tickers.json", "sec_company_tickers.json",
                    {"User-Agent": UA}, ttl_days=7)
        _CIK_MAP = {}
        for row in json.loads(txt).values():
            _CIK_MAP[row["ticker"].upper()] = int(row["cik_str"])
    t = ticker.upper().replace(".", "-")
    return _CIK_MAP.get(t)


US_CONCEPTS = {
    "ebit": ["OperatingIncomeLoss"],
    "revenue": ["Revenues", "RevenueFromContractWithCustomerExcludingAssessedTax", "SalesRevenueNet"],
    "pretax": ["IncomeLossFromContinuingOperationsBeforeIncomeTaxesExtraordinaryItemsNoncontrollingInterest",
               "IncomeLossFromContinuingOperationsBeforeIncomeTaxesMinorityInterestAndIncomeLossFromEquityMethodInvestments"],
    "tax": ["IncomeTaxExpenseBenefit"],
    "da": ["DepreciationDepletionAndAmortization", "DepreciationAndAmortization", "DepreciationAmortizationAndAccretionNet"],
    "capex": ["PaymentsToAcquirePropertyPlantAndEquipment", "PaymentsToAcquireProductiveAssets"],
    "assets": ["Assets"],
    "cur_liab": ["LiabilitiesCurrent"],
    "cash": ["CashAndCashEquivalentsAtCarryingValue"],
    "st_inv": ["ShortTermInvestments", "MarketableSecuritiesCurrent", "AvailableForSaleSecuritiesDebtSecuritiesCurrent"],
    "debt_cur": ["DebtCurrent"],
    "ltd_cur": ["LongTermDebtCurrent"],
    "st_borrow": ["ShortTermBorrowings", "CommercialPaper"],
    "ltd": ["LongTermDebtNoncurrent", "LongTermDebt"],
    "lease": ["OperatingLeaseLiability"],
    "lease_cur": ["OperatingLeaseLiabilityCurrent"],
    "lease_nc": ["OperatingLeaseLiabilityNoncurrent"],
    "minority": ["MinorityInterest"],
    "equity": ["StockholdersEquity"],
}
US_FLOWS = {"ebit", "revenue", "pretax", "tax", "da", "capex"}


def _dedupe_latest(entries):
    """companyfacts repeats a period in every later filing; keep the most
    recently filed value per (start,end)."""
    best = {}
    for e in entries:
        k = (e.get("start"), e["end"])
        if k not in best or e["filed"] > best[k]["filed"]:
            best[k] = e
    return best


def us_series(facts, concept_names, flow):
    """Flows: dict end_date -> quarterly value plus dict end -> annual value.
    Instants: dict end -> value.  Every candidate concept is merged (first
    hit per period wins) because filers switch tags between years.
    10-Q cash-flow tags are year-to-date: within one start date the
    consecutive ends are de-cumulated (Q2 = 6M - Q1, Q3 = 9M - 6M,
    Q4 = FY - 9M)."""
    gaap = facts.get("facts", {}).get("us-gaap", {})
    inst, quarters, annual = {}, {}, {}
    for n in concept_names:
        c = gaap.get(n)
        if not c or "USD" not in c.get("units", {}):
            continue
        best = _dedupe_latest(c["units"]["USD"])
        if not flow:
            for (s, e), v in best.items():
                inst.setdefault(parse_date(e), v["val"])
            continue
        by_start = {}
        for (s, e), v in best.items():
            if s:
                by_start.setdefault(parse_date(s), {})[parse_date(e)] = v["val"]
        for sd, ends in by_start.items():
            prev_end, prev_val = None, 0.0
            for ed in sorted(ends):
                days = (ed - sd).days
                val = ends[ed]
                if 80 <= days <= 100:
                    quarters.setdefault(ed, val)
                elif 350 <= days <= 380:
                    annual.setdefault(ed, val)
                if prev_end is not None and 80 <= (ed - prev_end).days <= 100 and 80 <= (prev_end - sd).days <= 290:
                    quarters.setdefault(ed, val - prev_val)
                prev_end, prev_val = ed, val
    if not flow:
        return inst, {}
    # Q4 from FY minus the three quarters inside it (when de-cumulation could not)
    for fy_end, fy_val in annual.items():
        if fy_end in quarters:
            continue
        inside = sorted(q for q in quarters if fy_end - timedelta(days=370) < q < fy_end)
        if len(inside) >= 3:
            quarters[fy_end] = fy_val - sum(quarters[q] for q in inside[-3:])
    return quarters, annual


def us_shares(facts):
    """dei EntityCommonStockSharesOutstanding (cover page; multi-class rows
    of the same filing are summed), else us-gaap CommonStockSharesOutstanding,
    else the latest diluted weighted average (GOOGL / META tag neither of
    the first two in companyfacts)."""
    dei = facts.get("facts", {}).get("dei", {}).get("EntityCommonStockSharesOutstanding")
    if dei and dei.get("units", {}).get("shares"):
        entries = dei["units"]["shares"]
        latest_end = max(parse_date(e["end"]) for e in entries)
        same = [e for e in entries if parse_date(e["end"]) == latest_end]
        accn = max(same, key=lambda e: e["filed"])["accn"]
        vals = {}
        for e in same:
            if e["accn"] == accn:
                vals[e.get("frame", "") + str(e["val"])] = e["val"]
        return sum(vals.values()), "dei EntityCommonStockSharesOutstanding as of %s" % latest_end
    gaap = facts.get("facts", {}).get("us-gaap", {})
    for tag, note in (("CommonStockSharesOutstanding", "us-gaap CommonStockSharesOutstanding as of %s"),
                      ("WeightedAverageNumberOfDilutedSharesOutstanding", "diluted weighted average shares, quarter to %s")):
        c = gaap.get(tag)
        if c and c.get("units", {}).get("shares"):
            best = _dedupe_latest(c["units"]["shares"])
            if tag.startswith("Weighted"):
                best = {k: v for k, v in best.items() if k[0] and 80 <= (parse_date(k[1]) - parse_date(k[0])).days <= 100}
            if best:
                k = max(best, key=lambda k: k[1])
                return best[k]["val"], note % k[1]
    return None, "no share count tag"


def load_us(ticker, history=True):
    cik = cik_for(ticker)
    if cik is None:
        raise RuntimeError("no CIK for " + ticker)
    txt = fetch("https://data.sec.gov/api/xbrl/companyfacts/CIK%010d.json" % cik,
                "sec_facts_%d.json" % cik, {"User-Agent": UA}, ttl_days=1)
    facts = json.loads(txt)
    co = {"ticker": ticker, "market": "US", "name": facts.get("entityName", ticker), "unit": "USD"}
    q, inst = {}, {}
    for k, names in US_CONCEPTS.items():
        s, a = us_series(facts, names, k in US_FLOWS)
        if k in US_FLOWS:
            q[k] = s
        else:
            inst[k] = s
    if not q["ebit"]:
        raise RuntimeError("no OperatingIncomeLoss for " + ticker)
    ends = sorted(q["ebit"].keys())
    periods = []
    for ed in ends[-(HIST_QUARTERS + 4):]:
        row = {"end": ed}
        for k in US_FLOWS:
            row[k] = q[k].get(ed)
        for k in inst:
            row[k] = inst[k].get(ed)
        periods.append(row)
    # debt composition
    for row in periods:
        st = row.get("debt_cur")
        if not num(st):
            st = (row.get("ltd_cur") or 0) + (row.get("st_borrow") or 0)
            if st == 0 and not num(row.get("ltd_cur")) and not num(row.get("st_borrow")):
                st = None
        row["st_debt"] = st
        row["lt_debt"] = row.get("ltd")
        lease = row.get("lease")
        if not num(lease):
            l1, l2 = row.get("lease_cur"), row.get("lease_nc")
            lease = (l1 or 0) + (l2 or 0) if (num(l1) or num(l2)) else None
        row["lease_total"] = lease
        # IC std adds back short-term interest-bearing debt sitting in current liabilities
        row["st_borrow_for_ic"] = st
    co["periods"] = periods
    sh, sh_note = us_shares(facts)
    co["shares"] = sh
    co["shares_note"] = sh_note
    co["src"] = "SEC companyfacts CIK %d" % cik
    return co


# ----------------------------------------------------------------------------
# TW: MOPS t164sb01 inline XBRL
# ----------------------------------------------------------------------------
IX_RE = re.compile(r"<ix:nonfraction\s+([^>]*)>([^<]*)</ix:nonfraction>", re.I)
ATTR_RE = re.compile(r'(\w[\w:-]*)\s*=\s*"([^"]*)"')
NONNUM_RE = re.compile(r'<ix:nonnumeric\s+name="tifrs-notes:CompanyChineseName"[^>]*>([^<]*)</ix:nonnumeric>', re.I)


def mops_report(co_id, year, season):
    url = "https://mopsov.twse.com.tw/server-java/t164sb01"
    key = "mops_%s_%d_%d.html" % (co_id, year, season)
    raw = fetch(url, key, {"User-Agent": YAHOO_UA}, ttl_days=30, binary=True,
                params={"step": "1", "CO_ID": co_id, "SYEAR": str(year), "SSEASON": str(season), "REPORT_ID": "C"})
    html = raw.decode("cp950", "replace")
    if "ix:nonfraction" not in html.lower():
        # do not keep an empty page in the cache
        try:
            os.remove(cache_path(key))
        except OSError:
            pass
        return None
    return html


def parse_ix(html):
    facts = {}
    for m in IX_RE.finditer(html):
        attrs = dict((k.lower(), v) for k, v in ATTR_RE.findall(m.group(1)))
        name, ctx = attrs.get("name"), attrs.get("contextref")
        if not name or not ctx or "_" in ctx:
            continue
        raw = m.group(2).strip().replace(",", "")
        if raw in ("", "-"):
            continue
        try:
            v = float(raw)
        except ValueError:
            continue
        sc = attrs.get("scale")
        if sc and sc.lstrip("-").isdigit():
            v *= 10 ** int(sc)
        if attrs.get("sign") == "-":
            v = -v
        facts[(name, ctx)] = v
    return facts


def q_bounds(y, q):
    sm = (q - 1) * 3 + 1
    start = date(y, sm, 1)
    end = date(y + (1 if sm + 3 > 12 else 0), (sm + 3 - 1) % 12 + 1, 1) - timedelta(days=1)
    return start, end


def ctx_dur(s, e):
    return "From%sTo%s" % (s.strftime("%Y%m%d"), e.strftime("%Y%m%d"))


def ctx_inst(e):
    return "AsOf" + e.strftime("%Y%m%d")


TW_FLOW = {
    "ebit": ["ifrs-full:ProfitLossFromOperatingActivities"],
    "revenue": ["ifrs-full:Revenue"],
    "pretax": ["ifrs-full:ProfitLossBeforeTax", "tifrs-SCF:ProfitLossBeforeTax"],
    "tax": ["ifrs-full:IncomeTaxExpenseContinuingOperations"],
    "da": ["ifrs-full:AdjustmentsForDepreciationExpense", "+ifrs-full:AdjustmentsForAmortisationExpense"],
    "capex": ["ifrs-full:PurchaseOfPropertyPlantAndEquipmentClassifiedAsInvestingActivities"],
}
TW_INST = {
    "assets": ["ifrs-full:Assets"],
    "cur_liab": ["ifrs-full:CurrentLiabilities"],
    "cash": ["ifrs-full:CashAndCashEquivalents"],
    "st_inv": ["tifrs-bsci-ci:CurrentFinancialAssetsAtFairValueThroughProfitOrLoss",
               "ifrs-full:CurrentFinancialAssetsAtFairValueThroughProfitOrLoss"],
    "st_borrow": ["ifrs-full:ShorttermBorrowings", "+tifrs-bsci-ci:ShorttermNotesAndBillsPayable",
                  "+tifrs-bsci-ci:LongtermBorrowingsCurrentPortion", "+ifrs-full:CurrentPortionOfLongtermBorrowings",
                  "+tifrs-bsci-ci:BondsPayableCurrentPortion"],
    "lt_debt": ["ifrs-full:LongtermBorrowings", "+ifrs-full:BondsIssued", "+tifrs-bsci-ci:BondsPayable"],
    "lease_total": ["ifrs-full:LeaseLiabilities", "+ifrs-full:CurrentLeaseLiabilities", "+ifrs-full:NoncurrentLeaseLiabilities"],
    "minority": ["ifrs-full:NoncontrollingInterests"],
    "equity": ["ifrs-full:EquityAttributableToOwnersOfParent"],
    "share_cap": ["tifrs-bsci-ci:OrdinaryShare", "ifrs-full:IssuedCapital"],
}


def tw_get(facts, names, ctx):
    """First plain name wins; '+name' entries are added on top (for split
    line items).  None when nothing matched."""
    total, found = 0.0, False
    for n in names:
        add = n.startswith("+")
        key = (n[1:] if add else n, ctx)
        if key in facts:
            v = facts[key]
            if add:
                total += v
                found = True
            elif not found:
                total = v
                found = True
    return total if found else None


def tw_flow_val(facts, k, s, e):
    names = TW_FLOW[k]
    if k == "da":
        a = tw_get(facts, [names[0]], ctx_dur(s, e))
        b = tw_get(facts, [names[1][1:]], ctx_dur(s, e))
        if a is None and b is None:
            return None
        return (a or 0) + (b or 0)
    return tw_get(facts, names, ctx_dur(s, e))


def latest_tw_period(today=None):
    """Newest quarter whose statutory deadline has passed (5/15, 8/14,
    11/14, 3/31)."""
    today = today or date.today()
    y = today.year
    cands = [(y - 1, 4, date(y, 3, 31)), (y, 1, date(y, 5, 15)), (y, 2, date(y, 8, 14)), (y, 3, date(y, 11, 14))]
    avail = [(yy, q) for yy, q, dl in cands if today >= dl]
    return avail[-1] if avail else (y - 1, 3)


def prev_q(y, q):
    return (y - 1, 4) if q == 1 else (y, q - 1)


def load_tw(co_id, history=True):
    co = {"ticker": co_id, "market": "TW", "name": co_id, "unit": "TWD"}
    y, q = latest_tw_period()
    reports = {}      # (y,q) -> facts

    def rep(yy, qq):
        if (yy, qq) not in reports:
            html = mops_report(co_id, yy, qq)
            reports[(yy, qq)] = parse_ix(html) if html else None
            if html and co["name"] == co_id:
                m = NONNUM_RE.search(html)
                if m:
                    co["name"] = m.group(1).strip()
        return reports[(yy, qq)]

    # walk back until the newest available report
    tries = 0
    while rep(y, q) is None and tries < 3:
        y, q = prev_q(y, q)
        tries += 1
    if rep(y, q) is None:
        raise RuntimeError("MOPS has no report for %s" % co_id)
    co["latest"] = (y, q)

    # quarterly history: the 3-month contexts, Q4 = FY - Q3 YTD
    periods = []
    yy, qq = y, q
    for _ in range((HIST_QUARTERS + 1) if history else 1):
        f = rep(yy, qq)
        if f is None:
            break
        qs, qe = q_bounds(yy, qq)
        row = {"end": qe}
        for k in TW_FLOW:
            if qq == 4:
                fy_v = tw_flow_val(f, k, date(yy, 1, 1), qe)
                f3 = rep(yy, 3)
                ytd3 = tw_flow_val(f3, k, date(yy, 1, 1), date(yy, 9, 30)) if f3 else None
                row[k] = (fy_v - ytd3) if (num(fy_v) and num(ytd3)) else None
                row[k + "_ytd"] = fy_v
            else:
                row[k] = tw_flow_val(f, k, qs, qe)
                row[k + "_ytd"] = tw_flow_val(f, k, date(yy, 1, 1), qe)
                if row[k] is None and qq == 1:
                    row[k] = row[k + "_ytd"]
        for k, names in TW_INST.items():
            row[k] = tw_get(f, names, ctx_inst(qe))
        row["st_debt"] = row.get("st_borrow")
        row["st_borrow_for_ic"] = row.get("st_borrow")
        periods.append(row)
        yy, qq = prev_q(yy, qq)
    periods.reverse()
    co["periods"] = periods
    # TTM flows via YTD arithmetic (works for the cash-flow items that only
    # carry YTD contexts in Q2/Q3 reports)
    cur = rep(y, q)
    ttm = {}
    if q == 4:
        for k in TW_FLOW:
            ttm[k] = tw_flow_val(cur, k, date(y, 1, 1), date(y, 12, 31))
    else:
        py = rep(y - 1, 4)
        pq = rep(y - 1, q)
        qs, qe = q_bounds(y, q)
        pqs, pqe = q_bounds(y - 1, q)
        for k in TW_FLOW:
            a = tw_flow_val(cur, k, date(y, 1, 1), qe)
            b = tw_flow_val(py, k, date(y - 1, 1, 1), date(y - 1, 12, 31)) if py else None
            # the current report carries the prior-year comparative YTD too
            c = tw_flow_val(cur, k, date(y - 1, 1, 1), pqe)
            if c is None and pq:
                c = tw_flow_val(pq, k, date(y - 1, 1, 1), pqe)
            ttm[k] = (a + b - c) if (num(a) and num(b) and num(c)) else None
    co["ttm_override"] = ttm
    last = periods[-1]
    sc = last.get("share_cap")
    co["shares"] = sc / 10.0 if num(sc) else None
    co["shares_note"] = "ordinary share capital / NT$10 par (MOPS balance sheet)"
    co["src"] = "MOPS t164sb01 %dQ%d" % (y, q)
    return co


# ----------------------------------------------------------------------------
# metrics
# ----------------------------------------------------------------------------
def ic_readings(row):
    ic = None
    if num(row.get("assets")) and num(row.get("cur_liab")):
        ic = row["assets"] - row["cur_liab"]
    ic_ex = None
    if num(ic):
        ic_ex = ic - (row.get("cash") or 0) - (row.get("st_inv") or 0)
        if ic_ex <= 0.1 * ic:
            ic_ex = None          # cash (nearly) equals capital employed: the reading is meaningless
    ic_std = None
    if num(ic):
        ic_std = ic + (row.get("st_borrow_for_ic") or 0)
    return ic, ic_ex, ic_std


def tax_rate(tax, pretax, market):
    r = div(tax, pretax)
    if num(r) and 0.0 <= r <= 0.40:
        return r, "TTM"
    return TAX_FALLBACK[market], "fallback"


def sum_last(periods, k, n=4):
    vals = [p.get(k) for p in periods[-n:]]
    if len(vals) < n or any(not num(v) for v in vals):
        return None
    return sum(vals)


def compute_company(co, price):
    m = co["market"]
    periods = co["periods"]
    last = periods[-1]
    ttm = {}
    override = co.get("ttm_override", {})
    for k in ("ebit", "revenue", "pretax", "tax", "da", "capex"):
        v = override.get(k)
        if not num(v):
            v = sum_last(periods, k)
        ttm[k] = v
    t, t_src = tax_rate(ttm["tax"], ttm["pretax"], m)
    ic, ic_ex, ic_std = ic_readings(last)
    nopat = ttm["ebit"] * (1 - t) if num(ttm["ebit"]) else None
    out = {
        "ticker": co["ticker"], "name": co["name"], "market": m, "unit": co["unit"],
        "period_end": last["end"].isoformat(), "src": co["src"],
        "ebit_ttm": ttm["ebit"], "revenue_ttm": ttm["revenue"], "tax_rate": t, "tax_src": t_src,
        "nopat": nopat, "ic": ic, "ic_excash": ic_ex, "ic_std": ic_std,
        "roic": div(nopat, ic), "roic_excash": div(nopat, ic_ex), "roic_std": div(nopat, ic_std),
        "margin": div(nopat, ttm["revenue"]), "turnover": div(ttm["revenue"], ic),
        "da": ttm["da"], "capex": ttm["capex"],
        "equity": last.get("equity"), "shares": co.get("shares"), "shares_note": co.get("shares_note", ""), "price": price,
    }
    capex = abs(ttm["capex"]) if num(ttm["capex"]) else None
    reinv = None
    if num(capex) and num(ttm["da"]) and num(nopat) and nopat > 0:
        reinv = (capex - ttm["da"]) / nopat
    out["reinvest"] = reinv
    out["g_internal"] = out["roic"] * reinv if (num(out["roic"]) and num(reinv)) else None
    # enterprise value
    debt = (last.get("st_debt") or 0) + (last.get("lt_debt") or 0)
    lease = last.get("lease_total") or 0
    minority = last.get("minority") or 0
    cash = (last.get("cash") or 0) + (last.get("st_inv") or 0)
    mcap = price * co["shares"] if (num(price) and num(co.get("shares"))) else None
    out["mcap"] = mcap
    out["net_debt_adj"] = debt + lease + minority - cash
    out["ev"] = (mcap + out["net_debt_adj"]) if num(mcap) else None
    out["ev_ebit"] = div(out["ev"], ttm["ebit"]) if (num(ttm["ebit"]) and ttm["ebit"] > 0) else None
    out["pb"] = div(mcap, last.get("equity"))
    out["pe"] = None
    return out


def history_rows(co):
    """Target only: quarterly x4 ROIC (Excel column convention) and rolling
    TTM ROIC, plus the margin x turnover split."""
    rows = []
    periods = co["periods"]
    m = co["market"]
    for i, p in enumerate(periods):
        ic, ic_ex, ic_std = ic_readings(p)
        t, _ = tax_rate(p.get("tax"), p.get("pretax"), m)
        ebit = p.get("ebit")
        nopat_q = ebit * (1 - t) if num(ebit) else None
        roic_q = div(nopat_q * 4, ic) if num(nopat_q) else None
        ttm_ebit = sum_last(periods[: i + 1], "ebit") if i >= 3 else None
        ttm_rev = sum_last(periods[: i + 1], "revenue") if i >= 3 else None
        ttm_tax, _ = tax_rate(sum_last(periods[: i + 1], "tax") if i >= 3 else None,
                              sum_last(periods[: i + 1], "pretax") if i >= 3 else None, m)
        nopat_ttm = ttm_ebit * (1 - ttm_tax) if num(ttm_ebit) else None
        rows.append({
            "period": p["end"].isoformat(),
            "ebit_q": ebit, "nopat_q": nopat_q, "ic": ic, "ic_excash": ic_ex, "ic_std": ic_std,
            "roic_q4x": roic_q,
            "roic_ttm": div(nopat_ttm, ic),
            "roic_ttm_excash": div(nopat_ttm, ic_ex),
            "roic_ttm_std": div(nopat_ttm, ic_std),
            "margin_ttm": div(nopat_ttm, ttm_rev),
            "turnover_ttm": div(ttm_rev, ic),
        })
    return rows[-HIST_QUARTERS:]


# ----------------------------------------------------------------------------
# peers, regression, fair value
# ----------------------------------------------------------------------------
def percentile(sorted_vals, p):
    if not sorted_vals:
        return None
    if len(sorted_vals) == 1:
        return sorted_vals[0]
    k = (len(sorted_vals) - 1) * p
    f, c = math.floor(k), math.ceil(k)
    if f == c:
        return sorted_vals[int(k)]
    return sorted_vals[f] + (sorted_vals[c] - sorted_vals[f]) * (k - f)


def pct_rank(vals, x):
    vals = [v for v in vals if num(v)]
    if not vals or not num(x):
        return None
    below = sum(1 for v in vals if v < x)
    return below / len(vals)


def ols(xs, ys):
    n = len(xs)
    if n < 3:
        return None
    mx, my = sum(xs) / n, sum(ys) / n
    sxx = sum((x - mx) ** 2 for x in xs)
    if sxx == 0:
        return None
    b = sum((x - mx) * (y - my) for x, y in zip(xs, ys)) / sxx
    a = my - b * mx
    ss_tot = sum((y - my) ** 2 for y in ys)
    ss_res = sum((y - (a + b * x)) ** 2 for x, y in zip(xs, ys))
    r2 = 1 - ss_res / ss_tot if ss_tot else 0.0
    return a, b, r2


def build(target, market, peers, group):
    market = market.upper()
    loader = load_tw if market == "TW" else load_us
    tickers = [target] + [p for p in peers if p.upper() != target.upper()]
    companies = []
    for i, tk in enumerate(tickers):
        try:
            co = loader(tk, history=(i == 0))
            px, sym = price_for(tk, market)
            if not num(px):
                log("%s: no price" % tk)
            row = compute_company(co, px)
            row["sym"] = sym
            companies.append(row)
            log("%s ok  ROIC=%s EV/EBIT=%s" % (tk, fmt(row["roic"], "pct"), fmt(row["ev_ebit"])))
        except Exception as e:  # noqa
            log("%s skipped: %s" % (tk, e))
            if tk.upper() == target.upper():
                raise
    tgt = companies[0]
    peers_ok = [c for c in companies[1:] if num(c["ev_ebit"]) and 0 < c["ev_ebit"] < 100 and num(c["roic"])]
    # ---- relative position
    roic_all = [c["roic"] for c in companies if num(c["roic"])]
    tgt["roic_pct_rank"] = pct_rank(roic_all, tgt["roic"])
    mult = sorted(c["ev_ebit"] for c in peers_ok)
    p25, p50, p75 = percentile(mult, 0.25), percentile(mult, 0.5), percentile(mult, 0.75)
    reg = ols([c["roic"] for c in peers_ok], [c["ev_ebit"] for c in peers_ok])
    pred = None
    if reg and num(tgt["roic"]):
        pred = reg[0] + reg[1] * tgt["roic"]
    premium = (tgt["ev_ebit"] / pred - 1) if (num(pred) and pred > 0 and num(tgt["ev_ebit"]) and reg[1] > 0) else None
    # ---- fair value band (peer EV/EBIT percentile picked by the target's ROIC rank)
    rank = tgt["roic_pct_rank"]
    if num(rank):
        chosen = p75 if rank >= 0.75 else (p25 if rank < 0.25 else p50)
        chosen_lbl = "P75" if rank >= 0.75 else ("P25" if rank < 0.25 else "P50")
    else:
        chosen, chosen_lbl = p50, "P50"

    def px_from_mult(mu):
        if not (num(mu) and num(tgt["ebit_ttm"]) and tgt["ebit_ttm"] > 0 and num(tgt["shares"]) and tgt["shares"] > 0):
            return None
        eq = mu * tgt["ebit_ttm"] - tgt["net_debt_adj"]
        return eq / tgt["shares"]

    fair_lo, fair_mid, fair_hi, fair = px_from_mult(p25), px_from_mult(p50), px_from_mult(p75), px_from_mult(chosen)
    margin = (1 - tgt["price"] / fair) if (num(fair) and fair > 0 and num(tgt["price"])) else None
    # ---- history
    hist = history_rows(_last_loaded[target.upper()])
    floor = ROIC_FLOOR[market]
    roic_hist = [h["roic_ttm"] for h in hist if num(h["roic_ttm"])]
    hist_med = statistics.median(roic_hist[:-4]) if len(roic_hist) > 5 else (statistics.median(roic_hist) if roic_hist else None)
    recent = statistics.mean(roic_hist[-4:]) if len(roic_hist) >= 4 else (roic_hist[-1] if roic_hist else None)
    trend = ""
    if num(hist_med) and num(recent):
        d = recent - hist_med
        trend = "IMPROVING" if d > 0.02 else ("DETERIORATING" if d < -0.02 else "FLAT")
    return {
        "target": tgt, "peers": companies[1:], "peers_used": len(peers_ok), "group": group,
        "p25": p25, "p50": p50, "p75": p75, "reg": reg, "pred_mult": pred, "premium": premium,
        "chosen": chosen, "chosen_lbl": chosen_lbl,
        "fair_lo": fair_lo, "fair_mid": fair_mid, "fair_hi": fair_hi, "fair": fair, "margin": margin,
        "hist": hist, "floor": floor, "hist_med": hist_med, "recent": recent, "trend": trend,
    }


_last_loaded = {}
_orig_load_us, _orig_load_tw = load_us, load_tw


def load_us(ticker, history=True):  # noqa: F811  (remember the loaded target for the history block)
    co = _orig_load_us(ticker, history)
    if history:
        _last_loaded[ticker.upper()] = co
    return co


def load_tw(ticker, history=True):  # noqa: F811
    co = _orig_load_tw(ticker, history)
    if history:
        _last_loaded[ticker.upper()] = co
    return co


# ----------------------------------------------------------------------------
# output
# ----------------------------------------------------------------------------
def write_out(res, path, elapsed):
    t = res["target"]
    lines = []

    def block(name):
        lines.append("#" + name)

    def row(*cells):
        lines.append("\t".join("" if c is None else str(c) for c in cells))

    block("META")
    row("ticker", t["ticker"])
    row("name", t["name"])
    row("market", t["market"])
    row("unit", t["unit"])
    row("group", res["group"])
    row("period_end", t["period_end"])
    row("src", t["src"])
    row("price", fmt(t["price"]))
    row("sym", t.get("sym", ""))
    row("shares", fmt(t["shares"]))
    row("built", datetime.now().strftime("%Y-%m-%d %H:%M"))
    row("elapsed_s", "%.0f" % elapsed)
    row("peers_used", res["peers_used"])
    row("floor", fmt(res["floor"], "pct"))

    block("SUMMARY")     # label, value, fmt, note
    row("ROIC (modValuation)", fmt(t["roic"], "pct"), "pct", "NOPAT TTM / (assets - current liabilities)")
    row("ROIC ex-cash", fmt(t["roic_excash"], "pct"), "pct", "IC less cash and short-term investments")
    row("ROIC std", fmt(t["roic_std"], "pct"), "pct", "IC plus short-term borrowings")
    row("Absolute floor", fmt(res["floor"], "pct"), "pct", "US 10% / TW 8% - fixed, do not tune per ticker")
    row("History median", fmt(res["hist_med"], "pct"), "pct", "TTM ROIC, quarters before the last four")
    row("Last 4Q avg", fmt(res["recent"], "pct"), "pct", res["trend"])
    row("Peer pct rank", fmt(t["roic_pct_rank"], "pct"), "pct", "share of group (incl. target) with lower ROIC")
    row("NOPAT margin", fmt(t["margin"], "pct"), "pct", "NOPAT / revenue TTM")
    row("IC turnover", fmt(t["turnover"]), "x", "revenue TTM / IC")
    row("Reinvestment rate", fmt(t["reinvest"], "pct"), "pct", "(capex - D&A) / NOPAT  (working capital not included)")
    row("Internal growth", fmt(t["g_internal"], "pct"), "pct", "ROIC x reinvestment rate")
    row("Tax rate", fmt(t["tax_rate"], "pct"), "pct", t["tax_src"])
    row("EBIT TTM", fmt(t["ebit_ttm"]), "num", "")
    row("NOPAT TTM", fmt(t["nopat"]), "num", "")
    row("Invested capital", fmt(t["ic"]), "num", "")
    row("Market cap", fmt(t["mcap"]), "num", t.get("shares_note", ""))
    row("EV", fmt(t["ev"]), "num", "mcap + debt + leases + minority - cash & ST inv")
    row("EV/EBIT", fmt(t["ev_ebit"]), "x", "")
    row("P/B", fmt(t["pb"]), "x", "")

    block("RELATIVE")
    row("Peer EV/EBIT P25", fmt(res["p25"]), "x", "")
    row("Peer EV/EBIT P50", fmt(res["p50"]), "x", "")
    row("Peer EV/EBIT P75", fmt(res["p75"]), "x", "")
    reg = res["reg"]
    if reg:
        warn = "" if reg[1] > 0 else " - NEGATIVE SLOPE, multiples do not track ROIC in this group; ignore the premium line"
        row("Regression", "EV/EBIT = %.2f + %.2f x ROIC" % (reg[0], reg[1]), "txt", "R2 = %.2f, n = %d%s" % (reg[2], res["peers_used"], warn))
    else:
        row("Regression", "n/a", "txt", "fewer than 3 usable peers")
    row("Predicted EV/EBIT", fmt(res["pred_mult"]), "x", "from the peer regression at the target's ROIC")
    row("Premium / discount", fmt(res["premium"], "pct"), "pct", "actual / predicted - 1")

    block("FAIR")
    row("Multiple used", res["chosen_lbl"], "txt", "peer percentile picked by ROIC rank (>=75% -> P75, <25% -> P25, else P50)")
    row("Fair price (P25)", fmt(res["fair_lo"]), "px", "")
    row("Fair price (P50)", fmt(res["fair_mid"]), "px", "")
    row("Fair price (P75)", fmt(res["fair_hi"]), "px", "")
    row("Fair price (chosen)", fmt(res["fair"]), "px", "")
    row("Current price", fmt(t["price"]), "px", t.get("sym", ""))
    row("Safety margin", fmt(res["margin"], "pct"), "pct", "1 - price / fair; need >= 25% (cyclicals >= 35%)")

    block("HISTORY")
    row("period", "ebit_q", "nopat_q", "ic", "roic_q4x", "roic_ttm", "roic_ttm_excash", "roic_ttm_std", "margin_ttm", "turnover_ttm")
    for h in res["hist"]:
        row(h["period"], fmt(h["ebit_q"]), fmt(h["nopat_q"]), fmt(h["ic"]), fmt(h["roic_q4x"], "pct"),
            fmt(h["roic_ttm"], "pct"), fmt(h["roic_ttm_excash"], "pct"), fmt(h["roic_ttm_std"], "pct"),
            fmt(h["margin_ttm"], "pct"), fmt(h["turnover_ttm"]))

    block("PEERS")
    row("ticker", "name", "period_end", "roic", "ev_ebit", "pb", "margin", "turnover", "mcap", "ev", "ebit_ttm", "ic", "note")
    for c in [t] + res["peers"]:
        note = ""
        if not num(c["ev_ebit"]):
            note = "no EV/EBIT (negative EBIT or no price)"
        elif c["ev_ebit"] >= 100:
            note = "excluded: EV/EBIT >= 100"
        if c["tax_src"] == "fallback":
            note = (note + "; " if note else "") + "tax fallback"
        row(c["ticker"], c["name"], c["period_end"], fmt(c["roic"], "pct"), fmt(c["ev_ebit"]), fmt(c["pb"]),
            fmt(c["margin"], "pct"), fmt(c["turnover"]), fmt(c["mcap"]), fmt(c["ev"]), fmt(c["ebit_ttm"]), fmt(c["ic"]), note)

    block("THESIS")
    rk = t["roic_pct_rank"]
    lines.append("[4] Three numbers")
    lines.append("  1. ROIC %s (most conservative of the three readings: %s / %s / %s), peer rank %s" % (
        pct(min_ok([t["roic"], t["roic_excash"], t["roic_std"]])), pct(t["roic"]), pct(t["roic_excash"]), pct(t["roic_std"]), pct(rk)))
    lines.append("  2. ROIC split: NOPAT margin %s x IC turnover %s; last-4Q avg %s vs history median %s (%s)" % (
        pct(t["margin"]), xval(t["turnover"]), pct(res["recent"]), pct(res["hist_med"]), res["trend"] or "n/a"))
    lines.append("  3. EV/EBIT %s vs peer regression %s -> %s premium/discount" % (
        xval(t["ev_ebit"]), xval(res["pred_mult"]), pct(res["premium"])))
    lines.append("[5] Falsification thresholds")
    lines.append("  - TTM ROIC below the %s floor for 2 consecutive quarters, or peer rank falling under 50%%" % pct(res["floor"]))
    lines.append("  - IC turnover down 3 quarters running while NOPAT margin does not rise")
    lines.append("  - revenue momentum more than 10pp under internal growth (%s)" % pct(t["g_internal"]))
    lines.append("[6] Decision")
    lines.append("  Fair band %s - %s (P25-P75 peer EV/EBIT x EBIT TTM), chosen %s = %s, price %s, safety margin %s" % (
        pxv(res["fair_lo"]), pxv(res["fair_hi"]), res["chosen_lbl"], pxv(res["fair"]), pxv(t["price"]), pct(res["margin"])))

    block("LOG")
    for m in LOG:
        lines.append(m)
    with open(path, "w", encoding="utf-8", newline="\n") as f:
        f.write("\n".join(lines) + "\n")


def min_ok(vals):
    vals = [v for v in vals if num(v)]
    return min(vals) if vals else None


def pct(v):
    return ("%.1f%%" % (v * 100)) if num(v) else "n/a"


def xval(v):
    return ("%.1fx" % v) if num(v) else "n/a"


def pxv(v):
    return ("%.2f" % v) if num(v) else "n/a"


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--req", help="request json written by VBA: {ticker, market, peers[], group}")
    ap.add_argument("--ticker")
    ap.add_argument("--market", default="")
    ap.add_argument("--peers", default="")
    ap.add_argument("--group", default="")
    ap.add_argument("--out", required=True)
    a = ap.parse_args()
    if a.req:
        with open(a.req, encoding="utf-8") as f:
            req = json.load(f)
        ticker, market = req["ticker"], req.get("market", "")
        peers, group = req.get("peers", []), req.get("group", "")
    else:
        ticker, market, group = a.ticker, a.market, a.group
        peers = [p.strip() for p in a.peers.split(",") if p.strip()]
    ticker = ticker.strip().upper()
    for suf in (".TWO", ".TW"):
        if ticker.endswith(suf):
            ticker, market = ticker[: -len(suf)], "TW"
    peers = [p.upper().replace(".TWO", "").replace(".TW", "") if market.upper() == "TW" else p.upper() for p in peers]
    if not market:
        market = "TW" if ticker.isdigit() else "US"
    t0 = time.time()
    try:
        res = build(ticker, market, peers, group)
        write_out(res, a.out, time.time() - t0)
    except Exception as e:  # noqa
        with open(a.out, "w", encoding="utf-8") as f:
            f.write("#ERROR\n%s\n#LOG\n%s\n" % (e, "\n".join(LOG)))
        raise


if __name__ == "__main__":
    main()
