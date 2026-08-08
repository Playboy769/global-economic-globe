"""Fetch an option chain from Yahoo Finance and emit raw IV quotes as JSON.

Deliberately does NOT interpolate: it exports the surviving quotes at their true
log-moneyness. The viewer does shape-preserving interpolation from these points,
which it cannot do if the strike axis has already been flattened onto a grid here.
"""

import argparse
import datetime as dt
import json
import math
import pathlib

import yfinance as yf

MONEYNESS_MIN = -0.35
MONEYNESS_MAX = 0.35

MIN_IV = 0.01
MAX_IV = 3.0
# A handful of quotes cannot pin down a smile; such an expiry only adds noise
# that the surface fit then has to reach across.
MIN_POINTS_PER_EXPIRY = 8
MIN_DTE = 4
# Minimum extrinsic value, as a fraction of spot, for an IV to mean anything.
MIN_TIME_VALUE = 0.002


def clean_quotes(frame, spot, side):
    """Rows -> [(log_moneyness, iv)] sorted, keeping only liquid, sane quotes."""
    points = {}
    for row in frame.itertuples():
        iv = float(row.impliedVolatility)
        strike = float(row.strike)
        if not (MIN_IV < iv < MAX_IV) or strike <= 0:
            continue
        # A zero-OI, zero-volume contract has a stale IV that distorts the surface.
        if (row.openInterest or 0) <= 0 and (row.volume or 0) <= 0:
            continue
        k = math.log(strike / spot)
        if not (MONEYNESS_MIN <= k <= MONEYNESS_MAX):
            continue

        bid, ask = float(row.bid or 0), float(row.ask or 0)
        # No bid means no one is willing to buy it at any price: not a market.
        if bid <= 0:
            continue
        mid = (bid + ask) / 2 if ask > 0 else float(row.lastPrice or 0)
        intrinsic = max(0.0, spot - strike) if side == "calls" else max(0.0, strike - spot)
        # Deep ITM options are almost entirely intrinsic value. Inverting a price
        # that is 99.9% intrinsic gives an IV driven by rounding, not by the
        # market's view of volatility -- these produce the fake near-dated spikes.
        if mid - intrinsic < MIN_TIME_VALUE * spot:
            continue

        # Same strike can appear twice; keep the more liquid quote.
        liquidity = (row.openInterest or 0) + (row.volume or 0)
        if k not in points or liquidity > points[k][1]:
            points[k] = (iv, liquidity)
    return sorted((k, v[0]) for k, v in points.items())


def build_surface(ticker, side, max_expiries):
    tk = yf.Ticker(ticker)
    spot = tk.fast_info.get("lastPrice")
    if not spot:
        raise SystemExit(f"no spot price available for {ticker}")
    spot = float(spot)

    expiries = list(tk.options)[:max_expiries]
    if not expiries:
        raise SystemExit(f"no listed options for {ticker}")

    today = dt.date.today()
    rows = []

    for expiry in expiries:
        chain = tk.option_chain(expiry)
        frame = chain.calls if side == "calls" else chain.puts
        points = clean_quotes(frame, spot, side)
        if len(points) < MIN_POINTS_PER_EXPIRY:
            continue
        dte = (dt.date.fromisoformat(expiry) - today).days
        # Options inside a few days of expiry are driven by pin and gamma effects
        # rather than by a view on volatility; their IV is erratic and drags the
        # near edge of the surface into noise.
        if dte < MIN_DTE:
            continue
        rows.append({
            "expiry": expiry,
            "dte": dte,
            "points": [[round(k, 6), round(iv, 6)] for k, iv in points],
        })

    if not rows:
        raise SystemExit(f"no usable {side} quotes for {ticker}")

    return {
        "ticker": ticker.upper(),
        "side": side,
        "spot": round(spot, 4),
        "fetched": today.isoformat(),
        "expiries": rows,
    }


def write_manifest(out_dir):
    """List tickers that have both sides on disk, so the viewer can discover them."""
    have = {}
    for p in out_dir.glob("*_calls.json"):
        t = p.name[:-len("_calls.json")]
        if (out_dir / f"{t}_puts.json").exists():
            try:
                have[t] = json.loads(p.read_text(encoding="utf-8"))["fetched"]
            except (ValueError, KeyError):
                pass
    manifest = {"tickers": sorted(have), "fetched": have}
    (out_dir / "index.json").write_text(json.dumps(manifest, indent=1), encoding="utf-8")
    return sorted(have)


def main():
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("ticker", nargs="+")
    ap.add_argument("--side", choices=["calls", "puts", "both"], default="both")
    ap.add_argument("--max-expiries", type=int, default=12)
    args = ap.parse_args()

    out_dir = pathlib.Path(__file__).parent / "data"
    out_dir.mkdir(exist_ok=True)
    sides = ["calls", "puts"] if args.side == "both" else [args.side]

    for ticker in args.ticker:
        for side in sides:
            try:
                surface = build_surface(ticker, side, args.max_expiries)
            except SystemExit as e:
                print(f"{ticker} {side}: skipped ({e})")
                continue
            except Exception as e:                      # noqa: BLE001
                print(f"{ticker} {side}: failed ({type(e).__name__}: {e})")
                continue
            path = out_dir / f"{surface['ticker']}_{side}.json"
            path.write_text(json.dumps(surface, indent=1), encoding="utf-8")
            total = sum(len(e["points"]) for e in surface["expiries"])
            print(f"{path.name}: spot {surface['spot']}, "
                  f"{len(surface['expiries'])} expiries, {total} quotes")

    print("manifest:", ", ".join(write_manifest(out_dir)))


if __name__ == "__main__":
    main()
