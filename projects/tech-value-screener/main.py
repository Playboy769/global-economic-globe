# -*- coding: utf-8 -*-
import sys, io, os
# Windows 終端機強制使用 UTF-8 輸出
if sys.platform == "win32":
    os.environ.setdefault("PYTHONUTF8", "1")
    try:
        sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")
        sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding="utf-8", errors="replace")
    except AttributeError:
        pass

"""
main.py — 科技股潛在盈利篩選系統 入口

用法範例：
  python main.py                          # 預設分析美股清單
  python main.py --market tw              # 分析台股清單
  python main.py --market all             # 美股 + 台股
  python main.py --tickers MSFT GOOGL NVDA  # 自訂股票清單
  python main.py --output console         # 只輸出終端機（不存Excel）
  python main.py --top 15 --min-score 60  # 顯示前15名、最低60分
"""
import argparse
import logging
import sys
from pathlib import Path

import pandas as pd

from config import US_STOCKS, TW_STOCKS
from fetcher import StockFetcher
import adjuster
import screener
import reporter

logging.basicConfig(
    level=logging.WARNING,          # 隱藏 yfinance 的除錯訊息
    format="%(levelname)s  %(message)s",
)
logger = logging.getLogger(__name__)


def parse_args():
    p = argparse.ArgumentParser(
        description="科技股潛在盈利篩選器｜揭露 GAAP 隱藏的真實獲利能力"
    )
    p.add_argument("--market",    choices=["us", "tw", "all"], default="us",
                   help="市場：us=美股  tw=台股  all=兩者")
    p.add_argument("--tickers",   nargs="+",
                   help="自訂股票（覆蓋 --market）；台股加 .TW 後綴")
    p.add_argument("--output",    choices=["console", "excel", "both"], default="both",
                   help="輸出方式")
    p.add_argument("--top",       type=int, default=25,
                   help="顯示前 N 名（預設 25）")
    p.add_argument("--min-score", type=float, default=0.0,
                   help="評分門檻（0-100，預設 0 = 全部顯示）")
    p.add_argument("--out-file",  default="tech_screener_results.xlsx",
                   help="Excel 輸出路徑")
    return p.parse_args()


def run_screen(symbols: list[str], label: str) -> list[dict]:
    """對股票清單執行抓取→調整→評分，回傳結果清單"""
    fetcher = StockFetcher()
    results = []
    total = len(symbols)

    print(f"\n▶ {label} — 分析 {total} 檔股票...")

    for i, sym in enumerate(symbols, 1):
        print(f"  [{i:>3}/{total}] {sym:<12}", end="\r", flush=True)
        raw = fetcher.fetch(sym)
        if not raw:
            continue

        metrics = adjuster.calculate(raw)
        if not metrics:
            continue

        ok, fails = screener.passes_filter(metrics)
        if not ok:
            logger.debug(f"[{sym}] 篩除：{' | '.join(fails)}")
            # 仍計算分數，但標記
            metrics["_filtered"] = True
        else:
            metrics["_filtered"] = False

        metrics["score"]  = screener.score(metrics)
        metrics["market"] = label
        results.append(metrics)

    print(f"  完成 {label}，取得 {len(results)} 筆資料{' ' * 20}")
    return results


def main():
    args = parse_args()

    print("=" * 60)
    print("  科技股潛在盈利篩選系統  |  Tech Value Screener")
    print("  資料來源：Yahoo Finance（免費，不需 API Key）")
    print("=" * 60)

    # ── 決定分析清單 ──────────────────────────────────────────────────────────
    all_data: list[dict] = []

    if args.tickers:
        all_data += run_screen(args.tickers, "自訂清單")
    else:
        if args.market in ("us", "all"):
            all_data += run_screen(US_STOCKS, "美股")
        if args.market in ("tw", "all"):
            all_data += run_screen(TW_STOCKS, "台股")

    if not all_data:
        print("\n❌ 未取得任何資料，請確認網路連線或股票代碼是否正確。")
        sys.exit(1)

    # ── 建立 DataFrame，篩選並排序 ──────────────────────────────────────────
    df = pd.DataFrame(all_data)
    df_pass = df[~df["_filtered"]].sort_values("score", ascending=False).reset_index(drop=True)
    df_all  = df.sort_values("score",     ascending=False).reset_index(drop=True)

    if args.min_score > 0:
        df_pass = df_pass[df_pass["score"] >= args.min_score]

    passed = len(df_pass)
    total  = len(df_all)
    print(f"\n✅ 分析完成：{total} 支有效資料，{passed} 支通過篩選條件")

    # ── 輸出 ──────────────────────────────────────────────────────────────────
    if args.output in ("console", "both"):
        reporter.print_console(df_pass, top=args.top)

    if args.output in ("excel", "both"):
        out = Path(args.out_file)
        reporter.export_excel(df_pass, out)
        print(f"\n📁 Excel 已儲存：{out.resolve()}")

    # ── 快速摘要 ──────────────────────────────────────────────────────────────
    if not df_pass.empty:
        top3 = df_pass.head(3)
        print("\n🏆 前三名：")
        for _, r in top3.iterrows():
            print(
                f"   {r['symbol']:>8}  {r['name'][:22]:22}"
                f"  評分:{r['score']:.0f}  "
                f"GAAP利潤:{r['op_margin']:.1f}%  "
                f"調整後利潤:{r['adj_op_margin_a']:.1f}%  "
                f"FCF:{r['fcf_margin']:.1f}%"
            )


if __name__ == "__main__":
    main()
