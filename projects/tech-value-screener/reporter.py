"""
reporter.py — 輸出格式化
  - 終端機：使用 rich 套件顯示彩色排名表
  - Excel：多工作表詳細報告（openpyxl）
"""
import logging
from datetime import datetime
from pathlib import Path
from typing import Optional

import pandas as pd

logger = logging.getLogger(__name__)

try:
    from rich.console import Console
    from rich.table import Table
    from rich.text import Text
    from rich import box as rbox
    _RICH = True
    _console = Console()
except ImportError:
    _RICH = False
    _console = None


# ─── 格式化小工具 ──────────────────────────────────────────────────────────────
def _pct(v, d=1) -> str:
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return "N/A"
    return f"{v:.{d}f}%"

def _x(v, d=1) -> str:
    if v is None or (isinstance(v, float) and pd.isna(v)) or v <= 0:
        return "N/A"
    return f"{v:.{d}f}x"

def _b(v) -> str:
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return "N/A"
    if abs(v) >= 1000:
        return f"${v/1000:.1f}T"
    return f"${v:.1f}B"

def _color(score: float) -> str:
    return "bright_green" if score >= 70 else ("yellow" if score >= 55 else "red")


# ─── 終端機輸出 ────────────────────────────────────────────────────────────────
def print_console(df: pd.DataFrame, top: int = 25):
    if df.empty:
        print("無符合條件的股票。")
        return

    now = datetime.now().strftime("%Y-%m-%d %H:%M")
    header = f"  科技股潛在盈利篩選結果 ─ {now}  "
    print("\n" + "=" * len(header))
    print(header)
    print("=" * len(header))

    sub = df.head(top)

    if _RICH:
        _print_rich(sub)
    else:
        _print_plain(sub)

    _print_legend()


def _print_rich(df: pd.DataFrame):
    t = Table(box=rbox.ROUNDED, header_style="bold cyan", show_lines=False)
    t.add_column("#",        style="dim",         width=3,  justify="right")
    t.add_column("代碼",                           width=9,  justify="left")
    t.add_column("公司",                           width=20, justify="left")
    t.add_column("評分",                           width=5,  justify="right")
    t.add_column("市值",                           width=8,  justify="right")
    t.add_column("毛利%",                          width=7,  justify="right")
    t.add_column("GAAP利潤%",                      width=10, justify="right")
    t.add_column("調整利潤%▲",style="bold",        width=11, justify="right")
    t.add_column("研發%",                          width=7,  justify="right")
    t.add_column("FCF%",                           width=7,  justify="right")
    t.add_column("P/FCF",                          width=7,  justify="right")
    t.add_column("R40",                            width=6,  justify="right")

    for rank, (_, r) in enumerate(df.iterrows(), 1):
        sc = r.get("score", 0)
        c  = _color(sc)
        t.add_row(
            str(rank),
            Text(r["symbol"], style=f"bold {c}"),
            r["name"][:20],
            Text(f"{sc:.0f}", style=f"bold {c}"),
            _b(r.get("market_cap_b")),
            _pct(r.get("gross_margin")),
            _pct(r.get("op_margin")),
            Text(_pct(r.get("adj_op_margin_a")), style="bold green"),
            _pct(r.get("rd_pct")),
            _pct(r.get("fcf_margin")),
            _x(r.get("p_to_fcf")),
            f"{r.get('rule_of_40', 0):.0f}",
        )

    _console.print(t)


def _print_plain(df: pd.DataFrame):
    hdr = f"{'#':>3} {'代碼':>9} {'公司':20} {'評分':>5} {'市值':>8} {'毛利%':>7} {'調整利潤%':>9} {'研發%':>7} {'FCF%':>6} {'P/FCF':>6} {'R40':>5}"
    print(hdr)
    print("-" * len(hdr))
    for rank, (_, r) in enumerate(df.iterrows(), 1):
        print(
            f"{rank:>3} {r['symbol']:>9} {r['name'][:20]:20}"
            f" {r.get('score', 0):>5.0f}"
            f" {_b(r.get('market_cap_b')):>8}"
            f" {_pct(r.get('gross_margin')):>7}"
            f" {_pct(r.get('adj_op_margin_a')):>9}"
            f" {_pct(r.get('rd_pct')):>7}"
            f" {_pct(r.get('fcf_margin')):>6}"
            f" {_x(r.get('p_to_fcf')):>6}"
            f" {r.get('rule_of_40', 0):>5.0f}"
        )


def _print_legend():
    print("\n─── 指標說明 ───────────────────────────────────────────────────────────")
    print("  調整利潤%▲  = GAAP 營業利益率 + 全部 R&D 費用（揭露隱藏盈利上限）")
    print("  R40         = 營收成長% + FCF利潤%（SaaS黃金指標，>40優 >60卓越）")
    print("  P/FCF       = 市值 / 自由現金流（越低越便宜；<20x 屬低估）")
    print("  評分 ≥70 綠色 ■  55-69 黃色 ■  <55 紅色 ■")


# ─── Excel 匯出 ────────────────────────────────────────────────────────────────
def export_excel(df: pd.DataFrame, path: Path):
    try:
        with pd.ExcelWriter(path, engine="openpyxl") as writer:
            _sheet_ranking(df, writer)
            _sheet_detail(df, writer)
            _sheet_guide(writer)
        logger.info(f"Excel 已儲存：{path}")
    except Exception as e:
        logger.error(f"匯出 Excel 失敗：{e}")


def _sheet_ranking(df: pd.DataFrame, writer):
    col_map = {
        "symbol":           "股票代碼",
        "name":             "公司名稱",
        "score":            "綜合評分",
        "market_cap_b":     "市值(十億USD)",
        "revenue_growth":   "營收成長%",
        "gross_margin":     "毛利率%",
        "op_margin":        "GAAP營業利益率%",
        "adj_op_margin_a":  "調整後利益率%(全R&D)",
        "adj_op_margin_b":  "調整後利益率%(超額R&D)",
        "rd_pct":           "研發費用率%",
        "sm_pct":           "行銷費用率%",
        "fcf_margin":       "FCF利潤率%",
        "p_to_fcf":         "P/FCF倍數",
        "adj_pe_a":         "調整後PE(全R&D)",
        "adj_pe_b":         "調整後PE(超額R&D)",
        "rule_of_40":       "Rule of 40",
        "hidden_ratio":     "隱藏盈利比(倍)",
        "trailing_pe":      "GAAP本益比",
        "fcf_yield":        "FCF殖利率%",
    }
    avail = [c for c in col_map if c in df.columns]
    out = df[avail].rename(columns=col_map).reset_index(drop=True)
    out.index += 1
    out.to_excel(writer, sheet_name="篩選排名", index=True, index_label="排名")
    _autofit(writer.sheets["篩選排名"])


def _sheet_detail(df: pd.DataFrame, writer):
    col_map = {
        "symbol": "代碼", "name": "公司",
        "_rev":        "年營收(USD)",
        "_rd":         "研發費用(USD)",
        "_excess_rd":  "超額研發費用(USD)",
        "_sm":         "行銷費用(USD)",
        "_op_income":  "GAAP營業利益(USD)",
        "_net_income": "GAAP淨利(USD)",
        "_fcf":        "自由現金流(USD)",
        "_mktcap":     "市值(USD)",
    }
    avail = [c for c in col_map if c in df.columns]
    out = df[avail].rename(columns=col_map)
    out.to_excel(writer, sheet_name="財務明細", index=False)
    _autofit(writer.sheets["財務明細"])


def _sheet_guide(writer):
    rows = [
        ("評分架構（滿分100）", "隱藏盈利25 + 現金流品質20 + 成長性20 + 估值20 + 研發強度15"),
        ("", ""),
        ("毛利率%",           "高毛利=軟體/高值產品特徵；>60% 理想，>45% 可接受"),
        ("研發費用率%",        "研發投入越高，被壓低的盈利越多；>15% 為高研發"),
        ("GAAP營業利益率%",    "傳統會計利潤，受研發/行銷超額支出壓低"),
        ("調整後利益率% A",    "加回全部R&D費用後的利潤率（最樂觀上限）"),
        ("調整後利益率% B",    "只加回超出傳統企業3%基準的R&D（保守估計）"),
        ("FCF利潤率%",        "自由現金流/營收；>15% 優秀，是最真實的獲利衡量"),
        ("P/FCF倍數",         "市值/FCF；<15x 低估，15-30x 合理，>50x 昂貴"),
        ("調整後PE",          "市值/調整後淨利；比傳統PE更適合評估高研發科技股"),
        ("Rule of 40",       "=營收成長%+FCF利潤%；SaaS業黃金指標；>40優秀，>60卓越"),
        ("隱藏盈利比(倍)",     "調整後淨利/GAAP淨利；倍數越高代表隱藏越多（投資機會）"),
        ("FCF殖利率%",        "FCF/市值×100；>3% 視為具吸引力"),
        ("", ""),
        ("資料來源",           "Yahoo Finance (yfinance)，完全免費、不需API Key"),
        ("美股格式",           "直接輸入代碼：MSFT, GOOGL, NVDA ..."),
        ("台股格式",           "代碼後加 .TW：2330.TW, 2454.TW ..."),
        ("更新日期",           datetime.now().strftime("%Y-%m-%d")),
    ]
    pd.DataFrame(rows, columns=["指標", "說明"]).to_excel(
        writer, sheet_name="指標說明", index=False
    )
    ws = writer.sheets["指標說明"]
    ws.column_dimensions["A"].width = 22
    ws.column_dimensions["B"].width = 72


def _autofit(ws):
    for col in ws.columns:
        width = max((len(str(c.value or "")) for c in col), default=10)
        ws.column_dimensions[col[0].column_letter].width = min(width + 2, 40)
