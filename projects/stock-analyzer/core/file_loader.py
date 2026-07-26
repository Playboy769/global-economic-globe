"""
file_loader.py — 讀取 CSV / XLSX / XLSM
修正 V1 問題：
  - XLSX 外部連結 KeyError → 改用 pandas.ExcelFile（完全繞過 openpyxl workbook 層）
  - 多 Sheet 全部讀入並回傳
"""
import os
import csv as _csv
import chardet
import pandas as pd


def load_file(path: str) -> dict:
    """
    回傳:
    {
      "meta": { filename, rows, cols, encoding, sheets, date, strategy, file_type },
      "df":   pd.DataFrame,           # 主 sheet（第一個）
      "sheets": { name: df, ... }     # 所有 sheet
    }
    """
    ext = os.path.splitext(path)[1].lower()
    if ext == ".csv":
        return _load_csv(path)
    elif ext in (".xlsx", ".xlsm"):
        return _load_excel(path)
    else:
        raise ValueError(f"不支援的格式：{ext}")


# ── CSV ───────────────────────────────────────────────────────────────────────

_TRY_ENC = ["big5", "cp950", "gb18030", "utf-8-sig", "utf-8"]


def _detect_encoding(path: str) -> str:
    with open(path, "rb") as f:
        raw = f.read()
    for enc in _TRY_ENC:
        try:
            raw.decode(enc)
            return enc
        except (UnicodeDecodeError, LookupError):
            continue
    return chardet.detect(raw[:8192]).get("encoding") or "utf-8"


def _load_csv(path: str) -> dict:
    enc = _detect_encoding(path)
    with open(path, "rb") as f:
        text = f.read().decode(enc, errors="replace")
    lines = text.splitlines()

    header_idx = _find_header_row(lines)
    date_str = strategy_str = ""
    for ml in lines[:header_idx]:
        if "年" in ml or "日" in ml:
            date_str = ml.strip()
        if "策略" in ml:
            strategy_str = ml.strip()

    header_fields = next(_csv.reader([lines[header_idx]]))
    n_cols = len(header_fields)

    rows = []
    for raw_line in lines[header_idx + 1:]:
        if not raw_line.strip():
            continue
        try:
            fields = next(_csv.reader([raw_line]))
        except Exception:
            continue
        if len(fields) < 2:
            continue
        if len(fields) > n_cols:
            fields = fields[:n_cols - 1] + [",".join(fields[n_cols - 1:])]
        elif len(fields) < n_cols:
            fields += [""] * (n_cols - len(fields))
        rows.append(fields)

    df = _clean_df(pd.DataFrame(rows, columns=header_fields))
    sheet_name = os.path.splitext(os.path.basename(path))[0]
    meta = {
        "filename": os.path.basename(path),
        "rows": len(df), "cols": len(df.columns),
        "encoding": enc, "sheets": [sheet_name],
        "date": date_str, "strategy": strategy_str,
        "file_type": "csv",
    }
    return {"meta": meta, "df": df, "sheets": {sheet_name: df}}


def _find_header_row(lines: list) -> int:
    keywords = {"排名", "代碼", "代號", "商品", "成交", "漲幅", "close", "open", "symbol"}
    for i, line in enumerate(lines):
        try:
            fields = next(_csv.reader([line]))
        except Exception:
            continue
        if len(fields) >= 5 and any(k in "".join(fields).lower() for k in keywords):
            return i
    return 3


# ── Excel / XLSM ─────────────────────────────────────────────────────────────

def _load_excel(path: str) -> dict:
    """
    使用 pandas.ExcelFile 讀取，完全繞過 openpyxl 的外部連結解析，
    避免 xlsm/xlsx 含有失效外部連結時的 KeyError。
    """
    ext = os.path.splitext(path)[1].lower().lstrip(".")

    # pandas ExcelFile 不會解析外部連結，安全讀取
    xl = pd.ExcelFile(path, engine="openpyxl")
    sheet_names = xl.sheet_names

    sheets = {}
    for name in sheet_names:
        try:
            df = xl.parse(name)
            df = _clean_df(df)
            if not df.empty:
                sheets[name] = df
        except Exception:
            continue

    if not sheets:
        raise ValueError("檔案內沒有可讀取的工作表")

    main = list(sheets.keys())[0]
    df = sheets[main]

    meta = {
        "filename": os.path.basename(path),
        "rows": len(df), "cols": len(df.columns),
        "encoding": "N/A", "sheets": list(sheets.keys()),
        "date": "", "strategy": "",
        "file_type": ext,
    }
    return {"meta": meta, "df": df, "sheets": sheets}


# ── 共用清理 ──────────────────────────────────────────────────────────────────

def _clean_df(df: pd.DataFrame) -> pd.DataFrame:
    df = df.dropna(how="all").dropna(axis=1, how="all").reset_index(drop=True)
    # 嘗試轉數值（超過 60% 成功才改型態）
    for col in df.columns:
        try:
            converted = pd.to_numeric(df[col], errors="coerce")
            if converted.notna().mean() > 0.6:
                df[col] = converted
        except Exception:
            pass
    return df
