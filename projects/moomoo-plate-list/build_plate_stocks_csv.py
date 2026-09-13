"""
把瀏覽器抓下來的板塊成分股原始資料（mmscrape.json）整理成 moomoo_us_plate_stocks.csv。

mmscrape.json 的來源（2026-09-13）：用 Claude Code 內建瀏覽器逐一開
https://www.moomoo.com/sectors/<plateCode>，讀頁面表格（含翻頁）存進 localStorage，
最後經剪貼簿落地。格式 {"BK2072": "<tab 分隔列>\n<...>", ...}，每列 27 欄，欄位順序同網頁表頭。
頁面當時被轉到繁中版，所以名稱是中文、數量用「萬／億」單位；本腳本把數量欄換算成純數字。

用法:
    python build_plate_stocks_csv.py
輸出:
    moomoo_us_plate_stocks.csv —— 每列一檔成分股，前面帶板塊資訊與依市值排名的 rank。
"""

import csv
import json
import re

RAW_PATH = "mmscrape.json"
PLATE_LIST_PATH = "moomoo_us_plate_list.csv"
OUT_PATH = "moomoo_us_plate_stocks.csv"

# 網頁表頭順序（前 27 欄；之後的 Industry / Watchlist / Paper Trade 沒抓）
COLS = [
    "symbol", "name", "price", "chg", "chgPct", "volume", "turnover", "open", "prevClose",
    "high", "low", "marketCap", "floatCap", "shares", "sharesFloat",
    "chg5D", "chg10D", "chg20D", "chg60D", "chg120D", "chg250D", "chgYTD",
    "divYieldTTM", "turnoverRatio", "peTTM", "peStatic", "rangePct",
]
# 這些欄位要把「196.14萬 / 1.93億 / 8.97M / 590.64B」換算成純數字
QTY_COLS = ["volume", "turnover", "marketCap", "floatCap", "shares", "sharesFloat"]
PCT_COLS = ["chgPct", "chg5D", "chg10D", "chg20D", "chg60D", "chg120D", "chg250D", "chgYTD",
            "divYieldTTM", "turnoverRatio", "rangePct"]

UNITS = {"K": 1e3, "M": 1e6, "B": 1e9, "T": 1e12, "萬": 1e4, "億": 1e8, "兆": 1e12}


def to_number(s):
    """'196.14萬' → 1961400.0；'+7.16%' → 7.16；'--' / 'Loss' / '虧損' → None。"""
    if s is None:
        return None
    s = str(s).strip().replace(",", "")
    if s in ("", "--", "-", "Loss", "虧損", "亏损"):
        return None
    m = re.fullmatch(r"([+-]?\d+(?:\.\d+)?)\s*([KMBT萬億兆]*)%?", s)
    if not m:
        return None
    val = float(m.group(1))
    for u in m.group(2):  # 「萬億」= 1e4 × 1e8 = 1e12，逐字相乘即可
        val *= UNITS[u]
    return val


def main():
    with open(RAW_PATH, encoding="utf-8") as f:
        raw = json.load(f)
    with open(PLATE_LIST_PATH, newline="", encoding="utf-8-sig") as f:
        plates = {p["plateCode"]: p for p in csv.DictReader(f)}

    out_cols = ["plateType", "plateCode", "plateName", "plateId", "rank"] + COLS + \
               [c + "Num" for c in QTY_COLS] + [c + "Num" for c in PCT_COLS] + ["peTTMNum"]

    rows = []
    for code, blob in raw.items():
        p = plates.get(code, {})
        items = []
        for line in blob.split("\n"):
            if not line.strip():
                continue
            cells = line.split("\t")
            cells += [""] * (len(COLS) - len(cells))
            item = dict(zip(COLS, cells[:len(COLS)]))
            for c in QTY_COLS + PCT_COLS:
                item[c + "Num"] = to_number(item[c])
            item["peTTMNum"] = to_number(item["peTTM"])
            items.append(item)
        # 依市值由大到小排名；市值缺的排最後
        items.sort(key=lambda it: it["marketCapNum"] if it["marketCapNum"] is not None else -1, reverse=True)
        for rank, it in enumerate(items, 1):
            rows.append({
                "plateType": p.get("plateType", ""), "plateCode": code,
                "plateName": p.get("plateName", ""), "plateId": p.get("plateId", ""),
                "rank": rank, **it,
            })

    with open(OUT_PATH, "w", newline="", encoding="utf-8-sig") as f:
        w = csv.DictWriter(f, fieldnames=out_cols)
        w.writeheader()
        w.writerows(rows)
    print(f"{len(raw)} 個板塊、{len(rows)} 檔，已存成 {OUT_PATH}")


if __name__ == "__main__":
    main()
