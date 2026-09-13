"""
依 moomoo_us_plate_list.csv 的每個板塊，抓成分股前 N 檔。
API 端點: https://www.moomoo.com/quote-api/quote-v2/get-plate-stock
  （2026-09-13 從瀏覽器 DevTools 抄下：
   get-plate-stock?marketType=2&plateId=10002072&page=0&pageSize=30&_=<毫秒時間戳>）
  伺服器沒有排序參數，只回預設順序；本腳本把整頁抓下來後在本機依市值欄位排序取前 N。

用法:
    python fetch_plate_stocks.py                 # 每個板塊前 10 檔
    python fetch_plate_stocks.py --top 20        # 每個板塊前 20 檔
    python fetch_plate_stocks.py --only industry # 只跑產業板塊
    python fetch_plate_stocks.py --sort-key marketVal   # 指定排序欄位（第一次跑會印出可用欄位）
    python fetch_plate_stocks.py --fresh         # 忽略既有輸出，全部重抓

輸出:
    moomoo_us_plate_stocks.csv
    每列 = 一檔成分股，前面帶 plateType/plateCode/plateName/plateId/rank，後面是 API 回傳的全部欄位。
    每抓完一個板塊就寫一次；再跑會跳過已完成的板塊（依 plateId）。

簽章、cookie 暖身、限流退避都沿用 fetch_moomoo_plate_list.py（直接 import）。
請勿高頻或大量重複呼叫，僅供個人查詢使用。
"""

import argparse
import csv
import sys
import time

import requests

from fetch_moomoo_plate_list import (
    HEADERS, MARKET_TYPE_US, MAX_RETRIES, REQUEST_INTERVAL, RETRY_WAIT,
    TooFrequent, parse_json, quote_token, warm_up,
)
from fetch_moomoo_plate_list import OUT_PATH as PLATE_LIST_PATH

STOCK_URL = "https://www.moomoo.com/quote-api/quote-v2/get-plate-stock"
OUT_PATH = "moomoo_us_plate_stocks.csv"
PAGE_SIZE = 30  # 網頁自己用 30；要更多檔就翻頁，但這裡只取前 N，一頁通常夠

META_FIELDS = ["plateType", "plateCode", "plateName", "plateId", "rank"]

# 市值欄位的候選名稱（API 實際叫什麼要看第一次跑印出的欄位清單）
MARKET_CAP_KEYS = ["marketVal", "marketValue", "totalMarketValue", "marketCap", "totalMarketVal"]


def parse_number(v):
    """把 '590.64B' / '1,234.5M' / '12.3K' / '3.2%' 這類字串轉成 float；轉不了回 None。"""
    if v is None:
        return None
    if isinstance(v, (int, float)):
        return float(v)
    s = str(v).strip().replace(",", "").replace("+", "").rstrip("%")
    mult = {"K": 1e3, "M": 1e6, "B": 1e9, "T": 1e12}
    if s and s[-1].upper() in mult:
        try:
            return float(s[:-1]) * mult[s[-1].upper()]
        except ValueError:
            return None
    try:
        return float(s)
    except ValueError:
        return None


def fetch_plate_page(sess, plate_id, page):
    # 鍵順序照瀏覽器實際請求：marketType, plateId, page, pageSize, _（簽章依此序列化）
    params = {
        "marketType": MARKET_TYPE_US,
        "plateId": plate_id,
        "page": page,
        "pageSize": PAGE_SIZE,
        "_": int(time.time() * 1000),
    }
    headers = {"quote-token": quote_token(params)}
    wait = RETRY_WAIT
    for attempt in range(MAX_RETRIES + 1):
        try:
            resp = sess.get(STOCK_URL, headers=headers, params=params, timeout=10)
        except requests.exceptions.ConnectionError as e:
            raise SystemExit(
                f"伺服器主動斷線 (plateId={plate_id}, page={page})：{e}"
                + chr(10)
                + "這代表封鎖已升級為拒絕連線，短時間內重試無效，請隔數小時再跑。"
            ) from None
        resp.raise_for_status()
        try:
            payload = parse_json(resp, f"plateId={plate_id}, page={page}")
            break
        except TooFrequent:
            if attempt == MAX_RETRIES:
                raise RuntimeError(
                    f"連續 {MAX_RETRIES} 次被限流 (plateId={plate_id})，請隔幾分鐘再跑一次。"
                ) from None
            print(f"  plateId={plate_id} 被限流（too frequent），等 {wait} 秒後重試…")
            time.sleep(wait)
            wait *= 2
    if payload.get("code") != 0:
        raise RuntimeError(f"API 回傳錯誤 (plateId={plate_id}, page={page}): {payload}")
    return payload["data"]


def read_plates(only):
    with open(PLATE_LIST_PATH, newline="", encoding="utf-8-sig") as f:
        plates = list(csv.DictReader(f))
    if only:
        plates = [p for p in plates if p["plateType"] == only]
    return plates


def read_existing():
    try:
        with open(OUT_PATH, newline="", encoding="utf-8-sig") as f:
            reader = csv.DictReader(f)
            rows = list(reader)
            return rows, list(reader.fieldnames or [])
    except FileNotFoundError:
        return [], []


def write_csv(rows, fieldnames):
    with open(OUT_PATH, "w", newline="", encoding="utf-8-sig") as f:
        w = csv.DictWriter(f, fieldnames=fieldnames, extrasaction="ignore")
        w.writeheader()
        w.writerows(rows)


def pick_sort_key(item, requested):
    if requested:
        if requested not in item:
            raise SystemExit(f"--sort-key {requested!r} 不在 API 欄位裡。可用欄位：{sorted(item)}")
        return requested
    for k in MARKET_CAP_KEYS:
        if k in item:
            return k
    return None


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--top", type=int, default=10, help="每個板塊取前幾檔（預設 10）")
    ap.add_argument("--only", choices=["industry", "concept"], help="只跑這一類板塊")
    ap.add_argument("--sort-key", help="用哪個 API 欄位排序（預設自動找市值欄位；找不到就照 API 順序）")
    ap.add_argument("--fresh", action="store_true", help="忽略既有輸出，全部重抓")
    args = ap.parse_args()

    plates = read_plates(args.only)
    if not plates:
        raise SystemExit(f"{PLATE_LIST_PATH} 裡沒有板塊，請先跑 fetch_moomoo_plate_list.py")

    rows, fieldnames = ([], []) if args.fresh else read_existing()
    done_ids = {r["plateId"] for r in rows}
    todo = [p for p in plates if p["plateId"] not in done_ids]
    print(f"板塊共 {len(plates)} 個，已完成 {len(plates) - len(todo)}，本次要抓 {len(todo)}")
    if not todo:
        return

    sess = warm_up()
    sort_key = None
    announced = False

    for i, p in enumerate(todo, 1):
        data = fetch_plate_page(sess, p["plateId"], 0)
        items = data.get("list", [])
        if items and not announced:
            print(f"API 回傳欄位：{sorted(items[0])}")
            sort_key = pick_sort_key(items[0], args.sort_key)
            print("排序欄位：" + (sort_key if sort_key else "（找不到市值欄位，照 API 順序）"))
            announced = True
            # 欄位順序：meta 在前、API 欄位照第一次看到的順序在後
            for k in items[0]:
                if k not in fieldnames and k not in META_FIELDS:
                    fieldnames.append(k)
            fieldnames = META_FIELDS + [k for k in fieldnames if k not in META_FIELDS]

        if sort_key:
            items.sort(key=lambda it: parse_number(it.get(sort_key)) or 0, reverse=True)
        for rank, it in enumerate(items[: args.top], 1):
            row = {
                "plateType": p["plateType"], "plateCode": p["plateCode"],
                "plateName": p["plateName"], "plateId": p["plateId"], "rank": rank,
            }
            row.update(it)
            for k in it:
                if k not in fieldnames:
                    fieldnames.append(k)
            rows.append(row)

        write_csv(rows, fieldnames)
        print(f"[{i}/{len(todo)}] {p['plateCode']} {p['plateName']}：取 {min(len(items), args.top)} 檔"
              f"（該頁共 {len(items)}，板塊合計 {data.get('pagination', {}).get('total', '?')}）")
        if i < len(todo):
            time.sleep(REQUEST_INTERVAL)

    print(f"\n完成！共 {len(rows)} 列，已存成 {OUT_PATH}")


if __name__ == "__main__":
    main()
