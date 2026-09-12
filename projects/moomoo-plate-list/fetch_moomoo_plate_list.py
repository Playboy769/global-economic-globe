"""
抓取 moomoo 美股「板塊」代碼清單 (plateCode / plateName)
資料來源: https://www.moomoo.com/quote/us/sector-industry
API 端點: https://www.moomoo.com/quote-api/quote-v2/get-plate-list

用法:
    python fetch_moomoo_plate_list.py                 # 產業＋概念都抓
    python fetch_moomoo_plate_list.py --only industry # 只抓產業（或 concept）
  已存在的 CSV 裡若某分類已抓齊，會自動跳過那個分類，只補缺的；被限流後再跑不必從頭來。

輸出:
    moomoo_us_plate_list.csv  (plateType, plateCode, plateName, plateEnName, plateId, leaderStock)

注意:
- 這是網站前端內部使用的未公開 API，並非官方文件化的接口，欄位與網址未來可能異動。
- 2026-09-12 實測：沒帶簽章時 API 回 {"code":500,"message":"Params Error"}，不是 403。
  前端 axios 攔截器對每個請求加 `quote-token` 標頭，算法照抄自 app.*.js：
      s = JSON.stringify({每個參數值轉字串})           # 鍵順序 = 參數順序
      quote-token = SHA256( HmacSHA512(s, "quote_web").hex[:10] ).hex[:10]
  之後若再出現 Params Error，優先懷疑站方換了金鑰或算法。
- 2026-09-12 第二次實測：直接打 API 會回 HTTP 200 但 body 是空的（resp.json() 拋 JSONDecodeError）。
  站方要先「開過頁面」拿到 cookie（含 csrfToken）才肯回資料，所以改用 Session 先 GET 一次
  sector-industry 頁面暖身，並把 cookie 裡的 csrfToken 放進 `futu-x-csrf-token` 標頭。
- 2026-09-12 第三次實測：每秒 1 次連打 6 頁後，第 7 頁回 HTML「403 - Operations too frequent」
  （HTTP 200、Content-Type text/html）。所以請求間隔拉到 REQUEST_INTERVAL 秒，碰到 too frequent
  就等 RETRY_WAIT 秒再重試（最多 MAX_RETRIES 次、每次等待加倍），且每抓完一頁就寫一次 CSV。
- 若暖身後仍拿不到資料，把瀏覽器開發者工具裡的 Cookie 整串貼到 COOKIE 變數。
- 請勿高頻或大量重複呼叫，僅供個人查詢使用。
"""

import argparse
import csv
import hashlib
import hmac
import json
import time

import requests

BASE_URL = "https://www.moomoo.com/quote-api/quote-v2/get-plate-list"
PAGE_URL = "https://www.moomoo.com/quote/us/sector-industry"

# 如果不帶 Cookie 就被擋，把瀏覽器裡複製到的完整 Cookie 字串貼在這裡
COOKIE = ""

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/124.0.0.0 Safari/537.36"
    ),
    "Referer": PAGE_URL,
    "Accept": "application/json, text/plain, */*",
}
if COOKIE:
    HEADERS["Cookie"] = COOKIE

MARKET_TYPE_US = 2

# plateType 對照（來自前端 store）：1 = 產業 Industry、2 = 概念 Concept
PLATE_TYPES = {
    1: "industry",
    2: "concept",
}

PAGE_SIZE = 30

REQUEST_INTERVAL = 4   # 每個請求之間至少等幾秒（1 秒會踩到 too frequent）
RETRY_WAIT = 30        # 被限流時第一次等幾秒；之後每次加倍
MAX_RETRIES = 3

OUT_PATH = "moomoo_us_plate_list.csv"
FIELDNAMES = ["plateType", "plateCode", "plateName", "plateEnName", "plateId", "leaderStock"]


def quote_token(params: dict) -> str:
    """複製前端 axios 攔截器的簽章：參數值全轉字串後 JSON 化，HMAC-SHA512 再 SHA256 各取前 10 碼。"""
    s = json.dumps({k: str(v) for k, v in params.items()}, separators=(",", ":"))
    h = hmac.new(b"quote_web", s.encode("utf-8"), hashlib.sha512).hexdigest()[:10]
    return hashlib.sha256(h.encode("utf-8")).hexdigest()[:10]


def warm_up() -> requests.Session:
    """先像瀏覽器一樣載入頁面一次，讓伺服器發 cookie（csrfToken 等）；API 沒有這些 cookie 會回空 body。"""
    sess = requests.Session()
    sess.headers.update(HEADERS)
    resp = sess.get(PAGE_URL, timeout=15)
    print(f"暖身頁面 HTTP {resp.status_code}，取得 cookie: {sorted(sess.cookies.keys())}")
    csrf = sess.cookies.get("csrfToken", "")
    if csrf:
        sess.headers["futu-x-csrf-token"] = csrf
    return sess


class TooFrequent(Exception):
    pass


def parse_json(resp: requests.Response, ctx: str) -> dict:
    try:
        return resp.json()
    except ValueError:
        body = " ".join(resp.text[:300].split())
        if "too frequent" in body.lower():
            raise TooFrequent(ctx) from None
        raise RuntimeError(
            f"回應不是 JSON ({ctx}) — HTTP {resp.status_code}, "
            f"Content-Type={resp.headers.get('Content-Type')!r}, "
            f"Content-Length={resp.headers.get('Content-Length')!r}, body[:300]={body!r}"
        ) from None


def fetch_plate_type(sess: requests.Session, plate_type: int, label: str, on_page=None):
    rows = []
    page = 0
    total_pages = None

    while total_pages is None or page < total_pages:
        # 鍵順序要和前端一致：marketType, plateType, page, pageSize（簽章依此序列化）
        params = {
            "marketType": MARKET_TYPE_US,
            "plateType": plate_type,
            "page": page,
            "pageSize": PAGE_SIZE,
        }
        headers = {"quote-token": quote_token(params)}
        wait = RETRY_WAIT
        for attempt in range(MAX_RETRIES + 1):
            resp = sess.get(BASE_URL, headers=headers, params=params, timeout=10)
            resp.raise_for_status()
            try:
                payload = parse_json(resp, f"plateType={plate_type}, page={page}")
                break
            except TooFrequent:
                if attempt == MAX_RETRIES:
                    raise RuntimeError(
                        f"連續 {MAX_RETRIES} 次被限流 (plateType={plate_type}, page={page})，"
                        f"請隔幾分鐘再跑一次。"
                    ) from None
                print(f"[{label}] 第 {page + 1} 頁被限流（too frequent），等 {wait} 秒後重試…")
                time.sleep(wait)
                wait *= 2

        if payload.get("code") != 0:
            raise RuntimeError(f"API 回傳錯誤 (plateType={plate_type}, page={page}): {payload}")

        data = payload["data"]
        pagination = data["pagination"]
        total_pages = pagination["pageCount"]

        for item in data["list"]:
            rows.append({
                "plateType": label,
                "plateCode": item.get("plateCode"),
                "plateName": item.get("plateName"),
                "plateEnName": item.get("plateEnName"),
                "plateId": item.get("plateId"),
                "leaderStock": f"{item.get('stockName')} ({item.get('stockCode')})",
            })

        print(f"[{label}] 已抓取第 {page + 1}/{total_pages} 頁，累積 {len(rows)} 筆")
        if on_page:
            on_page(rows)
        page += 1
        time.sleep(REQUEST_INTERVAL)  # 禮貌性延遲；1 秒實測會被限流

    return rows


def read_existing():
    """讀取上一次留下的 CSV，回傳 {分類: rows}；沒有檔案就回空 dict。"""
    try:
        with open(OUT_PATH, newline="", encoding="utf-8-sig") as f:
            rows = list(csv.DictReader(f))
    except FileNotFoundError:
        return {}
    by_type = {}
    for r in rows:
        by_type.setdefault(r["plateType"], []).append(r)
    return by_type


def write_csv(rows):
    with open(OUT_PATH, "w", newline="", encoding="utf-8-sig") as f:
        writer = csv.DictWriter(f, fieldnames=FIELDNAMES)
        writer.writeheader()
        writer.writerows(rows)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--only", choices=sorted(PLATE_TYPES.values()),
                    help="只抓這一個分類（industry / concept）")
    ap.add_argument("--fresh", action="store_true", help="忽略既有 CSV，全部重抓")
    args = ap.parse_args()

    existing = {} if args.fresh else read_existing()
    for label, rows in list(existing.items()):
        if len(rows) % PAGE_SIZE == 0:
            # 剛好是整頁倍數，可能是上次抓到一半被擋；保守起見重抓這一類
            print(f"既有 CSV 的 [{label}] {len(rows)} 筆可能不完整，本次重抓")
            del existing[label]
        else:
            print(f"既有 CSV 已有 [{label}] {len(rows)} 筆，本次跳過")

    todo = {t: l for t, l in PLATE_TYPES.items()
            if (args.only is None or l == args.only) and l not in existing}
    if not todo:
        print("沒有需要抓的分類。要全部重抓請加 --fresh。")
        return

    sess = warm_up()
    done = []  # 已完成的分類（含既有 CSV 裡的）
    for label in PLATE_TYPES.values():
        done.extend(existing.get(label, []))

    def flush(partial):
        # 每抓完一頁就整份重寫，中途被擋也留得住已抓到的資料
        write_csv(done + partial)

    for plate_type, label in todo.items():
        done.extend(fetch_plate_type(sess, plate_type, label, on_page=flush))

    write_csv(done)
    print()
    print(f"完成！共 {len(done)} 筆，已存成 {OUT_PATH}")


if __name__ == "__main__":
    main()
