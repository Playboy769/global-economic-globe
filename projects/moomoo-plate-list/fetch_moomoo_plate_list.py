"""
抓取 moomoo 美股「板塊」代碼清單 (plateCode / plateName)
資料來源: https://www.moomoo.com/quote/us/sector-industry
API 端點: https://www.moomoo.com/quote-api/quote-v2/get-plate-list

用法:
    python fetch_moomoo_plate_list.py

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
- 若暖身後仍拿不到資料，把瀏覽器開發者工具裡的 Cookie 整串貼到 COOKIE 變數。
- 請勿高頻或大量重複呼叫，僅供個人查詢使用。
"""

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


def parse_json(resp: requests.Response, ctx: str) -> dict:
    try:
        return resp.json()
    except ValueError:
        body = " ".join(resp.text[:300].split())
        raise RuntimeError(
            f"回應不是 JSON ({ctx}) — HTTP {resp.status_code}, "
            f"Content-Type={resp.headers.get('Content-Type')!r}, "
            f"Content-Length={resp.headers.get('Content-Length')!r}, body[:300]={body!r}"
        ) from None


def fetch_plate_type(sess: requests.Session, plate_type: int, label: str):
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
        resp = sess.get(BASE_URL, headers=headers, params=params, timeout=10)
        resp.raise_for_status()
        payload = parse_json(resp, f"plateType={plate_type}, page={page}")

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
        page += 1
        time.sleep(1)  # 禮貌性延遲，避免打太快

    return rows


def main():
    sess = warm_up()
    rows = []
    for plate_type, label in PLATE_TYPES.items():
        rows.extend(fetch_plate_type(sess, plate_type, label))

    out_path = "moomoo_us_plate_list.csv"
    with open(out_path, "w", newline="", encoding="utf-8-sig") as f:
        writer = csv.DictWriter(
            f,
            fieldnames=["plateType", "plateCode", "plateName", "plateEnName", "plateId", "leaderStock"],
        )
        writer.writeheader()
        writer.writerows(rows)

    print(f"\n完成！共 {len(rows)} 筆，已存成 {out_path}")


if __name__ == "__main__":
    main()
