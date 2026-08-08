# 隱含波動率曲面 Volatility Surface Viewer

把一整條選擇權鏈畫成一片可旋轉的 3D 隱含波動率曲面。

- **線上**：`globe-invest.up.railway.app/volsurface`（與 globe-invest 其他頁一樣需先登入）
- **本目錄是 dev-source**，部署鏡像在 `globe-invest/app/volsurface/index.html`，
  兩邊必須一致（本鏡像對**不在** sync script 裡，改完要自己 copy 過去並推兩個 repo）

## 資料怎麼來

**線上版是即時抓取，不讀任何預存檔案。** 前端向同源的 `/api/volsurface?ticker=XXX` 要資料，
該端點實作在 `globe-invest/server.js`：

1. Yahoo 的 v7 選擇權端點沒有 cookie ＋ crumb 會回 401，所以先去 `fc.yahoo.com` 取 cookie、
   再用它換 crumb，兩者快取 30 分鐘，遇到 401 才重新取一次。
2. 逐一抓各到期日的鏈，套用過濾條件（見下），組成 `{ticker, spot, fetched, calls, puts}`。
3. 每個標的伺服器端快取 **10 分鐘**。

一次請求同時帶回 calls 與 puts，所以前端切換 Calls / Puts / OTM 合成**不會再發請求**。

## 本地預覽

這頁**不能**用純靜態 server 打開——它需要同源的 `/api/volsurface`。要在本地跑，必須跑
globe-invest 的 Node server（launch.json 的 `globe-invest-node`，port 8080）。

但整個 globe-invest 服務是 default-deny 的登入閘（只有 `/research/*` 公開），所以本地還要
給簽章金鑰並自備一張 session cookie：

```bash
AUTH_SIGNING_SECRET=localtestsecret AUTHORIZED_EMAILS=you@example.com \
PUBLIC_ORIGIN=http://localhost:8080 node globe-invest/server.js
```

再用 `globe-invest/auth.js` 的 `signFor()` 對 `{email, aud:'http://localhost:8080',
typ:TYP_SESSION, exp}` 簽一張 token，塞進 `gi_sid` cookie 即可。

## fetch_options.py 的定位

這支 script **不參與線上部署**，是離線分析／驗證用的：它把同樣的過濾規則套在 yfinance 上，
輸出 JSON 快照，方便在沒有伺服器的情況下比對曲面、或做研究用的定點快照。

```bash
python fetch_options.py TSM NVDA --max-expiries 11
```

⚠️ **它與 `globe-invest/server.js` 裡的過濾常數必須維持一致**（時間價值門檻、IV 上下限、
價性範圍、最小報價數、最小到期天數）。兩邊若漂移，離線快照與線上曲面會長得不一樣，
那正是這類工具最難察覺的 bug。server.js 該段有註解指回這支檔案。

## 過濾條件（兩邊共用）

| 條件 | 理由 |
|---|---|
| 時間價值 < 股價 0.2% → 丟棄 | 深度價內幾乎全是內含價值，倒推 IV 等於除以接近零的數，會長出假尖峰 |
| 買價為 0 → 丟棄 | 沒有人願意用任何價格買，就不算有市場 |
| 未平倉量與成交量同時為 0 → 丟棄 | IV 是陳舊殘值 |
| 到期 < 4 天 → 整個到期日丟棄 | 由 pin risk 與 gamma 主導，不是對波動率的看法 |
| 該到期日報價 < 8 筆 → 丟棄 | 撐不出一條微笑，只會讓擬合向外借點 |
| IV 不在 1%–300%、\|ln(K/S)\| > 0.35 → 丟棄 | 極端值與深度價外 |

## 曲面怎麼算

前端做一次**二維局部加權迴歸**（履約價方向二次、√時間方向線性，tricube 權重，
鄰域大小隨報價密度調整），擬合對象是**總變異數** w = σ²T 而非 σ。

之所以不是「先沿履約價平滑、再沿時間插值」的可分離兩段作法：每個到期日涵蓋的履約價範圍
不同（近月只有價平附近有時間價值），相鄰兩排會由**不同組到期日**擬合出來，曲面會沿著接縫
撕裂成鋸齒。實測可分離作法的算繪粗糙度是二維作法的 5 倍。

顯示層另有兩道處理：高度與顏色刻度取第 1／99 百分位（避免一小塊異常值獨占色階），
擬合結果限制在鄰近報價的 IV 範圍內（近月 σ = √(w/T) 會放大誤差）。

僅供資訊參考，非投資建議。
