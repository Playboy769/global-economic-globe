# Claudecode — Repo Map & Rules

This directory holds several **unrelated** personal projects side by side. Read this before
moving, deleting, or assuming anything about what's "dead" — a lot of it looks like clutter
but has an active deployment behind it.

## Mandatory pre-task clarification

Before starting **any** task in this repo — regardless of how clear the request seems — first
**re-read this entire CLAUDE.md file** (even if its content was already loaded into context
earlier in the session — read it again from disk every single time), then ask the user at least
10 clarifying questions and wait for answers before doing any work. `AskUserQuestion`'s per-call
max is 4 questions, so reaching 10 means multiple back-to-back calls (e.g. 4 + 4 + 2) — do not
stop at one call. This applies even when the request looks unambiguous. This overrides the
general "Auto Mode" bias toward proceeding without confirmation for this repo specifically.

## 功能位置對照表（改東西前先查這張表）

**觸發規則**：使用者只要說「要修改 / 改一下 / 調整 XX 功能」（或同義），**第一步先在本節
表格用功能名稱定位**，讀出：① dev repo、② folder、③ 要編輯的 file、④ 檔內錨點（id/行）、
⑤ route/URL、⑥ **有沒有部署鏡像要一起改**。不要每次重新 grep 找位置——先查表，查不到或
表可能過期才 grep 驗證，並順手更新本表。改完若該列標了鏡像，依對應 sync 規則同步（見
下方 Deployment topology 與 sync scripts）。所有路徑相對本 repo 根
`C:\Users\ryan9\OneDrive\桌面\Claudecode`。

> 「現況一致 ✅ / 已分歧 ⛔」是**製表當下（2026-07-18）**的快照，會過期；同步前仍以實際
> diff 為準，不要只信這欄。

### 1. OutsideFramework 作品集首頁
- 本 repo (`global-economic-globe.git`) · `app/OutsideFramework/index.html`（+`assets/`、`admin.html`）· 本地 `outside-framework` :8125
- **部署**：Dockerfile → `node /server.js`（**不是 nginx**）。**無鏡像、無需同步**。線上 `ofw.up.railway.app`。
- ⚠️ 本目錄根的 `nginx.conf` 已是**孤兒檔**，Dockerfile 沒有引用它；別再照它推論部署方式。

**這是一個有後端的服務，不是靜態站。** 根目錄三個檔案才是真正的執行體，改動時它們與 HTML 同等重要：

| 檔案 | 職責 | 動到時要注意 |
|---|---|---|
| `server.js` | HTTP 路由、Google OAuth、guest/full 兩種 HTML 變體、handoff token 簽發 | 請求標頭（Host / X-Forwarded-Proto / X-Forwarded-For）**一律不可信**，不得用來決定 cookie 旗標、origin 或速率限制分桶 |
| `auth.js` | HMAC token 簽發與驗證、cookie、速率限制 | **必須與 globe-invest / structural_holes / article_db 的副本逐字一致**（Python 版為 `auth.py`），且四邊共用同一組 `AUTH_SIGNING_SECRET` |
| `analytics.js` | 訪客流量 SQLite（`node:sqlite`，Railway volume `/data`） | 欄位名要進 `GROUPABLE` 白名單才能拿來 GROUP BY |

### Token 規則（2026-07-30 安全稽核後，五個服務一致）

**每張 token 都必須帶 `aud`（目標服務 origin）與 `typ`（`session` / `handoff` / `state`），
且簽章金鑰由 `aud` 衍生**：

```
service_key(aud) = HMAC-SHA256(AUTH_SIGNING_SECRET, "ofw-token-v2|" + aud)
```

所以為 globe-invest 簽的 token 在 article-db **簽章根本對不上**，不是靠某處記得比對欄位。
用 `signFor()` / `verifyFor()`（Python：`sign_for()` / `verify_for()`），**不要**直接用
`signToken` / `verifyToken` 做存取控制 —— 前者沒指定 `aud` 就無法呼叫（因為 `aud` 決定金鑰），
後者可以，那正是出事的地方。

稽核前的狀況：`aud` 有簽名卻從未在 session 路徑被檢查，而且**每個服務簽給自己的 session
cookie 完全沒有 `aud`** —— `sh_sid`/`gi_sid`/`adb_sid` 三張 cookie 彼此可互換，等於萬用鑰匙。
handoff token 又是掛在網址 `?auth=` 上傳遞，會經由 Referer、瀏覽器歷史與下游 log 外洩。

⚠️ **`auth.js` / `auth.py` 改動必須五個服務同步**（root、globe-invest、structural_holes、
article_db、projects/gaoye-mock-exam），否則跨網域登入會斷。且因為 Railway 無法原子性
同時部署五個服務，切換簽章方案時**先推四個下游、最後推 ofw** —— 下游都有迴圈斷路器，
這個順序讓中間窗口變成一次 503 說明頁；反過來推會變成瀏覽器無限重導向。

> 新增下游服務時要動的四個地方（2026-08-21 加 gaoye-mock-exam 時的實際清單）：
> ① 五份 auth 副本的 `ALLOWED_ORIGINS`（＋`DEV_ORIGINS` 的本地 port）、
> ② ofw `server.js` 的 `GATED_ORIGINS`、③ 新服務自己的登入閘門（照抄
> `globe-invest/server.js` 的 gate，含**迴圈斷路器**）、④ 新服務的 Railway 變數
> `AUTH_SIGNING_SECRET` 與 `AUTHORIZED_EMAILS`。
> 變數用跨服務引用 `${{Outside Framework.AUTH_SIGNING_SECRET}}` 設定，換金鑰時會跟著變；
> 但注意 **`railway variables` 會把引用解析後的值印出來**，別在共享畫面或紀錄裡跑它。

> 目前**尚未**做的是縮小爆炸半徑：四邊仍共用同一把 master secret，彼此都能推導出對方的金鑰。
> 要讓每個服務只持有自己的衍生金鑰，需要在三個 Railway 服務各設一組環境變數。

| 小功能 | 檔內錨點 | 對外連結目的地 |
|---|---|---|
| 導覽列 Home/About/Works/Philosophy/Quotes | 行 278–284 | 同頁錨點 |
| About（Structured/Interactive/Narrative/Open Source）| 行 351–388 | — |
| Works 分類標題 Earnings/Network/Causal/Knowledge | 行 408–411 | — |
| Works 卡片 · 14 篇 earnings 報告 | 行 421–506 | `…/research/<TICKER>_…Analysis.html` |
| Works 卡片 · Sankey / Options / Brownian | 行 510 / 517 / 520 | `…/sankey`、`…/options`、`…/brownian` |
| Works 卡片 · Structural Holes | 行 522 | `structural-holes-production.up.railway.app/` |
| Works 卡片 · Article Database | 行 528 | `articlebase.up.railway.app/` |
| Works 卡片 · Globe/Invest/Causal/Warning/High-price | 行 651–657 | `…/globe` … `/high-price` |

**導覽列雙語慣例（2026-07-31 起）**：桌面導覽列（`.nav-links` 內的 6 個 SPA 分頁按鈕）採
「英文在上、中文小字在下」堆疊，標記為 `<span class="nb-en">`＋`<span class="nb-zh">`，CSS
用 `.nav-btn:has(.nb-zh)` 限定套用範圍——刻意不用裸的 `.nav-btn`，因為 `.mobile-dock` 的圖示
按鈕與 Login/Admin/Logout（`.nav-auth-btn`）都共用 `.nav-btn` 這個 class，`:has()` 可以精準
只選到真的包中文 span 的那 6 個按鈕，不會因為 CSS specificity 意外影響到另外兩種。

翻譯對照：Home→首頁、About→關於、Works→作品、Timeline→時間軸、Philosophy→哲學、
Quotes→語錄（語錄二字是照抄站內既有的 `aria-label="語錄進度"`，不是新造詞）。

⚠️ **這暫時只套用在導覽列，全站其他地方的英文標籤（`.about-label` 顯示的
"About"/"Works"/"Philosophy"/"Timeline" 等 eyebrow 字樣、Works 分類標籤、頁尾連結）仍是
純英文，是刻意保留、不是忘記改。若之後要把雙語擴大到那些地方，套用同一套
`.nb-en`/`.nb-zh` span 命名＋`:has()` 選擇器手法即可，但那是後續的獨立改動。

### 2. GlobalEco（globe 3D 地球）
- Dev：本 repo · `app/GlobalEco/index.html` · 本地 `globe` :8124
- **鏡像**：`globe-invest.git` · `globe-invest/app/globe/index.html` → `/globe` · **sync：`scripts/sync-globe-invest.ps1`**

| 小功能 | 錨點 id |
|---|---|
| 地球畫布 / 搜尋 | `globe-canvas`、`search-box`/`search-input`/`search-results`（行 413–416）|
| 視覺模式鈕 熱力/網路/排名/比較 | `btn-heatmap`/`btn-network`/`btn-ranking`/`btn-compare`（行 420–423）|
| 熱力圖圖例/模式 | `heatmap-legend`、`hm-mode`（行 429）|
| 國家資訊側欄 | `sidebar`/`country-info`/`country-header`（行 432–450）|
| 6 資料 tab 概覽/人口/農業/貿易/金融/社會 | `tab-nav` + `panel-overview…social`（行 453–464）|
| 雙邊貿易比較 | `panel-trade` > `trade-detail`（行 ~468）|
| 相關後端 | `globe-invest/server.js`：`/api/oil-prices`、`/api/stock-history` |

### 3. InvestFrame（invest 投資框架）
- Dev：本 repo · `app/InvestFrame/index.html` · 本地 `investframe` :8126
- **鏡像**：`globe-invest/app/invest/index.html` → `/invest` · **sync：`scripts/sync-globe-invest.ps1`**

| 小功能 | 錨點 id |
|---|---|
| 四環框架 總經/風險/產業/持倉 | `iv-ring-macro`/`-risk`/`-industry`/`-portfolio`（行 510–532）|
| 資訊 feed（搜尋/刪除/清單/詳情）| `iv-feed`/`iv-search`/`iv-feed-list`/`iv-feed-detail`（行 537–562）|
| 觀察清單 watchlist | `iv-wl`/`iv-wl-list`（行 574–580）|
| 每日筆記/週條/立場 | `iv-daily-note`/`iv-week-strip`/`iv-stance-wrap`（行 585–589）|
| 相關後端 | `globe-invest/server.js`：`/api/invest-data`、`/api/invest-groups`、`/api/og-fetch`、`/api/upload-asset`、`/api/asset/*` |

### 4. CausalFrame（causal 因果圖）
- Dev：本 repo · `app/CausalFrame/index.html` · 本地 `causalframe` :8128
- **鏡像**：`globe-invest/app/causal/index.html` → `/causal` · **sync：`scripts/sync-globe-invest.ps1`**

| 小功能 | 錨點 id |
|---|---|
| 工具列 復原/重做/圖層 | `cf-undo`/`cf-redo`/`cf-layer-up`/`cf-layer-down`（行 819–834）|
| 迴圈/槓桿/迷你圖/搜尋/暗色 | `cf-loops`/`cf-leverage`/`cf-minimap-btn`/`cf-search-btn`/`cf-dark-btn`（行 845–874）|
| 新增節點 文字/圖/PDF/嵌入/表格/畫布/圖引用/分隔線/**earnings 報告連結** | `cf-add-text…-divider`、`cf-add-earnings`（行 900–981）|
| 插入 earnings call 報告節點（連到 `/research/<TICKER>_..._Analysis.html`）| 按鈕 `cf-add-earnings`（行 980）· 選取器 `showEarningsInsertPicker`/`addEarningsRefNode` · 節點型別 `isEarningsRef` |
| 檔案 feed（資料夾/檔案/範本）| `cf-feed`/`cf-new-folder`/`cf-new-file`/`cf-tpl-btn`（行 951–960）|
| 畫布層 | `cf-canvas`/`cf-edges`/`cf-world`/`cf-nodes`/`cf-labels`（行 963–968）|
| 相關後端 | `globe-invest/server.js`：`/api/causal-files`、`/api/research-reports`（掃描部署鏡像 `globe-invest/app/research/` 列出可插入的 earnings 報告）|

### 5. globe-invest 專屬頁（**只在 globe-invest repo，本 repo 無 dev-source — 直接改該 repo**）
folder 一律 `globe-invest/app/<x>/index.html`。本地整站 `globe-invest-app` :8136。

| 頁面 | folder → route | 內部小功能 | 後端 API（`globe-invest/server.js`）|
|---|---|---|---|
| 高價股追蹤 | `high-price/` → `/high-price` | 清單 / 即時報價 / 指標 | `/api/high-price/list`、`/realtime`、`/metrics` |
| 警示雷達 | `warning/` → `/warning` | TWSE/TPEX 公告·處置·三大法人·除息 | `/api/warning/{twse,tpex}-*`（12 端點）、`/company-info`、`/price-change` |

### 6. 鏡像對（**手動同步，不在 sync script — 改完要自己 copy 過去並推兩個 repo**）
| 功能 | Dev（本 repo） | 部署（globe-invest） | route |
|---|---|---|---|
| Brownian | `projects/brownian-motion-simulator/index.html`（本地 :8132）| `globe-invest/app/brownian/index.html` | `/brownian` |
| Options Guide | `projects/options_guide.html` | `globe-invest/app/options/index.html` | `/options` |
| Sankey 工具 | `projects/sankey-diagram-demo/index.html`（本地 `sankey-diagram-demo` :8137）| `globe-invest/app/sankey/index.html` | `/sankey` |
| Research 報告 | `research/<ticker>-…/*.html` | `globe-invest/app/research/*.html` | `/research/<FILE>.html` |
| Warning 舊 dev | `projects/market-warning-radar/`（本地 :8130）| `globe-invest/app/warning/` | `/warning` |
| 猛健樂完全解析（藥物專題雜誌）| `projects/mounjaro-magazine/index.html`（本地 `mounjaro-magazine` :8173）| `globe-invest/app/mounjaro/index.html` | `/mounjaro`（route 於 `globe-invest/server.js` 手動註冊，非資料夾自動對應）|
| Agentic AI 完全解析（技術專題雜誌）| `projects/agentic-ai-magazine/index.html`（本地 `agentic-ai-magazine` :8174）| `globe-invest/app/agentic-ai/index.html` | `/agentic-ai`（route 於 `globe-invest/server.js` 手動註冊，非資料夾自動對應）|

> ⚠️ `warning` 與 `research` 兩側已**實質分歧**（非落後幾個 commit），`sankey`/`options`/
> `brownian` 製表當下一致但同樣不在 sync script。同步前務必先 diff，勿盲 copy 覆蓋。

### 7. 獨立部署（各自 repo，本 repo 只是巢狀）
| 功能 | repo | folder | 部署檔 | URL | 本地 |
|---|---|---|---|---|---|
| article-db | 本 repo dev + `article-db-api.git`（remote `article-db`）| `article_db/` | `index.html`(前端) + `app.py`(FastAPI) | `articlebase.up.railway.app` | `article-db` :8127 |
| structural-holes | `structural-holes.git` | `structural_holes/` | `app.py`(uvicorn) + `graph.py` | `structural-holes-production.up.railway.app` | `structural-holes` :8129 |
| sector-rotation-system | `sector-rotation-system.git`（GitHub `Playboy769/sector-rotation-system`，私有）| `sector-rotation-system/` | `Dockerfile`（Caddy 供 `output/` 靜態圖表＋`scheduler.py` 背景排程重繪）| `sector-rotation-system-production.up.railway.app` | 無 dev server，本機直接雙擊開 `dashboard.html`（RRG 2D/3D、台股 3D、資金流向、季節性熱力圖、族群輪動、機構持倉 13F 等分頁）|
| twchips-report | `twchips-report.git`（GitHub `Playboy769/twchips-report`，私有）| `twchips-report/` | `app.py`(FastAPI，`Procfile`: `uvicorn app:app`) + `backfill.py`/`build.py` | `twchips-report-production.up.railway.app` | 無固定 dev server，本機 `uvicorn app:app --reload` |
| my-slide | `my-slide.git` | `my-slide/` | — | Netlify | — |

> **sector-rotation-system 補充**：Railway 綁 GitHub 自動部署（跟 article-db/structural-holes
> 同模式，**不是** gaoye-mock-exam/macro-tracker 那種手動 `railway up`）。已上架 OutsideFramework
> Works（Finance & Markets 分類，見 `d87fa67` commit「上架 TW Chips Report 與 Sector Rotation
> Dashboard」），純外部連結卡片，不比照 earnings 報告走鏡像流程。**部署容器有自己獨立的
> `scheduler.py` 背景排程**（Railway 上每天 18:00 Asia/Taipei 自動重繪全部圖表），跟本機
> README 提到的 Windows 工作排程器 `SectorRotationDailyUpdate` 是兩條互不相通的更新管道——
> 本機排程只更新本機 `dashboard.html` 讀的 `output/`，不會推上 Railway，反之亦然。改這個
> 專案的程式碼要進 `sector-rotation-system/` 自己的 repo 走它自己的 git 流程（GitHub push
> 觸發 Railway 重新部署），跟本 repo 的 commit 無關，不需要（也不能）用本 repo 的 sync
> script 同步。

> **twchips-report 補充**：三大法人籌碼透視站（現貨/期貨背離、自營商拆分、三方分歧、選擇權、
> 融資融券查詢）。資料抓取靠**第三方套件** [`twchips`](https://github.com/catcat222222/twchips)
> （`requirements.txt` 直接 `pip install git+https://github.com/catcat222222/twchips`，`app.py`
> 裡 `from twchips import twse`）——這不是本 repo 或使用者自己的專案，`git remote` 指向
> `catcat222222`，本機 `twchips/` 資料夾只是那個上游套件的唯讀參考副本（純讀原始碼/除錯用，
> 沒有自己的 Railway 部署，**不要**把它跟 `twchips-report` 這個部署混為一談——2026-08-31
> 曾一度誤會兩者是同一個已上架的東西，查 `railway list`／`railway status` 確認 Railway 上
> 只有 `twchips-report` 這一個 service，沒有獨立的 `twchips` 部署）。Railway 綁 GitHub 自動
> 部署（跟 article-db/structural-holes/sector-rotation-system 同模式）。已上架 OutsideFramework
> Works（Finance & Markets 分類，見 `d87fa67` commit「上架 TW Chips Report 與 Sector Rotation
> Dashboard」）。

> **article-db 改法**：編 `article_db/index.html` → 本 repo commit → 跑
> `scripts/sync-article-db.ps1` 推到 `article-db` remote，否則線上不更新（見下方專節）。

### 8. 非部署工具／資料（本 repo，標準）
| 功能 | folder | 本地 port |
|---|---|---|
| 股票分析器 | `projects/stock-analyzer/` | — |
| 科技估值篩選 | `projects/tech-value-screener/` | — |
| 食物熱量查詢 | `projects/food-calorie-lookup/` | `food-calorie-lookup` :8131 |
| **報告用圖表產生器**（折線／長條／甘特／圓餅／瀑布／散點／熱力圖／分層關係／堆疊面積／雷達／箱型）| `projects/chart-tools/` | `chart-tools` :8158 |
| 象限圖產生器（2×2 散點）| `projects/quadrant-chart-demo/`（單一 `index.html`）| `quadrant-chart-demo` :8160 |
| Earnings 分析（範本源 JNJ，見下方框架章節）| `research/<ticker>-analysis-*/`；各 ticker 一資料夾 | panw :8133 / sumco :8134 / dell :8135 / tsmc :8138 / ms :8139 / lrcx :8140 |
| SEC 抓取工具（VBA）| `SEC-Filing-Fetcher/`（`.bas`+`.xlsm`）| — |
| 產業結構圖 | `industry frame/`（PNG+SVG）| — |
| 圖庫素材（未被引用）| `ofwphoto/`（.jpg）| — |
| 交易/VBA 專案 | `RR4/`、`RR5/`、`EMA Bias Model/` | — |
| FinceptTerminal（空巢狀 repo，疑廢棄）| `FinceptTerminal/` | — |
| 舊版留存 | `archive/` | — |

#### SEC-Filing-Fetcher 慣例（2026-08-11 起）

- **Dashboard 只放快照表與圖表物件；任何圖表的來源資料一律寫到 `RawData` 工作表。** 每個
  系列佔相鄰兩欄（第 1 項 A:B、第 2 項 C:D…），第 1 列項目名、第 2 列欄名、第 3 列起資料，
  順序是股價 → EPS → 各指標的年度緊接季度。圖表用 `SeriesCollection.NewSeries` 指向那些
  範圍，**不要**把數字硬寫進 series（那樣資料不存在於任何格子、無法覆核），也不要把資料表
  搬回 Dashboard——那張 1250 列的日股價表曾把趨勢圖擠到幾千像素之下。
- **數值軸一律用 `modTheme.FitValueAxis` 依當次資料算範圍**，不要退回 Excel 自動縮放：
  自動縮放幾乎一律把軸錨在 0，實測 AXTI 那次 47 張圖有 29 張的資料用不到軸高 60%（股東權益
  179M–275M 畫在 0–300M 軸上就是一條平線）。它會在 4/5/6 格裡挑最貼合的取整刻度；雙軸的圖
  （股價 vs EPS）兩軸取相同格數，格線才會重合。全為零的序列（無股利、無長期負債）刻意不套用。
- **`build.ps1` 只注入不編譯**，印出 SUCCESS 不代表 VBA 跑得起來。驗證要用 Excel COM 開檔
  （`$excel.AutomationSecurity = 1`）設 `Input!A1` 觸發 `Worksheet_Change` 實跑一次（約 90 秒），
  再讀格子與 `SeriesCollection.Formula`／軸刻度。`Chart.Export` 在自動化模式下只會寫出 0 位元組
  PNG，別花時間截圖。

#### 估值欄位與台股股利（2026-09-05 起）

Filings / TW_Filings 兩張表的欄位 1–48 之後，接著 **49–59 的估值區塊**（`modValuation.bas`
定義欄位常數與表頭，美股台股共用一組標籤，避免兩邊漂移）：預收款流動/非流動/合計、
盈餘保留率、內生成長率、投入資本、NOPAT、ROIC、WACC、ROIC−WACC 價差、預收款來源。
刻意**接在第 48 欄（Items）之後而不是插進去**——48 這個欄號在 modSEC/modMOPS 兩邊都有
寫死的呼叫點，插入會把它們整排推掉。新欄位一併進了 `modCharts` 的 `metricCols`，所以
會照現行慣例自動寫進 RawData 並套 `FitValueAxis`。

- **口徑**（有多種合理定義，所以寫在 `modValuation.bas` 檔頭）：投入資本＝總資產−流動負債；
  NOPAT＝營業利益×(1−有效稅率)；內生成長率＝ROE(期末權益)×保留率。**季報列會把流量型
  比率年化（×4）**，否則拿單季 ROIC 去比年化 WACC，幾乎每家公司都會看起來在毀滅價值。
- **WACC 是估的，不是抓的**——沒有任何 filing 揭露資金成本。Beta 由月報酬對基準指數回歸，
  Rf／ERP／稅率備援放在 **`Input!B12:B20`**（美股台股各一組＋Beta 回歸月數），刻意放在
  工作表而非寫死成常數，這樣每個 WACC 背後的假設都看得到也改得動。
- **預收款欄位的覆蓋率陷阱**：窄口徑的 `CustomerAdvances*`／`CustomerDeposits*` 幾乎沒人
  用——實測 CY2024Q4 全市場只有約 **43 家**標記，`ContractWithCustomerLiability*` 則有
  **2,033 家**（SEC frames API）；台股 IFRS 更是**完全沒有**「預收貨款」這個資產負債表科目
  （實查 3017 FY2025Q1，只有合約負債、存入保證金、以及**預收股款**——最後這個是增資，
  不是客戶預付，別誤用）。所以是**窄口徑優先、缺了才退合約負債**，且第 59 欄逐列標示
  實際用的是哪一個，兩者語意不同不可混比。
- **台股股利走另一支查詢**（`modTWDividend.bas`，MOPS `ajax_t05st09_2`，POST）。三個會
  回「看起來正常但沒資料」的空殼頁、不會報錯的坑：① `firstin` 必須是字串 **`ture`**
  （MOPS 自己拼錯的 typo，送 `true`/`1` 都不行）② 必須帶 `encodeURIComponent=1`
  ③ 回應是**大寫 `<TR>/<TD>`**、每格補的是實體 `&nbsp;` 而非字元，且「股利年度」與
  「年度／上半年／下半年」**合併在同一格**，欄位索引會整排位移。`TYPEK=all` 一次涵蓋
  上市與上櫃（3017／6488 實測），不需要分市場探測；半年配的公司同一年度要加總。
  抓不到時退 TWSE `TWT49U` 除權息表，但那支沒有股利年度、要靠「除息年 N → 股利年度 N−1」
  推算，所以來源要標註。
- ⚠️ **不要用 Yahoo 的 `events=div` 當台股股利來源**：實測 3017 的 2026 除息事件回
  `17.898518`，官方是 `20.881604`——它會**漏掉資本公積發放的現金**那一段。

> 這裡的驗證數字都是實跑 Excel COM 得到的，不是推論；`build.ps1` 只注入不編譯的老問題
> 依舊——這次就是靠實跑才抓到「`Public Const` 放在程序之間」的編譯錯誤（VBA 要求模組層級
> 宣告必須在所有程序之前），而 build 照樣印 SUCCESS。

## Deployment topology (this is the part that bites)

There are **eight separate Railway deployments** sourced from **five separate git repos**, plus
this main repo itself:

| Deployment | Repo | Source path | Live URL |
|---|---|---|---|
| OutsideFramework (portfolio homepage + central auth) | this repo (`origin` = `global-economic-globe.git`) | `app/OutsideFramework/` + root `server.js`/`auth.js`/`analytics.js` → `Dockerfile` → **node** | `ofw.up.railway.app` |
| globe / invest / causal / brownian | **`globe-invest/` — its own repo** (`globe-invest.git`) | `globe-invest/app/{globe,invest,causal,brownian}/index.html` + `globe-invest/server.js` | `globe-invest.up.railway.app` |
| structural-holes | **`structural_holes/` — its own repo** | — | `structural-holes-production.up.railway.app` |
| article-db | **separate repo `article-db-api.git`** (this repo has it as remote `article-db`) | `index.html` at that repo's root — mirrors this repo's `article_db/index.html` | `articlebase.up.railway.app` |
| sector-rotation-system（RRG 2D/3D 板塊輪動、資金流向、季節性、族群輪動、13F 機構持倉，2026-08-30 新增）| **`sector-rotation-system/` — 其自己的 repo**（GitHub `Playboy769/sector-rotation-system`，私有，Railway 綁 GitHub 自動部署）| 全部 `*.py` + `Caddyfile`/`Dockerfile`/`entrypoint.sh` → 容器內 Caddy 供靜態 `output/` 圖表，`scheduler.py` 背景每日 18:00 Asia/Taipei 重繪 | `sector-rotation-system-production.up.railway.app` |
| twchips-report（三大法人籌碼透視，2026-08-30 新增）| **`twchips-report/` — 其自己的 repo**（GitHub `Playboy769/twchips-report`，私有，Railway 綁 GitHub 自動部署）| `app.py`(FastAPI/uvicorn) + `index.html`/`lookup.html` 靜態報告；資料抓取透過第三方套件 `twchips`（pip 依賴 `git+https://github.com/catcat222222/twchips`，**非**本 repo 專案，本機 `twchips/` 只是唯讀參考副本、無獨立部署）| `twchips-report-production.up.railway.app` |
| gaoye-mock-exam (高業模擬測驗＋錯題紀錄) | this repo | `projects/gaoye-mock-exam/` → `index.html`/`server.js`/`store.js`，**用 `railway up` 直接上傳，不綁 GitHub** | `gaoye-mock-exam-production.up.railway.app` |
| macro-tracker（總經黏性追蹤儀表板，2026-08-30 新增）| this repo | `research/macro-tracker-2026/` → `index.html`/`server.js`，**用 `railway up` 直接上傳，不綁 GitHub**；無登入、無資料庫，純唯讀代理政府公開 API（BLS/US Treasury），手動輸入的部分存前端 localStorage | `macro-tracker-production-f296.up.railway.app` |
| cosmetics-codex（日韓化妝品圖鑑，2026-09-03 新增）| this repo | `projects/cosmetics-codex/` → `index.html`/`server.js`/`store.js`＋`data/part*.json` 種子，**用 `railway up` 直接上傳，不綁 GitHub**；無登入（公開瀏覽），SQLite 寫在 volume `cosmetics-codex-volume`（`/data`）| `cosmetics-codex-production.up.railway.app` |
| my-slide | **`my-slide/` — its own repo** (Netlify) | — | — |

### gaoye-mock-exam、macro-tracker、cosmetics-codex：用 `railway up` 直接上傳的服務

其他服務（Outside Framework、globe-invest、Article Database、structural-holes）都是
Railway 綁 GitHub repo 自動部署，**這三個不是**——它們沒有綁任何 repo，都在同一個
Railway 專案 `dependable-charm` 底下當獨立 service，要更新一律：

```
railway up --detach --service <service-name>
```

（`--service` 可省略，前提是 CLI 目前已 `railway service <name>` link 到正確的服務；
這個專案底下有多個服務共用同一個 CLI link 狀態，省略前務必先 `railway status` 確認
目前 link 的是哪一個，避免對著別的服務部署。）

⚠️ **2026-08-30 更正**：本節先前寫「建置脈絡就是你執行 `railway up` 的那個資料夾（不是
git root），一定要在子資料夾底下執行」，**這是錯的**——依 `.dockerignore` 與兩份
`Dockerfile` 裡的實際註解、以及 macro-tracker 建立時的實測，真正的機制是：

- **建置脈絡固定是 git root**，`railway up` 一律在 **repo 根目錄**執行（不是子資料夾）。
- 每個服務用 service 變數 `RAILWAY_DOCKERFILE_PATH` 指到自己那份 `Dockerfile`
  （例：`research/macro-tracker-2026/Dockerfile`），Railway 才知道該用哪一份。
- 因為脈絡是 repo 根，根目錄的 `.dockerignore` 對所有服務生效（子資料夾自己的
  `.dockerignore` 不會被讀到）——新增一個這樣的服務時，記得在根 `.dockerignore` 加
  `!<該服務的資料夾路徑>` 例外，否則 `COPY` 會找不到檔案而建置失敗。
  `Dockerfile` 裡的 `COPY` 路徑因此要寫**相對 repo 根**的完整路徑，不能只寫檔名。
- 新增這類服務的完整流程：`railway add --service <name>` 建空服務 → `railway service
  <name>` link 過去 → `railway variables --service <name> --set
  "RAILWAY_DOCKERFILE_PATH=<path>/Dockerfile"` → 根 `.dockerignore` 開例外 →
  在 **repo 根**跑 `railway up --detach --service <name>`。

⚠️ **另外兩個一定會踩到的坑**：

1. **`railway logs --build` 不帶 deployment id 時，顯示的不一定是你剛觸發的那次建置**。
   排查失敗一律先 `railway deployment list` 取 id，再
   `railway logs --build <id>`，否則會對著別次的 log 診斷錯方向。
2. **Windows Git Bash 會把 `/data` 轉成 Windows 路徑**，`railway volume add --mount-path /data`
   會回「Mount path must start with a `/`」。前面加 `MSYS_NO_PATHCONV=1` 才過得去。

資料存在 volume `gaoye-mock-exam-volume`（掛載點 `/data`）的 SQLite `exam.db`，
由 `store.js` 以 `node:sqlite` 寫入，作法與 OutsideFramework 的 `analytics.js` 相同。
**登入預設關閉**（2026-08-21 當天先開後關）：一度改成全站強制走 ofw 中央登入，但
Google 回呼卡在「Invalid or expired login request」，人被鎖在自己的工具外面，當天即改回
不擋。整合本身**沒有刪掉**，只是變成 opt-in —— service 變數 `REQUIRE_LOGIN=1` 就會恢復
閘門（cookie 名 `gme_sid`，閘門與迴圈斷路器照抄 `globe-invest/server.js`，白名單沿用
同一組 `AUTHORIZED_EMAILS`）。要重開之前，得先查出上面那個回呼錯誤的成因。
`/healthz` 刻意留在閘門之前，否則 Railway 健康檢查永遠過不了。

⚠️ 現況是**知道網址就能讀寫紀錄**，且刻意不上架到 OutsideFramework Works（私人工具）。

紀錄的 `email` 欄位保留著：未啟用登入時線上一律記成 `OWNER_EMAIL`（預設 `owner@local`）、
本地記成 `DEV_EMAIL`（預設 `dev@localhost`），所以之後把 `REQUIRE_LOGIN` 打開不必再做
一次 schema 遷移。

作答紀錄**按 email 分帳**，`attempts`／`answers` 兩張表都有 `email` 欄且所有查詢都以它
過濾。`store.js` 的 `open()` 帶一段一次性遷移：偵測到舊表沒有 `email` 欄就 DROP 重建
（`CREATE TABLE IF NOT EXISTS` 不會改動既有表，留著會讓 INSERT 直接失敗）。

`index.html` 由 `高業考古/高業考古題-高頻觀念分析/scripts/s10_practice.py` 產生，
**不要手改**，改模板 `scripts/practice_template.html` 後重跑該腳本。

#### cosmetics-codex（日韓化妝品圖鑑，2026-09-03 新增）

`projects/cosmetics-codex/`：166 筆日韓彩妝／保養／香水／美髮美體的產品資料庫，
卡片牆＋篩選側欄前端＋零 npm 依賴的 Node 後端。公開瀏覽、無登入閘門。
volume `cosmetics-codex-volume` 掛在 `/data`。

**已上架 OutsideFramework Works**（2026-09-03）：`data-cat="knowledge"` 的「知識庫
Knowledge Base」小節，緊接在文章資料庫之後，`data-published="2026-09-03"`。卡片**沒有**
包在 `<!--GATE-->` 裡（站本身公開），但 Knowledge 這顆分類按鈕在 GATE 內——所以
**訪客在 Works 頁選不到 Knowledge 分類，只會在 Timeline 看到它**；要讓訪客也能從 Works
點進去，得把那顆 pill 移出 GATE（文章資料庫目前也是同樣情況，屬既有設計而非疏漏）。

- **寫入 API 預設開放**（知道網址就能改資料）。要擋住陌生人，設 service 變數
  `EDIT_TOKEN=<任意字串>`，之後 POST/PUT/DELETE 都必須帶 `x-edit-token`，前端會自動
  跳出輸入框。沒有走 ofw 中央登入，因此不必動 `auth.js` 那五份副本。
- **`data/part*.json` 是種子，不是資料來源**。`store.js` 只在 products 表為空時匯入，
  之後以 volume 裡的 SQLite 為準，不會覆蓋線上編輯的內容；要強制重匯設 `COSMETICS_RESEED=1`。
  改了種子檔記得線上不會自動跟著變（`tools/apply-images.js` 就是為了同時更新兩邊）。
- **`PUT /api/products/:id` 是整筆覆寫**：`store.upsert` 會把沒帶到的欄位設成 null，
  所以任何部分更新都必須先 GET 再合併再 PUT。
- 本機開發用 `COSMETICS_DATA_DIR` 把資料庫導到專案內的 `.localdata/`（已進 .gitignore），
  否則 `/data` 在 Windows 會解析成 `C:\data`。launch.json 名稱 `cosmetics-codex` :8183。

**圖片走品牌官網外連**，載入失敗時前端自動退回色卡卡片——那是預期行為不是 bug。
`tools/` 下有三支採集工具（`survey-sites.js` 快篩可採性、`harvest-images.js` 採集、
`apply-images.js` 套用）。實測品牌站分三類，只有前兩類抓得到圖：① Shopify 站
（`/products.json`）② 靜態站且 og:image 就是商品圖 ③ og:image 是全站通用圖或商品圖由
JS 載入（CANMAKE、ETUDE、CEZANNE、SHISEIDO…）。第三類唯一的例外是 CANMAKE：它是
WordPress 且 `wp-json` 開放，但**前 13 大品牌網域裡只有它是 WordPress，這招不通用**。
採集器刻意只產出提案不直接寫入，因為品名相似度比對必然誤判（實際踩過：配到男士線、
mini 版、組合包、品牌 banner、上妝後膚色特寫），**套用前一律把圖下載目視確認**。

### article-db is ALSO a two-repo split — same trap as globe-invest
`article_db/index.html` in this repo is the **dev-source copy only**. Railway's
`articlebase.up.railway.app` actually builds from a wholly separate GitHub repo,
`Playboy769/article-db-api.git`, which this repo already has registered as git remote
`article-db` (check with `git remote -v`). That repo has its own independent history — it is
**not** a fork or mirror set up via CI, just two copies kept in sync by hand.

Discovered 2026-07-07: a bugfix was committed and pushed to `origin` (this repo) and looked done,
but `articlebase.up.railway.app` never updated because Railway deploys from `article-db-api`, not
from this repo. Confirmed by diffing `article_db/index.html` against
`git show article-db/main:index.html` — they were byte-identical (mod line endings) except for
the missing fix, meaning the two copies really had been hand-synced up to that point.

**Whenever you edit `article_db/index.html`, `app.py` or `auth.py`: after committing here, run
`scripts/sync-article-db.ps1`** to push the same content to the `article-db` remote's `main`
branch — otherwise the live site silently stays on the old version indefinitely. `auth.py` was
added to the script's scope on 2026-07-30: it had been excluded, which meant syncing `app.py`
alone could ship a call into a function the deployed `auth.py` didn't define and 500 every
request. The script
fetches `article-db`, checks out its `main` into a throwaway worktree, copies the file in, and
lets `git status`/`git diff` inside that worktree (not a hand-extracted blob comparison) decide
whether there's a real change — so line-ending normalization follows the target repo's own git
config instead of getting mangled. It shows a diff and prompts before pushing; pass `-DryRun` to
only preview. It does not touch or push `origin` — commit here as usual first.

**`app/GlobalEco`, `app/InvestFrame`, `app/CausalFrame` are dev-source copies only — the
Dockerfile above does NOT deploy them.** The versions that actually go live are the mirrored
copies inside `globe-invest/app/`. Editing one side and forgetting the other is a real,
already-happened bug: an unrelated "fix sea routes" commit once silently dropped a commodity
and mislabeled another in GlobalEco's oil-price panel, and it went unnoticed for ~2 weeks
because nothing diffed the two repos.

**Whenever you edit GlobalEco / InvestFrame / CausalFrame: run
`scripts/sync-globe-invest.ps1`** to copy the change into `globe-invest/app/`, then commit +
push **both** repos (`origin` here, and `globe-invest`'s own `origin`). Don't rely on memory to
keep them in sync — the script diffs before copying.

### Known divergence — do not blind-sync these
`globe-invest/app/options`, `/warning`, and `/research` also started as copies of root-level
files (`options_guide.html` [now `projects/options_guide.html`], the old warning-radar HTML
[now `projects/market-warning-radar/`], and the MU analysis report [now
`research/mu-analysis-2026q3/`]). As of 2026-07-01, `options` is still identical; `warning`
and `research` have diverged into genuinely different content (different structure/theme, not
just "a few commits behind"). Don't add these to the sync script or overwrite either side until
a human reconciles which version is current.

### Another one: projects/brownian-motion-simulator
`globe-invest/app/brownian/index.html` (added 2026-07-01, deployed as the "Brownian Motion" work
on the OutsideFramework Works page) is a copy of `projects/brownian-motion-simulator/index.html`.
Same rule as above — it's **not** in the sync script. Edit the `projects/` copy for dev/testing,
then manually re-copy to `globe-invest/app/brownian/index.html` if it changes, and commit + push
both repos.

## Directory map

- `app/` — dev-source for the four "outside framework" apps (see table above)
- `globe-invest/`, `my-slide/`, `structural_holes/`, `sector-rotation-system/`, `twchips-report/`
  — **separate git repos**, nested here for convenience. Never `git add` their contents into
  this repo; they manage their own history.
- `twchips/` — **not this repo's or the user's project**: a read-only local reference clone of
  the third-party library [catcat222222/twchips](https://github.com/catcat222222/twchips) that
  `twchips-report/` depends on via pip (`requirements.txt`: `git+https://github.com/catcat222222/
  twchips`). Kept purely for reading/debugging its source — never `git add` it, and there is no
  Railway deployment tied to this folder (that's `twchips-report/`).
- `article_db/` — tracked in this repo as dev-source, but deployed from the separate
  `article-db-api` repo (remote `article-db`) to `articlebase.up.railway.app` — see deployment
  topology section above, must be pushed to both
- `projects/` — standalone tools/apps not part of the outside-framework family (stock analyzer,
  tech value screener, food calorie lookup, market warning radar, options guide, brownian motion
  simulator — the last one is also deployed via `globe-invest/app/brownian`, see above).
  ⚠️ 這個資料夾底下**有兩個是真的有部署的服務**，不要當成純本機工具：
  `gaoye-mock-exam/` 與 `cosmetics-codex/`，兩者都靠 `railway up` 上傳（見上方拓撲）
- `research/` — one-off analysis writeups (e.g. earnings-call breakdowns), not living apps
- `archive/` — superseded/legacy material kept for reference, not maintained
- `RR4/`, `RR5/`, `EMA Bias Model/` — active personal trading/VBA projects, not part of the web
  app suite. Left at the repo root deliberately — do not reorganize without asking.
- `shared-vba/` — VBA modules shared **verbatim** between the standalone `SEC-Filing-Fetcher`
  workbook and the RR4 workbook's "Company research" integration
  (`modHttp`/`modJsonUtil`/`modPrices` + the headless readers `modSECData`/`modMOPSData`).
  `shared-vba/` is the source of truth; push into both `.xlsm` with
  `scripts/sync-shared-vba.ps1` (programmatic `CodeModule.AddFromString`, **never** the VBE
  "Import File" menu — ANSI mojibake). See `shared-vba/README.md` and the plan/audit under
  `RR4/company-research-*`.
- `scripts/` — maintenance scripts: `sync-globe-invest.ps1` (GlobalEco/InvestFrame/CausalFrame →
  globe-invest/app/), `sync-article-db.ps1` (article_db/index.html → article-db-api remote),
  `sync-shared-vba.ps1` (`shared-vba/*.bas` → SEC-Filing-Fetcher + RR4 workbooks)

## Commit hygiene

Keep commits scoped to one concern. Several regressions in this repo's history came from
otherwise-correct commits that also carried an unrelated, unreviewed change (e.g. a routing-fix
commit that happened to rewrite the oil-price panel and dropped a field). If a change touches an
unrelated file, split it into its own commit even when working fast.

## Workflow preference

This repo is iterated on fast and pushed directly to the default branch on both `origin` and
`globe-invest` — note `globe-invest`'s default branch is **`master`**, not `main`
(`structural-holes` and `article-db` both use `main`) —
that's a deliberate choice for solo-project velocity, not an oversight. Don't propose branch
protection or PR gating; instead lean on the sync script above and `globe-invest`'s CI
(`.github/workflows/ci.yml`, runs `node --check` on push) as lightweight safety nets that don't
block pushes.

## Earnings Call 分析框架（常態功能）

**觸發**：使用者說「分析 XX earnings call」（或同義說法）時，依本框架執行。範本源自
`research/jnj-analysis-2026q2/JNJ_FY2026Q2_Analysis/JNJ_FY2026Q2_Analysis.html`（以 JNJ 為準，
AAOI/ATI/AXTI 交叉驗證）。

> **範本評選紀錄（2026-08-12）**：舊範本 MU 經複查發現供應鏈 tab 完全缺失（只有 9 個
> tab，且是框架自己的範本源）。派 10 個子代理對當時已確認 10-tab 完整的候選報告逐篇評分
> （標準：① 10-tab 結構完整度＋6 張必備圖全齊、② 分析深度與雙層互證/留白反推落實程度、
> ③ 視覺與版面規範遵循度，各 1–10 分），JNJ 27/30 奪冠（AAOI、ATI 並列 26/30 次之，
> AXTI 25/30）。評選同時揪出一個系統性問題：10 篇候選裡有 8 篇的區塊 margin-bottom
> 實際只有 10–14px，明顯低於下方視覺鐵則規定的 16–20px——只有 JNJ／ATI 兩篇真正達標，
> 見下方視覺鐵則一節的具體參考值。MU 本身已比照下方「7 家修正至 10-tab」流程補齊供應鏈
> tab，見鏡像對表。

### 輸入與輸出

- **輸入**（放進 `research/<ticker>-analysis-<fyQ>/` 根目錄）：
  `<TICKER> Qn Earnings Call Script.md`（逐字稿）＋ `<TICKER> Qn 10-Q Key Data.md`
  （財報關鍵數據；標準配置）；如有法說會簡報 PDF 一併保存。
- **SEC filings 交叉查詢（製作前必做）**：先自動從 SEC EDGAR 抓該公司的
  Form 4（內部人交易）、10-Q、10-K（外國發行人改抓 20-F），抓不到或非 SEC 申報公司
  才詢問使用者是否能提供。EDGAR 用官方 API
  （`https://data.sec.gov/submissions/CIK##########.json` 與 full-text search），
  查詢時帶 User-Agent。查到的檔案摘要存入同資料夾（如 `<TICKER> 10-K Supply Chain Notes.md`）。
  - Form 4 → 法說會前後的內部人買賣納入**風險矩陣/訊號解讀**（管理層信心交叉訊號）。
  - 10-K/20-F/10-Q → 供「供應鏈」tab 與財報數字覆核。
  - **若標的為美國公司或 ADR**：額外抓過去 6 期 10-Q（非僅本季），拉出逐季趨勢做財務比較
    （營收/毛利率/營業利益率/現金流等關鍵科目），不要只停留在單季 QoQ／YoY 兩個對比點。
- **台股標的的資料抓取（2026-08-14 起）**：標的為台股上市/上櫃公司時，沒有 SEC filings 可用，
  改走下面兩個來源，**兩個都要用、不是二選一**。這條同時適用完整 10-tab 報告與 Fallback 模式。

  1. **MOPS 公開資訊觀測站（主來源）** — 用 WebFetch 直接讀線上頁面即可，**不需要下載檔案**：
     - **合併財報（四大表＋全部附註）**：
       `https://mopsov.twse.com.tw/server-java/t164sb01?step=1&CO_ID=<代號>&SYEAR=<西元年>&SSEASON=<1-4>&REPORT_ID=C`
       —— 實測 3017 從 FY2022 到 2026Q2 逐季逐年都抓得到，含租賃、承諾及或有事項、
       關係人交易、合約負債等附註，量化五個 tab 與六項必備追蹤都靠這支。
     - **月營收**：`https://mopsov.twse.com.tw/nas/t21/sii/t21sc03_<民國年>_<月>.html`
       （上櫃改 `/otc/`）。民國年＝西元年−1911。
     - **年報（股東會年報）**：供「主要客戶及供應商」（占進貨/銷貨 10% 以上）、產銷量值、
       生產設備、轉投資 —— 供應鏈 tab 的主力來源。⚠️ 台灣年報慣例把客戶匿名成
       「客戶A/客戶B」，拼不出真名時標註信心度，不要用媒體推測冒充 filing 揭露。
     - **內部人持股轉讓事前申報**：等同美股 Form 4，一樣納入風險矩陣的管理層信心訊號。
  2. **政府資料開放平台 / TWSE Open API（輔助交叉驗證）** — `https://openapi.twse.com.tw/v1/`，
     如 `/opendata/t187ap06_L_ci`（上市綜合損益表）、`/opendata/t187ap07_L_ci`（資產負債表）、
     `/opendata/t187ap05_L`（月營收彙總）。用途是**覆核 MOPS 抓到的數字**，不是取代它。

  ⚠️ **兩個已實測踩過的坑，抓取時一律照做**：
  - **openapi 的全市場單一 JSON 會被 WebFetch 截斷**：實測 `t187ap05_L` 只讀到公司代號
    2374 就斷了，3017 完全沒出現在回應裡。**用 WebFetch 讀這類全表端點查個股就是不行**，
    個股一律走 MOPS 對應頁面。
    > ⚠️ **2026-09-05 更正**：這是 **WebFetch 的限制，不是端點本身的**。同一天實測
    > `curl -s "https://openapi.twse.com.tw/v1/opendata/t187ap45_L" -o t45.json` 抓下
    > 2.4MB 完整回應、1,226 筆全在，3017 也查得到。所以要用這類全表端點時，
    > **改成 curl 落地成檔再本地解析**（注意 Windows 主控台會用 cp950 解碼，Python 讀檔
    > 要明確指定 `encoding='utf-8'`），不要因為這條就直接放棄 openapi。
  - **MOPS 寬表逐欄擷取會串行**：實測 3017 月營收頁，「當月營收」與「累計」欄正確，但
    「去年同月營收／去年同月增減%」被讀成隔壁列的數字（回報 +5.51%，實際約 +57.4%）。
    **每個單月數字都要用累計欄反算交叉驗證**（本月累計 − 上月累計 ＝ 本月單月），
    對不上就重抓或改用另一個來源，不要直接把擷取結果寫進報告。

- **台股月營收輔助判斷（2026-08-14 起，台股標的必做）**：台股有美股沒有的月頻揭露，
  且公告時間遠早於季報，是「Earnings Cycle 概念框架」裡最接近 Execution 當下的可觀察訊號。
  台股報告一律額外做這一段，放在**損益表 tab**（與逐季趨勢圖並列）：
  - 抓**季報主軸期別往前推 24 個月**的月營收，畫成 `line-chart`（單月營收走左軸、
    YoY% 走右軸），⚑ 解讀註明資料來源為 MOPS 月營收頁哪一年月。
  - **季報結束後、報告產出前已公告的月份要一併納入**，並明確標示「這是季報期間之後的
    最新月份，尚未被任何一期財報涵蓋」—— 這正是超額資訊所在，不要因為它不在財報裡就略過。
  - 至少要回答一個問題：**最新月營收的動能，是支持還是打臉管理層在法說會給的下季展望？**
    這條要收進「估值觀察」tab 的互證對照表，左欄管理層展望、右欄月營收實績。

- **輸出一份合併的單檔 HTML**（tab 導覽、沿用既有 CSS 元件庫：
  `.tabs`/`sw()`、`.card`、`.stat-box`、`.risk-item`/`.risk-badge`、`.qa-item`/`toggleQA()`、
  `.kpi-row`/`.kpi`、`tr.section-head`、`.bu-card`、`.cf-row`、`.bar-row`）：
  - `<TICKER>_FYxxQn_Analysis/<TICKER>_FYxxQn_Analysis.html` —
    **質化 5 tabs ＋ 量化 5 tabs 共 10 個 tab，全部放在同一個檔案裡**

> **合併是強制步驟，不是選項。** 質化層與量化層可以分開起草（分開寫比較好控制品質），
> 但**交付前必須合併成上面那一個 `_Analysis.html`**，且合併後**不得留下獨立的
> `_Financials.html`**——研究資料夾與 `globe-invest/app/research/` 兩邊都只能有一個檔案。
> 理由有三：① `globe-invest/app/research/` 十幾篇既有報告全部是單檔 10 tabs，兩檔版本會
> 破壞一致性；② OutsideFramework 的 Works 卡片**一張只指向一個 URL**，拆兩檔會有一半內容
> 沒有入口、永遠不會被點到；③ 本框架的核心要求「雙層互證」（見下方分析手法）在讀者需要
> 切換檔案時形同虛設。
>
> **合併作業清單**（漏掉任一項就會壞版面或壞分頁）：
> 1. 把量化層 `<style>` 裡多出來的元件（`.kpi-row`/`.kpi*`、`tr.section-head`、`.bu-*`、
>    `.cf-*`、`.bar-*`）併進質化層的 `<style>`。
> 2. 5 個量化 tab 按鈕接在 `sw('risk',...)` 那顆之後，順序：損益表 → 業務部門 → 資產負債表
>    → 現金流量 → 估值觀察。
> 3. 5 個 `<div id="panel-...">` 貼在 `.footer` 之前，並把原本量化首頁的 `class="panel active"`
>    改回 `class="panel"`（`active` 全檔只能有一個，否則兩個 panel 會同時顯示）。
> 4. **class 名稱衝突自檢**：合併後全檔搜尋 `class="... panel"`，確認沒有任何非分頁元素
>    誤用 `panel` 這個 class——曾發生 `.tech-card.panel`（panel-level 封裝卡片）被
>    `.panel{display:none}` 吃掉、並被 `sw()` 當成分頁的狀況，已改名為 `.tech-card.pnl`。
> 5. 更新 `<title>`、`.page-header` 的 h1／p（p 要列出全部 10 個 tab），並**刪除兩檔時代的
>    互相指路文字**（「請見同系列 XXX_Financials」「詳見同系列 Analysis 報告的…」之類）。
> 6. 合併後在本地 launch.json port 逐一點過 10 個 tab，確認每次只有一個 panel 顯示、
>    Sankey 在切到「供應鏈」分頁後有畫出來、console 無錯誤、無橫向溢出。

### 質化層（tabs 1–5）

1. **Q&A 問答分析** — 開頭 3 個 stat-box（分析師提問數／核心問題數／邏輯缺口數）；
   每位分析師一張可折疊卡片：問題核心 → 管理層回應 → ⚑ 邏輯缺口/注意/觀察；
   卡片底部必附逐字稿原文引用框（溯源可覆核）。本 tab 的本質是**話術偵測**，不是摘要。
   **逐字稿原文引用框（`.orig-box`）規則**：每一位分析師的卡片都必須有，不可只做第一位、
   後面省略——曾發生過 GOOG Q1／Q2 兩篇報告都只有第一則 qa-item 附原文、其餘全部漏掉的
   狀況。原文要**逐字直接複製貼上**，不做摘錄／節錄／用「...」省略中段——分析師問題與
   管理層回答的完整交換都整段貼進去，一個字都不能改寫或精簡。摘要、解讀、邏輯缺口判斷
   都寫在 `.flag-box`/`.q-block` 裡，`.orig-box` 只放不經加工的原始逐字稿文字。
2. **本季特殊交易/事件深度拆解**（可替換槽：JNJ=Firefly Bio 併購整合、PANW=併購整合、
   AMD=認股權證、MU=SCA 長約）—
   四象限彩色卡片（規模/期限/財務承諾/尚待驗證）＋「核心疑問」盡職調查式追問清單收尾；
   有價格區間或利弊時用天花板/地板橫條與雙欄立場對照（對公司 vs 對客戶）。
   **資料中心專案融資條款（2026-08-13 起新增追蹤重點，泛用項目）**：若標的有重大資料
   中心/AI infra 資本支出，本 tab 額外納入融資結構（自有資本／合資 JV／SPV／租賃）、
   承諾金額、擔保條款、資產負債表認列方式（合併／表外）、對應的資本支出與自由現金流
   影響。來源：10-Q/10-K「Debt」或「Variable Interest Entities」或「Commitments」附註，
   輔以法說會逐字稿與簡報提及的合資/融資安排。非資料中心資本支出型公司則寫「無此資料/
   不適用」，不必額外著墨。
3. **技術與產品線（或業務平台）** — BU 營收橫條圖（QoQ%、佔比）＋分類卡片網格
   （技術節點或平台支柱，列時程/認證狀態）＋長期需求場景的因果鏈敘事。
4. **供應鏈** — 從 10-K/20-F/10-Q 提取的**全部**零組件供應鏈關係結構化列出：
   上游供應商（零組件/材料/設備，含集中度與單一來源風險）、下游客戶（含 10% 以上
   集中客戶揭露）、代工/封測外包關係；以上下游關係圖或分層卡片呈現，每項標注出處
   （哪份 filing 哪個 Item）。無 filing 可用時（非 SEC 申報公司、使用者無法提供），
   改用逐字稿＋公開資訊盡量拼出，並明確註記來源限制與信心度較低。
   - **關係圖畫法**：上下游關係（含營收流向/股權結構等其他 Sankey 適用的關係）用
     [projects/sankey-diagram-demo](projects/sankey-diagram-demo/index.html) 這個工具產生，
     不要每次手刻。流程：① 讀 filing 萃取節點/連結，整理成該工具吃的 JSON
     `{"nodes":[{"id","label","color"}],"links":[{"source","target","value"}]}`（value 用估計
     的相對規模即可，工具會自動依最大值縮放，不需要換算成固定像素）；② 開
     `sankey-diagram-demo`（launch.json 設定名稱同名，port 8137）的 JSON 分頁貼上/套用，
     檢查關係與顏色無誤；③ 在 JSON 分頁按「產生嵌入代碼」，複製產出的唯讀
     `<div>+<script>` 區塊；④ 貼進該公司 `Analysis.html` 的「供應鏈」tab，包在 `.card`
     裡即可 —— 這段嵌入代碼是自包含的，會自動偵測並注入 **D3（`d3@7` ＋
     `d3-sankey@0.12.3`，走 jsdelivr CDN）**，不需要另外修改 `Analysis.html` 的 `<head>`。
     每個節點/連結的數字來源仍要在旁邊文字註記出處（哪份 filing 哪個 Item），圖本身
     不取代那條「每表必註記」規則。
     > ⚠️ **chain-diagram 與 Sankey 兩張圖的節點必須是同一組實體**，不是「兩張都跟供應鏈
     > 有關」就算數——曾發生一篇報告 chain-diagram 畫「上游原料→產品線→下游客戶→終端
     > 應用」四層結構，Sankey 卻只畫「總營收→五大產品類別」的營收占比拆解，兩張圖談的
     > 根本是不同層級的實體，讀者會誤以為是兩套供應鏈。畫之前先確認兩張圖要用同一份
     > 節點清單。若供應鏈裡有法律/訴訟等例外支線（如產品責任訴訟牽動的原料商），可以在
     > chain-diagram 裡用不同顏色或虛線把那條支線從正常供應流視覺區隔開，不必為此拆成
     > 第三張圖。
     > ⚠️ 這個工具早期版本是 GoJS，2026-08 前的說明都寫 GoJS，**已經不對**。目前
     > 只剩兩處殘留：匯入時仍相容 GoJS 慣用的 `key/text/from/to/width` 欄位名，以及
     > 範例資料來源的註解。上面那組 JSON schema 本身沒變，照流程操作不受影響。
     > 另外要留意 Sankey 是這批工具裡**唯一**依賴外部 CDN 的，離線或 CDN 掛掉時
     > 圖會是空白；下面 chart-tools 那十一種則完全零外部依賴。
5. **風險矩陣**（收尾必備）— 高/中/低色彩徽章＋論證段落；至少一條是**獨創前瞻風險**
   （從分散答覆拼湊出的時間點/條款重疊，非管理層直接承認）；納入 Form 4 內部人
   交易訊號（若有）。

### 量化層（tabs 6–10，同檔）

損益表（KPI 四宮格 → 逐行 QoQ/YoY → GAAP/非GAAP 差異 → YTD → 財測 vs 實際 → 下年度展望）｜
業務部門（BU 卡片＋三季利潤率趨勢；**量價拆分**：出貨量 QoQ 與 ASP QoQ 分離）｜
資產負債表（三期別對比＋附註信評/股利/庫藏股）｜現金流量（四類＋FCF 橋接）｜
估值觀察（比率表每項附「意義」欄＋隱含 P/E 反推市場定價的懷疑程度）。

**業務部門 tab 不能只有財務數字**（2026-08-09 起）：每個 BU 卡片除了營收/margin
趨勢，還要詳細解釋該業務**產品本身是什麼、用在哪裡、為何客戶需要它**——例如
「高溫合金鍛件用於噴射引擎熱段，須耐受攝氏 XX 度並維持單晶結構」這種產品層級
的說明，而不是只停留在「BU 名稱＋營收數字」。寫法比照「業務平台與計畫」tab 的
技術卡片深度，但聚焦在財務數字背後對應的具體產品/技術，做詳細拆解與解析
（材料/製程/認證門檻/下游應用場景等），必要時自行繪製示意圖輔助理解（例如
產品在終端系統中的位置、製程流程、材料結構示意），不要求一定要用
chart-tools 的固定圖表類型，手繪 SVG 說明圖在此屬於例外允許。**這條規則要求的深度必須
落在業務部門 tab 本身**，不能寫在「技術與產品線」tab 就當作兩邊都滿足——曾發生報告把
產品層級解說整段放在技術與產品線 tab，業務部門 tab 卻只留財務數字卡片，形式上兩個 tab
都有內容、但業務部門 tab 本身沒有落實這條規則。

**手繪 SVG 示意圖必須自查文字是否重疊**（2026-08-09 起）：chart-tools 產生的圖表
會自動算文字寬度避開碰撞（見 `_shared.js` 的 `textWidth`/`truncateToWidth`），但手繪
SVG 沒有這層保護，元素排太密、標籤字數估太窄都會讓相鄰文字疊在一起變得無法閱讀
——曾發生核燃料棒剖面圖裡兩根相鄰的棒子中心距只有 25px，但各自的標籤有 4-5 個
中文字（換算約 40-50px 寬），兩組標籤因此左右重疊。畫完手繪示意圖後，**目測抓
每一組相鄰文字的中心點距離，換算成中文字數 × 字級大小估寬度**（CJK 字元寬度約
等於字級，例如 5 個中文字在 font-size 10 時約佔 50px），確認相鄰文字之間留有間隔
再定稿；元素排列太密就加大間距或拆成上下兩行標籤，不要硬塞在同一水平線上。

**資產負債表 tab 新增四項必備追蹤（2026-08-13 起）**：除既有三期別對比＋附註信評/股利/
庫藏股外，以下四項一律列入，缺資料時明確寫「無此資料/不適用」，不可略過不提：

- **應收帳款天數（DSO）**：DSO = 期末應收帳款（Accounts receivable, net）÷ 營收 × 天數
  （季報用 91、年報用 365）。來源：10-Q/10-K 資產負債表應收帳款餘額 ÷ 損益表營收；
  若揭露備抵壞帳（allowance for doubtful accounts），一併註記其占比變化。
- **未開始租賃（leases not yet commenced）**：來源：10-Q/10-K「Leases」附註（ASC 842）
  揭露的「lease that has not yet commenced」，記錄總金額與預計開始日期/期間，視為表外
  的未來義務先行揭露。
- **採購承諾（purchase commitments）**：來源：10-Q/10-K「Commitments and Contingencies」
  附註的 unconditional purchase obligations，依到期年度分層記錄（如 <1 年／1–3 年／
  3–5 年／3 年以上）；若承諾對象與「供應鏈」tab 已揭露的特定供應商重疊，兩個 tab
  互相加註交叉連結。
- **客戶預付款（customer prepayments）**：來源：10-Q/10-K 資產負債表「Contract
  liabilities」或「Deferred revenue」項目，及附註中 revenue recognition/contract
  balances 揭露的轉列營收時程，記錄逐季餘額變化。

四項的配圖規定見下方「圖表配置」表，一律標「視情況」——這是新增的追蹤重點，不改動
「分析完成後強制同步與上架」清單裡既有的 6 張必備圖。

**估值觀察 tab 新增利息保障倍數（2026-08-13 起）**：利息保障倍數 = EBIT（可用
Operating income，或 Net income + Interest expense + Tax expense 反推）÷ Interest
expense。來源：損益表營業利益與利息費用（10-Q/10-K 損益表或附註）。與既有比率表
同樣附「意義」欄，逐季列入趨勢，缺資料寫「無此資料」。

量化層各 tab 該搭什麼圖，見下方「圖表配置」專節（涵蓋 10 個 tab，不只量化層）。

### 圖表配置（10 個 tab 全適用，2026-08-09 起）

**這些圖一律用 `projects/chart-tools/` 的產生器做，不要手刻 SVG 或 CSS 長條。**
（本地 `chart-tools` :8158，入口頁列出全部十一種。手刻是這條偏好長期沒被落實的原因——
製表當下 46 篇報告裡只有 3 篇真的畫了折線圖。）

下表標 **必備** 的 tab **一定要有圖**，收尾清單會檢查；標「視情況」的依內容判斷。
**不要為了湊圖而畫沒有資訊量的圖**——一張只有兩根長條、旁邊文字已經講完的圖，
是在稀釋版面而不是幫助閱讀。

| tab | 畫什麼 | 工具 | 必備？ |
|---|---|---|---|
| 1 Q&A 問答分析 | 分析師提問主題分布（幾位問到同一件事） | `pie-chart` | 視情況 |
| 2 特殊交易/事件 | 合約期限、履約與交付時程 | `gantt-chart` | 視情況 |
| 2 特殊交易/事件 | 交易前後多面向對比（規模/期限/承諾/風險） | `radar-chart` | 視情況 |
| 2 特殊交易/事件 | 對價或財務承諾的金額拆解 | `waterfall-chart` | 視情況 |
| 2 特殊交易/事件 | 資料中心專案融資結構拆解（自有資本/JV/SPV/租賃） | `waterfall-chart` 或 `pie-chart` | 視情況 |
| 3 技術與產品線 | BU 營收橫條（QoQ%、佔比） | `bar-chart`（橫向） | 視情況 |
| 3 技術與產品線 | 技術節點、認證與量產時程 | `gantt-chart` | 視情況 |
| 4 供應鏈 | 上下游**結構**（誰供給誰、分幾層） | `chain-diagram` | **必備** |
| 4 供應鏈 | 上下游**流量**（線寬＝金額/比重） | `sankey-diagram-demo` | **必備** |
| 4 供應鏈 | 客戶或供應商集中度 | `pie-chart` 或 `heatmap-chart` | 視情況 |
| 5 風險矩陣 | 機率 × 衝擊矩陣 | `heatmap-chart` | 視情況 |
| 5 風險矩陣 | 各風險面向的強度輪廓（本季 vs 上季） | `radar-chart` | 視情況 |
| 6 損益表 | 逐季/逐年趨勢（金額左軸、比率右軸） | `line-chart` | **必備** |
| 6 損益表 | 近 24 個月月營收與 YoY%（**台股專用**，金額左軸、YoY% 右軸） | `line-chart` | **台股必備** |
| 6 損益表 | 營業利益 QoQ 橋接 | `waterfall-chart` | 視情況 |
| 6 損益表 | 成本結構逐季占比轉移 | `area-chart`（百分比堆疊） | 視情況 |
| 7 業務部門 | 各 BU 營收／利潤率對比 | `bar-chart` | **必備** |
| 7 業務部門 | 量價拆分（出貨量 QoQ vs ASP QoQ） | `bar-chart`（並排） | 視情況 |
| 7 業務部門 | 部門 × 季度利潤率矩陣 | `heatmap-chart` | 視情況 |
| 7 業務部門 | 營收結構逐季轉移（誰在長、誰在縮） | `area-chart` | 視情況 |
| 8 資產負債表 | 三期別科目對比 | `bar-chart`（並排） | 視情況 |
| 8 資產負債表 | 應收帳款天數（DSO）逐季趨勢 | `line-chart` | 視情況 |
| 8 資產負債表 | 採購承諾與未開始租賃合併到期時程 | `gantt-chart` | 視情況 |
| 8 資產負債表 | 客戶預付款（合約負債）逐季餘額變化 | `line-chart` 或 `area-chart` | 視情況 |
| 9 現金流量 | FCF 橋接 | `waterfall-chart` | **必備** |
| 9 現金流量 | 營運/投資/籌資逐季組成 | `area-chart` | 視情況 |
| 10 估值觀察 | 同業 P/E vs 成長率 | `scatter-chart` | **必備** |
| 10 估值觀察 | 同業估值倍數的分布與離散度 | `box-plot` | 視情況 |
| 10 估值觀察 | 利息保障倍數逐季趨勢 | `line-chart` | 視情況 |

**供應鏈 tab 兩張圖都要，分工要清楚**：`chain-diagram` 交代**結構**（有哪幾層、誰供給誰、
單一來源風險落在哪一層），Sankey 交代**流量**（線寬＝金額或比重）。兩張的節點命名要一致，
否則讀者會以為是兩套不同的供應鏈。沒有可量化流量時 Sankey 不要用猜的數字硬畫——
改在 `chain-diagram` 的連結標籤上寫明關係性質，並在圖說註明「無公開流量數據」。

**估值觀察 tab 的散點圖若真的找不到同業資料**：不要因此整張圖改題目就算了事。範本 JNJ、
交叉驗證的 AAOI／AXTI 都示範過同一個做法——先老實說明為什麼同業 P/E 比較沒有意義
（例如公司 GAAP 仍虧損、同業揭露的估值資料不可靠），**再補畫一張同業 P/E vs 成長率
散點圖並在圖說明確標註「此圖為框架要求之必備圖表」**，即使圖本身資訊量有限也照畫，
另外再加一張更有意義的替代圖（如公司自身估值倍數隨時間變化）。這樣兩頭都顧到：規格
要求的圖沒有被跳過，讀者也拿到真正有解讀價值的圖。不要只做替代圖、把必備圖整個換掉。

**風險矩陣的畫法**：列＝風險項、欄＝時間窗口或衝擊面向，格值＝強度（如 1–5），
色階選「綠→紅」。這張圖是現行高/中/低 `risk-badge` 卡片的**補充而非取代**——
矩陣負責一眼看出風險集中在哪一格，論證仍然寫在卡片裡。

**回溯範圍**：本節只適用於**新報告**。既有 46 篇不主動回頭補圖；日後若因其他原因
編輯到某一篇，才順手把那篇的手刻圖換成工具產出。不要為了套用本節而發起大規模改版。

#### 反查：手上有這種資料，該用哪個工具

上表是「哪個 tab 放什麼圖」，這張是反過來查——遇到上表沒列到的資料型態時看這張。

| 要畫什麼 | 用哪個工具 | 典型用途 |
|---|---|---|
| 逐季/逐年趨勢 | `chart-tools/line-chart.html` | 損益表 tab、10 季趨勢圖；金額走左軸、比率走右軸（右軸系列自動改虛線） |
| 類別比較 | `chart-tools/bar-chart.html` | BU 營收 QoQ、量價拆分、成本結構；並排／堆疊／百分比堆疊 × 直向／橫向 |
| 時程/區間重疊 | `chart-tools/gantt-chart.html` | 產能開出、客戶認證、長約履約期、專利到期；單一時間點用「里程碑」畫成菱形 |
| 占比分布 | `chart-tools/pie-chart.html` | 營收結構、成本結構、客戶集中度；中空比例可調 |
| 變動橋接 | `chart-tools/waterfall-chart.html` | 營業利益橋接、FCF 橋接、毛利率變動拆解 |
| 兩變數定位 | `chart-tools/scatter-chart.html` | 估值觀察 tab 的同業比較（P/E vs 成長率）；可自動用中位數分四象限，氣泡面積正比於第三變數 |
| 矩陣／密度 | `chart-tools/heatmap-chart.html` | 部門×季度矩陣、風險矩陣、供應商集中度；**可直接從 Excel 複製整塊貼上** |
| 分層關係 | `chart-tools/chain-diagram.html` | 供應鏈上下游、產業鏈情狀、組織架構、因果鏈 |
| 占比隨時間轉移 | `chart-tools/area-chart.html` | 營收結構逐季變化、成本組成、現金流組成；堆疊／百分比堆疊／重疊三種模式 |
| 多維度輪廓比較 | `chart-tools/radar-chart.html` | 同業競爭力多指標對比、風險面向評分；單位不同時切「各維度各自正規化」 |
| 分布與離散度 | `chart-tools/box-plot.html` | 同業估值倍數分布、毛利率離散度；也可只填五數摘要 |

> `.bar-row`／`.bar-fill` 這組手刻橫條在既有報告裡出現過 223 次，是最該改用 `bar-chart.html`
> 的一項；`.kpi-row` 四宮格則維持手刻（那是版面元件不是圖表，沒有做成工具）。

操作流程與 Sankey 一致：左側填資料 → 右側即時預覽 → 下方選輸出分頁 → 複製 → 貼進
`Analysis.html` 對應 tab 的 `.card` 裡。

**三項共通功能（2026-08-09 起；2026-08-12 新增的三個工具一併比照）**：

1. **貼上表格**：不要一格一格重打。直接從 10-Q／Excel 選取整塊複製，貼進工具的「貼上表格」
   欄位再按套用。解析器認得千分位、`%`、貨幣符號，以及 **`(1,234)` 括號負數**——財報的費用
   項用括號表示負數，沒處理的話整份損益表的費用會全部變成正的。`—` 與空白視為缺值。
   十一個工具裡只有 `chain-diagram` 例外（它走 JSON 而非表單）。
2. **跨圖配色一致**：同名系列在所有工具共用同一個顏色（存在 `localStorage`）。一篇報告六張
   圖以上，若不管它，「資料中心」會在每張圖拿到不同顏色。**換一篇報告時先按「清除配色記憶」**，
   否則會沿用上一家公司的對應。
3. **列印與色覺檢查**：每個工具的「配色檢查」面板會即時檢查目前顏色在灰階與三種色覺缺陷
   模擬下是否兩兩可辨，並提供「套用列印安全配色」。**預設調色盤是螢幕用的**，`#dc2626` 與
   `#2563eb` 灰階亮度只差 4，黑白列印會分不出來——報告要印或匯 PDF 時務必套用列印配色。
   列印調色盤前 6 色零問題；需要 7 個以上系列時面板會報警，該做的是拆圖而不是硬塞。

**兩種輸出擇一**：

- **靜態 SVG**：純 SVG，零 script、零外部連線。列印/PDF/離線都正常，預設用這個。
- **互動版**：同一張 SVG ＋ 一小段只做 hover 提示的 vanilla script，一樣零外部依賴。
  想讓讀者滑過去看逐點數值時才用。

兩者都不必改 `Analysis.html` 的 `<head>`，同一篇貼多張也不會互相干擾（每段嵌入代碼
帶自己的隨機 id）。因為圖形是實體 SVG 而不是 JS 當場算出來的，貼在還沒被點開的
`display:none` 分頁裡也不會有畫不出來的問題——這點跟 Sankey 不同，Sankey 是 JS 依
容器寬度重畫，才需要那套 observer/輪詢的保險。

⚠️ 圖本身**不取代**「每表必註記」規則：每張圖旁邊仍要有一條 ⚑ 人工解讀，並註明
數字出處（哪份 filing 哪個 Item）。

### 分析手法要求

- **雙層互證**：質化 tabs 的每個關鍵敘事判斷，須在量化 tabs 找到對應科目變動印證
  （例：SCA 押金 → 其他非流動負債暴增；「匯率貢獻 +0.6pp」→ 毛利率 QoQ 恰為 +0.6pp，
  即剝離匯率後為持平）。**「估值觀察」tab 收尾放一張互證對照表，這是強制交付物、不是
  視情況加選**——2026-08-12 評選新範本時發現，10 篇已確認 10-tab 完整的候選報告裡仍有
  2 篇整篇找不到這張表，對照關係雖然散落在各 tab 裡但沒有收攏成表，算沒有落實。左欄列
  管理層說法、右欄列對應的財報科目驗證。**範本 JNJ 的做法值得學：不是每一列都要驗證成功
  才能上表**——JNJ 把「法說會全程沒提到訴訟」這件事本身也當一列放進表裡，右欄對應
  10-Q 揭露的訴訟費用與準備金，等於把「管理層的沉默」也收進同一張驗證表，而不是只挑
  講出來、驗證得過的敘事上表。**新增範例（2026-08-13 起）**：管理層口頭強調「需求強勁、
  客戶提前鎖單」→ 客戶預付款/合約負債逐季餘額應同步放大；「擴產舉債但財務體質穩健」→
  利息保障倍數應維持健康區間而非惡化；「供應鏈」tab 揭露的長約供應商 → 「資產負債表」
  tab 的採購承諾金額應找得到對應分層數字；資料中心資本支出若對外宣稱「輕資產／合資
  分攤」→ 對應「特殊交易/事件」tab 揭露的融資結構應為表外或 JV 認列，而非全額併表負債。
- **留白反推**：把管理層「沒說的」當訊號（不拆客戶身份 → 反推客戶集中）；能反推的數字要反推。
- **每表必註記**：任何數字表格旁配一條 ⚑ 人工解讀，不讓表格自己說話。
- **Fallback 模式**（無逐字稿，如 SUMCO）：省略 Q&A tab、風險矩陣標題註明資料限制，
  退化為純財報＋財測推導。**台股標的走 Fallback 時，資料來源改用 MOPS ＋ 政府資料開放平台
  （見上方「台股標的的資料抓取」），並一律加做月營收輔助判斷**——少了逐字稿這層，月頻
  營收就是唯一能反推管理層口徑是否兌現的高頻證據，不做等於整份報告只剩落後指標。
- **視情況加選**：EN/繁中雙語切換（PANW/AMD 有）、營業利益橋接瀑布圖（SUMCO）、
  10 季趨勢圖、供給瓶頸嚴重度排序（DELL）、CEO 語言風格分析（PANW）。
- 頁尾標註「僅供資訊參考，非投資建議」。

### 視覺鐵則（2026-07-15 起適用所有報告）

**一律純白底＋高對比深色字**：整頁背景與所有卡片/區塊都用純白（`#fff`）底、
深色（`#111`–`#333`）文字。顏色只准用在：徽章（badge）、邊框、左側色條、小面積
強調元素。**禁止在飽和彩色底（橘/綠/紫/藍等）上放深色或低對比文字**——這曾造成
整批報告文字難以閱讀。既有深色主題報告遇到修改時應順手轉為白底。

**區塊之間一律留白，不准零間距貼在一起**（2026-08-09 起）：`.card`／`.tbl-wrap`／
`.note`／`.chart-wrap`／`.bu-grid`／`.tech-grid`／`.quad-grid`／`.pie-row`／
`.stat-row`／`.kpi-row`／`.risk-item` 這類區塊級元件，`margin-bottom` 一律抓
16–20px（`.divider` 前後留白更大，抓 20–22px）；曾發生 `.note` 只寫了
`margin-top`、漏了 `margin-bottom`，導致它與下一個標題／表格零間距貼在一起，
版面看起來很擠。新增元件時比照抓同一個量級的 `margin-bottom`，不要各自隨意
決定間距大小。

> ⚠️ **這條規則長期沒被落實**：2026-08-12 評選新範本時複查了 10 篇已確認 10-tab 完整
> 的候選報告，8 篇的區塊級 `margin-bottom` 實際只有 10–14px（`.divider` 常常只有
> 14px），系統性低於這裡寫的 16–20px／20–22px，只有範本 JNJ 與交叉驗證的 ATI 真正
> 達標。新報告或修改既有報告時，直接抄 JNJ／ATI 兩篇的實際 CSS 數值，不要憑感覺抓。

### 分析完成後強制同步與上架（缺一不可）

每次完成一份 earnings call 分析後，**收尾必做**下面五件事，做完框架本身的內容不算完工：

0. **合併**：確認質化 5 tabs 與量化 5 tabs 已合併成單一 `<TICKER>_FYxxQn_Analysis.html`
   （依上方「合併作業清單」六項逐項核對），且研究資料夾內**沒有殘留 `_Financials` 資料夾或檔案**。
   合併沒做完就往下走，後面的鏡像與上架都會是錯的。
0.5. **必備圖到齊＋互證對照表到齊**：對照「圖表配置」表逐項確認 6 張必備圖都在——供應鏈
   tab 的 `chain-diagram` 與 Sankey（兩張，**節點命名要一致**）、損益表的折線圖、業務部門的
   長條圖、現金流量的 FCF 橋接瀑布圖、估值觀察的散點圖。**台股標的再加第 7 張必備圖：
   損益表 tab 的近 24 個月月營收折線圖**（見「台股月營收輔助判斷」）。同時確認**每張圖旁都有一條 ⚑
   人工解讀並註明出處**（哪份 filing 哪個 Item），**估值觀察 tab 收尾有互證對照表**（見
   分析手法要求一節，這是強制交付物），以及**全篇沒有殘留手刻的 `.bar-row`／`.bar-fill`
   或手寫 `<svg>` 圖表**（版面元件如 `.kpi-row`、`.quad-grid` 不在此限，示意圖類例外見
   量化層一節）。缺圖或缺表就回頭補，不要先鏡像上架——線上與本地一旦分歧，後面兩步就白做了。
   同時確認**六項新增必備追蹤重點**（應收帳款天數、未開始租賃、採購承諾、利息保障倍數、
   資料中心專案融資條款［泛用，不適用時寫「無此資料/不適用」］、客戶預付款）皆已列入
   資產負債表／估值觀察／特殊交易-事件對應 tab，缺資料一律明確寫「無此資料/不適用」而
   非略過不提（此規則只適用新報告，見量化層一節，不回溯既有 46 篇）。
1. **鏡像**：把那一份 HTML 複製到 `globe-invest/app/research/`，檔名與資料夾內原檔名一致、直接攤平
   （不建 `<ticker>` 子資料夾），與現有十幾篇報告的擺法一致。這一步就是「Research 報告」鏡像對
   （見上方「鏡像對」表）唯一的同步方式，沒有 sync script。
2. **上架 OutsideFramework Works**：在 `app/OutsideFramework/index.html` 的 Works → Earnings Call
   底下，找到產業/類別最貼近的既有 `<div class="wk-group-title">` 分組（找不到就新建一組），
   仿照鄰近卡片格式新增一張 `.wk-card`，連結指向
   `https://globe-invest.up.railway.app/research/<TICKER>_..._Analysis.html`；同時把該公司的
   下次財報日期加進 `EARN_DATES` 陣列（依日期遞增排序插入正確位置），並在新卡片的 `.wk-info`
   內加對應的 `.wk-next-earn` 行（見該區塊上方的 Maintenance 註解）。**新卡片務必帶
   `data-published="YYYY-MM-DD"` 屬性**——Timeline 頁面（`#page-timeline`）沒有獨立資料來源，
   完全是 `buildTimelineRows()` 在讀取 `#page-works` 下所有帶 `data-published` 的 `.wk-card`
   自動產生排序清單，漏了這個屬性等於這篇報告不會出現在 Timeline。上架後應打開 Timeline
   頁面確認新報告有出現在清單最上方，而不是假設它會自動生效。
3. **兩邊都要 commit + push**：本 repo（origin）內，research 原始檔案（含新增的
   `.claude/launch.json` 本地預覽 port，若有）與 OutsideFramework 上架異動依 Commit hygiene
   規則切成獨立 commit；`globe-invest/app/research/` 的鏡像檔案則在 `globe-invest` 自己的 repo
   另開一個 commit。兩個 repo 都要 push——只 commit 不 push、或只做了鏡像沒上架、或上架了但
   忘記 push globe-invest，都算這篇報告還沒真正上線。

## Earnings Cycle 概念框架（白板筆記，2026-08-02）

這是 earnings call 分析背後的思考模型，不是操作步驟——上面「Earnings Call 分析框架」講的是
**怎麼做**一份分析，這裡講的是**為什麼**要那樣做。用途：分析時判斷哪個 tab 該深挖、哪段
管理層發言屬於「舊資訊」該打折扣、哪段才是真正有增量的 guidance。

核心論點：財報公布的數字是落後指標，超額資訊藏在「客戶意圖 → 下單承諾」與「CEO guidance →
下一輪客戶意圖」這兩段還沒被財報證實的區間裡。循環的 7 個步驟：

1. **FQ0 – Customer Intention**：起點是下游客戶產生購買意圖，這是需求鏈最早、最難觀察的一環。
2. **Place Order / Commitment**：意圖轉成正式下單、簽約承諾——這一步開始形成公司的 Backlog。
3. **Execution(s)**：公司交付訂單，才反映成財報上的 Revenue、EPS、Margin、市佔率；這些產出
   本質是過去累積的護城河（MOAT）與在手訂單（Backlog）撐出來的，不是當季新創造的。
4. **Press Release 校正**：headline numbers 公布時其實已經是數月前的舊資訊（訂單早已成立），
   但市場仍會立刻 re-price，造成股價當下的波動——這是市場對舊資訊即時反應的效率缺口。
5. **Earnings Call**：真正有增量資訊的地方是 CEO 對當下總體環境的 guidance/口風，而不是剛
   公布的舊數字。
6. **回到 Customer Intention（FQ1, FQ2, FQ3...）**：CEO 的 guidance 會回過頭去形塑/暗示下一
   季甚至未來好幾季的客戶意圖，形成迴圈。但這一步的性質**本質上是未經驗證的假設**——要等
   下一輪 Execution（步驟 3）真正發生才能證實，這是整個循環裡最大的不確定性缺口。
7. **下棋 / Critical Thinking / Assumptions**：收尾動作——不要被動接收前 6 步的資訊，要像
   下棋一樣主動質疑管理層沒明講的假設、往後推演好幾步，對步驟 6 的猜測做壓力測試（交叉比對
   供應鏈、上下游其他公司的說法），呼應「Earnings Call 分析框架」的**雙層互證**與**留白反推**。

## 500 字 Thesis 自檢清單（常態功能，2026-08-06 起適用）

**觸發**：使用者說「500字thesis」「500字版」（或指名某公司＋這兩個詞的組合，如「精材
500字thesis」）時，依本節執行。**觸發時不用套用本檔「Mandatory pre-task clarification」
的 4 題強制澄清**——這個觸發詞本身已經是明確指令，直接輸出即可，只有在使用者連目標公司
都沒指名時，才需要反問是哪一家。

**模板本體**：`research/_templates/500-word-investment-thesis-checklist-template.md`——
六段架構（一句話論點／商業模式與護城河／為什麼是現在／三個支撐數字／最大風險與證偽訊號／
決策與下一個驗證點），通用範圍（含財務／估值），與「Earnings Call 分析框架」10-tab 完整
報告及「產業轉型節點觀察模板」的不含財務範圍都不同，套用前先讀一次模板檔確認架構未被
之後的修訂調整過。

**素材優先順序**：
1. 先找 `research/<ticker或公司名>-analysis-*/` 底下**已有的完整分析報告或既有 thesis
   檔**（`*_Analysis.html`、`*_Investment_Thesis_*.md` 等），從裡面萃取濃縮成 500 字——
   不重新做研究，只做壓縮。
2. 若該公司完全沒有既有素材，才用 WebSearch／WebFetch 現查公開資訊補，比照本檔「產品線
   與產線產業結構」系列（3374／2059／8046／3711）遇到素材缺口時的做法。

**輸出**：直接在對話框給出 500 字版本，六段依序標註對應第幾點（比照模板檔裡精材 3374
的演練格式），不需要另存新檔、不需要鏡像上架。若字數明顯超支（如超過 550 字），依模板
「填寫規則」自我檢查是哪一段邏輯還沒收斂，而不是逐字砍到剛好 500。
