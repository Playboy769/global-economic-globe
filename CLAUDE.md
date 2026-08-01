# Claudecode — Repo Map & Rules

This directory holds several **unrelated** personal projects side by side. Read this before
moving, deleting, or assuming anything about what's "dead" — a lot of it looks like clutter
but has an active deployment behind it.

## Mandatory pre-task clarification

Before starting **any** task in this repo — regardless of how clear the request seems — first
ask the user at least 4 clarifying questions in a single `AskUserQuestion` call (that tool's
per-call max is 4, so this means using the full 4 slots) and wait for answers before doing any
work. This applies even when the request looks unambiguous. This overrides the general
"Auto Mode" bias toward proceeding without confirmation for this repo specifically.

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

### Token 規則（2026-07-30 安全稽核後，四個服務一致）

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

⚠️ **`auth.js` / `auth.py` 改動必須四個服務同步**，否則跨網域登入會斷。且因為 Railway 無法
原子性同時部署四個服務，切換簽章方案時**先推三個下游、最後推 ofw** —— 下游有迴圈斷路器，
這個順序讓中間窗口變成一次 503 說明頁；反過來推會變成瀏覽器無限重導向。

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

> ⚠️ `warning` 與 `research` 兩側已**實質分歧**（非落後幾個 commit），`sankey`/`options`/
> `brownian` 製表當下一致但同樣不在 sync script。同步前務必先 diff，勿盲 copy 覆蓋。

### 7. 獨立部署（各自 repo，本 repo 只是巢狀）
| 功能 | repo | folder | 部署檔 | URL | 本地 |
|---|---|---|---|---|---|
| article-db | 本 repo dev + `article-db-api.git`（remote `article-db`）| `article_db/` | `index.html`(前端) + `app.py`(FastAPI) | `articlebase.up.railway.app` | `article-db` :8127 |
| structural-holes | `structural-holes.git` | `structural_holes/` | `app.py`(uvicorn) + `graph.py` | `structural-holes-production.up.railway.app` | `structural-holes` :8129 |
| my-slide | `my-slide.git` | `my-slide/` | — | Netlify | — |

> **article-db 改法**：編 `article_db/index.html` → 本 repo commit → 跑
> `scripts/sync-article-db.ps1` 推到 `article-db` remote，否則線上不更新（見下方專節）。

### 8. 非部署工具／資料（本 repo，標準）
| 功能 | folder | 本地 port |
|---|---|---|
| 股票分析器 | `projects/stock-analyzer/` | — |
| 科技估值篩選 | `projects/tech-value-screener/` | — |
| 食物熱量查詢 | `projects/food-calorie-lookup/` | `food-calorie-lookup` :8131 |
| Earnings 分析（範本源 MU）| `research/mu-analysis-2026q3/`；各 ticker 一資料夾 | panw :8133 / sumco :8134 / dell :8135 / tsmc :8138 / ms :8139 / lrcx :8140 |
| SEC 抓取工具（VBA）| `SEC-Filing-Fetcher/`（`.bas`+`.xlsm`）| — |
| 產業結構圖 | `industry frame/`（PNG+SVG）| — |
| 圖庫素材（未被引用）| `ofwphoto/`（.jpg）| — |
| 交易/VBA 專案 | `RR4/`、`RR5/`、`EMA Bias Model/` | — |
| FinceptTerminal（空巢狀 repo，疑廢棄）| `FinceptTerminal/` | — |
| 舊版留存 | `archive/` | — |

## Deployment topology (this is the part that bites)

There are **four separate Railway deployments** sourced from **three separate git repos**, plus
this main repo itself:

| Deployment | Repo | Source path | Live URL |
|---|---|---|---|
| OutsideFramework (portfolio homepage + central auth) | this repo (`origin` = `global-economic-globe.git`) | `app/OutsideFramework/` + root `server.js`/`auth.js`/`analytics.js` → `Dockerfile` → **node** | `ofw.up.railway.app` |
| globe / invest / causal / brownian | **`globe-invest/` — its own repo** (`globe-invest.git`) | `globe-invest/app/{globe,invest,causal,brownian}/index.html` + `globe-invest/server.js` | `globe-invest.up.railway.app` |
| structural-holes | **`structural_holes/` — its own repo** | — | `structural-holes-production.up.railway.app` |
| article-db | **separate repo `article-db-api.git`** (this repo has it as remote `article-db`) | `index.html` at that repo's root — mirrors this repo's `article_db/index.html` | `articlebase.up.railway.app` |
| my-slide | **`my-slide/` — its own repo** (Netlify) | — | — |

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
- `globe-invest/`, `my-slide/`, `structural_holes/` — **separate git repos**, nested here for
  convenience. Never `git add` their contents into this repo; they manage their own history.
- `article_db/` — tracked in this repo as dev-source, but deployed from the separate
  `article-db-api` repo (remote `article-db`) to `articlebase.up.railway.app` — see deployment
  topology section above, must be pushed to both
- `projects/` — standalone tools/apps not part of the outside-framework family (stock analyzer,
  tech value screener, food calorie lookup, market warning radar, options guide, brownian motion
  simulator — the last one is also deployed via `globe-invest/app/brownian`, see above)
- `research/` — one-off analysis writeups (e.g. earnings-call breakdowns), not living apps
- `archive/` — superseded/legacy material kept for reference, not maintained
- `RR4/`, `RR5/`, `EMA Bias Model/` — active personal trading/VBA projects, not part of the web
  app suite. Left at the repo root deliberately — do not reorganize without asking.
- `scripts/` — maintenance scripts: `sync-globe-invest.ps1` (GlobalEco/InvestFrame/CausalFrame →
  globe-invest/app/), `sync-article-db.ps1` (article_db/index.html → article-db-api remote)

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
`research/mu-analysis-2026q3/`（以 MU 為準，PANW/DELL/SUMCO/AMD 交叉驗證）。

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
2. **本季特殊交易/事件深度拆解**（可替換槽：MU=SCA 長約、PANW=併購整合、AMD=認股權證）—
   四象限彩色卡片（規模/期限/財務承諾/尚待驗證）＋「核心疑問」盡職調查式追問清單收尾；
   有價格區間或利弊時用天花板/地板橫條與雙欄立場對照（對公司 vs 對客戶）。
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
     裡即可 —— 這段嵌入代碼是自包含的，會自動偵測並注入 GoJS，不需要另外修改
     `Analysis.html` 的 `<head>`。每個節點/連結的數字來源仍要在旁邊文字註記出處
     （哪份 filing 哪個 Item），圖本身不取代那條「每表必註記」規則。
5. **風險矩陣**（收尾必備）— 高/中/低色彩徽章＋論證段落；至少一條是**獨創前瞻風險**
   （從分散答覆拼湊出的時間點/條款重疊，非管理層直接承認）；納入 Form 4 內部人
   交易訊號（若有）。

### 量化層（tabs 6–10，同檔）

損益表（KPI 四宮格 → 逐行 QoQ/YoY → GAAP/非GAAP 差異 → YTD → 財測 vs 實際 → 下年度展望）｜
業務部門（BU 卡片＋三季利潤率趨勢；**量價拆分**：出貨量 QoQ 與 ASP QoQ 分離）｜
資產負債表（三期別對比＋附註信評/股利/庫藏股）｜現金流量（四類＋FCF 橋接）｜
估值觀察（比率表每項附「意義」欄＋隱含 P/E 反推市場定價的懷疑程度）。

**圖表輔助偏好**：損益表 tab 的逐季/逐年趨勢數字，盡量搭配折線圖輔助閱讀，不要只
放純表格；占比類數據（如營收結構、成本結構、客戶集中度等分布）盡量搭配圓餅圖。

### 分析手法要求

- **雙層互證**：質化 tabs 的每個關鍵敘事判斷，須在量化 tabs 找到對應科目變動印證
  （例：SCA 押金 → 其他非流動負債暴增；「匯率貢獻 +0.6pp」→ 毛利率 QoQ 恰為 +0.6pp，
  即剝離匯率後為持平）。**「估值觀察」tab 收尾放一張互證對照表**，左欄列管理層說法、
  右欄列對應的財報科目驗證，把散在十個 tab 的對照關係收在同一張表裡。
- **留白反推**：把管理層「沒說的」當訊號（不拆客戶身份 → 反推客戶集中）；能反推的數字要反推。
- **每表必註記**：任何數字表格旁配一條 ⚑ 人工解讀，不讓表格自己說話。
- **Fallback 模式**（無逐字稿，如 SUMCO）：省略 Q&A tab、風險矩陣標題註明資料限制，
  退化為純財報＋財測推導。
- **視情況加選**：EN/繁中雙語切換（PANW/AMD 有）、營業利益橋接瀑布圖（SUMCO）、
  10 季趨勢圖、供給瓶頸嚴重度排序（DELL）、CEO 語言風格分析（PANW）。
- 頁尾標註「僅供資訊參考，非投資建議」。

### 視覺鐵則（2026-07-15 起適用所有報告）

**一律純白底＋高對比深色字**：整頁背景與所有卡片/區塊都用純白（`#fff`）底、
深色（`#111`–`#333`）文字。顏色只准用在：徽章（badge）、邊框、左側色條、小面積
強調元素。**禁止在飽和彩色底（橘/綠/紫/藍等）上放深色或低對比文字**——這曾造成
整批報告文字難以閱讀。既有深色主題報告遇到修改時應順手轉為白底。

### 分析完成後強制同步與上架（缺一不可）

每次完成一份 earnings call 分析後，**收尾必做**下面四件事，做完框架本身的內容不算完工：

0. **合併**：確認質化 5 tabs 與量化 5 tabs 已合併成單一 `<TICKER>_FYxxQn_Analysis.html`
   （依上方「合併作業清單」六項逐項核對），且研究資料夾內**沒有殘留 `_Financials` 資料夾或檔案**。
   合併沒做完就往下走，後面的鏡像與上架都會是錯的。
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
