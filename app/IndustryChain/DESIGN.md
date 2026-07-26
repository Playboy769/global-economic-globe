# IndustryChain（產業鏈）— 設計文件

**狀態**：設計稿，未實作
**日期**：2026-07-26
**定位**：跨公司共用模組。不是某一篇 earnings 報告裡的 tab，而是一張所有報告共用的產業鏈全圖；
每篇 earnings 報告從自己的「供應鏈」tab 連過來，並用高亮標出該公司在鏈上的位置。

---

## 0. 現況盤點（2026-07-26 實際掃描 `globe-invest/app/research/`）

設計前先掃過全部 34 個 `*_Analysis.html`。結論直接改變了設計，先列出來：

### 0.1 供應鏈 tab 的覆蓋率只有 65%

| 狀態 | 篇數 | 檔案 |
|---|---|---|
| 有供應鏈 tab | 22 | AMZN, ARM, ASML, BE, CRWV, DELL, GOOG×2, INTC×2, LRCX, META, MSFT, NBIS, NVDA, ORCL, SNDK, STM, TSMC, TSM, TXN, WDC |
| **沒有供應鏈 tab** | 12 | ALAB, AMD, DDOG, GLW, GS, MRVL, MS, MU, NET, NOW, PANW, SUMCO |

`MU_FY26Q3` 是整套框架的範本源，但它**自己沒有供應鏈 tab**——這個 tab 是後來才加進框架的，
早期報告（MU/AMD/SUMCO/PANW）都在它之前。所以「從既有報告回填知識庫」不能假設每篇都有料。

### 0.2 tab id 有三種寫法

- `sw('chain',…)` → 10 篇（ARM, GOOG×2, INTC×2, MSFT, NBIS, NVDA, ORCL, STM）
- `sw('supply',…)` → 11 篇（AMZN, ASML, BE, CRWV, LRCX, META, SNDK, TSMC, TSM, TXN, WDC）
- `sw('ai',…)`「AI Backlog 與供應鏈」→ 1 篇（DELL，跟 backlog 合併成同一個 tab）

**影響**：第 8 節「在報告裡插入回鏈連結」不能寫一個 sed 一次改完，要逐篇認 panel id。

### 0.3 tab 內的呈現有三種格式

| 格式 | 篇數 | 說明 |
|---|---|---|
| `.chain-layer` ＋ `.node` 分層卡片 | 10 | 半結構化，`<div class="node conf-hi">名稱 — 角色 <span class="src">出處</span></div>`，**最好萃取** |
| Sankey 嵌入（`var DATA = {...}`） | 12 檔含 sankey（部分同時有分層） | JSON 內嵌，機器可讀，但語意是「流量」不是「環節」 |
| 自由排版卡片 | 其餘 | 只能人工讀 |

`.chain-layer` 那 10 篇的 `.node` 文字幾乎可以直接對應到本設計的 `player` 物件——
`conf-hi/mid/lo` 的信心徽章、`.src` 的出處註記，語意跟本設計完全一致。**這不是巧合，
是同一套規範的延伸，所以本設計刻意沿用同樣的 class 名與語意，讓回填是搬移而非重寫。**

### 0.4 28/34 篇落在同一條鏈上

把 34 篇按產業歸類：

- **AI 半導體產業鏈（28 篇）**：SUMCO, GLW, ASML, LRCX, TSMC, TSM, INTC×2, MU, SNDK, WDC,
  NVDA, AMD, MRVL, ALAB, ARM, STM, TXN, DELL, BE, MSFT, GOOG×2, AMZN, META, ORCL, CRWV, NBIS
- **鏈外（6 篇）**：DDOG, NOW, NET, PANW（應用軟體）、GS, MS（金融）

**這就是本模組的價值論證**：82% 的報告在講同一條鏈的不同片段，但目前它們彼此不知道對方存在。
第一條鏈就做 AI 半導體，一次點亮 28 篇。

### 0.5 已經萃取到的真實資料（可直接進知識庫）

掃描時順手撈出的、有出處的硬資料，P1 回填直接可用：

| 來源 | 內容 |
|---|---|
| TSMC_FY2025 | Customer A 佔 2025 淨營收 19%、Customer B 17%；轉投資 VIS ~19–20%、SSMC 38.8%、Xintec 41.0%、GUC 34.8% |
| TSM_FY2026Q2 | 2Q26 淨營收 NT$1,270,381M；HPC 66%(QoQ+20%)、Smartphone 22%(-4%)、IoT 5%、Auto 4%；北美客戶 78% |
| LRCX_FY2026Q3 | FY26Q3 營收 $5.84B；Foundry 54% / DRAM 27%(創紀錄) / NAND 12% / Logic 7%；廠區：馬來西亞（最大，籌建二廠）、Fremont/Livermore CA、Tualatin OR、Yongin KR、台灣、Salzburg/Villach AT、Ohio |
| SNDK_FY26Q3 | Flash Ventures（與 Kioxia 合資，持股 50%）；Datacenter $14.67億(+233% QoQ)、Edge $36.63億(+118%)、Consumer $8.20億(-10%) |
| WDC_FY26Q3 | 自製讀寫磁頭＋磁性介質（San Jose／Fremont 自有產能）；組裝測試泰／馬／菲／中；Cloud $29.72億佔 89%(+48% YoY) |
| NVDA_FY27Q1 | 具名供應商 TSMC / Samsung / SK Hynix / Micron / Foxconn / Wistron / Fabrinet（10-K Item 1）；Groq 非專屬 IP 授權應計 $39.57億（10-Q 附註 7）；3 家未具名直接客戶佔 21%/17%/16%（10-Q 附註 13）；Compute&Networking $745.50億(+88% YoY)、Graphics $70.65億(+58%) |
| INTC_FY26Q2 | 分層：前端製程節點／先進封裝／ASIC／AI 基礎設施夥伴生態 |

---

## 1. 設計目標與非目標

### 目標

1. **累積**：每篇報告的供應鏈萃取結果進同一個知識庫，第 29 篇只補差異而不是從零重畫
2. **定位**：任何一家公司都能被問「它在鏈上哪一層、上下游是誰、誰是同層對手」
3. **利潤池**：回答「這條鏈的錢被哪一段賺走」——這是單篇報告結構上無法回答的問題
4. **時間差**：把「下游砍單→上游反映」的落後季數顯性化，用來判讀本季法說會處在循環哪一段
5. **可回溯**：每一個數字都有出處，跟報告同一套 `conf-hi/mid/lo` 信心標準

### 非目標（明確不做）

- ❌ 不取代各報告的 Sankey——那是「這家公司的金流」，本模組是「這個產業的結構」，兩者互補
- ❌ 不做即時報價／市值連動（那是 high-price 頁的事）
- ❌ 不做自動抓取（EDGAR 全文檢索自動找同業）——資料策展品質 > 覆蓋率，先人工
- ❌ 不整合 CausalFrame（第二階段再議）
- ❌ 不做多語系（既有報告只有 PANW/AMD/ASML 有雙語，非主流）

---

## 2. 檔案拓撲與部署

### 2.1 路徑

| 角色 | 路徑 |
|---|---|
| dev-source 頁面 | `app/IndustryChain/index.html` |
| dev-source 資料 | `app/IndustryChain/data/<chain-id>.json` |
| dev-source 索引 | `app/IndustryChain/data/index.json`（列出所有鏈，供下拉選單） |
| 部署鏡像 | `globe-invest/app/industry-chain/index.html` ＋ `data/*.json` |
| route | `/industry-chain` |
| 本地 dev | launch.json `industry-chain` :8141 |

### 2.2 server.js 需要的改動（精確）

已確認的事實：

- `MIME` 表（server.js:17）**已含 `.json`**，資料檔不需要改 MIME
- 靜態 fallback（server.js:911 `else fp = path.join(APP_DIR, url)`）**不會**把目錄補成
  `index.html`，`/industry-chain` 會 404
- `APP_DIR = '/app'`（server.js:10，Railway 容器內絕對路徑）

所以只要在 route 表加一行，位置接在 `/sankey` 那行之後：

```js
  else if (url === '/sankey'         || url === '/sankey/')         fp = path.join(APP_DIR, 'sankey',         'index.html');
+ else if (url === '/industry-chain' || url === '/industry-chain/') fp = path.join(APP_DIR, 'industry-chain', 'index.html');
```

`data/*.json` 走既有靜態 fallback 即可，不需要新 API。

**本地 dev 不需要這行**：`globe-invest-app` :8136 是 `python -m http.server --directory
globe-invest/app`，python 會自動對目錄回傳 index.html。這是個容易踩的陷阱——本地測起來正常、
上線 404。P3 上線後務必實際打一次 `https://globe-invest.up.railway.app/industry-chain`。

### 2.3 sync script 的改動

`scripts/sync-globe-invest.ps1` 的 `$pairs`（行 48–52）是單檔對單檔結構
（`Get-FileHash` 比對、`git diff --no-index` 兩個檔）。資料檔是多檔，有兩條路：

**方案 A（建議，改動小）**：`$pairs` 逐檔列出，鏈數量少時完全夠用
```powershell
@{ Name = 'IndustryChain';      Src = 'app/IndustryChain/index.html';                  Dst = 'app/industry-chain/index.html' },
@{ Name = 'IndustryChain:index';Src = 'app/IndustryChain/data/index.json';             Dst = 'app/industry-chain/data/index.json' },
@{ Name = 'IndustryChain:ai-semi'; Src = 'app/IndustryChain/data/ai-semiconductor.json'; Dst = 'app/industry-chain/data/ai-semiconductor.json' }
```
缺點：每新增一條鏈要改 script。但新增鏈本來就是罕見事件，而且忘了改會在 script 輸出裡
看得出來（新檔不會被列出）——**失敗模式是「沒同步到」而不是「同步錯」**，可接受。

**方案 B**：`$pairs` 增加可選的 `SrcDir`/`DstDir` 欄位，逐檔 enumerate。彈性好但要動 script
的核心迴圈與 commit path 收集邏輯（行 111–112 的 `$srcPaths`/`$dstPaths`），
且 `git commit -- <paths>` 要能吃展開後的清單。

**P3 先用方案 A**，等鏈數 >3 再重構成 B。

---

## 3. 資料模型

### 3.1 索引檔 `data/index.json`

```json
{
  "chains": [
    {
      "id": "ai-semiconductor",
      "name": "AI 半導體產業鏈",
      "file": "ai-semiconductor.json",
      "layerCount": 9,
      "playerCount": 46,
      "reportCount": 28,
      "updated": "2026-07-26"
    }
  ]
}
```

單獨拆出來的理由：下拉選單不需要載入全部鏈的完整資料。`playerCount` 等統計由 lint script
（第 11 節）在 commit 前重算寫回，不手動維護。

### 3.2 鏈檔完整 schema

```jsonc
{
  // ── 鏈層級 metadata ──────────────────────────────────────────
  "id": "ai-semiconductor",            // string, 必填, kebab-case, 同檔名
  "name": "AI 半導體產業鏈",             // string, 必填
  "updated": "2026-07-26",             // ISO date, 必填, 任何欄位變更都要更新
  "scopeNote": "涵蓋從矽晶圓到雲端算力租賃。含資料中心電力設備（供電是 2026 年的實質瓶頸）；不含終端消費電子、不含應用軟體層（DDOG/NOW/NET/PANW 雖有報告但不在本鏈）。",
  "excluded": [                        // 明寫排除，避免「為什麼 PANW 不在上面」的重複疑問
    { "tickers": ["DDOG","NOW","NET","PANW"], "reason": "應用軟體層，與矽/製造無直接投入產出關係" },
    { "tickers": ["GS","MS"], "reason": "金融業，不屬本鏈" }
  ],

  // ── 環節分層（陣列順序 = 上游→下游）────────────────────────────
  "layers": [
    {
      "id": "materials",               // string, 必填, 鏈內唯一
      "name": "矽晶圓／材料",
      "short": "材料",                  // 時間軸／堆疊條上的短標籤，≤4 字
      "desc": "12吋拋光片、磊晶片、光阻、特殊氣體、CMP 耗材",
      "accent": "#6b7280",             // 左側色條色相；低飽和，不當背景用

      // 價值分配 / 利潤池 ──────────────────────────────
      "economics": {
        "status": "estimate",          // enum: verified | estimate | todo
        "grossMargin": [18, 32],       // [低, 高] %，null 表示未知
        "revenueShare": 4,             // 佔全鏈營收 %，null 表示未知
        "profitShare": 2,              // 佔全鏈營業利益 %，null 表示未知
        "bargaining": "low",           // enum: high | mid | low
        "barrier": "資本密集＋2–3 年客戶認證週期；但產品標準化，終端定價權在下游",
        "note": "⚑ 營收佔 4% 但利潤只佔 2%——材料端吃資本支出、賺不到 AI 溢價。這是判斷「AI 行情有沒有外溢」最靈敏的一層。",
        "src": "SUMCO FY26Q1 決算說明會 p.12；MU FY26Q3 10-K Item 1",
        "asOf": "2026Q1"               // 資料期別，用於第 4.3 節的時效灰標
      },

      // 景氣傳導與時間差 ─────────────────────────────
      "cycle": {
        "lagQuarters": 3,              // 相對 transmission.anchor 的落後季數；負值=領先
        "signal": "12吋長約報價與稼動率",
        "direction": "lagging",        // enum: leading | coincident | lagging（可由 lagQuarters 推導，但明寫以防手誤）
        "note": "⚑ 下游砍單後約 3 季才反映到矽晶圓報價。看到材料端轉弱時循環已走完大半——這一層是確認訊號，不是預警訊號。",
        "src": "SUMCO FY26Q1 Analysis 市場環境與展望 tab",
        "asOf": "2026Q1"
      },

      // 環節玩家 ─────────────────────────────────────
      "players": [
        {
          "ticker": "SUMCO",           // string|null, null = 未上市/非美股代碼
          "name": "SUMCO 勝高",         // string, 必填
          "role": "12吋拋光片、磊晶片",   // string, 必填, 在本層扮演什麼
          "share": 24,                 // number|null, 該層市佔 %
          "confidence": "mid",         // enum: hi | mid | lo  ← 同報告的 conf-hi/mid/lo
          "src": "SUMCO FY26Q1 決算說明會 p.12",
          "asOf": "2026Q1",
          "report": "SUMCO_FY26Q1_Analysis.html",   // string|null, null = 尚未分析
          "tags": ["日本", "純度玩家"],              // string[], 選填, 供未來篩選

          // 只有寫該公司報告時才填 ─────────────────
          "focus": {
            "position": "單一環節純度玩家，無上下游延伸；營收 100% 來自矽晶圓",
            "moat": "12吋認證週期 2–3 年、客戶轉換成本高；但產品規格標準化，議價力仍在客戶端",
            "extending": "無明顯往上下游延伸的動作（對照 GlobalWafers 赴美設廠）",
            "threat": "中國 12吋產能 2027 起放量；長約到期後重新議價風險",
            "src": "SUMCO FY26Q1 Analysis 觀察與風險 tab",
            "asOf": "2026Q1"
          }
        }
      ]
    }
  ],

  // ── 景氣傳導基準 ─────────────────────────────────────────────
  "transmission": {
    "anchor": "cloud",                 // layer id，以此層為 T+0
    "anchorNote": "以 hyperscaler capex 決策為 T+0 起算。選 cloud 而非終端消費，是因為 2026 年的需求脈衝來自資料中心而非手機。",
    "note": "訂單自 hyperscaler capex 起算往上游逐層遞延；反向的砍單訊號傳導更快（約為順向的 2/3），因為取消比下單容易。",
    "src": "跨報告綜合：NVDA FY27Q1、TSM FY26Q2、LRCX FY26Q3、SUMCO FY26Q1"
  }
}
```

### 3.3 欄位規則彙整

| 規則 | 說明 |
|---|---|
| **`null` 是合法值，代表「filing 未揭露」** | UI 顯示 `—`，hover 提示「未揭露」。**禁止**為了讓圖好看而填猜測值 |
| **每個有數字的物件必須有 `src` 與 `asOf`** | 缺 `src` 在載入時 `console.warn`，lint script 直接 fail |
| **`status: "estimate"` 的欄位 UI 會加 `~` 前綴** | 例：`~4%`。估計值與申報值視覺上必須可區分 |
| **`confidence` 只描述「這筆事實對不對」** | 不描述「這家公司好不好」。`hi` = filing 具名、`mid` = 法說會口頭或推算、`lo` = 外部資料 |
| **`focus` 一個 player 只有一份** | 寫的是「這家公司自己的處境」，跟誰在看無關。多篇報告提到同一家時，以最新 `asOf` 為準 |
| **`layers` 陣列順序即上下游順序** | 不另設 `order` 欄位——陣列順序就是單一事實來源，避免兩者不一致 |
| **同一家公司可出現在多層** | 例：TSMC 在 foundry 與 packaging（CoWoS）都有。這是真實情況，不是資料錯誤。lint 只警告不擋 |

### 3.4 為什麼 `focus` 內嵌在 player 而不是另開檔

替代方案是 `data/focus/<TICKER>.json`。內嵌的理由：

- 寫報告時本來就在讀該公司 10-K，順手寫進 player 即可，不需維護第二份索引與交叉引用
- 讀取端不需要 N+1 次 fetch
- 代價：同一家被多篇提到時只有一份 `focus`。但如前述，`focus` 的語意本來就是公司自身屬性

反例是若未來要做「同一家在不同時間點的立場變化」時序，那才需要拆檔。目前不需要。

### 3.5 AI 半導體鏈的九層（P1 實際要建的結構）

| # | layer id | 名稱 | 代表玩家（★ = 已有報告） | 已有報告數 |
|---|---|---|---|---|
| 1 | `materials` | 矽晶圓／材料 | ★SUMCO、信越、GlobalWafers、★GLW(玻璃/光纖) | 2 |
| 2 | `equipment` | 半導體設備 | ★ASML、★LRCX、AMAT、TEL、KLA | 2 |
| 3 | `foundry` | 晶圓代工 | ★TSMC/TSM、★INTC(Foundry)、Samsung Foundry、VIS、SSMC、UMC | 4 |
| 4 | `memory` | 記憶體與儲存 | ★MU、★SNDK、★WDC、SK Hynix、Samsung、Kioxia | 3 |
| 5 | `packaging` | 先進封裝與組裝 | TSMC(CoWoS)、ASE、Amkor、Xintec、Foxconn、Wistron、Fabrinet | 0 |
| 6 | `silicon` | 晶片設計與 IP | ★NVDA、★AMD、★INTC、★MRVL、★ALAB、★ARM、★STM、★TXN、AVGO、GUC | 8 |
| 7 | `system` | 伺服器系統與網通 | ★DELL、★GLW(光連接)、SMCI、Foxconn、Fabrinet | 2 |
| 8 | `power` | 資料中心電力與散熱 | ★BE、Vertiv、Eaton、Schneider | 1 |
| 9 | `cloud` | 雲端與算力租賃 | ★MSFT、★GOOG、★AMZN、★META、★ORCL、★CRWV、★NBIS | 8 |

**第 5 層是全鏈唯一 0 篇報告的空白**——這本身就是模組的第一個發現：CoWoS 封裝產能是
NVDA 報告裡反覆被點名的瓶頸（Vera CPU 與 Rubin GPU 共用同一條 CoWoS 產線），
但整個報告庫沒有任何一篇是從封裝端視角寫的。UI 應該讓這種空白**看得出來**（見 5.3）。

---

## 4. 資料治理

### 4.1 信心分級（沿用報告既有標準）

| 級別 | class | 定義 | 例 |
|---|---|---|---|
| `hi` | `conf-hi` | filing 白紙黑字具名 | NVDA 10-K Item 1 具名 TSMC/Samsung/SK Hynix |
| `mid` | `conf-mid` | 法說會口頭提及，或由申報數字推算 | NVDA「3 家未具名客戶佔 21/17/16%」的身份推測 |
| `lo` | `conf-lo` | 外部公開資料、產業共識 | 市佔率排名 |

### 4.2 估計值的標註義務

`revenueShare` / `profitShare` **本質上是估計值**——filing 不會給「全鏈」的分母。
必須做到：

1. `economics.status = "estimate"`
2. UI 在數字前加 `~`
3. 頁面底部固定一段方法論說明：**「全鏈營收／利潤池以該層已上市玩家合計財報數字近似，
   未含未上市與非美股掛牌者，故各層佔比為相對量級而非精確市佔。」**

不做這個標註就違反既有報告的「每表必註記」鐵則。

### 4.3 時效衰減

市佔、毛利帶這類數字半年就過期。機制：

- 每個 `economics` / `cycle` / `player` / `focus` 物件帶 `asOf`（格式 `YYYYQn`）
- 頁面載入時以 `chain.updated` 推算當前季，凡 `asOf` 落後 **≥4 季**者，該元素右上顯示灰色
  `舊` 小標（`.ic-stale`），hover 顯示「資料為 2026Q1，可能已過期」
- 灰標只是提示，不隱藏資料——過期資料仍比沒有資料好

### 4.4 出處顯示

`.src` 預設以 11px `#888` 顯示在元素內（跟既有報告的 `.src` 一致），
「顯示出處」開關關閉時 `display:none`，只留 hover title。預設**開啟**——
這套報告的核心價值就是可回溯，藏起來是本末倒置。

---

## 5. 版面與視覺規格

沿用 2026-07-15 起的視覺鐵則：**純白底 `#fff`＋深色字 `#111`–`#333`；顏色只出現在徽章、
邊框、左側色條、小面積強調**。整頁零外部依賴（不用 D3／GoJS——三個檢視都能純 CSS 做出來）。

### 5.1 設計 token

```css
:root{
  --ic-bg:#fff;  --ic-fg:#111;  --ic-fg-2:#333;  --ic-fg-3:#666;  --ic-fg-4:#888;
  --ic-line:#e5e5e5;  --ic-line-2:#d4d4d4;
  --ic-focus:#111;              /* 聚焦公司邊框 */
  --ic-hi:#15803d;              /* conf-hi  綠 */
  --ic-mid:#c2600a;             /* conf-mid 橘 */
  --ic-lo:#94a3b8;              /* conf-lo  灰藍 */
  --ic-gap:12px;  --ic-radius:6px;
  font-family:'Noto Serif TC',serif;   /* 與報告同字體 */
}
```

層色 `accent` 由上游到下游走一條低飽和色相漸變（灰→藍→靛→紫→洋紅→橘→琥珀→綠→青），
**只用在 3px 左側色條與堆疊條**，絕不當背景。

### 5.2 整體版面

```
┌────────────────────────────────────────────────────────────────────────────┐
│  產業鏈 · AI 半導體                                    28/34 篇報告在本鏈上 │
│  [AI 半導體 ▾]  聚焦：[NVDA ▾] [×]  🔍___  ☑顯示出處 ☐只顯示已分析        │ ← sticky
├────────────────────────────────────────────────────────────────────────────┤
│  [ 環節分層 ]  [ 價值分配 ]  [ 景氣傳導 ]                                   │ ← .tabs / sw()
├────────────────────────────────────────────────────────────────────────────┤
│  （檢視內容）                                                               │
├────────────────────────────────────────────────────────────────────────────┤
│  方法論說明：全鏈營收／利潤池以該層已上市玩家合計財報近似……                  │
│  僅供資訊參考，非投資建議。                                                 │
└────────────────────────────────────────────────────────────────────────────┘
```

控制列 `position:sticky; top:0; background:#fff; border-bottom:1px solid var(--ic-line)`。
右上角的「28/34 篇報告在本鏈上」是即時算的，點擊展開列出鏈外 6 篇與其排除理由
（來自 `excluded`）。

### 5.3 檢視一：環節分層（主視圖）

每層一條橫帶，`display:grid; grid-template-columns:200px 1fr`：

```
┌──────────────────┬───────────────────────────────────────────────────────┐
│ ① 矽晶圓／材料    │ ┌─────────┐┌─────────┐┌─────────┐┌─────────┐        │
│                  │ │▎SUMCO   ││▎信越     ││▎GlobalW ││▎GLW     │        │
│ 毛利 18–32%      │ │ 12吋拋光││ 矽晶圓   ││ 矽晶圓  ││ 玻璃基板│        │
│ 議價力 低         │ │ 24% ●mid││ 27% ●lo ││ — ●lo   ││ — ●hi   │        │
│ ▌營收~4%         │ │ 報告 ↗  ││         ││         ││ 報告 ↗  │        │
│ ▌利潤~2%         │ └─────────┘└─────────┘└─────────┘└─────────┘        │
│                  │ ⚑ 營收佔 4% 但利潤只佔 2%——材料端吃資本支出、賺不到    │
│ 資本密集＋2–3年   │   AI 溢價。這是判斷「AI 行情有沒有外溢」最靈敏的一層。 │
│ 認證週期…         │   〔SUMCO FY26Q1 決算說明會 p.12〕                     │
├──────────────────┼───────────────────────────────────────────────────────┤
│ ⑤ 先進封裝與組裝  │ ┌─────────┐┌─────────┐┌─────────┐                    │
│                  │ │▎TSMC    ││▎ASE     ││▎Amkor   │  ← 全部灰字         │
│ ⚠ 本層 0 篇報告   │ │ CoWoS   ││ 封測     ││ 封測    │                    │
│                  │ └─────────┘└─────────┘└─────────┘                    │
│                  │ ⚑ 全鏈唯一無報告覆蓋的環節，而 CoWoS 產能是 NVDA 報告  │
│                  │   反覆點名的瓶頸——這是報告庫的結構性盲點。            │
├──────────────────┼───────────────────────────────────────────────────────┤
│ ⑥ 晶片設計與 IP   │ ┏━━━━━━━━━┓┌─────────┐┌─────────┐┌─────────┐        │
│                  │ ┃▎NVDA 本篇┃│▎AMD     ││▎MRVL    ││▎ALAB    │        │
│  ...             │ ┗━━━━━━━━━┛└─────────┘└─────────┘└─────────┘        │
│                  │ ┌───────────────────────────────────────────────┐    │
│                  │ │ ▸ 位置：獨佔加速運算平台層，同時往上（Vera CPU）│    │
│                  │ │   往下（DGX Cloud）延伸                        │    │
│                  │ │ ▸ 護城河：CUDA 生態＋晶片/系統/網通/軟體協同設計│    │
│                  │ │ ▸ 延伸中：Vera CPU 進入 ⑥ 層 CPU 區塊          │    │
│                  │ │ ▸ 威脅：客戶自研 ASIC（Google TPU、Amazon      │    │
│                  │ │   Trainium）；CoWoS 產能與 Rubin 互相排擠      │    │
│                  │ │   〔NVDA FY27Q1 Analysis 風險矩陣〕            │    │
│                  │ └───────────────────────────────────────────────┘    │
└──────────────────┴───────────────────────────────────────────────────────┘
```

規格：

- **層資訊欄（左 200px）**：層序圓圈數字、層名、毛利帶、議價力、營收/利潤佔比雙色條
  （`▌` 用 3px 高的 div，寬度 = 佔比 ×2px，同層兩條並排時落差一眼可見）、進入門檻文字
- **玩家卡片列（右）**：`display:flex; gap:12px; overflow-x:auto`，
  **只有這一列可橫向捲動，頁面 body 永不橫向捲動**
- **卡片 `.ic-card`**：白底、`1px solid var(--ic-line)`、`border-radius:6px`、
  `padding:10px 12px`、`min-width:140px`、左側 3px 層 accent 色條
- **聚焦卡片 `.ic-card.is-focus`**：`border:2px solid var(--ic-focus)`、
  右上 `.ic-badge-focus`「本篇」黑底白字小徽章
- **無報告卡片 `.ic-card.no-report`**：文字 `var(--ic-fg-3)`，
  「只顯示已分析」開關開啟時 `display:none`
- **信心點 `.ic-conf`**：6px 圓點 ＋ 文字，色用 `--ic-hi/mid/lo`
- **層解讀 `.ic-note`**：⚑ 開頭，`border-left:3px solid var(--ic-line-2)`，
  尾端接 `.src`
- **空白層警示**：該層 `players` 中 `report != null` 的數量為 0 時，
  層資訊欄顯示 `⚠ 本層 0 篇報告`，並強制顯示一條 `.ic-note`
- **focus 展開區**：聚焦公司所在層，卡片列下方插入 `.ic-focus-panel`，
  四行 position / moat / extending / threat，`1px dashed var(--ic-line-2)` 框

### 5.4 檢視二：價值分配

```
營收池  ┌────┬──────┬─────────────┬────────┬────┬───────────────┬──────┬───┬────────────┐
        │材料│ 設備 │   代工      │ 記憶體 │封裝│  晶片設計     │ 系統 │電力│   雲端     │
        │~4% │ ~7%  │   ~11%      │  ~9%   │~3% │    ~24%       │ ~8%  │~2%│   ~32%     │
        └────┴──────┴─────────────┴────────┴────┴───────────────┴──────┴───┴────────────┘
利潤池  ┌──┬─────┬──────────┬──────┬──┬──────────────────────┬───┬──┬──────────────┐
        │~2│ ~9  │   ~13    │ ~7   │~4│       ~38            │~4 │~2│    ~21       │
        └──┴─────┴──────────┴──────┴──┴──────────────────────┴───┴──┴──────────────┘
         ↑ 材料：營收 4% → 利潤 2%        ↑ 晶片設計：營收 24% → 利潤 38%
```

- 兩條 100% 堆疊橫條上下對照，**同一層在兩條上以同一 accent 色對齊**，
  寬度落差就是「誰在賺錢」的視覺答案
- 條下方一張表：層名 / 毛利帶 / 議價力 / 進入門檻 / **⚑ 解讀**（每列必有）
- 聚焦公司所在層：兩條上的該段加 `outline:2px solid #111`
- 所有 `estimate` 數字加 `~` 前綴；表格上方固定方法論說明
- **不畫圓餅圖**——比較長度比比較角度準，且堆疊條能同時呈現順序（＝上下游位置）

### 5.5 檢視三：景氣傳導

```
       領先 ←─────────────────  T+0  ─────────────────→ 落後
        T-1        T+0        T+1        T+2        T+3        T+4
    ─────┼──────────┼──────────┼──────────┼──────────┼──────────┼───→
         │          │          │          │          │          │
      [電力]     [雲端]     [系統]    [晶片設計]  [代工]      [材料]
                  ▲anchor              [記憶體]   [封裝]     [設備]

  層          落後   觀察訊號                        方向
  ─────────────────────────────────────────────────────────────
  電力        -1 季  資料中心併網排隊時程            leading
  雲端         0 季  hyperscaler capex 指引         anchor
  系統        +1 季  ODM 月營收、backlog 轉換率      lagging
  晶片設計    +1 季  資料中心營收 QoQ               lagging
  ...
  材料        +3 季  12吋長約報價與稼動率            lagging
    ⚑ 看到材料端轉弱時循環已走完大半——確認訊號，不是預警訊號。
```

- 水平軸 `T-1 … T+4`，每層依 `cycle.lagQuarters` 定位（CSS grid 欄位，非絕對定位）
- anchor 層加粗並標 `▲anchor`
- 同一刻度多層時垂直堆疊
- 軸下方表格列每層的 `signal` / `direction` / `note` / `src`
- 聚焦公司所在層整列反白（`background:#fafafa`，不是彩色）
- 頂端固定一句 `transmission.note`（含反向傳導較快的說明）

### 5.6 響應式

- 桌面 ≥1024px：如上
- 平板 768–1023px：層資訊欄縮到 160px，卡片 `min-width:120px`
- 手機 <768px：層資訊欄改為橫帶（層名＋佔比條一行），卡片列維持橫向捲動；
  價值分配的兩條堆疊條改為上下各自佔滿寬度並加標籤換行；景氣傳導軸改為垂直時間軸
- 所有寬內容（卡片列、堆疊條、時間軸）各自 `overflow-x:auto`，body 永不橫向捲動

### 5.7 列印

`@media print`：控制列與 tab 隱藏，三個檢視全部展開依序列印，卡片 `break-inside:avoid`，
連結後綴顯示 URL。既有報告都能列印，本頁不該例外。

---

## 6. 互動與狀態

### 6.1 URL 狀態

`/industry-chain?chain=ai-semiconductor&focus=NVDA&view=value&src=0&only=1`

| 參數 | 值 | 預設 |
|---|---|---|
| `chain` | 鏈 id | `index.json` 第一條 |
| `focus` | ticker | 無 |
| `view` | `layers` \| `value` \| `cycle` | `layers` |
| `src` | `0` \| `1` 顯示出處 | `1` |
| `only` | `0` \| `1` 只顯示已分析 | `0` |

任何互動都 `history.replaceState` 更新網址（不塞 history entry——這是篩選不是導航）。
效果：從報告連過來的深連結、與使用者調整後想貼回報告的網址，是同一套格式。

### 6.2 鍵盤

| 鍵 | 行為 |
|---|---|
| `1` / `2` / `3` | 切換三個檢視 |
| `/` | 聚焦搜尋框 |
| `Esc` | 清除聚焦、清空搜尋 |
| `↑` / `↓` | 在搜尋結果中移動；`Enter` 設為 focus |

### 6.3 搜尋

即時比對 `ticker`、`name`、`role`、`tags`，不分大小寫、支援中文子字串。
命中唯一結果時自動設 focus 並捲到該層。搜尋不過濾畫面（不是 filter，是 locator）。

### 6.4 聚焦行為

設定 `focus=NVDA` 時：

1. 該卡片加 `is-focus` 樣式與「本篇」徽章
2. 該層下方展開 `.ic-focus-panel`
3. 頁面捲動到該層（`scrollIntoView({block:'center'})`）
4. 價值分配／景氣傳導檢視中該層加 outline / 反白
5. 若該 ticker 出現在多層（如 TSMC），**全部**標記，focus panel 只在第一個出現的層展開

---

## 7. 前端架構

單檔 `index.html`，`<script>` 內以 IIFE 分區塊，不引入框架（與既有報告一致）。

```
index.html
├── <style>              設計 token ＋ 元件 CSS（約 400 行）
├── <body>
│   ├── header.ic-bar    控制列（sticky）
│   ├── nav.tabs         三檢視切換（沿用報告的 sw() 慣例）
│   ├── main
│   │   ├── #view-layers
│   │   ├── #view-value
│   │   └── #view-cycle
│   └── footer.ic-method 方法論＋免責
└── <script>
    ├── STATE            { chain, data, focus, view, showSrc, onlyReported }
    ├── loadIndex()      fetch data/index.json → 填下拉
    ├── loadChain(id)    fetch data/<id>.json → validate() → STATE.data
    ├── validate(d)      第 11 節的 client-side 檢查，console.warn 不擋渲染
    ├── parseUrl() / syncUrl()
    ├── renderLayers()   檢視一
    ├── renderValue()    檢視二
    ├── renderCycle()    檢視三
    ├── renderAll()      依 STATE.view 分派 ＋ 共用的 focus/篩選套用
    ├── el(tag, cls, txt)  極簡 DOM helper
    ├── srcSpan(o)       產生 .src，依 STATE.showSrc 決定顯示
    ├── staleTag(asOf)   第 4.3 節時效灰標
    └── fmtPct(v, est)   null → '—'；estimate → '~4%'
```

**渲染策略**：每次狀態變更整段重繪（`innerHTML = ''` 後重建）。資料量級是數十個節點，
不需要 diff。這與既有報告的做法一致，且讓狀態同步不會出錯。

**無外部請求**：字體用系統 fallback（`'Noto Serif TC', serif`），不引 Google Fonts——
既有報告已是此做法。

---

## 8. 與 earnings 報告的雙向連結

### 8.1 產業鏈 → 報告

卡片上的 `報告 ↗` → `/research/<player.report>`，`target="_blank"`。
本地 dev（:8141）時 `/research/` 不存在，需在 dev 環境把連結指向
`http://localhost:8136/research/...`——用 `location.port === '8141'` 判斷即可。

### 8.2 報告 → 產業鏈

在每篇報告的供應鏈 tab 頂端、既有解讀註記下方插入一行：

```html
<p class="note"><a href="https://globe-invest.up.railway.app/industry-chain?chain=ai-semiconductor&amp;focus=NVDA" target="_blank">↗ 在 AI 半導體產業鏈全圖中檢視 NVIDIA 的位置</a>——含九個環節的利潤池分佈與景氣傳導時間差。</p>
```

**因為第 0.2 節的三種 tab id，這件事不能一次 sed 改完**。實際工序：

| 群組 | panel id | 篇數 | 插入點 |
|---|---|---|---|
| A | `#panel-chain` | 10 | `.section-title` 之後、第一個 `.flag-box` 之前 |
| B | `#panel-supply` | 11 | 同上 |
| C | DELL `#panel-ai` | 1 | 該 tab 內「供應鏈」小節標題之後 |
| D | 無供應鏈 tab | 12 | **不插入**——見下 |

**D 群組（12 篇無供應鏈 tab）怎麼辦**：不為了加連結而硬塞一個 tab。改為在該報告的
**風險矩陣 tab 末尾**加一行同樣的 `.note` 連結。理由：這 12 篇（MU/AMD/SUMCO/PANW 等）
是框架早期產物，補做完整供應鏈 tab 是另一件事，不該綁在本功能的 P3 裡。

### 8.3 OutsideFramework Works 上架

在 `app/OutsideFramework/index.html` 的 Works → 既有的
Globe/Invest/Causal/Warning/High-price 那組（行 651–657 附近）新增一張 `.wk-card`：

```
標題：Industry Chain — 產業鏈全圖
說明：把 28 篇 earnings 報告的供應鏈萃取結果疊成同一條產業鏈，
      呈現九個環節的玩家、利潤池分佈與景氣傳導時間差。
連結：https://globe-invest.up.railway.app/industry-chain
```

這張卡不需要 `.wk-next-earn`（非個股報告），也不需要動 `EARN_DATES`。

---

## 9. 回填計畫（P1 的主要工作量）

28 篇在鏈上的報告，依第 0.3 節的格式分三批處理：

### 批次 1：`.chain-layer` 格式（10 篇，最省力）

ARM, GOOG×2, INTC×2, MSFT, NBIS, NVDA, ORCL, STM

`.node` 的結構是 `<div class="node conf-hi">名稱 — 角色 <span class="src">出處</span></div>`，
與本設計的 player 物件一一對應：

```
class="node conf-hi"  →  confidence: "hi"
"TSMC — 主力邏輯晶圓代工"  →  name: "TSMC", role: "主力邏輯晶圓代工"
<span class="src">10-K Item 1</span>  →  src: "10-K Item 1"
```

可寫一支一次性 Node 腳本半自動轉出草稿（`scripts/extract-chain-nodes.js`，
**不進 sync script，是一次性工具**），人工覆核歸層與去重。

### 批次 2：Sankey 格式（12 檔含 sankey）

`var DATA = {...}` 是現成 JSON，但語意是流量不是環節：node 混雜了供應商、地區、
業務部門、股東結構（TSMC 那篇的 4 張 sankey 分別是供應鏈／地區×平台／股權／轉投資）。
**只能人工挑出「真的是上下游對手方」的 node**，其餘丟棄。已在第 0.5 節挑好的部分直接用。

### 批次 3：自由排版與無 tab（其餘 6 篇 + 12 篇無 tab 中在鏈上者）

人工讀。無 tab 的 12 篇中，MU/AMD/MRVL/ALAB/GLW/SUMCO 在鏈上，
需回到其 10-K/法說會補萃取——這是回填中最貴的一段，可延後到 P4。

### 回填順序建議

1. 先建九層骨架與 `economics`/`cycle`（用第 0.5 節已有的硬資料 ＋ 標 `todo` 的留白）
2. 批次 1 的 10 篇（半自動）→ 立刻能看到有意義的畫面
3. 批次 2 挑出的節點
4. 批次 3 人工補

**驗收：第 2 步做完就該 demo 一次**。若那時看起來沒價值，後面兩批不值得做。

---

## 10. 產出流程（每篇新 earnings 分析要多做的事）

現行框架第 4 tab「供應鏈」本來就要從 filing 萃取全部供應鏈關係。新增動作是把結果寫回知識庫：

1. 判斷這家公司屬於哪條鏈的哪一層；無對應鏈就新建 JSON（並更新 `index.json`）
2. 在該層 `players` 找到自己——
   - 已存在（被別篇建立過）：補 `report`、補 `focus`、必要時更新 `share`/`src`/`asOf`
   - 不存在：新增
3. 本篇新萃取到的上下游對手方，逐一比對是否已在鏈上；缺的補進對應層（`report: null`）
4. 若 filing 提供新的環節經濟數字（毛利、市佔、長約條款），更新該層
   `economics`/`cycle` 並換上新的 `src`/`asOf`
5. 在報告的供應鏈 tab 加第 8.2 節那一行回鏈連結
6. 跑 `node scripts/lint-industry-chain.js`（第 11 節）
7. 同步鏡像、commit、push（見第 13 節更新後的收尾清單）

產物是一個 JSON diff，跟該篇報告同一個 commit 進版控，出處可回溯。

---

## 11. 驗證：`scripts/lint-industry-chain.js`

純 Node、零依賴、可在 `globe-invest` 的 CI（`.github/workflows/ci.yml` 目前跑 `node --check`）
裡加一步。檢查項：

| # | 檢查 | 級別 |
|---|---|---|
| 1 | JSON 可解析 | error |
| 2 | 必填欄位齊全（`id`/`name`/`updated`/`layers`/`transmission`） | error |
| 3 | `layer.id` 鏈內唯一 | error |
| 4 | `transmission.anchor` 指向存在的 layer | error |
| 5 | enum 值合法（`status`/`bargaining`/`confidence`/`direction`） | error |
| 6 | 任何非 null 的數字欄位，其所屬物件必有 `src` 與 `asOf` | error |
| 7 | `player.report` 指向的檔案存在於 `globe-invest/app/research/` | error |
| 8 | `revenueShare` 合計、`profitShare` 合計各在 95–105% 之間 | warn |
| 9 | `cycle.direction` 與 `lagQuarters` 正負一致 | warn |
| 10 | 同一 ticker 出現在多層 | warn（合法，但值得確認） |
| 11 | `asOf` 落後 `updated` ≥4 季 | warn |
| 12 | `index.json` 的 `layerCount`/`playerCount`/`reportCount` 與實際相符 | **auto-fix** |

第 12 項用 `--fix` 直接寫回，避免手動維護統計數字。

---

## 12. 分階段實作與驗收

| 階段 | 內容 | 驗收條件 |
|---|---|---|
| **P1** | `index.html` ＋ 環節分層檢視 ＋ 九層骨架 ＋ 批次 1 回填（10 篇） | 本地 :8141 開得起來；NVDA/TSMC/LRCX 三家能正確定位並連到報告；空白層（封裝）看得出來 |
| **P2** | 價值分配 ＋ 景氣傳導兩檢視；顯示出處開關；時效灰標；lint script | 四個內容維度到齊；`lint` 全綠；列印正常 |
| **P3** | 鏡像 globe-invest、server.js 加 route、sync script 加 pair、OutsideFramework 上架、22 篇報告插入回鏈連結 | **實際打 `https://globe-invest.up.railway.app/industry-chain` 確認不是 404**（第 2.2 節的陷阱）；兩個 repo 都已 push |
| **P4** | 批次 2/3 回填；第二條鏈（金融？看 GS/MS 是否會增篇）；CLAUDE.md 更新 | 28 篇全數在鏈上定位 |

P1 的分層檢視是整個模組的價值所在，另外兩個檢視是加分。**若 P1 demo 後覺得沒價值，
就停在那裡，不要因為設計文件寫了四階段就把四階段做完。**

---

## 13. CLAUDE.md 需要的改動（P3/P4 完成後）

1. **功能位置對照表新增第 9 節**：IndustryChain（dev repo／folder／檔內錨點／route／鏡像），
   格式仿第 2–4 節
2. **第 6 節「鏡像對」表**：新增 IndustryChain 一列，並註明**已納入** sync script
   （與 brownian/options/sankey 那幾列的「不在 sync script」相反）
3. **Earnings 框架第 4 tab「供應鏈」說明**：末尾加一句
   > 萃取結果同時寫回產業鏈知識庫 `app/IndustryChain/data/<chain>.json`（見該資料夾 DESIGN.md
   > 第 10 節），並在本 tab 頂端加一行回鏈連結。
4. **「分析完成後強制同步與上架」從 3 件事變 4 件事**，新增：
   > 4. **更新產業鏈知識庫**：把本篇萃取到的環節／玩家／經濟數字寫回
   >    `app/IndustryChain/data/<chain>.json`，跑 `node scripts/lint-industry-chain.js`，
   >    鏡像到 `globe-invest/app/industry-chain/data/`。知識庫沒更新＝這篇報告的供應鏈
   >    萃取結果沒有被累積，等於白做。
5. **`.claude/launch.json`** 新增 `industry-chain` :8141

---

## 14. 風險與未解問題

### 風險

| 風險 | 影響 | 緩解 |
|---|---|---|
| **知識庫腐化** | 市佔／毛利半年過期，看起來精確其實過時 | `asOf` ＋ ≥4 季灰標（4.3）；lint 第 11 項 warn |
| **`revenueShare`/`profitShare` 是估計值** | 被當成精確市佔引用 | `status:"estimate"` ＋ `~` 前綴 ＋ 頁尾方法論（4.2）。這條**不可省** |
| **鏈的邊界主觀** | 「要不要含電力／散熱」沒有標準答案 | `scopeNote` ＋ `excluded` 明寫，比假裝客觀好 |
| **回填工作量被低估** | 批次 3 是純人工，可能比預期貴 3 倍 | 分階段驗收，P1 後可停 |
| **雙 repo 鏡像忘記同步** | 線上停在舊版（此 repo 的歷史慣犯） | 納入 sync script（方案 A）而非手動 copy |
| **維護成本轉嫁到每篇報告** | 每篇 earnings 分析多 20–40 分鐘 | 若 P1 demo 後覺得不值，就不要把第 13.4 條寫進 CLAUDE.md |

### 未解問題（需要決定）

1. **第 5 層（封裝）0 篇報告要不要補一篇？** ASE 或 Amkor 的 earnings call。
   這是全鏈唯一空白，而 CoWoS 是 NVDA 報告反覆點名的瓶頸。
2. **GLW 跨兩層（材料的玻璃基板、系統的光連接）怎麼呈現？** 目前設計允許同一家出現在多層，
   但 focus panel 只在第一層展開。若這種情況變多，需要更好的設計。
3. **`profitShare` 的分母怎麼算？** 用「該層已上市玩家合計營業利益」會漏掉 Samsung/SK Hynix
   這種不單獨揭露半導體部門的公司。可能需要在方法論裡明列納入了誰。
4. **TSMC 有兩篇報告（FY2025 年報、FY2026Q2 季報）**，`player.report` 只能指一個。
   → 建議改成 `reports: []` 陣列，卡片顯示最新一篇、hover 列出全部。**此改動建議直接進 P1**，
   改 schema 比事後遷移便宜。
