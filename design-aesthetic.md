# Outside FrameWork — 設計美學指南

## 設計哲學

以「編輯式白色美學」（Editorial White）為核心，追求極簡、留白、與克制。整體風格介於高端雜誌排版與當代藝廊之間——讓內容自身成為視覺焦點，設計退居幕後。

---

## 色彩系統

| 用途 | 色碼 | 說明 |
|------|------|------|
| 主背景 | `#fff` | 純白底，最大化留白感 |
| 主文字 | `#0a0a0a` | 近黑，避免純黑的壓迫感 |
| 選取反白 | `bg #0a0a0a / fg #fff` | `::selection` 反色處理 |
| 次要文字 | `#666` | 內文段落 |
| 輔助文字 | `#888` / `#999` | 描述、標籤 |
| 淡化文字 | `#aaa` / `#bbb` / `#ccc` | 導航未選中、編號、副標 |
| 分隔線 | `#eee` | 極淡灰，僅暗示結構 |
| 卡片背景 | `#fafafa` | 微灰，與白底產生層次 |
| 卡片邊框 | `rgba(0,0,0,0.05)` | 幾乎不可見的邊界 |
| 圖標描邊 | `#bbb` | SVG 線條色 |
| 強調淡化 | `#999` | Philosophy 頁 em 文字 |
| 滾動條 | thumb `#ddd` / hover `#bbb` | 6px 寬，圓角 |

**原則**：整個網站僅使用灰階。沒有品牌色、沒有強調色。透過字重、字號、與透明度建立視覺層級。

---

## 字體系統

### 字體堆疊

| 角色 | 字體 | 備註 |
|------|------|------|
| 英文標題/品牌 | `Cormorant Garant` | 細瘦優雅的襯線體，用於導航、品牌名、編號 |
| 英文大標題 | `Playfair Display` | 高對比襯線體，用於 Section Heading |
| 中文標題 | `Noto Serif TC` | 思源宋體，用於中文語句與哲學段落 |
| 品牌特殊 | `Orchid` | 僅用於首頁底部 "Outside FrameWork" 品牌字 |
| UI 元素/卡片標題 | `Segoe UI, Noto Sans TC, Helvetica Neue, Arial, sans-serif` | 無襯線，用於 Works 卡片名稱、About 數值區塊 |
| 通用回退 | `Georgia, serif` | body 層級回退 |

### 字體設定

| 元素 | 大小 | 字重 | 間距 | 其他 |
|------|------|------|------|------|
| 品牌名 SILAS | 22px | 300 | `.22em` | uppercase |
| 導航按鈕 | 13px | 400 | `.1em` | uppercase |
| 導航 Logo | 15px | 400 | `.18em` | uppercase |
| Section 標題 | `clamp(36px, 5vw, 56px)` | 300 | `.01em` | — |
| About 主語句 | `clamp(28px, 3.5vw, 44px)` | 700 | `-.01em` | 中文宋體 |
| About 內文 | 16px | 400 | — | `line-height: 2` |
| 數值區塊標題 | 20px | 600 | `-.01em` | 無襯線 |
| 數值區塊描述 | 14px | 400 | — | `line-height: 1.8` |
| Works 卡片名 | 16px | 500 | `.01em` | 無襯線 |
| Works 類型標 | 11px | 400 | `.06em` | uppercase |
| Philosophy 引言 | `clamp(24px, 3.5vw, 40px)` | 700 | `-.01em` | 中文宋體 |
| Philosophy 編號 | 48px | 300 | — | 斜體，色 `#eee` |
| Philosophy 段落 | 14px | 400 | — | `line-height: 1.9` |
| 首頁 CTA 按鈕 | 13px | 400 | `.1em` | uppercase, 圓角膠囊 |
| 首頁品牌字 | `clamp(28px, 5vw, 56px)` | 400 | `-.01em` | Orchid 字體 |
| 頁尾品牌 | 16px | 700 | `-.01em` | Playfair Display |
| 頁尾版權 | 12px | — | `.04em` | 色 `#bbb` |

---

## 佈局系統

### 容器
- 最大寬度：`1320px`
- 水平內距：`48px`（平板以下 `28px`）
- 置中：`margin: 0 auto`

### 間距原則
- 導航列內距：`32px 48px`
- 頁面頂部間距：`140px`（留出導航高度 + 呼吸空間）
- Section 之間：`80px ~ 100px`
- 卡片間距：`24px`（平板 `12px`）

### 網格
- Works 卡片：`3 欄`（平板 `2 欄`、手機 `1 欄`）
- About 雙欄：`1fr 1fr`，gap `80px`
- About 數值：`3 欄`，以邊框分隔
- Philosophy：`80px + 1fr` 兩欄（編號 + 內容）

---

## 元件設計

### 導航列
- `position: fixed`，透明背景
- 非首頁時：加上 `rgba(255,255,255,0.92)` 背景 + `blur(20px)` 毛玻璃效果
- 滾動隱藏：`transform: translateY(-100%)`，搭配 `0.5s cubic-bezier(.4,0,.2,1)` 過渡
- 首頁 SILAS 名稱與導航同高對齊，使用 `pointer-events: none/auto` 技巧避免遮擋

### Works 卡片
- 圓角：`16px`
- 視覺區：`aspect-ratio: 4/3`，背景 `#fafafa`
- Hover 效果：`translateY(-4px)` + `box-shadow: 0 12px 40px rgba(0,0,0,0.06)`
- 圖標 Hover：內部容器 `scale(1.02)`
- 所有過渡：`0.4s ~ 0.5s cubic-bezier(.4,0,.2,1)`

### 專案覆蓋層
- 全屏 `position: fixed`，`z-index: 200`
- 黑色背景，iframe 無邊框
- 關閉按鈕：圓形毛玻璃 `rgba(255,255,255,0.12)` + `blur(8px)`
- Hover：`scale(1.1)` + 提亮背景

### CTA 按鈕（首頁）
- 膠囊型：`border-radius: 100px`
- 透明背景 + `1px solid #0a0a0a` 邊框
- Hover：反轉為黑底白字
- 箭頭圖標 Hover 右移 `3px`

---

## 動畫與過渡

### 滾動顯現（Reveal）
- 初始：`opacity: 0; transform: translateY(40px)`
- 觸發：`.visible` class → `opacity: 1; translateY(0)`
- 時長：`0.8s ease`（opacity）+ `0.8s cubic-bezier(.4,0,.2,1)`（transform）
- 延遲層級：`.reveal-delay-1` (0.1s)、`-2` (0.2s)、`-3` (0.3s)

### 首頁插圖
- `fadeInIllus` 動畫：從 `translateY(12px)` 淡入
- 延遲 `0.4s`，時長 `1.2s ease`

### 通用過渡曲線
- 主要動態：`cubic-bezier(.4, 0, .2, 1)`（Material Design 標準減速曲線）
- 色彩變化：`ease`，`0.25s ~ 0.3s`

---

## SVG 圖標風格

### 首頁中心插圖
- 框架隱喻：矩形網格 + 四條曲線突破邊界
- 描邊色：`#2a2a2a`（近黑），寬度 `1.5px`
- 網格虛線：`#d8d8d8`，`stroke-dasharray: 4 6`
- 箭頭：實心三角 marker-end
- 交叉點：實心圓 `r=3`

### Works 卡片圖標
- ViewBox：`80 × 80`
- 描邊色：`#bbb`，寬度 `1 ~ 1.2px`
- 風格：極簡線條，無填充（`fill: none`）
- 部分元素使用 `opacity: 0.3 ~ 0.6` 製造層次

---

## 響應式策略

| 斷點 | 變化 |
|------|------|
| `≤ 900px` | 容器內距縮減；About 改單欄；數值改單欄去右框線；Works 改 2 欄 |
| `≤ 600px` | Works 改 1 欄；Philosophy 編號獨立一行；導航間距縮減 |

---

## 核心設計原則

1. **極致克制**：灰階配色，無裝飾元素，讓留白說話
2. **字體即層級**：透過字體家族、字重、字號建立視覺階層，而非顏色
3. **微妙動態**：所有互動回饋控制在 `4px` 位移、`0.06` 陰影透明度以內
4. **編輯式排版**：大量 uppercase 小字標籤 + 高對比大標題的雜誌風格
5. **中西融合**：英文襯線體（Cormorant / Playfair）與中文宋體（Noto Serif TC）自然並存
6. **框架隱喻**：SVG 插圖以「突破框架的曲線」呼應網站主題 "Outside FrameWork"
