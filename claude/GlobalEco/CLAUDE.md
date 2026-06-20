# Global Economics Globe — Project Rules

## Architecture

Single-file app: **`index.html`** (~3600 lines). All HTML, CSS, and JS live in one file — do not split into multiple files.

### Libraries (CDN, do not upgrade without asking)
- Three.js r128 + OrbitControls
- D3 v7.8.5 + TopoJSON v3.0.2
- GSAP (GreenSock Animation Platform) — for UI animations, transitions, scroll-linked effects

### GSAP Skills (AI reference)
Source: https://github.com/greensock/gsap-skills.git

When adding animations to this project, use GSAP best practices from the official skill set:

| Skill | Scope |
|-------|-------|
| `gsap-core` | `gsap.to()` / `from()` / `fromTo()`, easing, duration, stagger, defaults |
| `gsap-timeline` | Sequencing, position parameter, labels, nesting, playback control |
| `gsap-scrolltrigger` | Scroll-linked animations, pinning, scrub, triggers, refresh & cleanup |
| `gsap-plugins` | ScrollToPlugin, Flip, Draggable, SplitText, physics utilities |
| `gsap-utils` | Helpers: `clamp`, `mapRange`, `random`, `toArray`, `pipe` |
| `gsap-react` | `useGSAP` hook, context management (not used here — vanilla JS project) |
| `gsap-performance` | Transform aliases, `will-change`, `autoAlpha`, GPU-friendly patterns |
| `gsap-frameworks` | Vue/Svelte integration (not used here — vanilla JS project) |

**GSAP rules for this project:**
- GSAP is 100% free including all plugins — no license cost
- Prefer `gsap.to()` with transform aliases (`x`, `y`, `scale`, `rotation`) over CSS transforms
- Use `autoAlpha` instead of `opacity` (handles `visibility` automatically)
- Use `gsap.timeline()` for sequenced animations — not chained `setTimeout` or delays
- Use `gsap.ticker` for render-loop integration if needed alongside Three.js `requestAnimationFrame`
- Clean up animations on mode switch: `gsap.killTweensOf(target)` or `timeline.kill()`
- For scroll-linked UI panels, prefer `ScrollTrigger` over manual scroll event listeners

### Three Modes
The globe has three mutually exclusive modes controlled by `switchMode(mode)`:
| Mode | Variable | UI panel | Globe objects |
|------|----------|----------|---------------|
| `econ` | — | `#sidebar` | country polygons, heatmap, trade arcs |
| `energy` | `enMode` | `#energy-ui` (left + right panels) | energy hubs, routes, flow dots, crisis rings |
| `news` | `newsMode` | `#news-ui` (right panel) | white pillar markers (newsGlobeObjs) |

**Critical**: when entering a mode, always `clear*()` objects from the other two modes. See `switchMode()` for the canonical pattern.

### Scene Lighting
The scene uses `AmbientLight(0xffffff, 0.6)` + `DirectionalLight(0xffffff, 0.8)`. Required for `MeshPhongMaterial` to render — do not remove.

## Data

### Country Data (`CD` object)
- Keyed by **ISO 3166-1 numeric** code as string (e.g., `'840'` = USA)
- `nm`: Chinese name (繁體中文)
- English names are in `EN_NM` map (same keys)
- Extra fields (unemployment, flag emoji, GDP history) merged from `CD_X`
- Financial sector data in `CD_FIN`

### Economy Sidebar
- Country header shows: `中文名 / English Name` (no flag emoji)
- Sub-header: `ISO {code} · GDP 全球第 {rank} 大 · 2023 年估計`
- Tabs: 概覽, 人口, 農業, 貿易, 金融, 社會

## Style Rules

### Color Convention (Asian markets)
- **Red = up** (漲): `#dd4444` / class `op-up`
- **Green = down** (跌): `#44dd88` / class `op-down`
- This is the opposite of Western convention — do not "fix" it.

### Fonts
- UI: `'Noto Sans TC', 'Segoe UI', sans-serif` (weight 300 default)
- Data/monospace: `Consolas, monospace`
- Country names in sidebar: `'Calibri', 'Segoe UI', sans-serif`

### Design Language
- Dark glassmorphism: `rgba(10,14,20,0.92)` backgrounds, `backdrop-filter:blur(12px)`
- Earth tones for charts, blue accent `#4ad` / `#4a88dd`
- Borders: `rgba(255,255,255,0.06–0.15)`
- No emoji in UI text — use text labels only
- Scrollbar styling: thin, blue gradient thumb

## News Mode

### Data Source
RSS feeds via `rss2json.com` converter (fallback: `allorigins`, `corsproxy`):
- Google News RSS (business/economy search)
- BBC Business RSS
- CNBC RSS

### News Markers (3D)
- **White only** (`0xffffff`) — unified color for all countries
- Pillar: `CylinderGeometry`, base radius `0.008 + size*0.005`
- Head: `OctahedronGeometry` with glow sphere
- Only visible in NEWS mode — `clearNewsGlobe()` removes them in other modes
- Click detection uses recursive raycasting on `THREE.Group` objects

### Country Selector
- Dropdown select (`.news-select-wrap`) with search — not chips
- On selection: updates dropdown display, flies camera, loads country news
- Country-specific fallback: Google News RSS search for that country

### Article Classification
`_nmAlias` maps country codes to name variants (full name, abbreviations, cities, stock indices). Longest-match algorithm distributes bulk articles to countries.

## Energy Mode

### Oil Price Widget
- Yahoo Finance API via CORS proxies
- Symbols: CL=F (WTI), BZ=F (Brent), NG=F (Henry Hub), TTF=F, RB=F (RBOB), HO=F
- Timestamps: **local time** (user is in UTC+8 Taiwan)
- Market closed detection: `regularMarketTime` age + `currentTradingPeriod` — show "Market Closed" and freeze at closing price, no Brownian motion simulation

### Energy Routes
- CatmullRom curves for smooth arcs
- Route types: crude (blue), LNG (green), pipeline (orange), disrupted (red)
- Resizable left/right panels with drag handles

## Coordinate Helpers
- `ll2v(lat, lng, radius)` — lat/lng to THREE.Vector3
- `flyTo(lat, lng)` — smooth camera animation to coordinates
- `slerpArc(v1, v2, height)` — great circle arc between two points

## Preview
- Dev server: Python `http.server` on port 8124 (configured in `.claude/launch.json`)
- Three.js heavy rendering may cause screenshot timeouts — use `preview_eval` / `preview_snapshot` for verification
