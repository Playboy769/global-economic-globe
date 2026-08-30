'use strict';
// 總經黏性追蹤儀表板。零 npm 依賴（Node 內建 fetch），與本 repo 既有 server.js 家法一致。
//
// 職責：① 靜態送出 index.html（單一自足檔案）
//       ② /api/* 代理政府公開資料源（BLS、US Treasury），避開瀏覽器端 CORS 限制，
//          並用記憶體快取降低對上游的請求頻率（這些資料本身就是日頻/月頻，不需要即時）
//
// 六項指標對應：
//   ① 通膨黏性滲透   -> /api/cpi-stickiness   (BLS CPI API v2，免金鑰)
//   ② 實質消費支出   -> /api/real-pce         (FRED PCEC96，需 FRED_API_KEY)
//   ③ 長債殖利率韌性 -> /api/treasury-yield   (Treasury 每日殖利率曲線 XML + TreasuryDirect 標售結果)
//   ④ 公司債 vs 國債發行 -> /api/debt-issuance (Treasury Debt to Penny，國債側自動；科技公司債側手動輸入)
//   ⑤ FOMC 決議措辭   -> /api/fomc-statements (Fed 貨幣政策新聞稿 RSS)
//   ⑥ 利息支出/稅收比率 -> /api/fiscal-ratio  (Treasury interest_expense + MTS Table 4 receipts)
//
// 掛在指標 1 底下的輔助訊號（不算獨立的第 7 項，避免打亂原本 6 項框架）：
//   勞動市場鬆緊度   -> /api/jobless-claims (FRED ICSA 週度初請＋IC4WSA 4週移動平均，需 FRED_API_KEY)
//   PCE 通膨廣度     -> /api/pce-breadth    (BEA NIUnderlyingDetail Table 2.4.4U，需 BEA_API_KEY)
//
// 沒有使用者資料、沒有登入閘門——純唯讀公開資料代理，手動輸入的部分留在前端 localStorage。
//
// FRED_API_KEY 為必要環境變數才能啟用 ② 與勞動市場鬆緊度的自動抓取。
// BEA_API_KEY 啟用後改用來抓 PCE 通膨廣度（Table 2.4.4U 逐項價格指數），這是 FRED 沒有
// 對應細項序列的資料——FRED 只到「財貨/服務/耐久財」這種大分類，BEA 才有近 400 項細項。

const http = require('http');
const fs = require('fs');
const path = require('path');

const PORT = Number(process.env.PORT) || 8080;
const ROOT = __dirname;
const FRED_API_KEY = process.env.FRED_API_KEY || '';
const BEA_API_KEY = process.env.BEA_API_KEY || '';

const TYPES = {
  '.html': 'text/html; charset=utf-8',
  '.js': 'text/javascript; charset=utf-8',
  '.css': 'text/css; charset=utf-8',
  '.json': 'application/json; charset=utf-8',
  '.ico': 'image/x-icon',
};

// ── 簡易記憶體快取：政府資料是日/月頻，20 分鐘 TTL 已經很夠用，避免每次重整都打上游 ──
const CACHE_TTL_MS = 20 * 60 * 1000;
const cache = new Map();
async function cached(key, fn) {
  const hit = cache.get(key);
  if (hit && Date.now() - hit.at < CACHE_TTL_MS) return hit.data;
  const data = await fn();
  cache.set(key, { at: Date.now(), data });
  return data;
}

function sendJSON(res, code, obj) {
  const body = JSON.stringify(obj);
  res.writeHead(code, {
    'content-type': 'application/json; charset=utf-8',
    'content-length': Buffer.byteLength(body),
    'cache-control': 'no-store',
  });
  res.end(body);
}

function serveStatic(req, res, pathname) {
  const rel = pathname === '/' ? 'index.html' : pathname.replace(/^\/+/, '');
  const file = path.resolve(ROOT, rel);
  if (file !== ROOT && !file.startsWith(ROOT + path.sep)) {
    res.writeHead(403).end('forbidden');
    return;
  }
  if (path.basename(file) === 'server.js') {
    res.writeHead(404).end('not found');
    return;
  }
  fs.readFile(file, (err, buf) => {
    if (err) { res.writeHead(404).end('not found'); return; }
    res.writeHead(200, {
      'content-type': TYPES[path.extname(file).toLowerCase()] || 'application/octet-stream',
      'cache-control': 'no-cache',
    });
    res.end(buf);
  });
}

// ── ① BLS CPI：核心 CPI（supercore 代理）、OER、能源 CPI、平均時薪 YoY% ──
const BLS_SERIES = {
  core: 'CUSR0000SA0L1E',   // CPI-U less food & energy, SA
  oer: 'CUSR0000SEHC',      // Owners' equivalent rent, SA
  energy: 'CUSR0000SA0E',   // Energy, SA
  wages: 'CES0500000003',   // Average hourly earnings, total private, SA
};

function periodToDate(year, period) {
  // BLS period 格式 "M01".."M12"；"M13" 是年度平均，跳過
  const m = Number(period.slice(1));
  if (!m || m > 12) return null;
  return `${year}-${String(m).padStart(2, '0')}-01`;
}

async function fetchCpiStickiness() {
  const thisYear = new Date().getUTCFullYear();
  const r = await fetch('https://api.bls.gov/publicAPI/v2/timeseries/data/', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({
      seriesid: Object.values(BLS_SERIES),
      startyear: String(thisYear - 3),
      endyear: String(thisYear),
    }),
  });
  if (!r.ok) throw new Error('BLS API HTTP ' + r.status);
  const j = await r.json();
  if (j.status !== 'REQUEST_SUCCEEDED') throw new Error('BLS API: ' + (j.message || []).join('; '));

  const bySeries = {};
  for (const s of j.Results.series) {
    const pts = s.data
      .map(d => ({ date: periodToDate(d.year, d.period), value: Number(d.value) }))
      .filter(d => d.date)
      .sort((a, b) => a.date.localeCompare(b.date));
    bySeries[s.seriesID] = pts;
  }

  const out = {};
  for (const [key, seriesId] of Object.entries(BLS_SERIES)) {
    const pts = bySeries[seriesId] || [];
    const yoy = pts.map((d, i) => {
      const prior = pts[i - 12];
      if (!prior) return null;
      return { date: d.date, value: Number((((d.value / prior.value) - 1) * 100).toFixed(2)) };
    }).filter(Boolean);
    out[key] = yoy.slice(-24);

    if (key === 'core') {
      // 3 個月年化（3mo/3mo SAAR）：比 YoY 更快反映最近動能，拿來跟單月讀數對照，
      // 判斷「單月意外」是否代表底層趨勢真的轉向，還是雜訊——市場常用的 core CPI
      // 3mo annualized 指標，不是自創算法。
      const ann3 = pts.map((d, i) => {
        const prior = pts[i - 3];
        if (!prior || prior.value <= 0) return null;
        const ratio = d.value / prior.value;
        return { date: d.date, value: Number(((Math.pow(ratio, 4) - 1) * 100).toFixed(2)) };
      }).filter(Boolean);
      out.coreAnn3 = ann3.slice(-24);
    }
  }
  return { asOf: new Date().toISOString(), series: out };
}

// ── 勞動市場鬆緊度輔助訊號：初請失業金週度數據＋FRED 內建 4 週移動平均 ──
// 掛在指標 1 底下（不另開新卡片）：初請失業金是薪資/服務業黏性能否維持的領先指標，
// 4 週移動平均由 FRED 官方序列直接提供（IC4WSA），不必自己滾動計算。
async function fetchJoblessClaims() {
  if (!FRED_API_KEY) {
    const err = new Error('FRED_API_KEY 未設定');
    err.code = 'NO_KEY';
    throw err;
  }
  const [weeklyR, ma4R] = await Promise.all([
    fetch(`https://api.stlouisfed.org/fred/series/observations?series_id=ICSA&api_key=${FRED_API_KEY}&file_type=json&sort_order=asc&observation_start=2024-06-01`),
    fetch(`https://api.stlouisfed.org/fred/series/observations?series_id=IC4WSA&api_key=${FRED_API_KEY}&file_type=json&sort_order=asc&observation_start=2024-06-01`),
  ]);
  if (!weeklyR.ok) throw new Error('FRED ICSA HTTP ' + weeklyR.status);
  if (!ma4R.ok) throw new Error('FRED IC4WSA HTTP ' + ma4R.status);
  const weeklyJ = await weeklyR.json();
  const ma4J = await ma4R.json();
  const toPts = j => (j.observations || [])
    .map(o => ({ date: o.date, value: Number(o.value) }))
    .filter(o => Number.isFinite(o.value));
  return {
    asOf: new Date().toISOString(),
    weekly: toPts(weeklyJ).slice(-52),
    ma4: toPts(ma4J).slice(-52),
  };
}

function monthsBefore(dateStr, n) {
  const [y, m] = dateStr.split('-').map(Number);
  const idx = (y * 12 + (m - 1)) - n;
  const ny = Math.floor(idx / 12);
  const nm = (idx % 12) + 1;
  return `${ny}-${String(nm).padStart(2, '0')}-01`;
}

// ── PCE 通膨廣度：一籃子商品裡有多少比例仍在漲，用來判斷通膨壓力是集中還是普遍 ──
// 資料來源 BEA NIUnderlyingDetail Table 2.4.4U（Price Indexes for PCE by Type of Product），
// 一次 GetData 回傳全表所有具名項目 × 全部月份，不必逐項個別呼叫。
//
// ⚠️ 方法論限制（老實寫在這裡）：BEA API 沒有回傳這張表的階層/縮排資訊，程式無法區分
// 「大分類加總列」跟「真正的細項列」。這裡只排除唯一的總計列（Line 1「Personal
// consumption expenditures」）與整組「Market-based PCE ...」（換一種統計基礎重算同一批
// 項目，算進來會重複計算），其餘全部具名項目（含各層級的中間加總，如「Durable goods」
// 本身也會被算成一項）都納入計算——這代表這裡的廣度百分比不是特定研究機構「199項」
// 定義的逐項覆現，兩者會有落差，但排除規則整個攤在這裡，可自行覆核調整門檻或排除清單。
const BEA_PCE_TABLE = 'U20404';
const BEA_BREADTH_THRESHOLD = 3; // % YoY，常見的廣度分析參考門檻，非官方目標值

async function fetchPceBreadth() {
  if (!BEA_API_KEY) {
    const err = new Error('BEA_API_KEY 未設定或尚未啟用');
    err.code = 'NO_KEY';
    throw err;
  }
  const thisYear = new Date().getUTCFullYear();
  const years = [thisYear - 2, thisYear - 1, thisYear].join(',');
  const url = `https://apps.bea.gov/api/data/?UserID=${BEA_API_KEY}&method=GetData&datasetname=NIUnderlyingDetail&TableName=${BEA_PCE_TABLE}&Frequency=M&Year=${years}&ResultFormat=JSON`;
  const r = await fetch(url);
  if (!r.ok) throw new Error('BEA API HTTP ' + r.status);
  const j = await r.json();
  const results = j.BEAAPI && j.BEAAPI.Results;
  if (!results || results.Error) {
    throw new Error('BEA API: ' + (results && results.Error && results.Error.APIErrorDescription || 'unknown error'));
  }
  const rows = results.Data || [];

  // 按 LineNumber 分組：每組是一個具名項目的完整月度序列，byDate 供 O(1) 查值。
  const byLine = new Map();
  for (const row of rows) {
    const desc = row.LineDescription || '';
    if (row.LineNumber === '1' || desc.startsWith('Market-based PCE')) continue;
    const m = /^(\d{4})M(\d{2})$/.exec(row.TimePeriod);
    const value = Number(row.DataValue);
    if (!m || !Number.isFinite(value)) continue;
    const date = `${m[1]}-${m[2]}-01`;
    if (!byLine.has(row.LineNumber)) byLine.set(row.LineNumber, { description: desc, byDate: new Map() });
    byLine.get(row.LineNumber).byDate.set(date, value);
  }

  const allDates = [...new Set(rows.map(r => {
    const m = /^(\d{4})M(\d{2})$/.exec(r.TimePeriod);
    return m ? `${m[1]}-${m[2]}-01` : null;
  }).filter(Boolean))].sort();
  const latestDate = allDates[allDates.length - 1];

  const items = [];
  for (const entry of byLine.values()) {
    const cur = entry.byDate.get(latestDate);
    const prior12 = entry.byDate.get(monthsBefore(latestDate, 12));
    const prior6 = entry.byDate.get(monthsBefore(latestDate, 6));
    const yoy12 = (cur != null && prior12 != null) ? Number((((cur / prior12) - 1) * 100).toFixed(2)) : null;
    const ann6 = (cur != null && prior6 != null && prior6 > 0)
      ? Number(((Math.pow(cur / prior6, 2) - 1) * 100).toFixed(2))
      : null;
    if (yoy12 != null || ann6 != null) items.push({ description: entry.description, yoy12, ann6 });
  }

  function breadthSeriesFor(monthsBack, exponent) {
    const out = [];
    for (const date of allDates) {
      const priorDate = monthsBefore(date, monthsBack);
      let above = 0, total = 0;
      for (const entry of byLine.values()) {
        const cur = entry.byDate.get(date);
        const prior = entry.byDate.get(priorDate);
        if (cur == null || prior == null || prior <= 0) continue;
        total++;
        if ((Math.pow(cur / prior, exponent) - 1) * 100 > BEA_BREADTH_THRESHOLD) above++;
      }
      if (total > 0) out.push({ date, value: Number(((above / total) * 100).toFixed(1)) });
    }
    return out;
  }

  return {
    asOf: new Date().toISOString(),
    threshold: BEA_BREADTH_THRESHOLD,
    itemCount: items.length,
    breadth12: breadthSeriesFor(12, 1),
    breadth6: breadthSeriesFor(6, 2),
    items: items.sort((a, b) => (b.yoy12 ?? -999) - (a.yoy12 ?? -999)),
  };
}

// ── ② FRED PCEC96：實質消費支出（月頻，Chained 2017 Dollars, SAAR）──
async function fetchRealPce() {
  if (!FRED_API_KEY) {
    const err = new Error('FRED_API_KEY 未設定');
    err.code = 'NO_KEY';
    throw err;
  }
  const r = await fetch(`https://api.stlouisfed.org/fred/series/observations?series_id=PCEC96&api_key=${FRED_API_KEY}&file_type=json&sort_order=asc&observation_start=2018-01-01`);
  if (!r.ok) throw new Error('FRED API HTTP ' + r.status);
  const j = await r.json();
  const pts = (j.observations || [])
    .map(o => ({ date: o.date, value: Number(o.value) }))
    .filter(o => Number.isFinite(o.value));

  const mom = pts.map((d, i) => {
    const prior = pts[i - 1];
    if (!prior) return null;
    return { date: d.date, value: Number((((d.value / prior.value) - 1) * 100).toFixed(2)) };
  }).filter(Boolean);
  const yoy = pts.map((d, i) => {
    const prior = pts[i - 12];
    if (!prior) return null;
    return { date: d.date, value: Number((((d.value / prior.value) - 1) * 100).toFixed(2)) };
  }).filter(Boolean);

  return {
    asOf: new Date().toISOString(),
    level: pts.slice(-24),
    mom: mom.slice(-24),
    yoy: yoy.slice(-24),
  };
}

// ── ③ Treasury 每日殖利率曲線（2Y/5Y/10Y/20Y/30Y 比較）＋ 近期標售結果 ──
const YIELD_TENORS = { y2: 'BC_2YEAR', y5: 'BC_5YEAR', y10: 'BC_10YEAR', y20: 'BC_20YEAR', y30: 'BC_30YEAR' };

async function fetchTreasuryYield() {
  const year = new Date().getUTCFullYear();
  const years = [year, year - 1];
  const byTenor = {}; // { y2: [{date,value}], y5: [...], ... }
  for (const key of Object.keys(YIELD_TENORS)) byTenor[key] = [];

  for (const y of years) {
    const r = await fetch(`https://home.treasury.gov/resource-center/data-chart-center/interest-rates/pages/xml?data=daily_treasury_yield_curve&field_tdr_date_value=${y}`);
    if (!r.ok) continue;
    const xml = await r.text();
    const dates = [...xml.matchAll(/<d:NEW_DATE[^>]*>([^<]+)<\/d:NEW_DATE>/g)].map(m => m[1].slice(0, 10));
    for (const [key, field] of Object.entries(YIELD_TENORS)) {
      const re = new RegExp(`<d:${field}[^>]*>([^<]*)<\\/d:${field}>`, 'g');
      const vals = [...xml.matchAll(re)].map(m => m[1]);
      for (let i = 0; i < dates.length; i++) {
        if (vals[i]) byTenor[key].push({ date: dates[i], value: Number(vals[i]) });
      }
    }
  }
  const series = {};
  for (const key of Object.keys(YIELD_TENORS)) {
    series[key] = byTenor[key].sort((a, b) => a.date.localeCompare(b.date)).slice(-120);
  }
  let auctions = [];
  try {
    const ar = await fetch('https://www.treasurydirect.gov/TA_WS/securities/search?type=Note&term=10-Year&status=Auctioned&pagesize=6&format=json');
    if (ar.ok) {
      const aj = await ar.json();
      // TreasuryDirect 的 pagesize 參數不會生效（實測回傳全部歷史場次，回溯到 1979 年），
      // 一律自己排序後只取最近幾場。
      auctions = (Array.isArray(aj) ? aj : []).map(x => ({
        auctionDate: (x.auctionDate || '').slice(0, 10),
        highYield: x.highYield || x.averageMedianYield || null,
        bidToCover: x.bidToCoverRatio || null,
      })).filter(x => x.auctionDate).sort((a, b) => a.auctionDate.localeCompare(b.auctionDate)).slice(-8);
    }
  } catch (e) { /* 標售資料非必要，取不到就留空 */ }

  return { series, auctions };
}

// ── ④ Treasury Debt to Penny：國債存量日序列（可反推淨發行）──
async function fetchDebtIssuance() {
  const r = await fetch('https://api.fiscaldata.treasury.gov/services/api/fiscal_service/v2/accounting/od/debt_to_penny?sort=-record_date&page%5Bsize%5D=200&fields=record_date,tot_pub_debt_out_amt');
  if (!r.ok) throw new Error('fiscaldata debt_to_penny HTTP ' + r.status);
  const j = await r.json();
  const series = (j.data || [])
    .map(d => ({ date: d.record_date, value: Number(d.tot_pub_debt_out_amt) }))
    .sort((a, b) => a.date.localeCompare(b.date));
  return { series };
}

// ── ⑥ 利息支出／稅收比率：interest_expense（全類別加總）÷ MTS Table 4「Total -- Receipts」──
async function fetchFiscalRatio() {
  const [expR, rcptR] = await Promise.all([
    fetch('https://api.fiscaldata.treasury.gov/services/api/fiscal_service/v2/accounting/od/interest_expense?sort=-record_date&page%5Bsize%5D=500&fields=record_date,month_expense_amt'),
    fetch('https://api.fiscaldata.treasury.gov/services/api/fiscal_service/v1/accounting/mts/mts_table_4?filter=classification_desc:eq:Total%20--%20Receipts&sort=-record_date&page%5Bsize%5D=30&fields=record_date,current_month_net_rcpt_amt'),
  ]);
  if (!expR.ok) throw new Error('fiscaldata interest_expense HTTP ' + expR.status);
  if (!rcptR.ok) throw new Error('fiscaldata mts_table_4 HTTP ' + rcptR.status);
  const expJ = await expR.json();
  const rcptJ = await rcptR.json();

  const expByDate = new Map();
  for (const row of (expJ.data || [])) {
    const v = Number(row.month_expense_amt);
    if (!Number.isFinite(v)) continue;
    expByDate.set(row.record_date, (expByDate.get(row.record_date) || 0) + v);
  }
  const rcptByDate = new Map();
  for (const row of (rcptJ.data || [])) {
    const v = Number(row.current_month_net_rcpt_amt);
    if (Number.isFinite(v)) rcptByDate.set(row.record_date, v);
  }

  const dates = [...expByDate.keys()].filter(d => rcptByDate.has(d)).sort();
  const series = dates.map(d => {
    const interestExpense = expByDate.get(d);
    const receipts = rcptByDate.get(d);
    return {
      date: d,
      interestExpense,
      receipts,
      ratio: receipts ? Number(((interestExpense / receipts) * 100).toFixed(2)) : null,
    };
  }).slice(-24);
  return { series };
}

// ── ⑤ Fed 貨幣政策新聞稿 RSS ──
async function fetchFomcStatements() {
  const r = await fetch('https://www.federalreserve.gov/feeds/press_monetary.xml');
  if (!r.ok) throw new Error('Fed RSS HTTP ' + r.status);
  const xml = await r.text();
  const items = [...xml.matchAll(/<item>([\s\S]*?)<\/item>/g)]
    .map(m => m[1])
    .slice(0, 8)
    .map(block => {
      const title = ((block.match(/<title>([\s\S]*?)<\/title>/) || [])[1] || '')
        .replace(/&#39;/g, "'").replace(/&amp;/g, '&').trim();
      const link = ((block.match(/<link><!\[CDATA\[([\s\S]*?)\]\]><\/link>/) || [])[1] || '').trim();
      const pubDate = ((block.match(/<pubDate><!\[CDATA\[([\s\S]*?)\]\]><\/pubDate>/) || [])[1] || '').trim();
      return { title, link, pubDate };
    });
  return { items };
}

const ENDPOINTS = {
  '/api/cpi-stickiness': () => cached('cpi', fetchCpiStickiness),
  '/api/real-pce': () => cached('pce', fetchRealPce),
  '/api/jobless-claims': () => cached('jobless', fetchJoblessClaims),
  '/api/pce-breadth': () => cached('breadth', fetchPceBreadth),
  '/api/treasury-yield': () => cached('yield', fetchTreasuryYield),
  '/api/debt-issuance': () => cached('debt', fetchDebtIssuance),
  '/api/fiscal-ratio': () => cached('fiscal', fetchFiscalRatio),
  '/api/fomc-statements': () => cached('fomc', fetchFomcStatements),
};

const server = http.createServer((req, res) => {
  let pathname = '/';
  try { pathname = decodeURIComponent(new URL(req.url, 'http://x').pathname); }
  catch (e) { res.writeHead(400).end('bad request'); return; }

  if (pathname === '/healthz') {
    res.writeHead(200, { 'content-type': 'text/plain' }).end('ok');
    return;
  }

  if (ENDPOINTS[pathname]) {
    ENDPOINTS[pathname]()
      .then(data => sendJSON(res, 200, data))
      .catch(e => {
        console.error('[api] ' + pathname + ' 失敗:', e && e.message);
        sendJSON(res, 502, { error: '上游資料源暫時無法取得', detail: e && e.message });
      });
    return;
  }

  if (pathname.startsWith('/api/')) {
    sendJSON(res, 404, { error: 'no such endpoint' });
    return;
  }

  if (req.method !== 'GET' && req.method !== 'HEAD') {
    res.writeHead(405).end('method not allowed');
    return;
  }
  serveStatic(req, res, pathname);
});

server.listen(PORT, () => {
  console.log('macro-tracker listening on :' + PORT);
});
