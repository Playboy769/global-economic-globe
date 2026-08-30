'use strict';
// 總經黏性追蹤儀表板。零 npm 依賴（Node 內建 fetch），與本 repo 既有 server.js 家法一致。
//
// 職責：① 靜態送出 index.html（單一自足檔案）
//       ② /api/* 代理政府公開資料源（BLS、US Treasury），避開瀏覽器端 CORS 限制，
//          並用記憶體快取降低對上游的請求頻率（這些資料本身就是日頻/月頻，不需要即時）
//
// 六項指標對應：
//   ① 通膨黏性滲透   -> /api/cpi-stickiness   (BLS CPI API v2，免金鑰)
//   ② 實質消費支出   -> 前端手動輸入（無免金鑰的即時來源，見 index.html 內註記）
//   ③ 長債殖利率韌性 -> /api/treasury-yield   (Treasury 每日殖利率曲線 XML + TreasuryDirect 標售結果)
//   ④ 公司債 vs 國債發行 -> /api/debt-issuance (Treasury Debt to Penny，國債側自動；科技公司債側手動輸入)
//   ⑤ FOMC 決議措辭   -> /api/fomc-statements (Fed 貨幣政策新聞稿 RSS)
//   ⑥ 利息支出/稅收比率 -> /api/fiscal-ratio  (Treasury interest_expense + MTS Table 4 receipts)
//
// 沒有使用者資料、沒有登入閘門——純唯讀公開資料代理，手動輸入的部分留在前端 localStorage。

const http = require('http');
const fs = require('fs');
const path = require('path');

const PORT = Number(process.env.PORT) || 8080;
const ROOT = __dirname;

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
  }
  return { asOf: new Date().toISOString(), series: out };
}

// ── ③ Treasury 每日殖利率曲線（10Y）＋ 近期標售結果 ──
async function fetchTreasuryYield() {
  const year = new Date().getUTCFullYear();
  const years = [year, year - 1];
  const points = [];
  for (const y of years) {
    const r = await fetch(`https://home.treasury.gov/resource-center/data-chart-center/interest-rates/pages/xml?data=daily_treasury_yield_curve&field_tdr_date_value=${y}`);
    if (!r.ok) continue;
    const xml = await r.text();
    const dates = [...xml.matchAll(/<d:NEW_DATE[^>]*>([^<]+)<\/d:NEW_DATE>/g)].map(m => m[1].slice(0, 10));
    const y10s = [...xml.matchAll(/<d:BC_10YEAR[^>]*>([^<]*)<\/d:BC_10YEAR>/g)].map(m => m[1]);
    for (let i = 0; i < dates.length; i++) {
      if (y10s[i]) points.push({ date: dates[i], value: Number(y10s[i]) });
    }
  }
  points.sort((a, b) => a.date.localeCompare(b.date));

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

  return { series: points.slice(-120), auctions };
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
