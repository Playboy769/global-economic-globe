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
//   信用利差（掛指標4）-> /api/credit-spreads (FRED BAMLH0A0HYM2 高收益債＋BAMLC0A0CM 投資級 OAS，需 FRED_API_KEY)
//   企業利潤週期（掛指標2）-> /api/corp-profits (BEA NIPA Table 1.14 稅前/稅後利潤等，季頻，需 BEA_API_KEY)
//
// 獨立第 7 項（NFP，非六項框架原生成員，是後續加開的獨立卡片）：
//   ⑦ 非農就業與勞動市場斷層 -> /api/nfp (BLS CES 12產業別＋CPS家戶調查，免金鑰；FRED PAYEMS
//      初值/修訂值比對＋官方擴散指數，需 FRED_API_KEY)
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

// ── 指標 4 輔助：信用利差（HY／IG OAS）——AI 資本支出撐住需求時，金融情勢是否仍寬鬆 ──
async function fetchCreditSpreads() {
  if (!FRED_API_KEY) {
    const err = new Error('FRED_API_KEY 未設定');
    err.code = 'NO_KEY';
    throw err;
  }
  const [hyR, igR] = await Promise.all([
    fetch(`https://api.stlouisfed.org/fred/series/observations?series_id=BAMLH0A0HYM2&api_key=${FRED_API_KEY}&file_type=json&sort_order=asc&observation_start=2024-06-01`),
    fetch(`https://api.stlouisfed.org/fred/series/observations?series_id=BAMLC0A0CM&api_key=${FRED_API_KEY}&file_type=json&sort_order=asc&observation_start=2024-06-01`),
  ]);
  if (!hyR.ok) throw new Error('FRED BAMLH0A0HYM2 HTTP ' + hyR.status);
  if (!igR.ok) throw new Error('FRED BAMLC0A0CM HTTP ' + igR.status);
  const hyJ = await hyR.json();
  const igJ = await igR.json();
  const toPts = j => (j.observations || [])
    .map(o => ({ date: o.date, value: Number(o.value) }))
    .filter(o => Number.isFinite(o.value));
  return {
    asOf: new Date().toISOString(),
    hy: toPts(hyJ).slice(-260),
    ig: toPts(igJ).slice(-260),
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

// ── 指標 2 輔助：企業利潤週期（BEA NIPA Table 1.14，季頻）──
// 資料來源 BEA NIPA Table 1.14「Gross Value Added of Domestic Corporate Business」，取整體
// 企業（金融＋非金融合計）的稅前利潤（Line 11「Corporate profits with IVA and CCAdj」，即
// 一般引用的headline 企業利潤數字）及其下游分解（Line 12-15：企業所得稅、稅後利潤、淨股利、
// 未分配利潤）。曾用 GetParameterValues 逐一核對 NIPA 表清單才挑到這張——Table 1.14 標題
// 沒有「Profits」字樣（叫 Gross Value Added），容易被誤認成別的表；真正的 headline「Corporate
// profits with IVA and CCAdj」數字實際上就藏在這張表的 Line 11。
//
// ⚠️ 是季頻，且公布時間落後：本季資料通常要等到該季結束後約 2 個月的 GDP「第二次估計」
// 才會出現在 API 裡（第一次的「預估」估計還沒有完整利潤細項），不像儀表板其他多數指標
// 是月頻甚至週頻——nextUpdateNote 是依此估算的概略下次更新月份，不是 BEA 官方公告的日期。
const BEA_CORP_PROFITS_TABLE = 'T11400';
const BEA_CORP_PROFITS_LINES = {
  11: 'profitsBeforeTax', // Corporate profits with IVA and CCAdj（headline 數字）
  12: 'taxes',            // Taxes on corporate income
  13: 'profitsAfterTax',  // Profits after tax with IVA and CCAdj
  14: 'dividends',        // Net dividends
  15: 'undistributed',    // Undistributed profits with IVA and CCAdj
};

async function fetchCorpProfits() {
  if (!BEA_API_KEY) {
    const err = new Error('BEA_API_KEY 未設定或尚未啟用');
    err.code = 'NO_KEY';
    throw err;
  }
  const thisYear = new Date().getUTCFullYear();
  const years = [];
  for (let y = thisYear - 9; y <= thisYear; y++) years.push(y); // 拉 10 年，季頻資料點少，要拉長才看得出趨勢
  const url = `https://apps.bea.gov/api/data/?UserID=${BEA_API_KEY}&method=GetData&datasetname=NIPA&TableName=${BEA_CORP_PROFITS_TABLE}&Frequency=Q&Year=${years.join(',')}&ResultFormat=JSON`;
  const r = await fetch(url);
  if (!r.ok) throw new Error('BEA API HTTP ' + r.status);
  const j = await r.json();
  const results = j.BEAAPI && j.BEAAPI.Results;
  if (!results || results.Error) {
    throw new Error('BEA API: ' + (results && results.Error && results.Error.APIErrorDescription || 'unknown error'));
  }
  const rows = results.Data || [];

  const byField = {};
  for (const field of Object.values(BEA_CORP_PROFITS_LINES)) byField[field] = new Map();
  for (const row of rows) {
    const field = BEA_CORP_PROFITS_LINES[Number(row.LineNumber)];
    if (!field) continue;
    const m = /^(\d{4})Q([1-4])$/.exec(row.TimePeriod);
    if (!m) continue;
    const value = Number(String(row.DataValue).replace(/,/g, '')); // DataValue 是千分位字串，如 "1,757,834"
    if (!Number.isFinite(value)) continue;
    const date = `${m[1]}-${String((Number(m[2]) - 1) * 3 + 1).padStart(2, '0')}-01`; // 用該季起始月表示整季
    byField[field].set(date, value);
  }

  const dates = [...byField.profitsBeforeTax.keys()].sort();
  const seriesFor = field => dates.map(d => ({ date: d, value: byField[field].has(d) ? byField[field].get(d) : null }));
  const preTax = seriesFor('profitsBeforeTax');

  const qoq = preTax.map((d, i) => {
    const prior = preTax[i - 1];
    if (!prior || d.value == null || prior.value == null) return null;
    return { date: d.date, value: Number((((d.value / prior.value) - 1) * 100).toFixed(2)) };
  }).filter(Boolean);
  const yoy = preTax.map((d, i) => {
    const prior = preTax[i - 4];
    if (!prior || d.value == null || prior.value == null) return null;
    return { date: d.date, value: Number((((d.value / prior.value) - 1) * 100).toFixed(2)) };
  }).filter(Boolean);

  const latestDate = dates[dates.length - 1] || null;
  const breakdown = latestDate
    ? Object.entries(BEA_CORP_PROFITS_LINES).map(([lineNumber, field]) => ({
        lineNumber: Number(lineNumber),
        field,
        value: byField[field].has(latestDate) ? byField[field].get(latestDate) : null,
      }))
    : [];

  // 下次更新概略估算：latestDate 是「已有資料的最新一季」的起始月；下一季結束後約 2 個月
  // 才會有下一筆資料（見上方⚠️說明）——用 Date 物件算，讓月份跨年自動進位不必手動處理。
  let nextUpdateNote = null;
  if (latestDate) {
    const [ly, lm] = latestDate.split('-').map(Number);
    const d = new Date(Date.UTC(ly, (lm - 1) + 7, 1)); // +7 = 下一季(+3個月) 結束(+2) 再等公布(+2)
    nextUpdateNote = `${d.getUTCFullYear()}年${d.getUTCMonth() + 1}月`;
  }

  return {
    asOf: new Date().toISOString(),
    unit: '$M（Current Dollars, SAAR 季調年化）',
    latestQuarter: latestDate,
    nextUpdateNote,
    preTax: preTax.slice(-40),
    qoq: qoq.slice(-40),
    yoy: yoy.slice(-40),
    breakdown,
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

// ── ⑦ NFP：非農就業與勞動市場斷層（獨立卡片，非六項框架原生成員）──
// 三邊資料拼起來：
//   ① 12 個產業別本月 MoM 新增就業 + 家戶調查（失業率／勞動參與率／U-6）
//      -> BLS CES/CPS API，一次 POST 抓齊全部序列，免金鑰
//   ② Headline 總非農就業「初值 vs 修訂後值」-> FRED PAYEMS，同一序列分別用
//      output_type=1（目前最新、已修訂）與 output_type=4（當初第一次公布時的水準）各拉一次
//   ③ 官方 1 個月擴散指數（多少比例產業在新增雇用）-> FRED SMS00000000000000021
//      （BLS 原始擴散指數的 FRED 鏡像，不是自行依產業資料重算）
//
// ⚠️ 修訂方法論（老實寫在這裡）：這裡「初值 MoM 新增」是用「該月自己第一次公布時的水準」
// 減去「上一個月自己第一次公布時的水準」，是可重現的一致定義；財經媒體常引用的「當月發布時
// 比較的是『上月當下已知、可能已修訂一次的水準』」略有不同，兩者差異通常是幾千人等級，
// 不影響修訂方向的判讀，但精確數字可能對不上媒體當時報導的數字。
const NFP_INDUSTRIES = [
  { label: '礦業與伐木', ids: ['CES1000000001'] },
  { label: '營建業', ids: ['CES2000000001'] },
  { label: '製造業', ids: ['CES3000000001'] },
  { label: '貿易／運輸／公用事業', ids: ['CES4000000001'] },
  { label: '資訊業', ids: ['CES5000000001'] },
  { label: '金融業', ids: ['CES5500000001'] },
  { label: '專業與商業服務', ids: ['CES6000000001'] },
  { label: '教育與醫療服務', ids: ['CES6500000001'] },
  { label: '休閒與餐旅業', ids: ['CES7000000001'] },
  { label: '其他服務業', ids: ['CES8000000001'] },
  { label: '聯邦政府', ids: ['CES9091000001'] },
  { label: '州及地方政府', ids: ['CES9092000001', 'CES9093000001'] }, // BLS 無合併序列，加總州+地方兩條官方序列
];
const NFP_HOUSEHOLD_SERIES = {
  unemployment: 'LNS14000000', // 失業率 U-3
  participation: 'LNS11300000', // 勞動參與率
  u6: 'LNS13327709', // 廣義失業率 U-6
};
const NFP_DIFFUSION_SERIES = 'SMS00000000000000021'; // FRED：Diffusion Indexes, 1-Month Span, Total Nonfarm

async function fetchNfp() {
  if (!FRED_API_KEY) {
    const err = new Error('FRED_API_KEY 未設定');
    err.code = 'NO_KEY';
    throw err;
  }
  const thisYear = new Date().getUTCFullYear();
  const startYear = thisYear - 5;
  const obsStart = `${startYear}-01-01`;

  // ① BLS：12 個產業別（13 條底層序列，州+地方分開拉再加總）+ 家戶調查 3 條，一次 POST 抓齊
  const blsIds = [...new Set(NFP_INDUSTRIES.flatMap(i => i.ids))].concat(Object.values(NFP_HOUSEHOLD_SERIES));
  const blsR = await fetch('https://api.bls.gov/publicAPI/v2/timeseries/data/', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ seriesid: blsIds, startyear: String(startYear), endyear: String(thisYear) }),
  });
  if (!blsR.ok) throw new Error('BLS API HTTP ' + blsR.status);
  const blsJ = await blsR.json();
  if (blsJ.status !== 'REQUEST_SUCCEEDED') throw new Error('BLS API: ' + (blsJ.message || []).join('; '));

  const blsBySeries = {};
  for (const s of blsJ.Results.series) {
    const pts = s.data
      .map(d => ({ date: periodToDate(d.year, d.period), value: Number(d.value) }))
      .filter(d => d.date && Number.isFinite(d.value))
      .sort((a, b) => a.date.localeCompare(b.date));
    blsBySeries[s.seriesID] = pts;
  }

  const industries = NFP_INDUSTRIES.map(ind => {
    // 多條序列（州+地方）逐日期加總；單條序列直接沿用
    const dates = [...new Set(ind.ids.flatMap(id => (blsBySeries[id] || []).map(p => p.date)))].sort();
    const summed = dates.map(date => ({
      date,
      value: ind.ids.reduce((sum, id) => {
        const pt = (blsBySeries[id] || []).find(p => p.date === date);
        return pt ? sum + pt.value : sum;
      }, 0),
    }));
    const latest = summed[summed.length - 1];
    const prior = summed[summed.length - 2];
    const momChange = (latest && prior) ? Number((latest.value - prior.value).toFixed(1)) : null;
    return { label: ind.label, level: latest ? latest.value : null, momChange, latestDate: latest ? latest.date : null };
  });

  const household = {};
  for (const [key, id] of Object.entries(NFP_HOUSEHOLD_SERIES)) {
    household[key] = (blsBySeries[id] || []).slice(-60);
  }

  // ② FRED PAYEMS：現值（output_type=1）vs 初次公布值（output_type=4）
  const [curR, initR] = await Promise.all([
    fetch(`https://api.stlouisfed.org/fred/series/observations?series_id=PAYEMS&api_key=${FRED_API_KEY}&file_type=json&sort_order=asc&observation_start=${obsStart}&output_type=1`),
    fetch(`https://api.stlouisfed.org/fred/series/observations?series_id=PAYEMS&api_key=${FRED_API_KEY}&file_type=json&sort_order=asc&observation_start=${obsStart}&output_type=4&realtime_start=${obsStart}&realtime_end=9999-12-31`),
  ]);
  if (!curR.ok) throw new Error('FRED PAYEMS(current) HTTP ' + curR.status);
  if (!initR.ok) throw new Error('FRED PAYEMS(initial) HTTP ' + initR.status);
  const curJ = await curR.json();
  const initJ = await initR.json();
  const toLevelPts = j => (j.observations || [])
    .map(o => ({ date: o.date, value: Number(o.value) }))
    .filter(o => Number.isFinite(o.value));
  const curLevels = toLevelPts(curJ);
  const initLevels = toLevelPts(initJ);

  function toMomSeries(levels) {
    return levels.map((d, i) => {
      const prior = levels[i - 1];
      if (!prior) return null;
      return { date: d.date, value: Number((d.value - prior.value).toFixed(0)) };
    }).filter(Boolean);
  }
  const currentMoM = toMomSeries(curLevels);
  const initialMoM = toMomSeries(initLevels);

  const latestCur = currentMoM[currentMoM.length - 1] || null;
  const latestMonth = latestCur ? latestCur.date : null;
  const latestInit = latestMonth ? (initialMoM.find(p => p.date === latestMonth) || null) : null;
  // 上一個月的「初值 vs 現值」修訂差額：用現值序列倒數第二筆的日期，對照初值序列同日期那筆
  const priorDate = currentMoM.length > 1 ? currentMoM[currentMoM.length - 2].date : null;
  const priorCur = priorDate ? currentMoM.find(p => p.date === priorDate) : null;
  const priorInit = priorDate ? initialMoM.find(p => p.date === priorDate) : null;
  const priorRevision = (priorCur && priorInit) ? {
    date: priorDate,
    initialChange: priorInit.value,
    currentChange: priorCur.value,
    revisionDelta: priorCur.value - priorInit.value,
  } : null;

  // ③ FRED：BLS 官方 1 個月擴散指數（Total Nonfarm）
  const diffR = await fetch(`https://api.stlouisfed.org/fred/series/observations?series_id=${NFP_DIFFUSION_SERIES}&api_key=${FRED_API_KEY}&file_type=json&sort_order=asc&observation_start=${obsStart}`);
  if (!diffR.ok) throw new Error('FRED diffusion index HTTP ' + diffR.status);
  const diffJ = await diffR.json();
  const diffusion = toLevelPts(diffJ);

  return {
    asOf: new Date().toISOString(),
    headline: {
      latestMonth,
      latestInitialChange: latestInit ? latestInit.value : null,
      latestCurrentChange: latestCur ? latestCur.value : null,
      priorRevision,
      currentMoM: currentMoM.slice(-60),
      initialMoM: initialMoM.slice(-60),
    },
    household,
    diffusion: diffusion.slice(-60),
    industries,
  };
}

const ENDPOINTS = {
  '/api/cpi-stickiness': () => cached('cpi', fetchCpiStickiness),
  '/api/real-pce': () => cached('pce', fetchRealPce),
  '/api/corp-profits': () => cached('corpprofits', fetchCorpProfits),
  '/api/nfp': () => cached('nfp', fetchNfp),
  '/api/jobless-claims': () => cached('jobless', fetchJoblessClaims),
  '/api/pce-breadth': () => cached('breadth', fetchPceBreadth),
  '/api/credit-spreads': () => cached('spreads', fetchCreditSpreads),
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
