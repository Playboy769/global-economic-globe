'use strict';
// 化妝品圖鑑的資料層。零 npm 依賴，沿用本 repo 既有 analytics.js / gaoye-mock-exam 的 store.js
// 家法：node:sqlite（Node 22.13 起內建，Dockerfile 因此釘 node:24-alpine）寫到 Railway
// persistent volume 掛載點 /data。
//
// 為什麼一定要 volume：Railway 每次 redeploy 都會清掉容器檔案系統，寫在映像檔裡的任何路徑
// 都會在下次 push 時無聲重置。/data 是 analytics.js 與 gaoye-mock-exam 已經在用的同一個慣例
// 掛載點，在 Windows 開發機上會解析成 C:\data（本地開發請用 COSMETICS_DATA_DIR 覆寫）。
//
// 種子資料：data/part*.json 是這個圖鑑的初始 166 筆。第一次啟動（products 表為空）時匯入；
// 之後就以資料庫為準，不會每次啟動覆蓋使用者在線上編輯的內容。要強制重新匯入請設
// COSMETICS_RESEED=1（會覆寫同 id 的列，但不會刪掉使用者自己新增的產品）。

const fs = require('fs');
const path = require('path');
const { DatabaseSync } = require('node:sqlite');

const DATA_DIR = process.env.COSMETICS_DATA_DIR || '/data';
const DB_PATH = process.env.COSMETICS_DB_PATH || path.join(DATA_DIR, 'cosmetics.db');
const SEED_DIR = path.join(__dirname, 'data');

// 陣列/物件欄位在 SQLite 裡以 JSON 字串存放，取出時還原。
const JSON_FIELDS = ['efficacy', 'key_ingredients', 'free_of', 'skin_types', 'scenes', 'shades', 'pros', 'cons'];
const TEXT_FIELDS = [
  'brand', 'brand_zh', 'origin', 'category', 'subcategory', 'name_local', 'name_zh',
  'volume', 'shade_note', 'review', 'image_url', 'official_url', 'source_note',
  'notes_top', 'notes_mid', 'notes_base', 'longevity',
];
const NUM_FIELDS = ['price_twd', 'launch_year', 'rating'];
const ALL_FIELDS = ['id', ...TEXT_FIELDS, ...NUM_FIELDS, ...JSON_FIELDS];

let db = null;

function open() {
  if (db) return db;
  const dir = path.dirname(DB_PATH);
  // 真正掛上的 Railway volume 在行程啟動前就已經是個目錄（即使是空的），所以走到這裡
  // 還不存在，多半代表 volume 根本沒掛上，而不是「今天是第一次寫入」。
  if (!fs.existsSync(dir)) {
    console.warn('[store] ' + dir + ' 不存在，建立中；若這是 Railway 環境，'
      + '請確認 volume 已掛在 ' + dir + '，否則資料會在下次部署時消失。');
    fs.mkdirSync(dir, { recursive: true });
  }
  db = new DatabaseSync(DB_PATH);
  db.exec('PRAGMA journal_mode = WAL');
  db.exec(`
    CREATE TABLE IF NOT EXISTS products (
      id              TEXT PRIMARY KEY,
      brand           TEXT NOT NULL,
      brand_zh        TEXT,
      origin          TEXT,
      category        TEXT NOT NULL,
      subcategory     TEXT,
      name_local      TEXT,
      name_zh         TEXT NOT NULL,
      volume          TEXT,
      price_twd       INTEGER,
      launch_year     INTEGER,
      efficacy        TEXT,
      key_ingredients TEXT,
      free_of         TEXT,
      skin_types      TEXT,
      scenes          TEXT,
      shades          TEXT,
      shade_note      TEXT,
      rating          REAL,
      review          TEXT,
      pros            TEXT,
      cons            TEXT,
      image_url       TEXT,
      official_url    TEXT,
      source_note     TEXT,
      notes_top       TEXT,
      notes_mid       TEXT,
      notes_base      TEXT,
      longevity       TEXT,
      created_at      TEXT NOT NULL,
      updated_at      TEXT NOT NULL
    );
    CREATE INDEX IF NOT EXISTS idx_products_category ON products(category);
    CREATE INDEX IF NOT EXISTS idx_products_brand    ON products(brand);
    CREATE INDEX IF NOT EXISTS idx_products_origin   ON products(origin);
  `);
  seedIfEmpty();
  return db;
}

function readSeedFiles() {
  if (!fs.existsSync(SEED_DIR)) return [];
  const files = fs.readdirSync(SEED_DIR).filter(f => /^part.*\.json$/.test(f)).sort();
  let rows = [];
  for (const f of files) {
    try {
      const arr = JSON.parse(fs.readFileSync(path.join(SEED_DIR, f), 'utf8'));
      if (Array.isArray(arr)) rows = rows.concat(arr);
    } catch (err) {
      console.error('[store] 種子檔解析失敗：' + f + ' — ' + err.message);
    }
  }
  return rows;
}

function seedIfEmpty() {
  const n = db.prepare('SELECT COUNT(*) AS c FROM products').get().c;
  const reseed = process.env.COSMETICS_RESEED === '1';
  if (n > 0 && !reseed) return;
  const rows = readSeedFiles();
  if (!rows.length) {
    console.warn('[store] 找不到種子資料（data/part*.json），資料庫維持空的。');
    return;
  }
  let ok = 0;
  for (const r of rows) {
    try { upsert(r, { seeding: true }); ok++; } catch (err) {
      console.error('[store] 種子匯入失敗 id=' + r.id + '：' + err.message);
    }
  }
  console.log('[store] 種子資料已匯入 ' + ok + '/' + rows.length + ' 筆'
    + (reseed ? '（COSMETICS_RESEED=1 強制重匯）' : ''));
}

function nowISO() { return new Date().toISOString(); }

function normalize(input) {
  const out = { id: String(input.id || '').trim() };
  if (!out.id) throw new Error('缺少 id');
  if (!/^[a-z0-9][a-z0-9-]{0,80}$/.test(out.id)) throw new Error('id 只能用小寫英數與連字號');
  for (const f of TEXT_FIELDS) {
    const v = input[f];
    out[f] = (v === undefined || v === null || v === '') ? null : String(v);
  }
  for (const f of NUM_FIELDS) {
    const v = input[f];
    out[f] = (v === undefined || v === null || v === '') ? null : Number(v);
    if (out[f] !== null && !Number.isFinite(out[f])) out[f] = null;
  }
  for (const f of JSON_FIELDS) {
    const v = input[f];
    out[f] = JSON.stringify(Array.isArray(v) ? v : []);
  }
  if (!out.brand) throw new Error('缺少品牌');
  if (!out.name_zh) throw new Error('缺少中文品名');
  if (!out.category) throw new Error('缺少分類');
  // image_url 只接受 http(s)，避免有人塞 javascript: 進來後被前端當成 src。
  if (out.image_url && !/^https?:\/\//i.test(out.image_url)) out.image_url = null;
  if (out.official_url && !/^https?:\/\//i.test(out.official_url)) out.official_url = null;
  return out;
}

function upsert(input, opts = {}) {
  const d = open() && normalize(input);
  const existing = db.prepare('SELECT created_at FROM products WHERE id = ?').get(d.id);
  const ts = nowISO();
  const created = existing ? existing.created_at : ts;
  const cols = ALL_FIELDS.concat(['created_at', 'updated_at']);
  const sql = 'INSERT OR REPLACE INTO products (' + cols.join(',') + ') VALUES ('
    + cols.map(() => '?').join(',') + ')';
  const vals = ALL_FIELDS.map(f => d[f]).concat([created, ts]);
  db.prepare(sql).run(...vals);
  return { ...hydrate(db.prepare('SELECT * FROM products WHERE id = ?').get(d.id)), _seeded: !!opts.seeding };
}

function hydrate(row) {
  if (!row) return null;
  const out = { ...row };
  for (const f of JSON_FIELDS) {
    try { out[f] = JSON.parse(row[f] || '[]'); } catch { out[f] = []; }
  }
  return out;
}

function get(id) {
  open();
  return hydrate(db.prepare('SELECT * FROM products WHERE id = ?').get(String(id)));
}

function remove(id) {
  open();
  const r = db.prepare('DELETE FROM products WHERE id = ?').run(String(id));
  return r.changes > 0;
}

// 陣列欄位存成 JSON 字串，用 LIKE '%"值"%' 過濾。值本身來自資料而非自由輸入，
// 且一律走參數綁定，不會有注入問題。
function arrayLike(field, value) {
  return { clause: field + ' LIKE ?', param: '%"' + String(value).replace(/["%_]/g, '') + '"%' };
}

const SORTS = {
  'rating-desc': 'rating DESC NULLS LAST, name_zh ASC',
  'price-asc': 'price_twd ASC NULLS LAST, name_zh ASC',
  'price-desc': 'price_twd DESC NULLS LAST, name_zh ASC',
  'year-desc': 'launch_year DESC NULLS LAST, name_zh ASC',
  'brand-asc': 'brand ASC, name_zh ASC',
};

function list(q = {}) {
  open();
  const where = [];
  const params = [];
  if (q.category) { where.push('category = ?'); params.push(q.category); }
  if (q.subcategory) { where.push('subcategory = ?'); params.push(q.subcategory); }
  if (q.brand) { where.push('brand = ?'); params.push(q.brand); }
  if (q.origin) { where.push('origin = ?'); params.push(q.origin); }
  if (q.skin_type) { const a = arrayLike('skin_types', q.skin_type); where.push(a.clause); params.push(a.param); }
  if (q.efficacy) { const a = arrayLike('efficacy', q.efficacy); where.push(a.clause); params.push(a.param); }
  if (q.scene) { const a = arrayLike('scenes', q.scene); where.push(a.clause); params.push(a.param); }
  if (q.price_min) { where.push('price_twd >= ?'); params.push(Number(q.price_min)); }
  if (q.price_max) { where.push('price_twd <= ?'); params.push(Number(q.price_max)); }
  if (q.has_shades === '1') { where.push("shades != '[]'"); }
  if (q.q) {
    const like = '%' + String(q.q).trim() + '%';
    where.push('(brand LIKE ? OR brand_zh LIKE ? OR name_zh LIKE ? OR name_local LIKE ?'
      + ' OR key_ingredients LIKE ? OR review LIKE ? OR subcategory LIKE ?)');
    for (let i = 0; i < 7; i++) params.push(like);
  }
  const sql = 'SELECT * FROM products'
    + (where.length ? ' WHERE ' + where.join(' AND ') : '')
    + ' ORDER BY ' + (SORTS[q.sort] || SORTS['rating-desc']);
  return db.prepare(sql).all(...params).map(hydrate);
}

// 側欄篩選用的分面統計。分類/子分類/品牌/產地直接 GROUP BY；陣列欄位在 JS 端展開計數
// （166 筆的規模下不值得為此另建一張正規化表）。
function facets() {
  open();
  const rows = db.prepare('SELECT * FROM products').all().map(hydrate);
  const count = (map, key) => { if (key) map[key] = (map[key] || 0) + 1; };
  const category = {}, subcategory = {}, brand = {}, origin = {},
    skin_types = {}, efficacy = {}, scenes = {}, free_of = {};
  for (const r of rows) {
    count(category, r.category);
    count(subcategory, r.subcategory);
    count(brand, r.brand);
    count(origin, r.origin);
    for (const v of r.skin_types) count(skin_types, v);
    for (const v of r.efficacy) count(efficacy, v);
    for (const v of r.scenes) count(scenes, v);
    for (const v of r.free_of) count(free_of, v);
  }
  const prices = rows.map(r => r.price_twd).filter(v => typeof v === 'number' && v > 0);
  return {
    total: rows.length,
    with_image: rows.filter(r => r.image_url).length,
    category, subcategory, brand, origin, skin_types, efficacy, scenes, free_of,
    price_min: prices.length ? Math.min(...prices) : 0,
    price_max: prices.length ? Math.max(...prices) : 0,
    // 子分類要能跟著分類連動，所以另外給一份「分類 → 子分類」對照。
    subcategory_by_category: rows.reduce((acc, r) => {
      if (!r.category || !r.subcategory) return acc;
      (acc[r.category] = acc[r.category] || {});
      acc[r.category][r.subcategory] = (acc[r.category][r.subcategory] || 0) + 1;
      return acc;
    }, {}),
  };
}

module.exports = { open, list, get, upsert, remove, facets };
