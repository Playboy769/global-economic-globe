'use strict';
// 化妝品圖鑑（日韓）— HTTP 伺服器。零 npm 依賴，與本 repo 既有 server.js 家法一致。
//
// 職責：① 靜態送出 index.html
//       ② /api/* 讀寫 SQLite（store.js），供前端卡片牆與篩選側欄使用
//
// 沒有登入閘門：依需求設定為「公開瀏覽、不上架作品集」。寫入 API 預設同樣開放，
// 若要擋住陌生人改資料，設 service 變數 EDIT_TOKEN=<任意字串>，之後所有寫入
// （POST/PUT/DELETE）都必須帶 x-edit-token 標頭，前端會在需要時跳出輸入框。
//
// ⚠️ 請求標頭（Host / X-Forwarded-*）一律不可信，本檔沒有任何地方用它們決定授權。
// /healthz 刻意放在所有處理之前，否則 Railway 健康檢查會被其他邏輯影響。

const http = require('http');
const fs = require('fs');
const path = require('path');
const store = require('./store');

const PORT = Number(process.env.PORT) || 8080;
const ROOT = __dirname;
const EDIT_TOKEN = process.env.EDIT_TOKEN || '';
const MAX_BODY = 256 * 1024; // 單筆產品的 JSON 不該超過這個量級

const TYPES = {
  '.html': 'text/html; charset=utf-8',
  '.js': 'text/javascript; charset=utf-8',
  '.css': 'text/css; charset=utf-8',
  '.json': 'application/json; charset=utf-8',
  '.svg': 'image/svg+xml',
  '.ico': 'image/x-icon',
};

function sendJSON(res, code, obj) {
  const body = JSON.stringify(obj);
  res.writeHead(code, {
    'content-type': 'application/json; charset=utf-8',
    'content-length': Buffer.byteLength(body),
    'cache-control': 'no-store',
  });
  res.end(body);
}

function readBody(req) {
  return new Promise((resolve, reject) => {
    let size = 0;
    const chunks = [];
    req.on('data', c => {
      size += c.length;
      if (size > MAX_BODY) { reject(new Error('內容過大')); req.destroy(); return; }
      chunks.push(c);
    });
    req.on('end', () => {
      if (!chunks.length) { resolve({}); return; }
      try { resolve(JSON.parse(Buffer.concat(chunks).toString('utf8'))); }
      catch (err) { reject(new Error('JSON 格式錯誤：' + err.message)); }
    });
    req.on('error', reject);
  });
}

// 寫入權限：沒設 EDIT_TOKEN 就一律放行（預設公開可編輯）；設了就必須帶對的標頭。
function canWrite(req) {
  if (!EDIT_TOKEN) return true;
  const got = req.headers['x-edit-token'];
  return typeof got === 'string' && got.length === EDIT_TOKEN.length && got === EDIT_TOKEN;
}

function serveStatic(req, res, pathname) {
  const rel = pathname === '/' ? 'index.html' : pathname.replace(/^\/+/, '');
  const file = path.resolve(ROOT, rel);
  if (file !== ROOT && !file.startsWith(ROOT + path.sep)) {
    res.writeHead(403).end('forbidden');
    return;
  }
  // 伺服器程式與資料庫種子檔不對外提供（種子內容本來就會由 /api/products 送出）。
  const base = path.basename(file);
  if (base === 'server.js' || base === 'store.js' || file.startsWith(path.join(ROOT, 'data') + path.sep)) {
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

const server = http.createServer(async (req, res) => {
  let url;
  try { url = new URL(req.url, 'http://localhost'); }
  catch { res.writeHead(400).end('bad request'); return; }
  const p = url.pathname;

  if (p === '/healthz') { res.writeHead(200, { 'content-type': 'text/plain' }).end('ok'); return; }

  try {
    if (p === '/api/facets' && req.method === 'GET') {
      sendJSON(res, 200, store.facets());
      return;
    }

    if (p === '/api/products' && req.method === 'GET') {
      const q = Object.fromEntries(url.searchParams.entries());
      const items = store.list(q);
      sendJSON(res, 200, { count: items.length, items });
      return;
    }

    if (p === '/api/products' && req.method === 'POST') {
      if (!canWrite(req)) { sendJSON(res, 403, { error: '需要編輯權杖' }); return; }
      const body = await readBody(req);
      const saved = store.upsert(body);
      sendJSON(res, 200, { item: saved });
      return;
    }

    const m = p.match(/^\/api\/products\/([A-Za-z0-9._-]+)$/);
    if (m) {
      const id = m[1];
      if (req.method === 'GET') {
        const item = store.get(id);
        if (!item) { sendJSON(res, 404, { error: '找不到這筆產品' }); return; }
        sendJSON(res, 200, { item });
        return;
      }
      if (req.method === 'PUT') {
        if (!canWrite(req)) { sendJSON(res, 403, { error: '需要編輯權杖' }); return; }
        const body = await readBody(req);
        const saved = store.upsert({ ...body, id });
        sendJSON(res, 200, { item: saved });
        return;
      }
      if (req.method === 'DELETE') {
        if (!canWrite(req)) { sendJSON(res, 403, { error: '需要編輯權杖' }); return; }
        sendJSON(res, 200, { deleted: store.remove(id) });
        return;
      }
    }

    if (p.startsWith('/api/')) { sendJSON(res, 404, { error: '沒有這個端點' }); return; }

    serveStatic(req, res, p);
  } catch (err) {
    console.error('[server] ' + req.method + ' ' + p + ' — ' + err.stack);
    sendJSON(res, 400, { error: err.message || '請求處理失敗' });
  }
});

store.open();
server.listen(PORT, () => {
  console.log('[server] 化妝品圖鑑 listening on ' + PORT
    + (EDIT_TOKEN ? '（寫入需 x-edit-token）' : '（寫入未設權杖，公開可編輯）'));
});
