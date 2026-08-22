'use strict';
// 高業模擬測驗服務。零 npm 依賴，只用 node 內建模組 —— 與本 repo 既有
// server.js / auth.js / analytics.js 的家法一致。
//
// 職責：① 靜態送出 index.html（單一自足檔案，離線也能作答）
//       ② /api/* 把作答紀錄寫進 Railway volume 上的 SQLite（見 store.js）
//       ③ 選用的登入閘門（**預設關閉**，見下方 AUTH_REQUIRED）
//
// 登入若啟用則沿用中央登入系統：OutsideFramework 是唯一跟 Google 對話的服務，這裡只
// 驗證它簽發的 handoff token，驗過一次後自己發一張短期 session cookie。閘門邏輯（含
// 迴圈斷路器）刻意與 globe-invest/server.js 同一套寫法，出事時兩邊症狀一致好排查。

const http = require('http');
const fs = require('fs');
const path = require('path');
const auth = require('./auth');
const store = require('./store');

const PORT = Number(process.env.PORT) || 8080;
const ROOT = __dirname;

// ── Auth ──
const CENTRAL_AUTH_ORIGIN = process.env.CENTRAL_AUTH_ORIGIN || 'https://ofw.up.railway.app';
const AUTH_SECRET = process.env.AUTH_SIGNING_SECRET || '';
const AUTHORIZED_EMAILS = (process.env.AUTHORIZED_EMAILS || '')
  .split(',')
  .map(e => e.trim().toLowerCase())
  .filter(Boolean);
const SESSION_COOKIE = 'gme_sid';
// 標記「已經送這個瀏覽器去中央服務拿過 handoff token」，供下面的迴圈斷路器判斷。
const HANDOFF_TRY_COOKIE = 'gme_hs';
const SESSION_TTL_SEC = 12 * 60 * 60;

// cookie 旗標與本服務 origin 一律來自設定，絕不取自請求標頭：Host 與
// X-Forwarded-Proto 都是攻擊者可設的，用後者決定 Secure 旗標等於讓人把它拔掉。
const IS_PROD =
  process.env.NODE_ENV === 'production' ||
  !!process.env.RAILWAY_PROJECT_ID ||
  !!process.env.RAILWAY_ENVIRONMENT_NAME;
const SELF_ORIGIN = process.env.PUBLIC_ORIGIN || 'https://gaoye-mock-exam-production.up.railway.app';
const ALLOWED_HOSTS = new Set([
  new URL(SELF_ORIGIN).host.toLowerCase(),
  'gaoye-mock-exam-production.up.railway.app',
]);
const LOCAL_HOST_RE = /^(localhost|127\.0\.0\.1)(:\d+)?$/;

// 登入預設關閉（2026-08-21）。原本線上強制走 ofw 中央登入，但 Google 回呼一直卡在
// 「Invalid or expired login request」，使用者被鎖在自己的工具外面，故改為預設不擋。
//
// 整合本身沒有刪除，只是改成 opt-in：把 service 變數 REQUIRE_LOGIN 設成 1（且
// AUTH_SIGNING_SECRET / AUTHORIZED_EMAILS 有值）就會恢復閘門，不必再改一次程式。
// 本地開發一律免登入。
const AUTH_REQUIRED = IS_PROD && /^(1|true|yes)$/i.test(process.env.REQUIRE_LOGIN || '');

if (AUTH_REQUIRED && !AUTH_SECRET) {
  console.error('WARNING: AUTH_SIGNING_SECRET 未設定 —— 每個請求都會被當成未登入。');
}
if (AUTH_REQUIRED && !AUTHORIZED_EMAILS.length) {
  console.error('WARNING: AUTHORIZED_EMAILS 未設定 —— 沒有人登得進來。');
}

function nowSec() { return Math.floor(Date.now() / 1000); }

function selfOrigin(req) {
  const host = String(req.headers.host || '').toLowerCase();
  if (!IS_PROD && LOCAL_HOST_RE.test(host)) return 'http://' + host;
  if (ALLOWED_HOSTS.has(host)) return 'https://' + host;
  return SELF_ORIGIN;
}

function getSessionEmail(req) {
  // verifyFor 由 aud 衍生金鑰，所以為別的服務簽的 token 在這裡簽章根本對不上。
  const cookies = auth.parseCookies(req.headers.cookie);
  const payload = auth.verifyFor(cookies[SESSION_COOKIE], AUTH_SECRET, {
    aud: selfOrigin(req),
    typ: auth.TYP_SESSION,
  });
  if (!payload || !payload.email) return null;
  if (!AUTHORIZED_EMAILS.includes(payload.email.toLowerCase())) return null;
  return payload.email.toLowerCase();
}

const authFailLimiter = auth.makeRateLimiter({ windowMs: 10 * 60 * 1000, max: 30 });

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

function readBody(req, cb) {
  let n = 0;
  const chunks = [];
  req.on('data', c => {
    n += c.length;
    // 一份三科 150 題的作答明細大約 60–80KB，1MB 已經很寬鬆了。
    if (n > 1024 * 1024) { req.destroy(); return; }
    chunks.push(c);
  });
  req.on('end', () => {
    try { cb(null, JSON.parse(Buffer.concat(chunks).toString('utf8'))); }
    catch (e) { cb(e); }
  });
  req.on('error', e => cb(e));
}

function serveStatic(req, res, pathname) {
  const rel = pathname === '/' ? 'index.html' : pathname.replace(/^\/+/, '');
  // 路徑穿越防護：解析後必須仍在 ROOT 底下。auth.js 不對外送。
  const file = path.resolve(ROOT, rel);
  if (file !== ROOT && !file.startsWith(ROOT + path.sep)) {
    res.writeHead(403).end('forbidden');
    return;
  }
  if (path.basename(file) === 'auth.js' || path.basename(file) === 'store.js'
      || path.basename(file) === 'server.js') {
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

// 回傳 email 代表放行；回傳 null 代表這個請求已經被導向/擋掉了，呼叫端直接 return。
function gate(req, res) {
  if (!AUTH_REQUIRED) {
    // email 欄位保留著：之後若把 REQUIRE_LOGIN 打開，舊紀錄仍在這個身分底下，
    // 不需要再做一次 schema 遷移。
    return IS_PROD ? (process.env.OWNER_EMAIL || 'owner@local')
                   : (process.env.DEV_EMAIL || 'dev@localhost');
  }

  const email = getSessionEmail(req);
  if (email) return email;

  const incomingUrl = new URL(req.url, selfOrigin(req));
  const incomingToken = incomingUrl.searchParams.get('auth');
  const tokenPayload = incomingToken
    ? auth.verifyFor(incomingToken, AUTH_SECRET, { aud: selfOrigin(req), typ: auth.TYP_HANDOFF })
    : null;
  const tokenOk = tokenPayload && tokenPayload.email &&
    AUTHORIZED_EMAILS.includes(tokenPayload.email.toLowerCase());

  if (tokenOk) {
    const sessionToken = auth.signFor(
      { email: tokenPayload.email.toLowerCase(), aud: selfOrigin(req),
        typ: auth.TYP_SESSION, exp: nowSec() + SESSION_TTL_SEC },
      AUTH_SECRET
    );
    incomingUrl.searchParams.delete('auth');
    res.writeHead(302, {
      'Set-Cookie': [
        auth.cookieHeader(SESSION_COOKIE, sessionToken,
          { maxAgeSec: SESSION_TTL_SEC, httpOnly: true, secure: IS_PROD }),
        auth.clearCookieHeader(HANDOFF_TRY_COOKIE, { secure: IS_PROD }),
      ],
      // ?auth= 會經由 Referer、瀏覽器歷史與下游 log 外洩，所以拿到就立刻從網址拿掉。
      Location: incomingUrl.pathname + (incomingUrl.search || ''),
    });
    res.end();
    return null;
  }

  if (!authFailLimiter(auth.clientIp(req))) {
    res.writeHead(429, { 'content-type': 'text/plain; charset=utf-8' });
    res.end('Too many attempts. Try again later.');
    return null;
  }

  // 迴圈斷路器：沒有 token 或 token 過期時，彈回中央服務重簽是對的；但如果「帶了
  // token 卻仍驗不過」，再簽一張同樣會失敗，瀏覽器就會無限彈跳。跨服務部署的中間
  // 窗口正好會產生這種狀況，所以寧可大聲失敗一次。
  if (incomingToken && auth.parseCookies(req.headers.cookie)[HANDOFF_TRY_COOKIE]) {
    res.writeHead(503, {
      'content-type': 'text/html; charset=utf-8',
      'Set-Cookie': auth.clearCookieHeader(HANDOFF_TRY_COOKIE, { secure: IS_PROD }),
    });
    res.end('<h1>登入交接失敗</h1><p>中央登入服務簽發的憑證無法在此服務驗證。'
      + '若剛完成部署，請稍候一分鐘再重試。</p>'
      + '<p><a href="' + CENTRAL_AUTH_ORIGIN + '">回首頁</a></p>');
    return null;
  }

  incomingUrl.searchParams.delete('auth');
  const returnTo = selfOrigin(req) + incomingUrl.pathname + (incomingUrl.search || '');
  res.writeHead(302, {
    'Set-Cookie': auth.cookieHeader(HANDOFF_TRY_COOKIE, '1',
      { maxAgeSec: 120, httpOnly: true, secure: IS_PROD }),
    Location: CENTRAL_AUTH_ORIGIN + '/auth/handoff?return_to=' + encodeURIComponent(returnTo),
  });
  res.end();
  return null;
}

const server = http.createServer((req, res) => {
  let pathname = '/';
  try { pathname = decodeURIComponent(new URL(req.url, 'http://x').pathname); }
  catch (e) { res.writeHead(400).end('bad request'); return; }

  // Railway 健康檢查要能在未登入下通過，否則部署永遠不會轉為 healthy。
  if (pathname === '/healthz') {
    res.writeHead(200, { 'content-type': 'text/plain' }).end('ok');
    return;
  }

  if (pathname === '/auth/logout') {
    res.writeHead(302, {
      'Set-Cookie': [
        auth.clearCookieHeader(SESSION_COOKIE, { secure: IS_PROD }),
        auth.clearCookieHeader(HANDOFF_TRY_COOKIE, { secure: IS_PROD }),
      ],
      Location: CENTRAL_AUTH_ORIGIN + '/',
    });
    res.end();
    return;
  }

  // 預設全擋，登入後才往下走 —— 之後新增的路由預設就是受保護的。
  const email = gate(req, res);
  if (email === null) return;

  if (!pathname.startsWith('/api/')) {
    if (req.method !== 'GET' && req.method !== 'HEAD') {
      res.writeHead(405).end('method not allowed');
      return;
    }
    serveStatic(req, res, pathname);
    return;
  }

  try {
    if (pathname === '/api/health' && req.method === 'GET') {
      return sendJSON(res, 200, Object.assign({ ok: true, email: email }, store.summary(email)));
    }
    if (pathname === '/api/wrong' && req.method === 'GET') {
      const lim = new URL(req.url, 'http://x').searchParams.get('limit');
      return sendJSON(res, 200, { items: store.wrongList(email, lim) });
    }
    if (pathname === '/api/stats' && req.method === 'GET') {
      return sendJSON(res, 200, { items: store.tagStats(email) });
    }
    if (pathname === '/api/history' && req.method === 'GET') {
      return sendJSON(res, 200, { items: store.history(email) });
    }
    if (pathname === '/api/attempt' && req.method === 'POST') {
      return readBody(req, (err, body) => {
        if (err || !body || !body.paper) {
          return sendJSON(res, 400, { error: 'bad payload' });
        }
        try { sendJSON(res, 200, store.saveAttempt(email, body)); }
        catch (e) {
          console.error('[api] saveAttempt 失敗:', e && e.message);
          sendJSON(res, 500, { error: 'save failed' });
        }
      });
    }
    // 筆記練習題庫抽題：GET 查這個科目已經抽過哪些題（優先抽沒抽過的用），
    // POST 記一批新抽到的題目——只要抽到就算，不管有沒有交卷。
    if (pathname === '/api/notes-drawn' && req.method === 'GET') {
      const subject = new URL(req.url, 'http://x').searchParams.get('subject') || '';
      if (!subject) return sendJSON(res, 400, { error: 'subject required' });
      return sendJSON(res, 200, { qids: store.drawnQids(email, subject) });
    }
    if (pathname === '/api/notes-draw' && req.method === 'POST') {
      return readBody(req, (err, body) => {
        if (err || !body || !body.subject || !Array.isArray(body.qids)) {
          return sendJSON(res, 400, { error: 'bad payload' });
        }
        try { sendJSON(res, 200, store.recordDraws(email, body.subject, body.qids, !!body.reset)); }
        catch (e) {
          console.error('[api] recordDraws 失敗:', e && e.message);
          sendJSON(res, 500, { error: 'save failed' });
        }
      });
    }
    return sendJSON(res, 404, { error: 'no such endpoint' });
  } catch (e) {
    console.error('[api] 未預期錯誤:', e && e.stack);
    return sendJSON(res, 500, { error: 'internal' });
  }
});

server.listen(PORT, () => {
  console.log('gaoye-mock-exam listening on :' + PORT + '  db=' + store.DB_PATH
    + (AUTH_REQUIRED
        ? '  (需登入，中央認證 ' + CENTRAL_AUTH_ORIGIN + ')'
        : '  (免登入；設 REQUIRE_LOGIN=1 可恢復閘門)'));
});
