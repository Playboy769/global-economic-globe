// OutsideFramework — central login service for the cross-domain auth system.
// Zero npm dependencies — only Node's built-in http/https/fs/path/crypto, matching the
// house style already used by globe-invest/server.js. This is the ONLY service that talks
// to Google; globe-invest/structural_holes/article_db just verify tokens this service mints
// (see auth.js — copied identically into every service, auth.py for the Python ones).
//
// Serves the portfolio homepage in two variants (guest / full) built once at startup from
// the same source HTML, using HTML-comment markers as the split points:
//   <!--GATE-->...<!--/GATE-->             gated Works content — stripped for guests
//   <!--GUEST_ONLY-->...<!--/GUEST_ONLY--> guest-only content (login hint) — stripped when logged in
//   <!--NAV_AUTH-->                        replaced with a Login or Logout link
//   <!--AUTH_TOKENS_SCRIPT-->              (full variant only) replaced per-request with a
//                                          freshly signed handoff token per downstream origin

const http = require('http');
const https = require('https');
const fs = require('fs');
const path = require('path');
const auth = require('./auth');

const PORT = process.env.PORT || 8080;
// In the Docker image, index.html lives at /app (see Dockerfile). Locally there's no /app,
// so fall back to the dev-source copy relative to this file — lets `node server.js` run
// straight from a repo checkout with no extra configuration.
const APP_DIR = fs.existsSync('/app') ? '/app' : path.join(__dirname, 'app', 'OutsideFramework');

const GOOGLE_CLIENT_ID = process.env.GOOGLE_CLIENT_ID || '';
const GOOGLE_CLIENT_SECRET = process.env.GOOGLE_CLIENT_SECRET || '';
const AUTHORIZED_EMAIL = (process.env.AUTHORIZED_EMAIL || '').toLowerCase();
const SECRET = process.env.AUTH_SIGNING_SECRET || '';

if (!SECRET) console.error('WARNING: AUTH_SIGNING_SECRET is not set — auth will not work.');
if (!AUTHORIZED_EMAIL) console.error('WARNING: AUTHORIZED_EMAIL is not set — no one will be able to log in.');
if (!GOOGLE_CLIENT_ID || !GOOGLE_CLIENT_SECRET) console.error('WARNING: GOOGLE_CLIENT_ID/GOOGLE_CLIENT_SECRET not set.');

const COOKIE_NAME = 'ofw_sid';
const SESSION_TTL_SEC = 12 * 60 * 60; // 12h local session
const HANDOFF_TTL_SEC = 5 * 60;       // 5min cross-service handoff token
const STATE_TTL_SEC = 10 * 60;        // 10min to complete the Google consent screen

// The three downstream services that need a handoff token minted for them whenever an
// already-logged-in visitor is served the full homepage (so the first cross-domain click
// already carries proof of login — see openGatedWork() in index.html).
const GATED_ORIGINS = [
  'https://globe-invest.up.railway.app',
  'https://structural-holes-production.up.railway.app',
  'https://articlebase.up.railway.app',
];

const MIME = {
  '.html': 'text/html; charset=utf-8', '.js': 'application/javascript', '.css': 'text/css',
  '.json': 'application/json', '.png': 'image/png', '.jpg': 'image/jpeg', '.jpeg': 'image/jpeg',
  '.svg': 'image/svg+xml', '.ico': 'image/x-icon', '.webp': 'image/webp', '.gif': 'image/gif',
  '.woff': 'font/woff', '.woff2': 'font/woff2',
};

function nowSec() {
  return Math.floor(Date.now() / 1000);
}

function stripBlock(html, tag) {
  const re = new RegExp('<!--' + tag + '-->[\\s\\S]*?<!--\\/' + tag + '-->', 'g');
  return html.replace(re, '');
}

function unwrapBlock(html, tag) {
  return html.split('<!--' + tag + '-->').join('').split('<!--/' + tag + '-->').join('');
}

const RAW_HTML = fs.readFileSync(path.join(APP_DIR, 'index.html'), 'utf8');

let GUEST_HTML = stripBlock(RAW_HTML, 'GATE');
GUEST_HTML = unwrapBlock(GUEST_HTML, 'GUEST_ONLY');
GUEST_HTML = GUEST_HTML.replace('<!--NAV_AUTH-->', '<a class="nav-btn nav-auth-btn" href="/auth/google">Login</a>');

let FULL_TEMPLATE = unwrapBlock(RAW_HTML, 'GATE');
FULL_TEMPLATE = stripBlock(FULL_TEMPLATE, 'GUEST_ONLY');
FULL_TEMPLATE = FULL_TEMPLATE.replace('<!--NAV_AUTH-->', '<a class="nav-btn nav-auth-btn" href="/auth/logout">Logout</a>');
// <!--AUTH_TOKENS_SCRIPT--> is left in place here — substituted per-request in renderFullHtml()
// because handoff tokens are short-lived and must be freshly signed on every load.

function renderFullHtml(email) {
  const tokens = {};
  const exp = nowSec() + HANDOFF_TTL_SEC;
  for (const origin of GATED_ORIGINS) {
    tokens[origin] = auth.signToken({ email, aud: origin, exp }, SECRET);
  }
  const script = 'window.__AUTH_TOKENS__=' + JSON.stringify(tokens) + ';';
  return FULL_TEMPLATE.replace('<!--AUTH_TOKENS_SCRIPT-->', script);
}

function selfOrigin(req) {
  const host = req.headers.host || 'ofw.up.railway.app';
  const xfProto = req.headers['x-forwarded-proto'];
  const isLocal = host.startsWith('localhost') || host.startsWith('127.0.0.1');
  const proto = xfProto || (isLocal ? 'http' : 'https');
  return proto + '://' + host;
}

function getSessionEmail(req) {
  const cookies = auth.parseCookies(req.headers.cookie);
  const payload = auth.verifyToken(cookies[COOKIE_NAME], SECRET);
  if (!payload || !payload.email) return null;
  if (payload.email.toLowerCase() !== AUTHORIZED_EMAIL) return null;
  return payload.email;
}

function decodeIdTokenUnsafe(idToken) {
  // Never called with a client-supplied id_token — this one came back from our own
  // server-to-server HTTPS call to Google's token endpoint, so we trust it without an
  // extra signature check (Google already authenticated the request that produced it).
  // We still sanity-check aud/iss/exp as cheap defense in depth.
  try {
    const parts = String(idToken).split('.');
    if (parts.length !== 3) return null;
    return JSON.parse(auth.b64urlDecode(parts[1]).toString('utf8'));
  } catch {
    return null;
  }
}

function exchangeCodeForToken(code, redirectUri) {
  return new Promise((resolve, reject) => {
    const body = new URLSearchParams({
      code,
      client_id: GOOGLE_CLIENT_ID,
      client_secret: GOOGLE_CLIENT_SECRET,
      redirect_uri: redirectUri,
      grant_type: 'authorization_code',
    }).toString();
    const reqOut = https.request(
      'https://oauth2.googleapis.com/token',
      {
        method: 'POST',
        headers: { 'Content-Type': 'application/x-www-form-urlencoded', 'Content-Length': Buffer.byteLength(body) },
      },
      (resIn) => {
        let data = '';
        resIn.on('data', (c) => (data += c));
        resIn.on('end', () => {
          if (resIn.statusCode !== 200) {
            reject(new Error('token endpoint ' + resIn.statusCode + ': ' + data));
            return;
          }
          try {
            resolve(JSON.parse(data));
          } catch (e) {
            reject(e);
          }
        });
      }
    );
    reqOut.on('error', reject);
    reqOut.write(body);
    reqOut.end();
  });
}

function serveStatic(res, filePath) {
  const resolved = path.resolve(filePath);
  if (!resolved.startsWith(path.resolve(APP_DIR))) {
    res.writeHead(403);
    res.end('Forbidden');
    return;
  }
  try {
    const content = fs.readFileSync(resolved);
    const ext = path.extname(resolved);
    res.writeHead(200, { 'Content-Type': MIME[ext] || 'application/octet-stream' });
    res.end(content);
  } catch {
    res.writeHead(404);
    res.end('Not found');
  }
}

// Simple IP-based rate limits for the auth endpoints — raises the cost of naive automated
// abuse. See auth.js for why this is deliberately not a distributed limiter.
const callbackLimiter = auth.makeRateLimiter({ windowMs: 10 * 60 * 1000, max: 20 });
const handoffLimiter = auth.makeRateLimiter({ windowMs: 10 * 60 * 1000, max: 60 });

const server = http.createServer(async (req, res) => {
  const parsedUrl = new URL(req.url, selfOrigin(req));
  const url = parsedUrl.pathname;
  const method = req.method;
  const isProd = selfOrigin(req).startsWith('https://');

  try {
    if (url === '/auth/google' && method === 'GET') {
      let returnTo = parsedUrl.searchParams.get('return_to');
      if (returnTo && !auth.isAllowedReturnTo(returnTo)) returnTo = null;
      const state = auth.signToken({ n: auth.randomToken(8), returnTo: returnTo || null, exp: nowSec() + STATE_TTL_SEC }, SECRET);
      const redirectUri = selfOrigin(req) + '/auth/google/callback';
      const params = new URLSearchParams({
        client_id: GOOGLE_CLIENT_ID,
        redirect_uri: redirectUri,
        response_type: 'code',
        scope: 'openid email',
        state,
        prompt: 'select_account',
      });
      res.writeHead(302, { Location: 'https://accounts.google.com/o/oauth2/v2/auth?' + params.toString() });
      res.end();
      return;
    }

    if (url === '/auth/google/callback' && method === 'GET') {
      const ip = auth.clientIp(req);
      if (!callbackLimiter(ip)) {
        res.writeHead(429, { 'Content-Type': 'text/plain; charset=utf-8' });
        res.end('Too many attempts. Try again later.');
        return;
      }
      const code = parsedUrl.searchParams.get('code');
      const statePayload = auth.verifyToken(parsedUrl.searchParams.get('state'), SECRET);
      if (!code || !statePayload) {
        res.writeHead(400, { 'Content-Type': 'text/plain; charset=utf-8' });
        res.end('Invalid or expired login request. Please try logging in again.');
        return;
      }
      const redirectUri = selfOrigin(req) + '/auth/google/callback';
      let tokenResp;
      try {
        tokenResp = await exchangeCodeForToken(code, redirectUri);
      } catch {
        res.writeHead(502, { 'Content-Type': 'text/plain; charset=utf-8' });
        res.end('Google login failed. Please try again.');
        return;
      }
      const idPayload = decodeIdTokenUnsafe(tokenResp.id_token);
      const validIss = idPayload && (idPayload.iss === 'accounts.google.com' || idPayload.iss === 'https://accounts.google.com');
      if (!idPayload || !validIss || idPayload.aud !== GOOGLE_CLIENT_ID || !idPayload.email || !idPayload.email_verified) {
        res.writeHead(403, { 'Content-Type': 'text/html; charset=utf-8' });
        res.end('<h1>登入失敗</h1><p>Google 帳號驗證未通過。</p><p><a href="/">回首頁</a></p>');
        return;
      }
      const email = String(idPayload.email).toLowerCase();
      if (email !== AUTHORIZED_EMAIL) {
        res.writeHead(403, { 'Content-Type': 'text/html; charset=utf-8' });
        res.end('<h1>此帳號未被授權</h1><p><a href="/">回首頁</a></p>');
        return;
      }
      const sessionToken = auth.signToken({ email, aud: selfOrigin(req), exp: nowSec() + SESSION_TTL_SEC }, SECRET);
      const headers = {
        'Set-Cookie': auth.cookieHeader(COOKIE_NAME, sessionToken, { maxAgeSec: SESSION_TTL_SEC, secure: isProd }),
      };
      const returnTo = statePayload.returnTo;
      if (returnTo && auth.isAllowedReturnTo(returnTo)) {
        const origin = auth.originOf(returnTo);
        const handoffToken = auth.signToken({ email, aud: origin, exp: nowSec() + HANDOFF_TTL_SEC }, SECRET);
        const sep = returnTo.includes('?') ? '&' : '?';
        headers.Location = returnTo + sep + 'auth=' + encodeURIComponent(handoffToken);
      } else {
        headers.Location = '/';
      }
      res.writeHead(302, headers);
      res.end();
      return;
    }

    if (url === '/auth/handoff' && method === 'GET') {
      const ip = auth.clientIp(req);
      if (!handoffLimiter(ip)) {
        res.writeHead(429, { 'Content-Type': 'text/plain; charset=utf-8' });
        res.end('Too many attempts. Try again later.');
        return;
      }
      const returnTo = parsedUrl.searchParams.get('return_to');
      if (!returnTo || !auth.isAllowedReturnTo(returnTo)) {
        res.writeHead(400, { 'Content-Type': 'text/plain; charset=utf-8' });
        res.end('Invalid return_to.');
        return;
      }
      const email = getSessionEmail(req);
      if (email) {
        const origin = auth.originOf(returnTo);
        const token = auth.signToken({ email, aud: origin, exp: nowSec() + HANDOFF_TTL_SEC }, SECRET);
        const sep = returnTo.includes('?') ? '&' : '?';
        res.writeHead(302, { Location: returnTo + sep + 'auth=' + encodeURIComponent(token) });
        res.end();
        return;
      }
      res.writeHead(302, { Location: '/auth/google?return_to=' + encodeURIComponent(returnTo) });
      res.end();
      return;
    }

    if (url === '/auth/logout' && method === 'GET') {
      res.writeHead(302, {
        'Set-Cookie': auth.clearCookieHeader(COOKIE_NAME, { secure: isProd }),
        Location: '/',
      });
      res.end();
      return;
    }

    if (url === '/api/auth/status' && method === 'GET') {
      const email = getSessionEmail(req);
      res.writeHead(200, { 'Content-Type': 'application/json' });
      res.end(JSON.stringify({ loggedIn: !!email }));
      return;
    }

    if (url.startsWith('/assets/') && method === 'GET') {
      serveStatic(res, path.join(APP_DIR, url));
      return;
    }

    if (method === 'GET') {
      const email = getSessionEmail(req);
      const html = email ? renderFullHtml(email) : GUEST_HTML;
      res.writeHead(200, { 'Content-Type': 'text/html; charset=utf-8', 'Cache-Control': 'no-cache' });
      res.end(html);
      return;
    }

    res.writeHead(404);
    res.end('Not found');
  } catch (e) {
    console.error(e);
    res.writeHead(500, { 'Content-Type': 'text/plain; charset=utf-8' });
    res.end('Internal error');
  }
});

server.listen(PORT, () => console.log('OutsideFramework listening on port ' + PORT));
