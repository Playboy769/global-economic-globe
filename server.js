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
//
// SECURITY NOTE — request headers are not configuration. Host, X-Forwarded-Proto and
// X-Forwarded-For are all attacker-controlled on any request. Nothing that decides a
// cookie flag, an origin, or a rate-limit bucket may be derived from them without
// validation; see IS_PROD, selfOrigin() and auth.clientIp().

const http = require('http');
const https = require('https');
const fs = require('fs');
const path = require('path');
const auth = require('./auth');
const analytics = require('./analytics');

// process.argv[2] lets .claude/launch.json pin the local dev port (8125) without needing
// to inject an env var — Railway sets PORT itself in production, which takes priority.
const PORT = process.env.PORT || process.argv[2] || 8080;
// In the Docker image, index.html lives at /app (see Dockerfile). Locally there's no /app,
// so fall back to the dev-source copy relative to this file — lets `node server.js` run
// straight from a repo checkout with no extra configuration.
const APP_DIR = fs.existsSync('/app') ? '/app' : path.join(__dirname, 'app', 'OutsideFramework');
const APP_ROOT = path.resolve(APP_DIR);

const GOOGLE_CLIENT_ID = process.env.GOOGLE_CLIENT_ID || '';
const GOOGLE_CLIENT_SECRET = process.env.GOOGLE_CLIENT_SECRET || '';
const AUTHORIZED_EMAIL = (process.env.AUTHORIZED_EMAIL || '').toLowerCase();
const SECRET = process.env.AUTH_SIGNING_SECRET || '';

if (!SECRET) console.error('WARNING: AUTH_SIGNING_SECRET is not set — auth will not work.');
if (!AUTHORIZED_EMAIL) console.error('WARNING: AUTHORIZED_EMAIL is not set — no one will be able to log in.');
if (!GOOGLE_CLIENT_ID || !GOOGLE_CLIENT_SECRET) console.error('WARNING: GOOGLE_CLIENT_ID/GOOGLE_CLIENT_SECRET not set.');

// Whether to mark cookies Secure, and which origin to fall back to. Both come from the
// environment, never from a request header: X-Forwarded-Proto is client-settable, and
// trusting it let anyone strip the Secure flag off an issued cookie just by sending
// "X-Forwarded-Proto: http".
const IS_PROD =
  process.env.NODE_ENV === 'production' ||
  !!process.env.RAILWAY_PROJECT_ID ||
  !!process.env.RAILWAY_ENVIRONMENT_NAME;
const CANONICAL_ORIGIN = process.env.PUBLIC_ORIGIN || 'https://ofw.up.railway.app';

// Hosts this service will echo back into redirect URIs and token audiences. Anything
// else falls back to CANONICAL_ORIGIN rather than reflecting whatever Host was sent.
const ALLOWED_HOSTS = new Set(
  [auth.originOf(CANONICAL_ORIGIN), 'https://ofw.up.railway.app']
    .filter(Boolean)
    .map((o) => new URL(o).host.toLowerCase())
);
const LOCAL_HOST_RE = /^(localhost|127\.0\.0\.1)(:\d+)?$/;

const COOKIE_NAME = 'ofw_sid';
const STATE_COOKIE = 'ofw_oauth_state';
const SESSION_TTL_SEC = 12 * 60 * 60; // 12h local session
const HANDOFF_TTL_SEC = 5 * 60;       // 5min cross-service handoff token
const STATE_TTL_SEC = 10 * 60;        // 10min to complete the Google consent screen

// Analytics visitor cookie. Deliberately just an opaque random id with no signature and
// no personal data in it: it only has to be stable enough to tell "same browser came
// back" apart from "two different browsers", and a forged value can at worst skew the
// owner's own private stats. A year is long enough to measure return visits.
const VISITOR_COOKIE = 'ofw_vid';
const VISITOR_TTL_SEC = 365 * 24 * 60 * 60;
const VISITOR_ID_RE = /^[A-Za-z0-9_-]{8,64}$/;

// The three downstream services that need a handoff token minted for them whenever an
// already-logged-in visitor is served the full homepage (so the first cross-domain click
// already carries proof of login — see openGatedWork() in index.html).
const GATED_ORIGINS = [
  'https://globe-invest.up.railway.app',
  'https://structural-holes-production.up.railway.app',
  'https://articlebase.up.railway.app',
];

// Local dev ports that the in-page preview iframe points at when running from a checkout.
const DEV_FRAME_ORIGINS = [
  'http://localhost:8124', 'http://localhost:8126',
  'http://localhost:8128', 'http://localhost:8132',
];

const MIME = {
  '.html': 'text/html; charset=utf-8', '.js': 'application/javascript', '.css': 'text/css',
  '.json': 'application/json', '.png': 'image/png', '.jpg': 'image/jpeg', '.jpeg': 'image/jpeg',
  '.svg': 'image/svg+xml', '.ico': 'image/x-icon', '.webp': 'image/webp', '.gif': 'image/gif',
  '.woff': 'font/woff', '.woff2': 'font/woff2',
};

// ── Security response headers ────────────────────────────────────────────────────────
// script-src has to keep 'unsafe-inline': index.html is built around inline <script>
// blocks and inline onclick= handlers, and a nonce cannot authorize an attribute event
// handler. What this policy still buys is real — it pins *which external origins* may
// supply script, style, font, image and frame content, and frame-ancestors kills the
// clickjacking path to /admin. Tightening script-src further means removing every inline
// handler from index.html first.
const CSP = [
  "default-src 'self'",
  "script-src 'self' 'unsafe-inline' https://cdn.jsdelivr.net",
  "style-src 'self' 'unsafe-inline' https://fonts.googleapis.com https://fonts.cdnfonts.com",
  "font-src 'self' https://fonts.gstatic.com https://fonts.cdnfonts.com data:",
  "img-src 'self' data: blob:",
  "connect-src 'self'",
  'frame-src ' + GATED_ORIGINS.concat(IS_PROD ? [] : DEV_FRAME_ORIGINS).join(' '),
  "frame-ancestors 'none'",
  "base-uri 'none'",
  "object-src 'none'",
  "form-action 'self'",
].join('; ');

function securityHeaders(extra) {
  const h = Object.assign(
    {
      'Content-Security-Policy': CSP,
      'X-Content-Type-Options': 'nosniff',
      'X-Frame-Options': 'DENY',
      // Handoff tokens ride in ?auth=... query strings on the destination URL. Without a
      // referrer policy, any subresource that page loads sends the whole URL — token and
      // all — to a third party in the Referer header.
      'Referrer-Policy': 'strict-origin-when-cross-origin',
      'Permissions-Policy': 'geolocation=(), microphone=(), camera=(), payment=(), usb=()',
      'Cross-Origin-Opener-Policy': 'same-origin',
    },
    extra || {}
  );
  if (IS_PROD) h['Strict-Transport-Security'] = 'max-age=31536000; includeSubDomains';
  return h;
}

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
const TOKENS_MARKER = '<!--AUTH_TOKENS_SCRIPT-->';

let GUEST_HTML = stripBlock(RAW_HTML, 'GATE');
GUEST_HTML = unwrapBlock(GUEST_HTML, 'GUEST_ONLY');
GUEST_HTML = GUEST_HTML.replace('<!--NAV_AUTH-->', '<a class="nav-btn nav-auth-btn" href="/auth/google">Login</a>');
// Guests get no handoff tokens, so the marker is removed rather than left in the served
// HTML. It sat inside a <script> block where JS parses "<!--" as a line comment, so it
// was inert — but shipping a template marker to visitors is how the next person concludes
// the substitution runs for guests too.
GUEST_HTML = GUEST_HTML.split(TOKENS_MARKER).join('');

let FULL_TEMPLATE = unwrapBlock(RAW_HTML, 'GATE');
FULL_TEMPLATE = stripBlock(FULL_TEMPLATE, 'GUEST_ONLY');
FULL_TEMPLATE = FULL_TEMPLATE.replace(
  '<!--NAV_AUTH-->',
  '<a class="nav-btn nav-auth-btn" href="/admin">Admin</a>' +
    '<a class="nav-btn nav-auth-btn" href="/auth/logout">Logout</a>'
);
// <!--AUTH_TOKENS_SCRIPT--> is left in place here — substituted per-request in renderFullHtml()
// because handoff tokens are short-lived and must be freshly signed on every load.
if (!FULL_TEMPLATE.includes(TOKENS_MARKER)) {
  console.error(
    'WARNING: ' + TOKENS_MARKER + ' is missing from index.html — logged-in visitors will ' +
    'get no handoff tokens and every gated Works card will bounce them back through login.'
  );
}

function renderFullHtml(email, origin) {
  const tokens = {};
  const exp = nowSec() + HANDOFF_TTL_SEC;
  for (const target of GATED_ORIGINS) {
    // signFor derives the key from `aud`, so each service's token is signed under a key
    // only that service's audience produces.
    tokens[target] = auth.signFor({ email, aud: target, typ: auth.TYP_HANDOFF, exp }, SECRET);
  }
  const script = 'window.__AUTH_TOKENS__=' + JSON.stringify(tokens) + ';';
  // Function replacer: a plain string replacement would interpret $& / $1 inside the
  // substitution text. Token charset can't produce those today, which is exactly the kind
  // of assumption that breaks silently when the payload shape changes.
  return FULL_TEMPLATE.replace(TOKENS_MARKER, () => script);
}

// Derives this service's own origin WITHOUT trusting the request. An unrecognized Host
// falls back to the canonical origin instead of being echoed back, so an attacker cannot
// steer the OAuth redirect_uri or a token audience by sending "Host: evil.example".
function selfOrigin(req) {
  const host = String(req.headers.host || '').toLowerCase();
  if (!IS_PROD && LOCAL_HOST_RE.test(host)) return 'http://' + host;
  if (ALLOWED_HOSTS.has(host)) return 'https://' + host;
  return CANONICAL_ORIGIN;
}

function getSessionEmail(req) {
  const cookies = auth.parseCookies(req.headers.cookie);
  // Audience AND purpose are both asserted, and the verification key is derived from the
  // audience — so a handoff token minted for a different service (which travels through
  // URLs, Referer headers and downstream logs) does not even produce a matching signature
  // here, let alone pass the aud check.
  const payload = auth.verifyFor(cookies[COOKIE_NAME], SECRET, {
    typ: auth.TYP_SESSION,
    aud: selfOrigin(req),
  });
  if (!payload || !payload.email) return null;
  if (String(payload.email).toLowerCase() !== AUTHORIZED_EMAIL) return null;
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
  // path.relative gives a proper containment test. A startsWith() prefix check treats
  // "/app-secrets" as living inside "/app" because it only compares characters.
  const rel = path.relative(APP_ROOT, resolved);
  if (rel === '' || rel.startsWith('..') || path.isAbsolute(rel)) {
    res.writeHead(403, securityHeaders());
    res.end('Forbidden');
    return;
  }
  try {
    const content = fs.readFileSync(resolved);
    const ext = path.extname(resolved);
    res.writeHead(
      200,
      securityHeaders({
        'Content-Type': MIME[ext] || 'application/octet-stream',
        'Cache-Control': 'public, max-age=3600',
      })
    );
    res.end(content);
  } catch {
    res.writeHead(404, securityHeaders());
    res.end('Not found');
  }
}

// Simple IP-based rate limits for the auth endpoints — raises the cost of naive automated
// abuse. See auth.js for why this is deliberately not a distributed limiter, and why the
// client IP is now taken from the last X-Forwarded-For hop rather than the first.
const callbackLimiter = auth.makeRateLimiter({ windowMs: 10 * 60 * 1000, max: 20 });
const handoffLimiter = auth.makeRateLimiter({ windowMs: 10 * 60 * 1000, max: 60 });
// /api/track is the one write endpoint open to anonymous visitors, so it carries its own
// (looser, but real) limit — a single reader legitimately fires one beacon per page switch.
const trackLimiter = auth.makeRateLimiter({ windowMs: 10 * 60 * 1000, max: 300 });
// Backstop across all clients. The per-IP limit alone bounds one abuser; this bounds the
// whole endpoint, because the analytics DB sits on a persistent volume with 400-day
// retention and unbounded anonymous writes are a disk-exhaustion path, not just noise.
const trackGlobalLimiter = auth.makeRateLimiter({ windowMs: 10 * 60 * 1000, max: 5000 });

// Railway exposes no visitor-country header; x-railway-edge names the edge PoP that served
// the request (e.g. "hkg1"), which is a coarse regional hint, not a country. The dashboard
// labels it as such rather than dressing it up as geolocation.
function visitorRegion(req) {
  return (
    req.headers['cf-ipcountry'] ||
    req.headers['x-vercel-ip-country'] ||
    req.headers['x-railway-edge'] ||
    ''
  );
}

function visitorIdFrom(req) {
  const raw = auth.parseCookies(req.headers.cookie)[VISITOR_COOKIE];
  return VISITOR_ID_RE.test(raw || '') ? raw : null;
}

function visitorCookieHeader(id) {
  // HttpOnly on purpose: nothing client-side reads this, so keeping it HttpOnly means page
  // scripts (and anything injected into them) can't see or spoof it.
  return auth.cookieHeader(VISITOR_COOKIE, id, {
    maxAgeSec: VISITOR_TTL_SEC,
    httpOnly: true,
    secure: IS_PROD,
  });
}

function trackEvent(req, { visitorId, type, path: evPath, label, isOwner }) {
  const ua = req.headers['user-agent'] || '';
  if (analytics.isBot(ua)) return;
  analytics.record({
    visitorId,
    type,
    path: evPath,
    label,
    referrer: analytics.parseReferrer(req.headers.referer || '', req.headers.host),
    device: analytics.parseDevice(ua),
    browser: analytics.parseBrowser(ua),
    region: visitorRegion(req),
    isOwner,
  });
}

// The GET catch-all below serves the SPA shell for *every* unmatched path, so without this
// gate /favicon.ico (and any other asset or probe URL) is counted as a page view and shows
// up in "top pages". Extension check catches assets; the Accept check catches the rest.
function looksLikeAsset(urlPath) {
  return /\.[a-z0-9]{1,8}$/i.test(urlPath);
}

function isPageRequest(req, urlPath) {
  if (looksLikeAsset(urlPath)) return false;
  const accept = req.headers.accept || '';
  return accept === '' || accept === '*/*' || accept.includes('text/html');
}

// Distinct from `null` (unparsable) so the caller can answer 413 rather than 400.
const BODY_TOO_LARGE = Symbol('body_too_large');

// Counts BYTES, not JS string length. `data += chunk` decodes to UTF-16, so a body of
// multibyte characters could carry well past maxBytes before the old length check tripped.
//
// Two tiers on purpose. Past the soft cap the buffer is dropped (memory stays bounded)
// but the request is still allowed to finish, so the client gets a real 413 back. Only
// past the hard cap is the socket destroyed — killing it at the soft cap means the caller
// writes a response into a closed socket and the client just sees a connection reset.
function readJsonBody(req, maxBytes) {
  const hardLimit = Math.max(maxBytes * 32, 262144);
  return new Promise((resolve) => {
    const chunks = [];
    let size = 0;
    let tooBig = false;
    let done = false;
    const finish = (v) => {
      if (done) return;
      done = true;
      resolve(v);
    };
    req.on('data', (c) => {
      size += c.length;
      if (size > maxBytes) {
        tooBig = true;
        chunks.length = 0;
        if (size > hardLimit) {
          finish(BODY_TOO_LARGE);
          req.destroy();
        }
        return;
      }
      chunks.push(c);
    });
    req.on('end', () => {
      if (tooBig) return finish(BODY_TOO_LARGE);
      try {
        finish(JSON.parse(Buffer.concat(chunks).toString('utf8')));
      } catch {
        finish(null);
      }
    });
    req.on('error', () => finish(null));
    req.on('aborted', () => finish(null));
  });
}

// Blocks the cross-site half of CSRF on state-changing requests while still allowing a
// user typing the URL or following a bookmark (Sec-Fetch-Site: none) and same-origin page
// requests. Browsers that don't send the header at all fall through to the Origin check.
function isSameOriginRequest(req) {
  const site = req.headers['sec-fetch-site'];
  if (site) return site === 'same-origin' || site === 'none';
  const origin = req.headers.origin;
  if (origin) return origin === selfOrigin(req);
  return true;
}

const ADMIN_HTML_PATH = path.join(APP_DIR, 'admin.html');

const server = http.createServer(async (req, res) => {
  const parsedUrl = new URL(req.url, selfOrigin(req));
  const url = parsedUrl.pathname;
  const method = req.method;

  try {
    if (url === '/auth/google' && method === 'GET') {
      let returnTo = parsedUrl.searchParams.get('return_to');
      if (returnTo && !auth.isAllowedReturnTo(returnTo)) returnTo = null;
      // The nonce is signed into the state AND stored in a cookie, so completing the flow
      // requires possession of both. Signing alone made state a bearer value: any valid
      // state worked in any browser, which is the shape of an OAuth login-CSRF.
      const nonce = auth.randomToken(16);
      const state = auth.signFor(
        { n: nonce, typ: auth.TYP_STATE, returnTo: returnTo || null, aud: selfOrigin(req), exp: nowSec() + STATE_TTL_SEC },
        SECRET
      );
      const redirectUri = selfOrigin(req) + '/auth/google/callback';
      const params = new URLSearchParams({
        client_id: GOOGLE_CLIENT_ID,
        redirect_uri: redirectUri,
        response_type: 'code',
        scope: 'openid email',
        state,
        prompt: 'select_account',
      });
      res.writeHead(
        302,
        securityHeaders({
          Location: 'https://accounts.google.com/o/oauth2/v2/auth?' + params.toString(),
          'Set-Cookie': auth.cookieHeader(STATE_COOKIE, nonce, {
            maxAgeSec: STATE_TTL_SEC,
            httpOnly: true,
            secure: IS_PROD,
          }),
        })
      );
      res.end();
      return;
    }

    if (url === '/auth/google/callback' && method === 'GET') {
      const ip = auth.clientIp(req);
      if (!callbackLimiter(ip)) {
        res.writeHead(429, securityHeaders({ 'Content-Type': 'text/plain; charset=utf-8' }));
        res.end('Too many attempts. Try again later.');
        return;
      }
      const code = parsedUrl.searchParams.get('code');
      const statePayload = auth.verifyFor(parsedUrl.searchParams.get('state'), SECRET, {
        typ: auth.TYP_STATE,
        aud: selfOrigin(req),
      });
      const stateCookie = auth.parseCookies(req.headers.cookie)[STATE_COOKIE];
      const clearState = auth.clearCookieHeader(STATE_COOKIE, { secure: IS_PROD });
      if (!code || !statePayload || !auth.safeEqual(statePayload.n, stateCookie)) {
        res.writeHead(
          400,
          securityHeaders({ 'Content-Type': 'text/plain; charset=utf-8', 'Set-Cookie': clearState })
        );
        res.end('Invalid or expired login request. Please try logging in again.');
        return;
      }
      const redirectUri = selfOrigin(req) + '/auth/google/callback';
      let tokenResp;
      try {
        tokenResp = await exchangeCodeForToken(code, redirectUri);
      } catch {
        res.writeHead(
          502,
          securityHeaders({ 'Content-Type': 'text/plain; charset=utf-8', 'Set-Cookie': clearState })
        );
        res.end('Google login failed. Please try again.');
        return;
      }
      const idPayload = decodeIdTokenUnsafe(tokenResp.id_token);
      const validIss =
        idPayload && (idPayload.iss === 'accounts.google.com' || idPayload.iss === 'https://accounts.google.com');
      // exp is checked too — the comment above claimed it was, and it wasn't.
      const notExpired = idPayload && typeof idPayload.exp === 'number' && nowSec() < idPayload.exp;
      if (
        !idPayload || !validIss || !notExpired ||
        idPayload.aud !== GOOGLE_CLIENT_ID || !idPayload.email || !idPayload.email_verified
      ) {
        res.writeHead(
          403,
          securityHeaders({ 'Content-Type': 'text/html; charset=utf-8', 'Set-Cookie': clearState })
        );
        res.end('<h1>登入失敗</h1><p>Google 帳號驗證未通過。</p><p><a href="/">回首頁</a></p>');
        return;
      }
      const email = String(idPayload.email).toLowerCase();
      if (email !== AUTHORIZED_EMAIL) {
        res.writeHead(
          403,
          securityHeaders({ 'Content-Type': 'text/html; charset=utf-8', 'Set-Cookie': clearState })
        );
        res.end('<h1>此帳號未被授權</h1><p><a href="/">回首頁</a></p>');
        return;
      }
      const sessionToken = auth.signFor(
        { email, aud: selfOrigin(req), typ: auth.TYP_SESSION, exp: nowSec() + SESSION_TTL_SEC },
        SECRET
      );
      const setCookies = [
        auth.cookieHeader(COOKIE_NAME, sessionToken, { maxAgeSec: SESSION_TTL_SEC, secure: IS_PROD }),
        clearState,
      ];
      const headers = securityHeaders({ 'Set-Cookie': setCookies });
      const returnTo = statePayload.returnTo;
      if (returnTo && auth.isAllowedReturnTo(returnTo)) {
        const origin = auth.originOf(returnTo);
        const handoffToken = auth.signFor(
          { email, aud: origin, typ: auth.TYP_HANDOFF, exp: nowSec() + HANDOFF_TTL_SEC },
          SECRET
        );
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
        res.writeHead(429, securityHeaders({ 'Content-Type': 'text/plain; charset=utf-8' }));
        res.end('Too many attempts. Try again later.');
        return;
      }
      const returnTo = parsedUrl.searchParams.get('return_to');
      if (!returnTo || !auth.isAllowedReturnTo(returnTo)) {
        res.writeHead(400, securityHeaders({ 'Content-Type': 'text/plain; charset=utf-8' }));
        res.end('Invalid return_to.');
        return;
      }
      const email = getSessionEmail(req);
      if (email) {
        const origin = auth.originOf(returnTo);
        const token = auth.signFor(
          { email, aud: origin, typ: auth.TYP_HANDOFF, exp: nowSec() + HANDOFF_TTL_SEC },
          SECRET
        );
        const sep = returnTo.includes('?') ? '&' : '?';
        res.writeHead(302, securityHeaders({ Location: returnTo + sep + 'auth=' + encodeURIComponent(token) }));
        res.end();
        return;
      }
      res.writeHead(302, securityHeaders({ Location: '/auth/google?return_to=' + encodeURIComponent(returnTo) }));
      res.end();
      return;
    }

    if (url === '/auth/logout' && method === 'GET') {
      // Logout stays a GET so the nav link keeps working, but a cross-site request can no
      // longer trigger it.
      if (!isSameOriginRequest(req)) {
        res.writeHead(403, securityHeaders({ 'Content-Type': 'text/plain; charset=utf-8' }));
        res.end('Forbidden');
        return;
      }
      res.writeHead(
        302,
        securityHeaders({
          'Set-Cookie': auth.clearCookieHeader(COOKIE_NAME, { secure: IS_PROD }),
          Location: '/',
        })
      );
      res.end();
      return;
    }

    if (url === '/api/auth/status' && method === 'GET') {
      const email = getSessionEmail(req);
      res.writeHead(200, securityHeaders({ 'Content-Type': 'application/json', 'Cache-Control': 'no-store' }));
      res.end(JSON.stringify({ loggedIn: !!email }));
      return;
    }

    // Anonymous-writable: the homepage is a single-page app, so every page switch after
    // the initial load is invisible to the server. This beacon fills that gap (and records
    // Works-card opens). It writes only to the owner's private stats and reflects nothing
    // back, but it does write to a persistent volume — so it is fenced by a same-origin
    // check, a JSON content-type requirement (which forces a CORS preflight that a
    // cross-site forgery cannot satisfy), a per-IP limit and a global limit.
    if (url === '/api/track' && method === 'POST') {
      if (!isSameOriginRequest(req)) {
        res.writeHead(403, securityHeaders({ 'Content-Type': 'application/json' }));
        res.end(JSON.stringify({ error: 'forbidden' }));
        return;
      }
      const ctype = String(req.headers['content-type'] || '').split(';')[0].trim().toLowerCase();
      if (ctype !== 'application/json') {
        res.writeHead(415, securityHeaders({ 'Content-Type': 'application/json' }));
        res.end(JSON.stringify({ error: 'unsupported_media_type' }));
        return;
      }
      if (!trackLimiter(auth.clientIp(req)) || !trackGlobalLimiter('__all__')) {
        res.writeHead(429, securityHeaders({ 'Content-Type': 'application/json' }));
        res.end(JSON.stringify({ error: 'rate_limited' }));
        return;
      }
      // A body that was unparsable or over the cap records nothing — earlier this fell
      // through to {} and wrote a phantom pageview with an empty path.
      const body = await readJsonBody(req, 2048);
      if (body === BODY_TOO_LARGE) {
        res.writeHead(413, securityHeaders({ 'Content-Type': 'application/json' }));
        res.end(JSON.stringify({ error: 'payload_too_large' }));
        return;
      }
      if (!body || typeof body.path !== 'string' || !body.path) {
        res.writeHead(400, securityHeaders({ 'Content-Type': 'application/json' }));
        res.end(JSON.stringify({ error: 'bad_request' }));
        return;
      }
      const type = body.type === 'work_click' ? 'work_click' : 'pageview';
      let visitorId = visitorIdFrom(req);
      const headers = securityHeaders({ 'Content-Type': 'application/json' });
      if (!visitorId) {
        visitorId = auth.randomToken(16);
        headers['Set-Cookie'] = visitorCookieHeader(visitorId);
      }
      trackEvent(req, {
        visitorId,
        type,
        path: body.path,
        label: typeof body.label === 'string' ? body.label : '',
        isOwner: !!getSessionEmail(req),
      });
      res.writeHead(204, headers);
      res.end();
      return;
    }

    if (url === '/api/admin/stats' && method === 'GET') {
      if (!getSessionEmail(req)) {
        res.writeHead(403, securityHeaders({ 'Content-Type': 'application/json' }));
        res.end(JSON.stringify({ error: 'forbidden' }));
        return;
      }
      const days = Math.min(Math.max(parseInt(parsedUrl.searchParams.get('days'), 10) || 30, 1), 400);
      const includeOwner = parsedUrl.searchParams.get('includeOwner') === '1';
      res.writeHead(
        200,
        securityHeaders({ 'Content-Type': 'application/json', 'Cache-Control': 'no-store' })
      );
      res.end(JSON.stringify(analytics.stats({ days, includeOwner })));
      return;
    }

    if (url === '/admin' && method === 'GET') {
      if (!getSessionEmail(req)) {
        res.writeHead(
          302,
          securityHeaders({ Location: '/auth/google?return_to=' + encodeURIComponent(selfOrigin(req) + '/admin') })
        );
        res.end();
        return;
      }
      let adminHtml;
      try {
        adminHtml = fs.readFileSync(ADMIN_HTML_PATH, 'utf8');
      } catch {
        res.writeHead(404, securityHeaders({ 'Content-Type': 'text/plain; charset=utf-8' }));
        res.end('admin.html not found');
        return;
      }
      res.writeHead(
        200,
        securityHeaders({ 'Content-Type': 'text/html; charset=utf-8', 'Cache-Control': 'no-store' })
      );
      res.end(adminHtml);
      return;
    }

    if (url.startsWith('/assets/') && method === 'GET') {
      serveStatic(res, path.join(APP_DIR, url));
      return;
    }

    if (method === 'GET') {
      // An unmatched path that looks like a file is a 404, not the homepage. Returning the
      // SPA shell with 200 for /favicon.ico or /whatever.js tells crawlers and browsers
      // that every missing asset exists, and (when logged in) minted three fresh handoff
      // tokens for each such probe.
      if (looksLikeAsset(url)) {
        res.writeHead(404, securityHeaders({ 'Content-Type': 'text/plain; charset=utf-8' }));
        res.end('Not found');
        return;
      }
      const email = getSessionEmail(req);
      const html = email ? renderFullHtml(email, selfOrigin(req)) : GUEST_HTML;
      // The logged-in variant embeds live handoff tokens, so it must not be stored by any
      // cache. no-cache still permits storage (it only forces revalidation).
      const headers = securityHeaders({
        'Content-Type': 'text/html; charset=utf-8',
        'Cache-Control': email ? 'no-store, private' : 'no-cache',
      });
      const isPage = isPageRequest(req, url);
      let visitorId = visitorIdFrom(req);
      if (!visitorId && isPage) {
        visitorId = auth.randomToken(16);
        headers['Set-Cookie'] = visitorCookieHeader(visitorId);
      }
      // Only real page requests are counted, and only they mint a visitor id — otherwise a
      // favicon fetch racing the page load would burn a second id for the same reader.
      if (isPage) trackEvent(req, { visitorId, type: 'pageview', path: url, isOwner: !!email });
      res.writeHead(200, headers);
      res.end(html);
      return;
    }

    res.writeHead(404, securityHeaders());
    res.end('Not found');
  } catch (e) {
    console.error(e);
    res.writeHead(500, securityHeaders({ 'Content-Type': 'text/plain; charset=utf-8' }));
    res.end('Internal error');
  }
});

server.listen(PORT, () => console.log('OutsideFramework listening on port ' + PORT));
