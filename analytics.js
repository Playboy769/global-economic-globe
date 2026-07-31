'use strict';
// Visitor analytics storage for OutsideFramework — zero npm dependencies, matching the
// house style of server.js/auth.js. Uses node:sqlite (built in since Node 22.13; the
// Dockerfile pins node:24-alpine for this reason).
//
// The DB lives on a Railway persistent volume mounted at /data — Railway wipes the
// container filesystem on every redeploy, so writing anywhere inside the image would
// silently reset the stats on each push. /data is the same mount path globe-invest's
// server.js already uses, which on a Windows dev box resolves to C:\data.

const fs = require('fs');
const path = require('path');
const https = require('https');
const { DatabaseSync } = require('node:sqlite');

const DATA_DIR = process.env.ANALYTICS_DATA_DIR || '/data';
const DB_PATH = process.env.ANALYTICS_DB_PATH || path.join(DATA_DIR, 'analytics.db');

// Events older than this are dropped at startup. Long enough for year-over-year
// comparison, short enough that the volume never needs attention.
const RETENTION_DAYS = 400;

// A visit belongs to the same session as the previous one from that visitor unless
// this many minutes elapsed — the standard web-analytics convention. Sessions are
// derived at query time from timestamps rather than tracked with a second cookie.
const SESSION_GAP_MIN = 30;

// Days are bucketed in this timezone, not UTC. Without it a Taipei-based reader sees
// each "day" start at 08:00 local, so an evening's traffic lands on the next day's bar.
const TZ_OFFSET_MIN = Number(process.env.ANALYTICS_TZ_OFFSET_MIN || 480); // +08:00 Taipei
const TZ_SQL = (TZ_OFFSET_MIN >= 0 ? '+' : '-') + Math.abs(TZ_OFFSET_MIN) + ' minutes';

let db = null;

function open() {
  if (db) return db;
  const dir = path.dirname(DB_PATH);
  // A real Railway volume mount already exists as a directory before this process starts,
  // even empty — so if it's missing here, on the default /data path this almost certainly
  // means the volume was never attached, not that today's traffic is the first write ever.
  // Silently mkdir-ing over it would just start writing into the container's throwaway
  // filesystem with no sign anything is wrong until the next deploy erases it all.
  if (!fs.existsSync(dir)) {
    if (!process.env.ANALYTICS_DATA_DIR && !process.env.ANALYTICS_DB_PATH) {
      console.error(
        'WARNING: ' + dir + ' does not exist — analytics is very likely NOT on a ' +
        'persistent volume and will reset on the next deploy. Mount a Railway Volume at ' +
        dir + ' (or set ANALYTICS_DATA_DIR) to fix this.'
      );
    }
    fs.mkdirSync(dir, { recursive: true });
  }
  db = new DatabaseSync(DB_PATH);
  db.exec('PRAGMA journal_mode = WAL');
  db.exec(`
    CREATE TABLE IF NOT EXISTS events (
      id         INTEGER PRIMARY KEY AUTOINCREMENT,
      visitor_id TEXT    NOT NULL,
      ts         INTEGER NOT NULL,
      type       TEXT    NOT NULL,
      path       TEXT    NOT NULL DEFAULT '',
      label      TEXT    NOT NULL DEFAULT '',
      referrer   TEXT    NOT NULL DEFAULT '',
      device     TEXT    NOT NULL DEFAULT '',
      browser    TEXT    NOT NULL DEFAULT '',
      region     TEXT    NOT NULL DEFAULT '',
      is_owner   INTEGER NOT NULL DEFAULT 0,
      ip         TEXT    NOT NULL DEFAULT ''
    )
  `);
  // CREATE TABLE IF NOT EXISTS is a no-op against the already-deployed table on the Railway
  // volume, so the `ip` column above only takes effect on a brand-new DB. This ALTER is what
  // actually adds it to the live one; the try/catch makes it a harmless no-op on every boot
  // after the first (SQLite has no "add column if missing" syntax).
  try {
    db.exec("ALTER TABLE events ADD COLUMN ip TEXT NOT NULL DEFAULT ''");
  } catch {
    // column already exists
  }
  // Caches IP -> country/city so the admin dashboard doesn't re-hit the geolocation API on
  // every view of the same visitor. Unbounded but tiny (a few dozen distinct IPs for a
  // personal-scale site over years) — not worth pruning alongside the 400-day event
  // retention below.
  db.exec(`
    CREATE TABLE IF NOT EXISTS ip_geo (
      ip          TEXT    PRIMARY KEY,
      country     TEXT    NOT NULL DEFAULT '',
      city        TEXT    NOT NULL DEFAULT '',
      resolved_at INTEGER NOT NULL
    )
  `);
  db.exec('CREATE INDEX IF NOT EXISTS idx_events_ts ON events(ts)');
  db.exec('CREATE INDEX IF NOT EXISTS idx_events_visitor ON events(visitor_id, ts)');
  db.exec('CREATE INDEX IF NOT EXISTS idx_events_type_ts ON events(type, ts)');
  prune();
  purgeOwnerRows();
  return db;
}

// Deletes every row ever flagged is_owner (visits recorded while the owner was actively
// logged in). Runs on every boot, not just once: OWNER_IPS (server.js) is the real ongoing
// exclusion for the owner's known devices, but is_owner still gets set for a login from an
// unrecognized device/network (travel, a friend's wifi), and the owner asked not to have
// their own visits recorded at all — so any that slip through get swept out at the next
// restart rather than lingering in the DB indefinitely.
function purgeOwnerRows() {
  if (!db) return;
  try {
    const result = db.prepare('DELETE FROM events WHERE is_owner = 1').run();
    if (result.changes) console.log('analytics: purged ' + result.changes + ' owner-flagged row(s)');
  } catch (e) {
    console.error('analytics: purge owner rows failed', e.message);
  }
}

// Retention used to be enforced only here, inside open() — which memoizes `db`, so it ran
// exactly once per process. A container that stays up for months therefore never pruned
// again and the "400 days" ceiling was really "400 days, or however long since the last
// deploy, whichever is longer". Now it also runs from record(), at most once a day.
let lastPruneMs = 0;
const PRUNE_INTERVAL_MS = 24 * 60 * 60 * 1000;

function prune() {
  if (!db) return;
  try {
    const cutoff = Math.floor(Date.now() / 1000) - RETENTION_DAYS * 86400;
    db.prepare('DELETE FROM events WHERE ts < ?').run(cutoff);
    lastPruneMs = Date.now();
  } catch (e) {
    console.error('analytics: retention prune failed', e.message);
  }
}

// Bots identify themselves in the UA far more often than not, and the ones that don't
// mostly never execute JS (so they only ever produce a server-side pageview, never a
// /api/track beacon). Filtering at write time keeps the stored data honest instead of
// making every query carry the exclusion.
const BOT_RE = /bot|crawler|spider|crawling|slurp|facebookexternalhit|preview|scrape|curl|wget|python-requests|headless|lighthouse|pingdom|uptime|monitor/i;

function isBot(ua) {
  return BOT_RE.test(String(ua || ''));
}

function parseDevice(ua) {
  const s = String(ua || '');
  if (/iPad|Tablet/i.test(s)) return 'Tablet';
  if (/Mobi|Android|iPhone|iPod/i.test(s)) return 'Mobile';
  if (!s) return 'Unknown';
  return 'Desktop';
}

function parseBrowser(ua) {
  const s = String(ua || '');
  // Order matters — Edge and Opera both claim "Chrome", Chrome claims "Safari".
  if (/Edg\//i.test(s)) return 'Edge';
  if (/OPR\/|Opera/i.test(s)) return 'Opera';
  if (/Firefox\//i.test(s)) return 'Firefox';
  if (/Chrome\//i.test(s)) return 'Chrome';
  if (/Safari\//i.test(s)) return 'Safari';
  return 'Other';
}

// Referrers are stored as a bare hostname, not the full URL: the path often carries
// search terms or campaign junk we have no use for, and collapsing to host is what
// makes the "traffic sources" chart readable.
function parseReferrer(ref, selfHost) {
  if (!ref) return '';
  let host;
  try {
    host = new URL(ref).hostname.replace(/^www\./, '');
  } catch {
    return '';
  }
  if (selfHost && host === String(selfHost).replace(/^www\./, '').split(':')[0]) return '';
  return host;
}

function clamp(s, max) {
  return String(s == null ? '' : s).slice(0, max);
}

function record(ev) {
  try {
    const d = open();
    if (Date.now() - lastPruneMs > PRUNE_INTERVAL_MS) prune();
    d.prepare(
      `INSERT INTO events (visitor_id, ts, type, path, label, referrer, device, browser, region, is_owner, ip)
       VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)`
    ).run(
      clamp(ev.visitorId, 64),
      Math.floor(Date.now() / 1000),
      clamp(ev.type || 'pageview', 24),
      clamp(ev.path, 300),
      clamp(ev.label, 200),
      clamp(ev.referrer, 200),
      clamp(ev.device, 24),
      clamp(ev.browser, 24),
      clamp(ev.region, 8),
      ev.isOwner ? 1 : 0,
      clamp(ev.ip, 64)
    );
  } catch (e) {
    // Analytics must never take the site down — a failed write is a lost data point,
    // not a failed page load.
    console.error('analytics: write failed', e.message);
  }
}

// Loopback/private ranges and Node's "socket had no address" fallback — never worth sending
// to a public geolocation API (they'd either error or resolve to the API provider's own
// network), so these skip the network call entirely rather than burning a lookup and a
// cache slot on them.
const PRIVATE_IP_RE = /^(127\.|10\.|192\.168\.|172\.(1[6-9]|2\d|3[01])\.|::1$|f[cd][0-9a-f]{0,2}:|fe80:|unknown$)/i;

function isPrivateIp(ip) {
  return !ip || PRIVATE_IP_RE.test(String(ip));
}

// One HTTPS call to ipapi.co's free JSON endpoint. Never called from a visitor-facing
// request path — only from the admin dashboard handler (server.js), which is the one
// place a live network round-trip on this scale is acceptable. Returns null (not cached)
// on any failure so a transient outage or rate limit gets retried next time the dashboard
// is opened, instead of a blank result sticking forever.
function fetchGeoFromApi(ip) {
  const GEO_TIMEOUT_MS = 5000;
  return new Promise((resolve) => {
    let done = false;
    const finish = (v) => {
      if (done) return;
      done = true;
      resolve(v);
    };
    const req = https.get(
      'https://ipapi.co/' + encodeURIComponent(ip) + '/json/',
      { headers: { 'User-Agent': 'ofw-analytics/1.0 (personal site, admin-only lookup)' }, timeout: GEO_TIMEOUT_MS },
      (res) => {
        let data = '';
        res.on('data', (c) => (data += c));
        res.on('end', () => {
          try {
            const j = JSON.parse(data);
            if (j.error) { finish(null); return; }
            finish({ country: clamp(j.country_name, 80), city: clamp(j.city, 80) });
          } catch {
            finish(null);
          }
        });
      }
    );
    req.on('timeout', () => req.destroy());
    req.on('error', () => finish(null));
  });
}

// Resolves an IP to {country, city}, checking the ip_geo cache first. Successful lookups
// are cached indefinitely (an IP's rough geolocation doesn't meaningfully change on the
// timescale this dashboard cares about); failures are not cached, so they retry on the
// next dashboard view rather than permanently showing blank for a real visitor.
async function resolveGeo(ip) {
  if (isPrivateIp(ip)) return { country: '', city: '' };
  const d = open();
  const cached = d.prepare('SELECT country, city FROM ip_geo WHERE ip = ?').get(ip);
  if (cached) return cached;
  const geo = await fetchGeoFromApi(ip);
  if (!geo) return { country: '', city: '' };
  try {
    d.prepare(
      'INSERT OR REPLACE INTO ip_geo (ip, country, city, resolved_at) VALUES (?, ?, ?, ?)'
    ).run(ip, geo.country, geo.city, Math.floor(Date.now() / 1000));
  } catch (e) {
    console.error('analytics: geo cache write failed', e.message);
  }
  return geo;
}

// The window is N whole days ending with today (inclusive), aligned to TZ_OFFSET_MIN
// midnight — not "N*24h ago from this instant", which would both clip today's traffic
// off the trend and make each bar a rolling partial day.
function windowClause(days, includeOwner) {
  const offsetMs = TZ_OFFSET_MIN * 60000;
  const todayLocalMidnightMs =
    Math.floor((Date.now() + offsetMs) / 86400000) * 86400000 - offsetMs;
  const since = Math.floor((todayLocalMidnightMs - (days - 1) * 86400000) / 1000);
  return { since, ownerSql: includeOwner ? '' : ' AND is_owner = 0' };
}

function summary(days, includeOwner) {
  const d = open();
  const { since, ownerSql } = windowClause(days, includeOwner);

  const totals = d
    .prepare(
      `SELECT COUNT(*) AS views, COUNT(DISTINCT visitor_id) AS visitors
       FROM events WHERE ts >= ? AND type = 'pageview'${ownerSql}`
    )
    .get(since) || { views: 0, visitors: 0 };

  // A returning visitor is one whose first-ever pageview predates this window.
  const returning = d
    .prepare(
      `SELECT COUNT(*) AS n FROM (
         SELECT visitor_id, MIN(ts) AS first_ts
         FROM events WHERE type = 'pageview'${ownerSql}
         GROUP BY visitor_id
       ) WHERE first_ts < ? AND visitor_id IN (
         SELECT DISTINCT visitor_id FROM events
         WHERE ts >= ? AND type = 'pageview'${ownerSql}
       )`
    )
    .get(since, since) || { n: 0 };

  const sessions = countSessions(days, includeOwner);

  return {
    views: totals.views || 0,
    visitors: totals.visitors || 0,
    sessions,
    returningVisitors: returning.n || 0,
  };
}

// Sessions are derived rather than stored: pull each visitor's pageview timestamps in
// order and start a new session whenever the gap exceeds SESSION_GAP_MIN.
function countSessions(days, includeOwner) {
  const d = open();
  const { since, ownerSql } = windowClause(days, includeOwner);
  const rows = d
    .prepare(
      `SELECT visitor_id, ts FROM events
       WHERE ts >= ? AND type = 'pageview'${ownerSql}
       ORDER BY visitor_id, ts`
    )
    .all(since);
  let sessions = 0;
  let prevVisitor = null;
  let prevTs = 0;
  for (const r of rows) {
    if (r.visitor_id !== prevVisitor || r.ts - prevTs > SESSION_GAP_MIN * 60) sessions++;
    prevVisitor = r.visitor_id;
    prevTs = r.ts;
  }
  return sessions;
}

// Returns one row per day across the whole window, including days with zero traffic —
// a line chart with missing days silently compresses the gaps and misreads as steady.
function dailyTrend(days, includeOwner) {
  const d = open();
  const { since, ownerSql } = windowClause(days, includeOwner);
  const rows = d
    .prepare(
      `SELECT date(ts, 'unixepoch', '${TZ_SQL}') AS day,
              COUNT(*) AS views,
              COUNT(DISTINCT visitor_id) AS visitors
       FROM events WHERE ts >= ? AND type = 'pageview'${ownerSql}
       GROUP BY day ORDER BY day`
    )
    .all(since);
  const byDay = new Map(rows.map((r) => [r.day, r]));
  const out = [];
  const startMs = since * 1000 + TZ_OFFSET_MIN * 60000;
  for (let i = 0; i < days; i++) {
    const day = new Date(startMs + i * 86400000).toISOString().slice(0, 10);
    const hit = byDay.get(day);
    out.push({ day, views: hit ? hit.views : 0, visitors: hit ? hit.visitors : 0 });
  }
  return out;
}

// A column name cannot be a bound parameter, so it is interpolated — which means it must
// come from this list and nowhere else. Every call site below passes a literal today; the
// whitelist is what keeps that true after someone wires a column name to a query string.
const GROUPABLE = new Set(['path', 'label', 'referrer', 'device', 'browser', 'region']);

function topBy(column, days, includeOwner, type, limit) {
  if (!GROUPABLE.has(column)) throw new Error('analytics: refusing to group by ' + column);
  const d = open();
  const { since, ownerSql } = windowClause(days, includeOwner);
  return d
    .prepare(
      `SELECT ${column} AS name, COUNT(*) AS count, COUNT(DISTINCT visitor_id) AS visitors
       FROM events
       WHERE ts >= ? AND type = ? AND ${column} != ''${ownerSql}
       GROUP BY ${column} ORDER BY count DESC LIMIT ?`
    )
    .all(since, type, limit);
}

function recentVisits(includeOwner, limit) {
  const d = open();
  const ownerSql = includeOwner ? '' : ' WHERE is_owner = 0';
  return d
    .prepare(
      `SELECT ts, type, path, label, referrer, device, browser, region, is_owner, ip
       FROM events${ownerSql} ORDER BY ts DESC LIMIT ?`
    )
    .all(limit);
}

function stats({ days = 30, includeOwner = false } = {}) {
  return {
    days,
    includeOwner,
    summary: summary(days, includeOwner),
    trend: dailyTrend(days, includeOwner),
    topPages: topBy('path', days, includeOwner, 'pageview', 12),
    topWorks: topBy('label', days, includeOwner, 'work_click', 12),
    referrers: topBy('referrer', days, includeOwner, 'pageview', 10),
    devices: topBy('device', days, includeOwner, 'pageview', 6),
    browsers: topBy('browser', days, includeOwner, 'pageview', 6),
    regions: topBy('region', days, includeOwner, 'pageview', 10),
    recent: recentVisits(includeOwner, 40),
  };
}

module.exports = {
  DB_PATH,
  RETENTION_DAYS,
  open,
  record,
  stats,
  isBot,
  parseDevice,
  parseBrowser,
  parseReferrer,
  resolveGeo,
};
