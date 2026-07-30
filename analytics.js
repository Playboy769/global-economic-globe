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
      is_owner   INTEGER NOT NULL DEFAULT 0
    )
  `);
  db.exec('CREATE INDEX IF NOT EXISTS idx_events_ts ON events(ts)');
  db.exec('CREATE INDEX IF NOT EXISTS idx_events_visitor ON events(visitor_id, ts)');
  db.exec('CREATE INDEX IF NOT EXISTS idx_events_type_ts ON events(type, ts)');
  prune();
  return db;
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
      `INSERT INTO events (visitor_id, ts, type, path, label, referrer, device, browser, region, is_owner)
       VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)`
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
      ev.isOwner ? 1 : 0
    );
  } catch (e) {
    // Analytics must never take the site down — a failed write is a lost data point,
    // not a failed page load.
    console.error('analytics: write failed', e.message);
  }
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
      `SELECT ts, type, path, label, referrer, device, browser, region, is_owner
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
};
