'use strict';
// 高業模擬測驗的作答紀錄儲存。零 npm 依賴，沿用本 repo 既有 analytics.js 的家法：
// node:sqlite（Node 22.13 起內建，Dockerfile 因此釘 node:24-alpine）寫到 Railway
// persistent volume 掛載點 /data。
//
// 為什麼一定要 volume：Railway 每次 redeploy 都會清掉容器檔案系統，寫在映像檔裡的
// 任何路徑都會在下次 push 時無聲重置。/data 是 analytics.js 與 globe-invest 已經
// 在用的同一個慣例掛載點，在 Windows 開發機上會解析成 C:\data。

const fs = require('fs');
const path = require('path');
const { DatabaseSync } = require('node:sqlite');

const DATA_DIR = process.env.EXAM_DATA_DIR || '/data';
const DB_PATH = process.env.EXAM_DB_PATH || path.join(DATA_DIR, 'exam.db');

let db = null;

function open() {
  if (db) return db;
  const dir = path.dirname(DB_PATH);
  // 真正掛上的 Railway volume 在行程啟動前就已經是個目錄（即使是空的），
  // 所以走到這裡還不存在，多半代表 volume 根本沒掛上，而不是「今天是第一次寫入」。
  if (!fs.existsSync(dir)) {
    console.warn('[store] ' + dir + ' 不存在，建立中；若這是 Railway 環境，'
      + '請確認 volume 已掛在 ' + dir + '，否則紀錄會在下次部署時消失。');
    fs.mkdirSync(dir, { recursive: true });
  }
  db = new DatabaseSync(DB_PATH);
  db.exec('PRAGMA journal_mode = WAL');

  // 一次性 schema 遷移：2026-08-21 加上登入後，兩張表多了 email 欄用來分帳。
  // CREATE TABLE IF NOT EXISTS 不會改動既有的表，所以舊表留著會讓 INSERT 直接失敗。
  // 舊資料全是上線前的測試紀錄，依使用者決定直接重建而不是搬移。
  for (const t of ['attempts', 'answers']) {
    const exists = db.prepare(
      "SELECT name FROM sqlite_master WHERE type='table' AND name=?").get(t);
    if (!exists) continue;
    const cols = db.prepare('PRAGMA table_info(' + t + ')').all();
    if (!cols.some(c => c.name === 'email')) {
      console.warn('[store] ' + t + ' 是登入前的舊 schema（無 email 欄），重建該表。');
      db.exec('DROP TABLE ' + t);
    }
  }

  db.exec(`
    CREATE TABLE IF NOT EXISTS attempts (
      id        INTEGER PRIMARY KEY AUTOINCREMENT,
      email     TEXT NOT NULL,
      ts        TEXT NOT NULL,
      paper     TEXT NOT NULL,
      paperName TEXT NOT NULL,
      subject   TEXT NOT NULL,
      score     INTEGER NOT NULL,
      nright    INTEGER NOT NULL,
      nblank    INTEGER NOT NULL,
      ntot      INTEGER NOT NULL,
      secs      INTEGER
    );
    CREATE TABLE IF NOT EXISTS answers (
      id        INTEGER PRIMARY KEY AUTOINCREMENT,
      attemptId INTEGER NOT NULL,
      email     TEXT NOT NULL,
      ts        TEXT NOT NULL,
      paper     TEXT NOT NULL,
      subject   TEXT NOT NULL,
      qid       TEXT NOT NULL,
      head      TEXT NOT NULL,
      tag       TEXT NOT NULL,
      chosen    TEXT,
      correct   TEXT NOT NULL,
      ok        INTEGER NOT NULL
    );
    CREATE INDEX IF NOT EXISTS ix_ans_qid ON answers(email, qid);
    CREATE INDEX IF NOT EXISTS ix_ans_tag ON answers(email, subject, tag);
    CREATE INDEX IF NOT EXISTS ix_ans_ok  ON answers(email, ok);
    CREATE INDEX IF NOT EXISTS ix_att_ts  ON attempts(email, ts);
  `);
  return db;
}

// 一次交卷 = 三科各一列 attempts + 該卷全部作答明細。整批包在單一交易裡，
// 中途失敗不會留下半份紀錄。
function saveAttempt(email, p) {
  if (!email) throw new Error('saveAttempt 需要 email');
  const d = open();
  const ts = new Date().toISOString();
  const insAtt = d.prepare(
    'INSERT INTO attempts (email,ts,paper,paperName,subject,score,nright,nblank,ntot,secs)'
    + ' VALUES (?,?,?,?,?,?,?,?,?,?)');
  const insAns = d.prepare(
    'INSERT INTO answers (attemptId,email,ts,paper,subject,qid,head,tag,chosen,correct,ok)'
    + ' VALUES (?,?,?,?,?,?,?,?,?,?,?)');

  d.exec('BEGIN');
  try {
    const ids = {};
    for (const s of Object.keys(p.subjects || {})) {
      const r = p.subjects[s];
      const info = insAtt.run(email, ts, String(p.paper),
        String(p.paperName || p.paper), s,
        Number(r.score) | 0, Number(r.right) | 0, Number(r.blank) | 0,
        Number(r.tot) | 0, r.secs == null ? null : Number(r.secs) | 0);
      ids[s] = Number(info.lastInsertRowid);
    }
    for (const a of (p.answers || [])) {
      insAns.run(ids[a.subject] || 0, email, ts, String(p.paper), String(a.subject),
        String(a.qid), String(a.head), String(a.tag),
        a.chosen == null ? null : String(a.chosen),
        String(a.correct), a.ok ? 1 : 0);
    }
    d.exec('COMMIT');
    return { ok: true, ts: ts, attempts: Object.keys(ids).length,
             answers: (p.answers || []).length };
  } catch (e) {
    d.exec('ROLLBACK');
    throw e;
  }
}

// 錯題清單：同一題（qid）跨考卷、跨次數合併，依錯過幾次排序。
function wrongList(email, limit) {
  return open().prepare(`
    SELECT qid, subject,
           MAX(head) AS head, MAX(tag) AS tag,
           SUM(CASE WHEN ok=0 THEN 1 ELSE 0 END) AS nwrong,
           COUNT(*) AS nseen,
           MAX(ts) AS lastTs,
           GROUP_CONCAT(DISTINCT CASE WHEN ok=0 THEN chosen END) AS wrongPicks
    FROM answers
    WHERE email = ?
    GROUP BY qid
    HAVING nwrong > 0
    ORDER BY nwrong DESC, lastTs DESC
    LIMIT ?`).all(email, Number(limit) || 300);
}

// 觀念弱點：依科目與觀念累計正確率，只算真的作答過的（未作答不計入分母）。
function tagStats(email) {
  return open().prepare(`
    SELECT subject, tag,
           SUM(ok) AS nright,
           SUM(CASE WHEN chosen IS NOT NULL THEN 1 ELSE 0 END) AS nanswered,
           COUNT(*) AS nseen
    FROM answers
    WHERE email = ?
    GROUP BY subject, tag
    HAVING nanswered > 0
    ORDER BY (CAST(SUM(ok) AS REAL) / SUM(CASE WHEN chosen IS NOT NULL THEN 1 ELSE 0 END)) ASC,
             nseen DESC`).all(email);
}

// 進步曲線：逐次交卷的分數。
function history(email) {
  return open().prepare(
    'SELECT ts,paper,paperName,subject,score,nright,nblank,ntot,secs'
    + ' FROM attempts WHERE email = ? ORDER BY ts ASC, subject ASC').all(email);
}

function summary(email) {
  const d = open();
  const a = d.prepare('SELECT COUNT(*) n FROM attempts WHERE email = ?').get(email);
  const b = d.prepare('SELECT COUNT(*) n, SUM(ok) k FROM answers WHERE email = ?').get(email);
  const w = d.prepare(
    'SELECT COUNT(*) n FROM (SELECT qid FROM answers WHERE email = ? GROUP BY qid'
    + ' HAVING SUM(CASE WHEN ok=0 THEN 1 ELSE 0 END) > 0)').get(email);
  return {
    attempts: a.n, answers: b.n, right: b.k || 0,
    wrongQuestions: w.n, dbPath: DB_PATH,
  };
}

module.exports = { saveAttempt, wrongList, tagStats, history, summary, DB_PATH };
