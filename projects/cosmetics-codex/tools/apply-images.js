'use strict';
// 把 harvest-images.js 產出的 proposals.json 套用到「本地種子檔」與「線上資料庫」。
//
// 兩邊都要寫的原因：線上 SQLite 在 volume 裡（使用者編輯的內容以它為準），但如果哪天
// volume 重建或換環境，重新匯入的是 data/part*.json；只寫一邊會讓圖片在某次重建後消失。
//
// ⚠️ 線上更新一定要「先 GET 再合併」：server.js 的 PUT 走 store.upsert + normalize()，
// normalize 會把沒帶到的欄位一律設成 null，直接 PUT {image_url} 會把整筆洗掉。
//
// DENY 是人工複核後判定為誤配的項目——採集器靠品名相似度配對，遇到官網已改版
// （原品項下架、只剩男士線或後繼品）時會配到「同系列但不同商品」，那種寧可不放圖。
//
// 用法：node tools/apply-images.js            → 只預覽（dry run）
//       node tools/apply-images.js --write    → 實際寫入本地種子與線上資料庫

const fs = require('fs');
const path = require('path');

const API = process.env.CODEX_API || 'https://cosmetics-codex-production.up.railway.app';
const DATA_DIR = path.join(__dirname, '..', 'data');
const PROPOSALS = path.join(__dirname, 'harvest', 'proposals.json');
const WRITE = process.argv.includes('--write');

// 人工複核後排除的誤配（理由寫清楚，之後回頭看才知道為什麼跳過）
const DENY = {
  'laneige-cream-skin-toner': '官網現行品項只剩「크림 스킨 세라펩타이드 미스트」，是噴霧不是化妝水本品',
  'roundlab-birch-moisturizing-cream': '官網 sitemap 只有「포 맨」男士版 75ml，非本品',
  // 下載後目視確認：kao.co.jp 的 og:image 路徑含 cmn/share，是子品牌共用的社群分享圖
  // （藍底 Bioré 商標與插畫人物），不是商品照。通用圖偵測沒抓到是因為它每個子品牌一張、
  // 佔比沒過三成門檻。
  'biore-u-body-wash': 'Kao 的 og:image 是 Bioré 品牌 banner，非商品照',
  'biore-uv-aqua-rich-watery-essence': 'Kao 的 og:image 是 Bioré UV 品牌 banner，非商品照',
  // CANMAKE 這輪：wp-json 標題配對是準的，但該商品頁的靜態 HTML 裡沒有包裝照。
  'canmake-powder-cheeks': '該頁只有上妝後的膚色特寫，沒有商品包裝照',
  'canmake-perfect-stylist-eyes': '官網已無パーフェクトスタイリストアイズ，配到的是不同商品パーフェクトマルチアイズ',
};

async function api(pathname, init) {
  const r = await fetch(API + pathname, init);
  const j = await r.json().catch(() => ({}));
  return { status: r.status, body: j };
}

(async () => {
  if (!fs.existsSync(PROPOSALS)) { console.error('找不到 ' + PROPOSALS + '，請先跑 harvest-images.js'); process.exit(1); }
  const proposals = JSON.parse(fs.readFileSync(PROPOSALS, 'utf8'));

  const accepted = [], skipped = [];
  for (const p of proposals) (DENY[p.id] ? skipped : accepted).push(p);

  console.log('提案 ' + proposals.length + ' 筆 → 採用 ' + accepted.length + '，人工排除 ' + skipped.length);
  for (const s of skipped) console.log('  ✗ ' + s.name_zh + '：' + DENY[s.id]);
  console.log('');
  for (const a of accepted) console.log('  ✓ ' + a.name_zh.padEnd(16) + a.image_url.slice(0, 78));
  if (!WRITE) { console.log('\n（預覽模式。加 --write 才會實際寫入）'); return; }

  // ── 1. 更新本地種子檔 ──
  const byId = Object.fromEntries(accepted.map(a => [a.id, a.image_url]));
  let localCount = 0;
  for (const f of fs.readdirSync(DATA_DIR).filter(x => /^part.*\.json$/.test(x))) {
    const fp = path.join(DATA_DIR, f);
    const arr = JSON.parse(fs.readFileSync(fp, 'utf8'));
    let touched = false;
    for (const item of arr) {
      if (byId[item.id] && item.image_url !== byId[item.id]) {
        item.image_url = byId[item.id];
        touched = true; localCount++;
      }
    }
    if (touched) {
      // 種子檔的格式是「一行一筆產品」，方便 diff。逐筆 stringify 再自行組裝，
      // 不要對整包 JSON 字串跑 regex 換行——產品描述裡若剛好出現 `},{` 會被改壞。
      fs.writeFileSync(fp, '[\n' + arr.map(o => JSON.stringify(o)).join(',\n') + '\n]\n');
    }
  }
  console.log('\n本地種子檔已更新 ' + localCount + ' 筆');

  // ── 2. 更新線上資料庫（先 GET 再合併，避免 PUT 洗掉其他欄位）──
  let ok = 0, fail = 0;
  for (const a of accepted) {
    const cur = await api('/api/products/' + encodeURIComponent(a.id));
    if (cur.status !== 200 || !cur.body.item) { console.log('  線上讀取失敗：' + a.id); fail++; continue; }
    const merged = { ...cur.body.item, image_url: a.image_url };
    delete merged.created_at; delete merged.updated_at;
    const res = await api('/api/products/' + encodeURIComponent(a.id), {
      method: 'PUT',
      headers: { 'content-type': 'application/json', ...(process.env.EDIT_TOKEN ? { 'x-edit-token': process.env.EDIT_TOKEN } : {}) },
      body: JSON.stringify(merged),
    });
    if (res.status === 200 && res.body.item && res.body.item.image_url === a.image_url) ok++;
    else { console.log('  線上寫入失敗：' + a.id + ' → ' + JSON.stringify(res.body).slice(0, 90)); fail++; }
  }
  console.log('線上資料庫已更新 ' + ok + ' 筆' + (fail ? '，失敗 ' + fail + ' 筆' : ''));
})();
