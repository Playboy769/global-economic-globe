'use strict';
// 品牌官方商品圖採集器（維護工具，不會被 Dockerfile COPY 進映像檔）。
//
// 圖鑑的 image_url 走品牌官網外連，逐一手動找 166 個商品頁不切實際，這支負責自動化。
// 實測品牌站分成三種，本檔只處理得了前兩種（第三種請見 README 的說明）：
//   ① Shopify 站 → /products.json 一次給出全部商品與圖片，最可靠
//   ② 靜態站且 og:image 就是商品圖（LANEIGE、ROUND LAB、rom&nd、Kao…）→ 走 sitemap
//   ③ og:image 是全站通用圖，或商品圖由 JS 動態載入（CANMAKE、ETUDE…）→ 抓不到
// ③ 的偵測方式：同一張 og:image 出現在超過三成的頁面就判定為通用圖並整批丟棄，
// 否則會發生「第一個產品配走通用圖、其餘全被去重擋掉」的假成功（實際踩過）。
//
// 刻意只產出提案、不直接寫入資料庫：品名比對必然有誤判，要人看過再套用（apply-images.js）。
//
// 用法：node tools/harvest-images.js            → 跑設定檔裡全部來源
//       node tools/harvest-images.js --only COSRX

const fs = require('fs');
const path = require('path');

const UA = 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/126 Safari/537.36';
const OUT_DIR = path.join(__dirname, 'harvest');
const DATA_DIR = path.join(__dirname, '..', 'data');
const MIN_SCORE = 0.45;

// 來源清單由 survey-sites.js 快篩結果決定，不要憑猜測往裡加。
const SOURCES = [
  { brand: 'COSRX', type: 'shopify', base: 'https://www.cosrx.com' },
  { brand: 'Beauty of Joseon', type: 'shopify', base: 'https://beautyofjoseon.com' },
  { brand: 'dasique', type: 'shopify', base: 'https://dasique.com' },
  { brand: 'MEDIHEAL', type: 'shopify', base: 'https://www.mediheal.com' },
  { brand: 'LANEIGE', type: 'sitemap', sitemap: 'https://www.laneige.com/sitemap.xml', pat: /\/product|\/p\// },
  { brand: 'ROUND LAB', type: 'sitemap', sitemap: 'https://roundlab.co.kr/sitemap.xml', pat: /\/product\// },
  { brand: 'rom&nd', type: 'sitemap', sitemap: 'https://romand.co.kr/sitemap.xml', pat: /\/product\// },
  { brand: 'Curél', type: 'sitemap', sitemap: 'https://www.kao.co.jp/sitemap.xml', pat: /curel/ },
  { brand: 'Bioré', type: 'sitemap', sitemap: 'https://www.kao.co.jp/sitemap.xml', pat: /biore/ },
  { brand: 'ATRIX', type: 'sitemap', sitemap: 'https://www.kao.co.jp/sitemap.xml', pat: /atrix/ },
];

async function get(url, ms = 20000, asJson = false) {
  const c = new AbortController();
  const t = setTimeout(() => c.abort(), ms);
  try {
    const r = await fetch(url, { headers: { 'user-agent': UA, 'accept-language': 'ja,ko,en' }, signal: c.signal, redirect: 'follow' });
    if (!r.ok) return null;
    return asJson ? await r.json() : await r.text();
  } catch { return null; } finally { clearTimeout(t); }
}

async function locs(url, depth = 0) {
  const xml = await get(url);
  if (!xml) return [];
  const l = [...xml.matchAll(/<loc>\s*([^<\s]+)\s*<\/loc>/gi)].map(m => m[1]);
  if (!/<sitemapindex/i.test(xml) || depth >= 2) return l;
  let out = [];
  for (const s of l.slice(0, 20)) out = out.concat(await locs(s, depth + 1));
  return out;
}

function ogMeta(html, prop) {
  const tag = (html.match(new RegExp('<meta[^>]+(?:property|name)=["\']' + prop + '["\'][^>]*>', 'i')) || [])[0];
  if (!tag) return '';
  return ((tag.match(/content=["']([^"']*)["']/i) || [])[1] || '').trim();
}

function decode(s) {
  return String(s || '').replace(/&amp;/g, '&').replace(/&lt;/g, '<').replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"').replace(/&#0?39;|&apos;/g, "'").replace(/&nbsp;/g, ' ');
}

function norm(s) {
  return decode(s).toLowerCase().normalize('NFKC')
    .replace(/[\s　・･,，、。.()（）\[\]【】「」『』<>《》|｜/／\\+＋&＆'"’”_-]/g, '');
}
function words(s) {
  return decode(s).toLowerCase().normalize('NFKC')
    .replace(/[^a-z0-9぀-ヿ一-鿿가-힯]+/g, ' ')
    .split(' ').filter(w => w.length > 1);
}
function bigrams(s) { const o = new Set(); for (let i = 0; i < s.length - 1; i++) o.add(s.slice(i, i + 2)); return o; }

function dice(a, b) {
  const A = norm(a), B = norm(b);
  if (!A || !B) return 0;
  if (A.length >= 4 && B.includes(A)) return 1;
  if (B.length >= 5 && A.includes(B)) return 0.9;
  const ga = bigrams(A), gb = bigrams(B);
  if (!ga.size || !gb.size) return 0;
  let i = 0; for (const g of ga) if (gb.has(g)) i++;
  return (2 * i) / (ga.size + gb.size);
}
// 英文標題（Shopify 多為英文）用詞彙重疊，中日韓用 bigram，取兩者最高分。
function tokenOverlap(a, b) {
  const A = new Set(words(a)), B = new Set(words(b));
  if (!A.size || !B.size) return 0;
  let i = 0; for (const w of A) if (B.has(w)) i++;
  return i / A.size;
}
// 變體關鍵字不該搶走本品的圖。這份清單是實跑後補出來的——第一版漏了男士線與大容量，
// 結果 LANEIGE 化妝水配到「크림 스킨 옴므」（男士全效）、ROUND LAB 面霜配到「포 맨」，
// rom&nd 唇釉配到「미니」。判斷依據是這些詞代表「同系列但不同商品」。
const VARIANT = new RegExp([
  'mini', 'ミニ', '미니',                                   // 迷你版
  '替え', '詰替', 'つめかえ', 'リフィル', 'refill', '리필',   // 補充包
  'セット', 'set', 'キット', 'kit', 'duo', '세트', '기획', '증정', // 組合／贈品
  'ブラシ', 'brush', 'パフ', 'puff', 'リムーバー', 'remover',  // 周邊配件
  'homme', 'for ?men', '옴므', '포 ?맨', '男性', 'メンズ',      // 男士線
  '대용량', '大容量', '徳用',                                // 大容量版
].join('|'), 'i');

function scorePair(product, cand) {
  const idText = product.id.replace(/-/g, ' ');
  const s = Math.max(
    dice(product.name_local, cand.title),
    tokenOverlap(idText, cand.title),
    tokenOverlap(product.name_local || '', cand.title)
  );
  return s - (VARIANT.test(cand.title) ? 0.18 : 0);
}

function loadProducts() {
  let all = [];
  for (const f of fs.readdirSync(DATA_DIR).filter(x => /^part.*\.json$/.test(x))) {
    all = all.concat(JSON.parse(fs.readFileSync(path.join(DATA_DIR, f), 'utf8')));
  }
  return all;
}

async function fromShopify(src) {
  const out = [];
  for (let page = 1; page <= 4; page++) {
    const j = await get(src.base + '/products.json?limit=250&page=' + page, 25000, true);
    if (!j || !Array.isArray(j.products) || !j.products.length) break;
    for (const p of j.products) {
      const img = (p.images && p.images[0] && p.images[0].src) || (p.image && p.image.src);
      if (!img) continue;
      out.push({ title: p.title, image: img, url: src.base + '/products/' + p.handle });
    }
    if (j.products.length < 250) break;
  }
  return out;
}

async function fromSitemap(src) {
  let urls = (await locs(src.sitemap)).filter(u => src.pat.test(u));
  urls = [...new Set(urls)].slice(0, 400);
  const pages = [];
  const queue = urls.slice();
  await Promise.all(Array.from({ length: 6 }, async () => {
    while (queue.length) {
      const u = queue.shift();
      const h = await get(u, 18000);
      if (!h) continue;
      const image = ogMeta(h, 'og:image');
      if (!image) continue;
      const title = ogMeta(h, 'og:title') || decode(((h.match(/<title[^>]*>([^<]*)<\/title>/i) || [])[1] || '').trim());
      pages.push({ title, image, url: u });
    }
  }));
  // 通用圖偵測：同一張圖佔超過三成頁面就是版型預設圖，不是商品圖。
  const freq = {};
  for (const p of pages) freq[p.image] = (freq[p.image] || 0) + 1;
  // 只用「佔比」判斷通用圖。先前多加了一條 freq>=8 的絕對門檻，結果把 Kao 站
  // 整批誤殺（同系列多品項共用一張系列圖，8 次很容易達到但那仍是有效的商品圖）。
  const generic = new Set(Object.keys(freq).filter(k => freq[k] / pages.length > 0.3));
  const kept = pages.filter(p => !generic.has(p.image));
  if (generic.size) console.log('    （排除 ' + generic.size + ' 張通用圖，涵蓋 ' + (pages.length - kept.length) + ' 頁）');
  return kept;
}

async function imageLoads(url) {
  try {
    const c = new AbortController();
    const t = setTimeout(() => c.abort(), 15000);
    const r = await fetch(url, { headers: { 'user-agent': UA, referer: 'https://cosmetics-codex-production.up.railway.app/' }, signal: c.signal });
    clearTimeout(t);
    return r.ok && (r.headers.get('content-type') || '').startsWith('image/');
  } catch { return false; }
}

(async () => {
  const only = process.argv.includes('--only') ? process.argv[process.argv.indexOf('--only') + 1] : null;
  const all = loadProducts();
  fs.mkdirSync(OUT_DIR, { recursive: true });
  const proposals = [];

  for (const src of SOURCES) {
    if (only && src.brand.toLowerCase() !== only.toLowerCase()) continue;
    const mine = all.filter(p => p.brand.toLowerCase() === src.brand.toLowerCase());
    if (!mine.length) continue;
    process.stdout.write(src.brand + '（圖鑑內 ' + mine.length + ' 筆，來源 ' + src.type + '）… ');
    const cands = src.type === 'shopify' ? await fromShopify(src) : await fromSitemap(src);
    console.log('候選 ' + cands.length + ' 個');
    if (!cands.length) continue;

    const ranked = [];
    for (const p of mine) for (const c of cands) {
      const s = scorePair(p, c);
      if (s >= MIN_SCORE) ranked.push({ s, p, c });
    }
    // 同分時偏好標題較短的（變體名通常更長）
    ranked.sort((a, b) => b.s - a.s || a.c.title.length - b.c.title.length);
    const takenP = new Set(), takenI = new Set();
    for (const r of ranked) {
      if (takenP.has(r.p.id) || takenI.has(r.c.image)) continue;
      if (!(await imageLoads(r.c.image))) { console.log('    圖片載入失敗，跳過：' + r.c.title.slice(0, 30)); continue; }
      takenP.add(r.p.id); takenI.add(r.c.image);
      proposals.push({
        id: r.p.id, name_zh: r.p.name_zh, brand: r.p.brand,
        score: Number(r.s.toFixed(3)), matched_title: r.c.title,
        image_url: r.c.image, source_page: r.c.url,
      });
      console.log('    ✓ ' + r.s.toFixed(2) + '  ' + r.p.name_zh + '  ←  ' + r.c.title.slice(0, 44));
    }
    for (const p of mine) if (!takenP.has(p.id)) console.log('    – 未配對：' + p.name_zh);
  }

  const out = path.join(OUT_DIR, 'proposals.json');
  fs.writeFileSync(out, JSON.stringify(proposals, null, 2) + '\n');
  console.log('\n合計配對 ' + proposals.length + ' 筆 → ' + path.relative(process.cwd(), out));
})();
