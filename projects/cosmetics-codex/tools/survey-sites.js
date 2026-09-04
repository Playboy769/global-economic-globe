'use strict';
// 品牌站可採性快篩。
//
// 動機：不是每個品牌站都能用 curl 抓到商品圖。實測有三種情況：
//   ① og:image 就是商品圖（如 Kanebo/KATE）→ 最理想
//   ② og:image 是全站通用圖（如 CANMAKE，124 頁只有 2 種 og:image）→ 無用
//   ③ 商品圖由 JS 動態載入，靜態 HTML 裡根本沒有 → 無用
// 全量抓完才發現是 ② 或 ③ 很浪費，所以先每站抽樣十幾頁，用「不同 og:image 的比例」
// 判斷可採性，再決定要不要跑完整採集。

const UA = 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/126 Safari/537.36';
const SAMPLE = 12;

const SITES = [
  { brand: 'LANEIGE', sitemap: 'https://www.laneige.com/sitemap.xml', pat: /\/product|\/p\// },
  { brand: 'ROUND LAB', sitemap: 'https://roundlab.co.kr/sitemap.xml', pat: /\/product\// },
  { brand: 'rom&nd', sitemap: 'https://romand.co.kr/sitemap.xml', pat: /\/product\// },
  { brand: 'ETUDE', sitemap: 'https://www.etude.com/sitemap.xml', pat: /\/product|\/p\// },
  { brand: 'COSRX', sitemap: 'https://www.cosrx.com/sitemap.xml', pat: /\/product/ },
  { brand: 'Beauty of Joseon', sitemap: 'https://beautyofjoseon.com/sitemap.xml', pat: /\/product/ },
  { brand: 'CLIO', sitemap: 'https://www.clubclio.com/sitemap.xml', pat: /\/product/ },
  { brand: 'SHISEIDO', sitemap: 'https://www.shiseido.co.jp/sitemap.xml', pat: /\/products?\// },
  { brand: 'Kao', sitemap: 'https://www.kao.co.jp/sitemap.xml', pat: /\/products?\// },
  { brand: 'TORRIDEN', sitemap: 'https://torriden.com/sitemap.xml', pat: /\/product/ },
];

async function get(url, ms = 20000) {
  const c = new AbortController();
  const t = setTimeout(() => c.abort(), ms);
  try {
    const r = await fetch(url, { headers: { 'user-agent': UA }, signal: c.signal, redirect: 'follow' });
    return r.ok ? await r.text() : null;
  } catch { return null; } finally { clearTimeout(t); }
}

async function locs(url, depth = 0) {
  const xml = await get(url);
  if (!xml) return [];
  const l = [...xml.matchAll(/<loc>\s*([^<\s]+)\s*<\/loc>/gi)].map(m => m[1]);
  if (!/<sitemapindex/i.test(xml) || depth >= 2) return l;
  let out = [];
  for (const s of l.slice(0, 12)) out = out.concat(await locs(s, depth + 1));
  return out;
}

function og(html, prop) {
  const tag = (html.match(new RegExp('<meta[^>]+(?:property|name)=["\']' + prop + '["\'][^>]*>', 'i')) || [])[0];
  if (!tag) return '';
  return ((tag.match(/content=["']([^"']*)["']/i) || [])[1] || '').trim();
}

(async () => {
  for (const s of SITES) {
    let urls = await locs(s.sitemap);
    const prod = urls.filter(u => s.pat.test(u));
    if (!prod.length) {
      console.log(String(s.brand).padEnd(18) + '✗ sitemap 無商品頁（總 URL ' + urls.length + '）');
      continue;
    }
    // 均勻取樣，避免全部落在同一個分類
    const step = Math.max(1, Math.floor(prod.length / SAMPLE));
    const pick = [];
    for (let i = 0; i < prod.length && pick.length < SAMPLE; i += step) pick.push(prod[i]);

    const imgs = [], titles = [];
    await Promise.all(pick.map(async u => {
      const h = await get(u);
      if (!h) return;
      const i = og(h, 'og:image');
      if (i) imgs.push(i);
      const t = og(h, 'og:title');
      if (t) titles.push(t);
    }));
    const distinct = new Set(imgs).size;
    const ratio = imgs.length ? distinct / imgs.length : 0;
    const verdict = ratio >= 0.8 ? '✓ 可採' : (distinct <= 2 ? '✗ og:image 是通用圖' : '△ 部分可用');
    console.log(String(s.brand).padEnd(18) + verdict
      + '  商品頁 ' + String(prod.length).padEnd(5)
      + ' 抽樣 ' + pick.length + ' 頁，og:image ' + imgs.length + ' 個 / 相異 ' + distinct
      + (titles[0] ? '  例：' + titles[0].slice(0, 30) : ''));
  }
})();
