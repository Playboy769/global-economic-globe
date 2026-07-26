// ═══════════════════════════════════════════════════
// 在瀏覽器打開「文章資料庫.html」後，按 F12 開 Console
// 把整段程式碼貼進去並按 Enter 執行
// ═══════════════════════════════════════════════════

(async function migrateToRailway() {
  const RAILWAY = 'https://web-production-2210c.up.railway.app';

  // ── 讀取 localStorage 資料 ──
  const articles = JSON.parse(localStorage.getItem('articles_db') || '[]');
  const folders  = JSON.parse(localStorage.getItem('folders_db')  || '[]');
  const reports  = JSON.parse(localStorage.getItem('reports_db')  || '[]');

  console.log(`📦 資料總覽：文章 ${articles.length} 篇 ／ 資料夾 ${folders.length} 個 ／ 報告 ${reports.length} 份`);
  if (!articles.length && !reports.length) {
    console.warn('⚠️  localStorage 沒有資料，請確認已在正確的頁面執行。');
    return;
  }

  // ── 工具函式 ──
  async function post(path, body) {
    const r = await fetch(RAILWAY + path, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(body)
    });
    if (!r.ok) {
      const txt = await r.text().catch(() => '');
      throw new Error(`HTTP ${r.status}：${txt.slice(0, 120)}`);
    }
    const text = await r.text();
    return text ? JSON.parse(text) : null;
  }

  // ── 1. 建立資料夾 → Railway categories ──
  console.log('\n─── [1/3] 轉移資料夾 ───');
  const catMap = {};  // folderName → cloudId
  for (const name of folders) {
    try {
      const res = await post('/categories', { name });
      catMap[name] = res.id;
      console.log(`  ✓ 「${name}」→ id ${res.id}`);
    } catch (e) {
      console.error(`  ✕ 「${name}」失敗：`, e.message);
    }
  }
  console.log(`  完成：${Object.keys(catMap).length} / ${folders.length} 個資料夾`);

  // ── 2. 上傳文章 ──
  console.log('\n─── [2/3] 轉移文章 ───');
  let artOk = 0, artFail = 0;
  const total = articles.length;

  for (let i = 0; i < articles.length; i++) {
    const a = articles[i];
    const imgs = Array.isArray(a.images) ? a.images : (a.image ? [a.image] : []);

    try {
      await post('/articles', {
        title:       a.title   || '未命名',
        content:     a.content || '',
        author:      a.author  || null,
        source:      a.source  || null,
        language:    'zh',
        category_id: catMap[a.folder] ?? null,
        summary:     null,
        date:        a.date    || null,
        images:      JSON.stringify(imgs),
        starred:     a.starred || 0,
        tags:        []
      });
      artOk++;
      const star = a.starred ? ' ⭐' : '';
      console.log(`  [${i+1}/${total}] ✓ ${a.title}${star}`);
    } catch (e) {
      artFail++;
      console.error(`  [${i+1}/${total}] ✕ ${a.title}：${e.message}`);
    }
  }
  console.log(`  完成：${artOk} / ${total}（失敗 ${artFail}）`);

  // ── 3. 上傳報告 ──
  console.log('\n─── [3/3] 轉移公司報告 ───');
  let repOk = 0, repFail = 0;
  const rtotal = reports.length;

  for (let i = 0; i < reports.length; i++) {
    const r = reports[i];
    const imgs = Array.isArray(r.images) ? r.images : (r.image ? [r.image] : []);

    try {
      await post('/reports', {
        company:  r.company  || '未命名',
        ticker:   r.ticker   || null,
        sector:   r.sector   || null,
        rating:   r.rating   || null,
        target:   r.target   || null,
        date:     r.date     || null,
        analyst:  r.analyst  || null,
        source:   r.source   || null,
        content:  r.content  || '',
        images:   JSON.stringify(imgs)
      });
      repOk++;
      console.log(`  [${i+1}/${rtotal}] ✓ ${r.company}（${r.ticker || '—'}）`);
    } catch (e) {
      repFail++;
      console.error(`  [${i+1}/${rtotal}] ✕ ${r.company}：${e.message}`);
    }
  }
  console.log(`  完成：${repOk} / ${rtotal}（失敗 ${repFail}）`);

  // ── 結果 ──
  console.log('\n═══════════════════════════════════════');
  console.log(`✅ 全部完成！`);
  console.log(`   文章：${artOk} 成功 / ${artFail} 失敗`);
  console.log(`   報告：${repOk} 成功 / ${repFail} 失敗`);
  console.log(`   打開 Railway 前端按 F5 重整即可看到資料`);
  console.log('═══════════════════════════════════════');
})();
