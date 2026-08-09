/* ============================================================================
 * chart-tools 共用工具函式
 *
 * 四個圖表工具（折線／甘特／圓餅／瀑布）共用這一份。刻意做成獨立 .js 而不是
 * 每個檔案各自複製一份，因為「產生嵌入代碼」的序列化規則（inline 化樣式、
 * 逸出、互動 script 樣板）只要有一處不一致，貼進報告後就會出現只有某一種圖
 * 壞掉的狀況，很難查。
 *
 * ⚠️ 注意：這個檔案只在「工具頁面本身」被載入。產生出來的嵌入代碼是自包含的，
 * 不會反過來依賴這支檔案 —— 報告端不需要、也不應該引用 _shared.js。
 * ==========================================================================*/
(function (global) {
  'use strict';

  var SVGNS = 'http://www.w3.org/2000/svg';

  /* ---------- SVG 建構 ----------
   * 一律使用 presentation attribute（fill / font-size / stroke…）而不是 CSS class。
   * 嵌入代碼會被貼進報告，工具頁的 <style> 不會跟著過去，用 class 的話到那邊就變
   * 成沒有樣式的黑白線稿。 */
  function el(tag, attrs, text) {
    var e = document.createElementNS(SVGNS, tag);
    for (var k in attrs) {
      if (attrs[k] !== null && attrs[k] !== undefined) e.setAttribute(k, attrs[k]);
    }
    if (text !== undefined && text !== null) e.textContent = text;
    return e;
  }

  function clear(node) {
    while (node.firstChild) node.removeChild(node.firstChild);
  }

  /* ---------- 座標軸刻度 ----------
   * 取「好看的」刻度間距（1/2/2.5/5 × 10^n），避免出現 0, 3.7143, 7.4286 這種軸。 */
  function niceStep(rough) {
    if (!(rough > 0)) return 1;
    var mag = Math.pow(10, Math.floor(Math.log(rough) / Math.LN10));
    var norm = rough / mag;
    var step;
    if (norm <= 1) step = 1;
    else if (norm <= 2) step = 2;
    else if (norm <= 2.5) step = 2.5;
    else if (norm <= 5) step = 5;
    else step = 10;
    return step * mag;
  }

  /* 回傳 {min, max, step, ticks[]}。includeZero 為 true 時強制軸含 0
   * （金額類軸不含 0 會嚴重誇大波動幅度；比率類軸則通常不需要）。 */
  function axisScale(values, targetTicks, includeZero) {
    var nums = values.filter(function (v) { return typeof v === 'number' && isFinite(v); });
    if (!nums.length) return { min: 0, max: 1, step: 1, ticks: [0, 1] };
    var lo = Math.min.apply(null, nums);
    var hi = Math.max.apply(null, nums);
    if (includeZero) { lo = Math.min(lo, 0); hi = Math.max(hi, 0); }
    if (lo === hi) { // 全部同值：撐開一個區間，否則除以 0
      var pad = Math.abs(lo) > 0 ? Math.abs(lo) * 0.1 : 1;
      lo -= pad; hi += pad;
    }
    var step = niceStep((hi - lo) / Math.max(1, (targetTicks || 5)));
    var min = Math.floor(lo / step) * step;
    var max = Math.ceil(hi / step) * step;
    var ticks = [];
    // 用乘法而非累加，避免 0.1 這類 step 累積出 0.30000000000000004
    for (var i = 0; min + i * step <= max + step * 1e-9; i++) ticks.push(round(min + i * step, 10));
    return { min: min, max: max, step: step, ticks: ticks };
  }

  function round(v, d) {
    var m = Math.pow(10, d || 0);
    return Math.round(v * m) / m;
  }

  /* 數字顯示：自動決定小數位，並在超過千位時加逗號 */
  function fmt(v, decimals) {
    if (typeof v !== 'number' || !isFinite(v)) return '';
    var d = decimals;
    if (d === undefined || d === null) {
      var a = Math.abs(v);
      d = a >= 100 ? 0 : a >= 10 ? 1 : a >= 1 ? 2 : 3;
      // 本身就是整數就不要補 .00
      if (Math.abs(v - Math.round(v)) < 1e-9) d = 0;
    }
    var s = v.toFixed(d);
    var parts = s.split('.');
    parts[0] = parts[0].replace(/\B(?=(\d{3})+(?!\d))/g, ',');
    return parts.join('.');
  }

  /* ---------- 文字寬度估算 ----------
   * SVG 沒有 text-overflow:ellipsis，也沒辦法在畫之前問瀏覽器「這串字多寬」
   * （getComputedTextLength 要先進 DOM，成本高又會強制 reflow）。這裡用字元
   * 類別粗估：CJK/全形約等於字級，ASCII 約 0.55 字級。誤差幾 px 不影響版面決策。 */
  function charW(code, fontSize) {
    // 0x2e80 之後涵蓋 CJK 部首、標點、假名、漢字、全形英數
    return code > 0x2e80 ? fontSize : fontSize * 0.55;
  }

  function textWidth(s, fontSize) {
    var w = 0;
    s = String(s);
    for (var i = 0; i < s.length; i++) w += charW(s.charCodeAt(i), fontSize);
    return w;
  }

  /* 超過 maxW 就截斷並補省略號；放得下就原樣回傳 */
  function truncateToWidth(s, fontSize, maxW) {
    s = String(s);
    if (textWidth(s, fontSize) <= maxW) return s;
    var budget = maxW - charW(0x2026, fontSize);   // 預留省略號的位置
    var out = '', w = 0;
    for (var i = 0; i < s.length; i++) {
      var cw = charW(s.charCodeAt(i), fontSize);
      if (w + cw > budget) break;
      out += s.charAt(i);
      w += cw;
    }
    return out + '…';
  }

  /* ---------- 逸出 ---------- */
  function escapeHtml(s) {
    return String(s).replace(/[&<>"']/g, function (c) {
      return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c];
    });
  }

  /* ---------- 調色盤 ----------
   * 依報告視覺鐵則挑的：全部是能安全放在純白底上的中深色，彼此色相間隔夠大，
   * 前 4 色在灰階列印下亮度也有差異。 */
  var PALETTE = ['#2563eb', '#dc2626', '#16a34a', '#d97706',
                 '#7c3aed', '#0891b2', '#db2777', '#65a30d'];

  function paletteAt(i) { return PALETTE[i % PALETTE.length]; }

  /* ---------- 匯出：靜態 SVG ---------- */
  /* keepHooks=false（靜態版）會把互動專用的圖元整個拿掉：
   *   [data-hit]    透明熱區 —— 沒有 script 就只是一堆看不見的矩形
   *   [data-reveal] 初始 opacity=0 的輔助圖元（參考線、hover 圓點）—— 永遠不會被顯示
   * 留著不會壞版面，但會讓貼進報告的 SVG 憑空多出一兩倍節點，日後有人要手改時很困惑。 */
  function serializeSvg(svg, keepHooks) {
    var clone = svg.cloneNode(true);
    clone.setAttribute('xmlns', SVGNS);
    clone.removeAttribute('id');
    clone.removeAttribute('class');
    clone.setAttribute('style', 'width:100%;height:auto;display:block;');
    if (!keepHooks) {
      Array.prototype.forEach.call(clone.querySelectorAll('[data-hit],[data-reveal]'), function (n) {
        n.parentNode.removeChild(n);
      });
    }
    return new XMLSerializer().serializeToString(clone);
  }

  function buildStaticEmbed(svg, caption) {
    var xml = serializeSvg(svg, false);
    var cap = caption
      ? '\n  <div style="font-size:12px;color:#666;text-align:center;margin-top:6px;">' +
        escapeHtml(caption) + '</div>'
      : '';
    return '<div style="max-width:900px;margin:0 auto;">\n  ' + xml + cap + '\n</div>';
  }

  /* ---------- 匯出：互動版 ----------
   * 產出 = 同一份靜態 SVG（已經是完整實體，不靠 JS 畫）+ 一段只負責 hover 的
   * vanilla script。因為圖形不是 JS 算出來的，即使被貼在 display:none 的分頁裡
   * 也不會有「量不到寬度所以畫不出來」的問題；最壞情況只是 script 沒跑，圖仍在。 */
  function buildInteractiveEmbed(svg, opts) {
    opts = opts || {};
    // uid 讓同一篇報告可以貼多張互動圖而不互相搶 DOM
    var uid = 'ct' + Math.random().toString(36).slice(2, 9);
    var xml = serializeSvg(svg, true);

    var cap = opts.caption
      ? '\n  <div style="font-size:12px;color:#666;text-align:center;margin-top:6px;">' +
        escapeHtml(opts.caption) + '</div>'
      : '';

    var lines = [];
    lines.push('<div id="' + uid + '" style="max-width:900px;margin:0 auto;position:relative;">');
    lines.push('  ' + xml);
    lines.push('  <div data-tip style="position:absolute;pointer-events:none;opacity:0;transition:opacity .12s;' +
               'background:#111;color:#fff;font-size:12px;line-height:1.5;padding:7px 10px;border-radius:6px;' +
               'white-space:nowrap;z-index:10;box-shadow:0 2px 8px rgba(0,0,0,.18);"></div>');
    lines.push(cap ? cap.replace(/^\n/, '') : '');
    lines.push('</div>');
    lines.push('<script>');
    lines.push('(function(){');
    lines.push('  var root = document.getElementById("' + uid + '");');
    lines.push('  if (!root) return;');
    lines.push('  var tip = root.querySelector("[data-tip]");');
    lines.push('  var hits = root.querySelectorAll("[data-hit]");');
    lines.push('  function show(e, html){');
    lines.push('    tip.innerHTML = html;');
    lines.push('    tip.style.opacity = "1";');
    lines.push('    var r = root.getBoundingClientRect();');
    lines.push('    var x = e.clientX - r.left, y = e.clientY - r.top;');
    lines.push('    var tw = tip.offsetWidth, th = tip.offsetHeight;');
    // 靠右／靠上時翻邊，避免 tooltip 被容器裁掉
    lines.push('    tip.style.left = Math.max(0, Math.min(r.width - tw, x - tw / 2)) + "px";');
    lines.push('    tip.style.top = (y - th - 12 < 0 ? y + 16 : y - th - 12) + "px";');
    lines.push('  }');
    lines.push('  function hide(){ tip.style.opacity = "0"; }');
    lines.push('  Array.prototype.forEach.call(hits, function(h){');
    lines.push('    var html = h.getAttribute("data-tip-html") || "";');
    lines.push('    var dim = h.getAttribute("data-dim");');
    lines.push('    var show_ = h.getAttribute("data-show");');
    lines.push('    h.addEventListener("mousemove", function(e){ show(e, html); });');
    lines.push('    h.addEventListener("mouseenter", function(){');
    // data-show / data-reveal：滑到熱區時顯示對應的輔助圖元（折線圖的垂直參考線、
    // 該期別的資料點圓圈）。這些圖元本身就畫在 SVG 裡、初始 opacity=0，
    // script 只切換 opacity，不做任何幾何計算。
    lines.push('      if (show_) {');
    lines.push('        Array.prototype.forEach.call(');
    lines.push('          root.querySelectorAll(\'[data-reveal="\' + show_ + \'"]\'), function(o){ o.style.opacity = "1"; });');
    lines.push('      }');
    lines.push('      if (!dim) return;');
    lines.push('      Array.prototype.forEach.call(root.querySelectorAll("[data-dim]"), function(o){');
    lines.push('        o.style.opacity = (o.getAttribute("data-dim") === dim) ? "1" : "0.25";');
    lines.push('      });');
    lines.push('    });');
    lines.push('    h.addEventListener("mouseleave", function(){');
    lines.push('      hide();');
    lines.push('      if (show_) {');
    lines.push('        Array.prototype.forEach.call(');
    lines.push('          root.querySelectorAll(\'[data-reveal="\' + show_ + \'"]\'), function(o){ o.style.opacity = "0"; });');
    lines.push('      }');
    lines.push('      Array.prototype.forEach.call(root.querySelectorAll("[data-dim]"), function(o){');
    lines.push('        o.style.opacity = "1";');
    lines.push('      });');
    lines.push('    });');
    lines.push('  });');
    lines.push('})();');
    lines.push('</' + 'script>');
    return lines.filter(function (l) { return l !== ''; }).join('\n');
  }

  /* ---------- 匯出 PNG ---------- */
  function exportPng(svg, filename, done) {
    var clone = svg.cloneNode(true);
    clone.setAttribute('xmlns', SVGNS);
    var vb = (svg.getAttribute('viewBox') || '0 0 900 500').split(/\s+/);
    var w = parseFloat(vb[2]) || 900, h = parseFloat(vb[3]) || 500;
    clone.setAttribute('width', w);
    clone.setAttribute('height', h);
    var xml = new XMLSerializer().serializeToString(clone);
    var url = URL.createObjectURL(new Blob([xml], { type: 'image/svg+xml;charset=utf-8' }));
    var img = new Image();
    img.onload = function () {
      var scale = 2; // 2x 供簡報/列印使用
      var canvas = document.createElement('canvas');
      canvas.width = w * scale; canvas.height = h * scale;
      var ctx = canvas.getContext('2d');
      ctx.fillStyle = '#ffffff';           // 透明底在 PPT/PDF 會變黑，一律填白
      ctx.fillRect(0, 0, canvas.width, canvas.height);
      ctx.drawImage(img, 0, 0, canvas.width, canvas.height);
      URL.revokeObjectURL(url);
      canvas.toBlob(function (blob) {
        var a = document.createElement('a');
        a.href = URL.createObjectURL(blob);
        a.download = filename || 'chart.png';
        a.click();
        setTimeout(function () { URL.revokeObjectURL(a.href); }, 1000);
        if (done) done();
      });
    };
    img.onerror = function () { URL.revokeObjectURL(url); if (done) done('PNG 匯出失敗'); };
    img.src = url;
  }

  /* ---------- 剪貼簿 ----------
   * 非 https / 非 localhost 下 navigator.clipboard 不存在，退回 execCommand。 */
  function copyText(text, done) {
    function fallback() {
      var ta = document.createElement('textarea');
      ta.value = text;
      ta.style.position = 'fixed';
      ta.style.opacity = '0';
      document.body.appendChild(ta);
      ta.select();
      var ok = false;
      try { ok = document.execCommand('copy'); } catch (e) { ok = false; }
      document.body.removeChild(ta);
      if (done) done(ok);
    }
    if (navigator.clipboard && navigator.clipboard.writeText) {
      navigator.clipboard.writeText(text).then(function () { if (done) done(true); }, fallback);
    } else {
      fallback();
    }
  }

  /* ---------- toast ---------- */
  function toast(msg) {
    var t = document.getElementById('toast');
    if (!t) return;
    t.textContent = msg;
    t.classList.add('show');
    clearTimeout(toast._t);
    toast._t = setTimeout(function () { t.classList.remove('show'); }, 1800);
  }

  /* ---------- localStorage 草稿 ----------
   * 工具頁重整不該讓輸入到一半的資料消失。存不進去（無痕模式/配額滿）時靜默略過，
   * 不要因為草稿功能壞掉就讓整個工具報錯。 */
  function saveDraft(key, data) {
    try { localStorage.setItem(key, JSON.stringify(data)); } catch (e) { /* 忽略 */ }
  }
  function loadDraft(key) {
    try {
      var raw = localStorage.getItem(key);
      return raw ? JSON.parse(raw) : null;
    } catch (e) { return null; }
  }

  global.CT = {
    SVGNS: SVGNS,
    el: el,
    clear: clear,
    axisScale: axisScale,
    niceStep: niceStep,
    round: round,
    fmt: fmt,
    escapeHtml: escapeHtml,
    textWidth: textWidth,
    truncateToWidth: truncateToWidth,
    PALETTE: PALETTE,
    paletteAt: paletteAt,
    buildStaticEmbed: buildStaticEmbed,
    buildInteractiveEmbed: buildInteractiveEmbed,
    exportPng: exportPng,
    copyText: copyText,
    toast: toast,
    saveDraft: saveDraft,
    loadDraft: loadDraft
  };
})(window);
