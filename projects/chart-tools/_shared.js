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
      var v = attrs[k];
      if (v === null || v === undefined) continue;
      // 座標算出來常是 255.20099999999996 這種浮點雜訊，直接寫進屬性會讓每段
      // 嵌入代碼平白膨脹一到兩成，貼進報告後也很難讀。SVG 沒有次像素以下的
      // 顯示差別，統一收斂到小數第 2 位。
      if (typeof v === 'number' && isFinite(v)) v = Math.round(v * 100) / 100;
      e.setAttribute(k, v);
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
      // 整數就不要補 .00
      if (Math.abs(v - Math.round(v)) < 1e-9) d = 0;
      // 有小數的就保留一位。早期版本對 >=100 的數字一律取整，結果圖例把使用者
      // 輸入的 107.7 顯示成 108，跟旁邊表格的 $107.7 對不上——圖與表互相矛盾
      // 是最傷報告可信度的錯誤，寧可多印一位小數。
      else if (a >= 1000) d = 0;      // 上千的數字多一位小數沒有資訊量，反而變雜訊
      else if (a >= 10) d = 1;
      else if (a >= 1) d = 2;
      else d = 3;
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
   * 依報告視覺鐵則挑的：全部是能安全放在純白底上的中深色，彼此色相間隔夠大。
   *
   * ⚠️ 這組是**螢幕用**的。它的色相分得開，但灰階亮度擠在一起——實測
   * #dc2626=92 / #2563eb=96 / #7c3aed=98 / #db2777=102，前四色在黑白列印下
   * 幾乎分不出來。報告會被印或匯 PDF，所以另備 PALETTE_PRINT，並提供
   * auditColors() 讓工具在使用者自選顏色時也能檢出這類問題。 */
  var PALETTE = ['#2563eb', '#dc2626', '#16a34a', '#d97706',
                 '#7c3aed', '#0891b2', '#db2777', '#65a30d'];

  /* 列印安全版：把灰階亮度**單調拉開**當作主要區分手段，色相只是輔助。
   * 實測亮度 38 / 65 / 85 / 106 / 125 / 144 / 165 / 187，任兩色至少差 18。
   *
   * 這是硬性取捨：8 色要在灰階下兩兩可辨，亮度就得橫跨約 150 的範圍，最後兩色
   * 因此偏淺。它們當長條/區塊填色沒問題，但當細折線會偏淡——超過 6 個系列時
   * 建議改用長條圖，或拆成兩張圖。
   *
   * 已知限制：**用到第 7、8 色時**，C3（洋紅 #b1408a）與 C6（淺綠 #8fbf5a）在色覺
   * 缺陷模擬下會接近。試過把 C6 換成偏藍的 #9ab8d8 可以解掉，但它亮度衝到 179、
   * 與 C7 只差 8，等於拿灰階去換色覺——對「列印安全」這個目的是划不來的交換，
   * 所以維持現狀。前 6 色在灰階與三種色覺模擬下皆零問題；真的需要 7 個以上系列時，
   * auditColors() 會即時報警，該做的是拆圖而不是硬塞。 */
  var PALETTE_PRINT = ['#182456', '#8f1f1f', '#0f6f8a', '#b1408a',
                       '#b3760f', '#57a0cf', '#8fbf5a', '#d9b48f'];

  function paletteAt(i) { return PALETTE[i % PALETTE.length]; }
  function palettePrintAt(i) { return PALETTE_PRINT[i % PALETTE_PRINT.length]; }

  /* ---------- 色彩工具 ---------- */
  function hexToRgb(hex) {
    var m = /^#?([a-f\d]{2})([a-f\d]{2})([a-f\d]{2})$/i.exec(String(hex).trim());
    if (m) return [parseInt(m[1], 16), parseInt(m[2], 16), parseInt(m[3], 16)];
    var r = /rgba?\((\d+)[,\s]+(\d+)[,\s]+(\d+)/.exec(String(hex));
    return r ? [+r[1], +r[2], +r[3]] : [0, 0, 0];
  }

  /* 感知亮度（ITU-R BT.601），與熱力圖決定黑白字用的是同一條公式 */
  function luminance(color) {
    var c = hexToRgb(color);
    return 0.299 * c[0] + 0.587 * c[1] + 0.114 * c[2];
  }

  /* 色覺缺陷模擬（Brettel/Viénot 線性近似）。
   * 用途只是「這兩色會不會撞在一起」的粗篩，不是色彩科學等級的精確模擬。 */
  var CVD_MATRIX = {
    deuteranopia: [[0.625, 0.375, 0], [0.7, 0.3, 0], [0, 0.3, 0.7]],   // 綠盲，最常見
    protanopia:   [[0.567, 0.433, 0], [0.558, 0.442, 0], [0, 0.242, 0.758]], // 紅盲
    tritanopia:   [[0.95, 0.05, 0], [0, 0.433, 0.567], [0, 0.475, 0.525]]    // 藍盲，罕見
  };

  function simulateCVD(color, kind) {
    var m = CVD_MATRIX[kind];
    if (!m) return hexToRgb(color);
    var c = hexToRgb(color);
    return m.map(function (row) {
      return Math.max(0, Math.min(255, Math.round(row[0] * c[0] + row[1] * c[1] + row[2] * c[2])));
    });
  }

  function rgbDist(a, b) {
    var dr = a[0] - b[0], dg = a[1] - b[1], db = a[2] - b[2];
    return Math.sqrt(dr * dr + dg * dg + db * db);
  }

  /* 檢查一組「名稱＋顏色」在灰階與三種色覺缺陷下是否兩兩可辨。
   * 回傳問題清單；沒問題就是空陣列。
   * 門檻：灰階亮度差 < 18（0-255 尺度）或模擬後 RGB 距離 < 55 視為難以區分。
   * 這兩個數字是拿現行 PALETTE 實測校出來的——調高會把可接受的組合也一起報警。 */
  function auditColors(items) {
    var issues = [];
    for (var i = 0; i < items.length; i++) {
      for (var j = i + 1; j < items.length; j++) {
        var a = items[i], b = items[j];
        if (!a.color || !b.color) continue;
        var dl = Math.abs(luminance(a.color) - luminance(b.color));
        if (dl < 18) {
          issues.push({ a: a.name, b: b.name, kind: '灰階', detail: '亮度差僅 ' + Math.round(dl) });
        }
        for (var k in CVD_MATRIX) {
          var d = rgbDist(simulateCVD(a.color, k), simulateCVD(b.color, k));
          if (d < 55) {
            issues.push({
              a: a.name, b: b.name,
              kind: k === 'deuteranopia' ? '綠色盲' : k === 'protanopia' ? '紅色盲' : '藍色盲',
              detail: '模擬後色差僅 ' + Math.round(d)
            });
          }
        }
      }
    }
    return issues;
  }

  /* ---------- 跨圖配色記憶 ----------
   * 一篇報告會有 6 張以上的圖，「資料中心」在折線圖是藍、在長條圖變紅，是因為
   * 每個工具各自從調色盤第 0 個開始排。這裡用一份共用登錄表：同名系列在任何
   * 工具裡都拿到同一個顏色。
   *
   * 刻意做成單一份全域記憶而不是「每份報告一組」——後者需要使用者維護報告代號，
   * 那就變成另一層要管的東西了。換報告時按「清除配色記憶」即可。 */
  var REG_KEY = 'chart-tools:colors';

  function colorRegistry() {
    try { return JSON.parse(localStorage.getItem(REG_KEY) || '{}'); } catch (e) { return {}; }
  }

  function saveRegistry(reg) {
    try { localStorage.setItem(REG_KEY, JSON.stringify(reg)); } catch (e) { /* 無痕/配額滿：略過 */ }
  }

  /* 取得某個具名系列該用的顏色。沒登錄過就配一個「還沒被用掉」的調色盤色並記下來。
   * usePrint=true 時改用列印安全調色盤。 */
  function colorFor(name, fallbackIndex, usePrint) {
    name = String(name || '').trim();
    var pal = usePrint ? PALETTE_PRINT : PALETTE;
    if (!name) return pal[fallbackIndex % pal.length];
    var reg = colorRegistry();
    if (reg[name]) return reg[name];
    var used = {};
    for (var k in reg) used[reg[k].toLowerCase()] = true;
    var pick = null;
    for (var i = 0; i < pal.length; i++) {
      if (!used[pal[i].toLowerCase()]) { pick = pal[i]; break; }
    }
    if (!pick) pick = pal[fallbackIndex % pal.length];   // 調色盤用完就開始重複
    reg[name] = pick;
    saveRegistry(reg);
    return pick;
  }

  /* 使用者在工具裡手動改色時呼叫，讓其他圖跟著一致 */
  function rememberColor(name, color) {
    name = String(name || '').trim();
    if (!name || !color) return;
    var reg = colorRegistry();
    reg[name] = color;
    saveRegistry(reg);
  }

  function forgetColors() {
    try { localStorage.removeItem(REG_KEY); } catch (e) { /* 略過 */ }
  }

  /* ---------- 表格解析 ----------
   * 讓每個工具都能直接從 Excel / Google Sheets / 10-Q 表格整塊貼上，
   * 不必一格一格重打。原本只有熱力圖有，現在集中在這裡給所有工具共用。 */

  /* 數字清洗。財報表格的慣例比一般 CSV 髒得多：
   *   1,234      千分位
   *   41.2%      百分比（去掉 % 保留數值，單位由使用者在工具裡填）
   *   $1,234     貨幣符號
   *   (1,234)    ← 括號代表負數，這是 10-Q/10-K 的標準寫法
   *   —, -, N/A  無資料
   * 括號負數若沒處理，整份損益表的費用項會全部變成正的，這是會直接畫錯圖的。 */
  function parseNumber(raw) {
    if (raw === null || raw === undefined) return null;
    var s = String(raw).trim();
    if (s === '') return null;
    if (/^(-|—|–|n\/?a|na|null)$/i.test(s)) return null;
    var neg = false;
    if (/^\((.*)\)$/.test(s)) { neg = true; s = s.replace(/^\(|\)$/g, ''); }
    s = s.replace(/[,%$＄¥€£\s ]/g, '');
    if (s === '' || s === '-') return null;
    var n = parseFloat(s);
    if (!isFinite(n)) return null;
    return neg ? -n : n;
  }

  /* 把貼上的文字切成二維陣列。優先用 Tab（Excel 複製的預設），沒有 Tab 才用逗號。
   * 逗號模式支援雙引號包住的欄位，否則 "台積電, 聯電" 這種會被切錯。 */
  function splitRow(line) {
    if (line.indexOf('\t') >= 0) return line.split('\t');
    var out = [], cur = '', inQ = false;
    for (var i = 0; i < line.length; i++) {
      var ch = line.charAt(i);
      if (ch === '"') {
        if (inQ && line.charAt(i + 1) === '"') { cur += '"'; i++; }
        else inQ = !inQ;
      } else if (ch === ',' && !inQ) { out.push(cur); cur = ''; }
      else cur += ch;
    }
    out.push(cur);
    return out;
  }

  /* 回傳 { header:[...], rows:[{ label, cells:[原字串], values:[數字或null] }] }
   * hasHeader=false 時 header 為 null，所有列都當資料列。 */
  function parseTable(text, hasHeader) {
    var lines = String(text || '').split(/\r?\n/).filter(function (l) { return l.trim() !== ''; });
    if (!lines.length) return null;
    var grid = lines.map(splitRow).map(function (r) {
      return r.map(function (c) { return String(c).trim(); });
    });
    var header = null;
    if (hasHeader !== false) { header = grid.shift(); }
    if (!grid.length) return null;
    var rows = grid.map(function (r) {
      return {
        label: r[0] || '',
        cells: r.slice(1),
        values: r.slice(1).map(parseNumber)
      };
    });
    return { header: header, rows: rows };
  }

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

  /* ---------- 共用 UI：貼上表格 ----------
   * 六個工具都要這塊，所以集中在這裡產生 DOM 並回呼，避免同樣的面板手寫六次
   * （那正是各工具解析行為會慢慢分歧的原因）。
   * opts: { placeholder, hint, headerDefault, onApply(parsed, hasHeader) } */
  function mountPastePanel(container, opts) {
    opts = opts || {};
    container.innerHTML =
      '<textarea data-paste spellcheck="false" style="height:120px;" placeholder="' +
        escapeHtml(opts.placeholder || '把表格整塊貼進來') + '"></textarea>' +
      '<div class="check-row"><input type="checkbox" data-hdr' +
        (opts.headerDefault === false ? '' : ' checked') +
        '><label>第一列是標題列</label></div>' +
      '<div class="btn-row"><button class="btn ghost small" data-apply>套用表格</button></div>' +
      '<div data-err style="display:none;font-size:11.5px;color:#c33;margin-top:6px;"></div>' +
      (opts.hint ? '<div class="hint-text">' + opts.hint + '</div>' : '');

    var ta = container.querySelector('[data-paste]');
    var hdr = container.querySelector('[data-hdr]');
    var err = container.querySelector('[data-err]');

    function fail(msg) { err.textContent = '⚠ ' + msg; err.style.display = ''; }

    container.querySelector('[data-apply]').addEventListener('click', function () {
      err.style.display = 'none';
      var parsed = parseTable(ta.value, hdr.checked);
      if (!parsed || !parsed.rows.length) { fail('看不出表格內容，至少要有一列資料'); return; }
      try {
        var r = opts.onApply(parsed, hdr.checked);
        if (r && r.error) { fail(r.error); return; }
        toast('已套用 ' + parsed.rows.length + ' 列');
      } catch (e) {
        fail('套用失敗：' + e.message);
      }
    });

    return {
      setText: function (t) { ta.value = t; },
      getText: function () { return ta.value; }
    };
  }

  /* ---------- 共用 UI：配色一致性與可辨識性 ----------
   * opts: { getItems() -> [{name,color}], onApplyPrint(), onReset() } */
  function mountColorPanel(container, opts) {
    opts = opts || {};
    container.innerHTML =
      '<div data-audit style="font-size:11.5px;line-height:1.7;"></div>' +
      '<div class="btn-row">' +
        '<button class="btn ghost small" data-print>套用列印安全配色</button>' +
        '<button class="btn ghost small" data-forget>清除配色記憶</button>' +
      '</div>' +
      '<div class="hint-text">同名系列在所有工具裡共用同一個顏色，換一篇報告時按「清除配色記憶」重新開始。</div>';

    var box = container.querySelector('[data-audit]');

    container.querySelector('[data-print]').addEventListener('click', function () {
      if (opts.onApplyPrint) opts.onApplyPrint();
      toast('已套用列印安全配色');
    });
    container.querySelector('[data-forget]').addEventListener('click', function () {
      forgetColors();
      if (opts.onReset) opts.onReset();
      toast('已清除配色記憶');
    });

    function refresh() {
      var items = (opts.getItems ? opts.getItems() : []).filter(function (x) { return x && x.color; });
      if (items.length < 2) { box.innerHTML = '<span style="color:#6b6b6b;">系列不足兩個，無須檢查。</span>'; return; }
      var issues = auditColors(items);
      if (!issues.length) {
        box.innerHTML = '<span style="color:#16a34a;">✓ 灰階與三種色覺模擬下皆可辨識</span>';
        return;
      }
      // 同一組系列可能同時在灰階與色覺模擬下撞色，只報一次最嚴重的
      var seen = {}, lines = [];
      issues.forEach(function (it) {
        var key = it.a + '|' + it.b;
        if (seen[key]) return;
        seen[key] = true;
        lines.push('<div style="color:#b45309;">⚠ <b>' + escapeHtml(it.a) + '</b> 與 <b>' +
          escapeHtml(it.b) + '</b> 在' + it.kind + '下難以區分（' + it.detail + '）</div>');
      });
      box.innerHTML = lines.join('');
    }

    refresh();
    return { refresh: refresh };
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
    PALETTE_PRINT: PALETTE_PRINT,
    paletteAt: paletteAt,
    palettePrintAt: palettePrintAt,
    luminance: luminance,
    simulateCVD: simulateCVD,
    auditColors: auditColors,
    colorFor: colorFor,
    rememberColor: rememberColor,
    forgetColors: forgetColors,
    colorRegistry: colorRegistry,
    parseNumber: parseNumber,
    parseTable: parseTable,
    mountPastePanel: mountPastePanel,
    mountColorPanel: mountColorPanel,
    buildStaticEmbed: buildStaticEmbed,
    buildInteractiveEmbed: buildInteractiveEmbed,
    exportPng: exportPng,
    copyText: copyText,
    toast: toast,
    saveDraft: saveDraft,
    loadDraft: loadDraft
  };
})(window);
