// Scans the deployed research-report mirror (globe-invest/app/research/*.html) for
// .qa-item blocks (the standard Earnings Call analysis framework's Q&A tab markup —
// see CLAUDE.md "Earnings Call 分析框架") and builds a quiz question bank for the
// Earnings Quiz game: globe-invest/app/earnings-quiz/questions.json.
//
// Not part of any CI/deploy step — rerun manually after publishing new reports:
//   node scripts/build-earnings-quiz-bank.js
//
// One quiz question per <div class="qa-item">: the full Q/A exchange (all q-block
// pairs inside that item, in order) is the prompt; all flag-box texts inside that
// same item (joined) are the "correct" logic-gap reading. Decoys are 3 flag-box
// texts sampled from other items (preferring a different ticker) so wrong answers
// read as plausibly-formatted but topically mismatched.

const fs = require('fs');
const path = require('path');

const REPO_ROOT = path.join(__dirname, '..');
const RESEARCH_DIR = path.join(REPO_ROOT, 'globe-invest', 'app', 'research');
// Written to the dev-source folder (source of truth) — the manual mirror-sync step for
// Earnings Quiz copies index.html and questions.json together into
// globe-invest/app/earnings-quiz/, same as every other hand-synced mirror pair in this repo.
const OUT_DIR = path.join(REPO_ROOT, 'app', 'EarningsQuiz');
const OUT_FILE = path.join(OUT_DIR, 'questions.json');

const FILENAME_RE = /^([A-Z0-9]+)_(.+)_(Analysis|Financials)\.html$/;

// Matches zero or more HTML attributes after a tag name, tolerating attribute
// VALUES that contain a literal '>' (the bilingual reports' data-en="..." often
// embeds raw <em>/<strong> tags inside the quoted string) — a plain [^>]* skip
// breaks the moment it hits that inner '>', truncating the match early.
const ATTRS = '(?:\\s+[a-zA-Z_:][-a-zA-Z0-9_:.]*(?:="[^"]*"|=\'[^\']*\'|=[^\\s>]+)?)*';

function divOpenRe(classAttr) {
  // classAttr e.g. 'class="qa-analyst"' — matched literally, then any further
  // attributes (data-en, style, onclick, ...) are consumed before the closing '>'.
  return new RegExp('<div ' + classAttr + ATTRS + '\\s*>', 'g');
}

function decodeEntities(s) {
  return s
    .replace(/&mdash;/g, '—')
    .replace(/&ldquo;/g, '「')
    .replace(/&rdquo;/g, '」')
    .replace(/&times;/g, '×')
    .replace(/&middot;/g, '·')
    .replace(/&amp;/g, '&')
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"')
    .replace(/&#39;/g, "'")
    .replace(/&nbsp;/g, ' ');
}

function stripTags(html) {
  return decodeEntities(
    String(html || '')
      .replace(/<br\s*\/?>/gi, '\n')
      .replace(/<\/?(b|strong|i|em|span)[^>]*>/gi, '')
      .replace(/<[^>]+>/g, '')
  ).trim();
}

// Walks <div ...>/</div> tokens with a depth counter so nested divs inside a
// qa-item (q-block, q-text, orig-box, ...) don't confuse the block boundary —
// a plain non-recursive regex can't find the matching close tag otherwise.
// Uses the same attribute-aware tag matcher as divOpenRe so a stray '>' inside
// a data-en="..." value doesn't misalign the stack for the rest of the file.
function extractTopLevelBlocks(html, className) {
  const tagRe = new RegExp('<div\\b' + ATTRS + '\\s*>|<\\/div>', 'gi');
  const classRe = new RegExp('class="[^"]*\\b' + className + '\\b[^"]*"');
  const stack = [];
  const blocks = [];
  let m;
  while ((m = tagRe.exec(html))) {
    const isClose = m[0] === '</div>';
    if (!isClose) {
      stack.push({ start: m.index, isMatch: classRe.test(m[0]) });
    } else {
      const entry = stack.pop();
      if (!entry) continue;
      if (entry.isMatch && !stack.some((s) => s.isMatch)) {
        blocks.push(html.slice(entry.start, m.index + m[0].length));
      }
    }
  }
  return blocks;
}

function matchAll(re, str) {
  const out = [];
  let m;
  re.lastIndex = 0;
  while ((m = re.exec(str))) out.push(m);
  return out;
}

function parseQaItem(block, ticker, period, sourceUrl) {
  const analystRe = new RegExp(divOpenRe('class="qa-analyst"').source + '([\\s\\S]*?)<\\/div>');
  const topicRe = new RegExp(divOpenRe('class="qa-topic"').source + '([\\s\\S]*?)<\\/div>');
  const analystM = block.match(analystRe);
  const topicM = block.match(topicRe);

  const qOpen = divOpenRe('class="q-label q"').source;
  const aOpen = divOpenRe('class="q-label a"').source;
  const textOpen = divOpenRe('class="q-text"').source;

  // q and a use different literal class strings so they're matched in two
  // passes, then merged back into document order (by match offset) for a
  // readable transcript.
  const withOffsets = [];
  matchAll(new RegExp(qOpen + '([\\s\\S]*?)<\\/div>\\s*' + textOpen + '([\\s\\S]*?)<\\/div>', 'g'), block)
    .forEach((m) => withOffsets.push({ offset: m.index, kind: 'q', text: stripTags(m[2]) }));
  matchAll(new RegExp(aOpen + '([\\s\\S]*?)<\\/div>\\s*' + textOpen + '([\\s\\S]*?)<\\/div>', 'g'), block)
    .forEach((m) => withOffsets.push({ offset: m.index, kind: 'a', text: stripTags(m[2]) }));
  if (!withOffsets.length) return null;
  withOffsets.sort((a, b) => a.offset - b.offset);
  const exchange = withOffsets.map((p) => (p.kind === 'q' ? '問：' : '答：') + p.text).join('\n\n');

  const flagOpen = divOpenRe('class="q-label flag"').source;
  const flagTexts = matchAll(new RegExp(flagOpen + '([\\s\\S]*?)<\\/div>\\s*' + textOpen + '([\\s\\S]*?)<\\/div>', 'g'), block)
    .map((m) => stripTags(m[2]));
  if (!flagTexts.length) return null;

  const origOpen = divOpenRe('class="orig-text"').source;
  const origM = block.match(new RegExp(origOpen + '([\\s\\S]*?)<\\/div>'));

  return {
    ticker,
    period,
    analyst: analystM ? stripTags(analystM[1]) : '',
    topic: topicM ? stripTags(topicM[1]) : '',
    exchange,
    flagText: flagTexts.join('\n\n'),
    origText: origM ? stripTags(origM[1]) : '',
    sourceUrl,
  };
}

function pickDecoys(pool, selfIndex, count) {
  const self = pool[selfIndex];
  const differentTicker = [];
  const rest = [];
  pool.forEach((item, i) => {
    if (i === selfIndex) return;
    if (item.ticker !== self.ticker) differentTicker.push(i);
    else rest.push(i);
  });
  const shuffled = differentTicker.concat(rest).sort(() => Math.random() - 0.5);
  return shuffled.slice(0, count).map((i) => pool[i].flagText);
}

function main() {
  if (!fs.existsSync(RESEARCH_DIR)) {
    console.error('Research mirror directory not found: ' + RESEARCH_DIR);
    process.exit(1);
  }
  const files = fs.readdirSync(RESEARCH_DIR).filter((f) => f.endsWith('.html'));

  const items = [];
  let skippedItems = 0;
  const skipReasons = {};
  const skipDetail = [];

  files.forEach((fn) => {
    const filePath = path.join(RESEARCH_DIR, fn);
    const html = fs.readFileSync(filePath, 'utf8');
    const nameM = fn.match(FILENAME_RE);
    const ticker = nameM ? nameM[1] : fn.split('_')[0];
    const period = nameM ? nameM[2] : '';
    const sourceUrl = '/research/' + fn;

    const qaBlocks = extractTopLevelBlocks(html, 'qa-item');
    qaBlocks.forEach((block, idx) => {
      const parsed = parseQaItem(block, ticker, period, sourceUrl);
      if (!parsed) {
        skippedItems++;
        const reason = /class="q-label flag"/.test(block) ? 'no q-block pairs matched' : 'no flag-box found';
        skipReasons[reason] = (skipReasons[reason] || 0) + 1;
        skipDetail.push(fn + ' #' + idx);
        return;
      }
      items.push(parsed);
    });
  });

  if (!items.length) {
    console.error('No qa-items extracted — check RESEARCH_DIR contents and markup assumptions.');
    process.exit(1);
  }

  const questions = items.map((item, i) => {
    const decoys = pickDecoys(items, i, 3);
    const options = [item.flagText, ...decoys];
    return {
      id: 'q' + (i + 1),
      ticker: item.ticker,
      period: item.period,
      exchange: item.exchange,
      origText: item.origText,
      sourceUrl: item.sourceUrl,
      correctAnswer: item.flagText,
      options,
    };
  });

  fs.mkdirSync(OUT_DIR, { recursive: true });
  fs.writeFileSync(OUT_FILE, JSON.stringify(questions, null, 2), 'utf8');

  console.log('Scanned ' + files.length + ' files.');
  console.log('Extracted ' + items.length + ' questions -> ' + path.relative(REPO_ROOT, OUT_FILE));
  if (skippedItems) {
    console.log('Skipped ' + skippedItems + ' qa-item block(s):');
    Object.entries(skipReasons).forEach(([reason, n]) => console.log('  - ' + n + 'x: ' + reason));
    console.log('  Files/indices: ' + skipDetail.slice(0, 20).join(', ') + (skipDetail.length > 20 ? ' ...' : ''));
  } else {
    console.log('Skipped 0 qa-item blocks.');
  }
}

main();
