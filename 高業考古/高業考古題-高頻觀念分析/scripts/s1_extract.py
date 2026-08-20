"""Extract every question from the 高業 past-paper PDFs into a normalized JSON corpus."""
import fitz, glob, os, re, json, unicodedata, sys

# repo-relative: scripts/ -> 高業考古題-高頻觀念分析/ -> 高業考古/
BASE = os.path.abspath(os.path.join(os.path.dirname(os.path.abspath(__file__)), '..', '..'))
OUT = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'out')
os.makedirs(OUT, exist_ok=True)

SUBJ_PAT = re.compile(r'專業科目[：:]\s*([^\n]{0,60})')
HEAD_PAT = re.compile(r'(\d{3})\s*年\s*第\s*(\d)\s*次')


def subject_of(label):
    if '投資學' in label:
        return '投資學'
    if '財務分析' in label:
        return '財務分析'
    if '法規' in label:
        return '法規'
    return None


def split_questions(block):
    """Walk sequentially for question markers 1..50 so stray numbers don't split."""
    spans = []
    pos = 0
    for n in range(1, 51):
        # marker at line start (or after whitespace) followed by . or 、
        pat = re.compile(r'(?:(?<=\n)|^)\s*' + str(n) + r'\s*[.、．]\s*')
        m = pat.search(block, pos)
        if not m:
            # fall back: allow mid-line marker
            pat2 = re.compile(r'(?<![\d.])' + str(n) + r'\s*[.、．]\s*(?=\S)')
            m = pat2.search(block, pos)
            if not m:
                spans.append((n, None, None))
                continue
        spans.append((n, m.end(), None))
        pos = m.end()
    out = []
    for i, (n, s, _) in enumerate(spans):
        if s is None:
            continue
        # end = start of next found question
        e = len(block)
        for n2, s2, _ in spans[i + 1:]:
            if s2 is not None:
                # back up over the marker itself
                e = s2 - len(str(n2)) - 2
                break
        out.append((n, block[s:max(s, e)]))
    return out


def norm_stem(raw):
    """Question stem = text before the first (A) option, aggressively normalized."""
    t = unicodedata.normalize('NFKC', raw)
    # cut at first option marker
    m = re.search(r'[（(]\s*[AＡ]\s*[)）]', t)
    if m:
        t = t[:m.start()]
    t = re.sub(r'\s+', '', t)
    # unify punctuation & common variants
    t = t.replace('，', ',').replace('。', '.').replace('：', ':').replace('；', ';')
    t = t.replace('？', '?').replace('（', '(').replace('）', ')').replace('、', ',')
    t = t.replace('「', '"').replace('」', '"').replace('『', '"').replace('』', '"')
    t = t.replace('─', '-').replace('—', '-').replace('－', '-').replace('～', '~')
    t = re.sub(r'[,.\s"\'()\[\]{}:;?!~*#_\-]', '', t)
    return t


def split_body(raw):
    """Return (題幹, [(letter, 選項文字), ...]) with layout whitespace collapsed.
    Needed by s6 to reprint a real, readable question — `stem` is normalized to
    death and `raw` is only for eyeballing."""
    t = re.sub(r'[ \t]*\n[ \t]*', '\n', raw).strip()
    t = re.sub(r'[ \t]{2,}', ' ', t)
    # Scan for A, then the first B after it, etc. A naive split breaks on
    # 「(D)選項(A)(B)(C)皆非」, where option D's own text contains (A)(B)(C).
    marks = [(m.start(), m.end(), unicodedata.normalize('NFKC', m.group(1)))
             for m in re.finditer(r'[（(]\s*([A-DＡ-Ｄ])\s*[)）]', t)]
    picked = []
    want = 'ABCD'
    pos = 0
    for letter in want:
        hit = next((mk for mk in marks if mk[0] >= pos and mk[2] == letter), None)
        if not hit:
            return re.sub(r'\s+', ' ', t).strip(), None
        picked.append(hit)
        pos = hit[1]
    head = re.sub(r'\s+', ' ', t[:picked[0][0]]).strip()
    opts = []
    for i, (s, e, letter) in enumerate(picked):
        end = picked[i + 1][0] if i + 1 < len(picked) else len(t)
        opts.append((letter, re.sub(r'\s+', ' ', t[e:end]).strip()))
    if not head or any(not b for _, b in opts):
        return head, None
    return head, opts


def parse_answers(path):
    """Answer sheets hold three consecutive 1..50 tables, in paper order
    (投資學 → 財務分析 → 法規). Key them as 1..150 by counting group restarts."""
    d = fitz.open(path)
    t = ''.join(pg.get_text() for pg in d)
    d.close()
    t = unicodedata.normalize('NFKC', t)
    ans = {}
    group = -1
    seen_in_group = set()
    for m in re.finditer(r'(?<!\d)(\d{1,2})\s*\n\s*((?:[A-D](?:\.[A-D])*)|送分|一律給分)', t):
        q = int(m.group(1))
        if not 1 <= q <= 50:
            continue
        if q in seen_in_group:      # table restarted -> next subject
            group += 1
            seen_in_group = set()
        if group < 0:
            group = 0
        seen_in_group.add(q)
        ans[group * 50 + q] = m.group(2)
    return ans


def main():
    records = []
    files = []
    for y in range(105, 116):
        files += sorted(glob.glob(f'{BASE}/{y}/*.pdf'))
    papers = 0
    for p in files:
        name = os.path.basename(p)
        if re.fullmatch(r'\d{5}a\.pdf', name):
            continue
        if '說明' in name:
            continue
        d = fitz.open(p)
        text = ''.join(pg.get_text() for pg in d)
        d.close()
        text = unicodedata.normalize('NFC', text)
        hm = HEAD_PAT.search(text)
        if not hm:
            print('NO HEADER', name, file=sys.stderr)
            continue
        year, sess = int(hm.group(1)), int(hm.group(2))

        # answer key: combined papers have a sibling <stem>a.pdf
        akey = {}
        cand = os.path.join(os.path.dirname(p), re.sub(r'\.pdf$', 'a.pdf', name))
        if os.path.exists(cand):
            akey = parse_answers(cand)

        marks = [(m.start(), subject_of(m.group(1))) for m in SUBJ_PAT.finditer(text)]
        marks = [(s, sub) for s, sub in marks if sub]
        if not marks:
            print('NO SUBJECT', name, file=sys.stderr)
            continue
        for i, (s, sub) in enumerate(marks):
            e = marks[i + 1][0] if i + 1 < len(marks) else len(text)
            block = text[s:e]
            qs = split_questions(block)
            papers += 1
            # answer offset: subject i occupies answers (i*50+1 .. i*50+50) in combined key
            off = i * 50 if len(marks) > 1 else 0
            for n, raw in qs:
                stem = norm_stem(raw)
                if len(stem) < 3:   # 3, not 6: 「失業率是指」 is a real (and repeated) stem
                    continue
                head, opts = split_body(raw)
                records.append({
                    'year': year, 'sess': sess, 'subject': sub, 'qno': n,
                    'exam': f'{year}-{sess}',
                    'stem': stem,
                    'raw': re.sub(r'\s+', ' ', raw).strip()[:400],
                    'head': head,          # printable question text
                    'opts': opts,          # [(A,…),(B,…),(C,…),(D,…)] or None
                    'ans': akey.get(off + n, ''),
                    'src': name,
                })
    with open(os.path.join(OUT, 'corpus.json'), 'w', encoding='utf-8') as f:
        json.dump(records, f, ensure_ascii=False)

    # report
    from collections import Counter
    c = Counter((r['year'], r['sess'], r['subject']) for r in records)
    lines = [f'papers(subject-blocks)={papers}  questions={len(records)}']
    exams = sorted({(r['year'], r['sess']) for r in records})
    lines.append(f'distinct exams={len(exams)}')
    bad = [(k, v) for k, v in sorted(c.items()) if v != 50]
    lines.append(f'blocks with !=50 questions: {len(bad)}')
    for k, v in bad:
        lines.append(f'  {k} -> {v}')
    withans = sum(1 for r in records if r['ans'])
    lines.append(f'questions with answer key = {withans}')
    open(os.path.join(OUT, 'parse_report.txt'), 'w', encoding='utf-8').write('\n'.join(lines))
    print('\n'.join(lines[:4]))


if __name__ == '__main__':
    main()
