"""Three option-form heuristics, tested against the answer cards:
  (2) 最長選項法則  — is the longest option the answer?
  (3) 絕對字詞 vs 緩和字詞 — do hedged options win and absolute ones lose?
  (4) 數字題中間值法則 — for all-numeric options, do the middle values win?

Sample: the 2,846 questions with a parsed A–D option set and a single-letter answer.
Baseline for every hit rate below is 25% (blind guessing).
"""
import json, os, re, io, math, unicodedata
from collections import Counter, defaultdict

HERE = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'out')
os.makedirs(HERE, exist_ok=True)
SUBJECTS = ['投資學', '財務分析', '法規']

# ---- shared: the all-of-the-above detector from s7, to exclude those questions ----
REFERS = re.compile(r'以上|選項|[（(]\s*[A-D]\s*[)）]')
QUANT = re.compile(r'(?:皆|均|都)\s*(?:非|錯誤|錯|不正確|不對|不是|不可|是|對|正確|可能|可|能|得|屬|有)')


def is_allofabove(text):
    return bool(REFERS.search(text) and QUANT.search(text))


# ---- (3) word lists -------------------------------------------------------
# 「僅」「不得」 deliberately NOT in the absolute list: 「僅甲、乙」 is combination-question
# scaffolding and 「不得超過…」 is ordinary statutory language, not an absolute qualifier.
ABSOLUTE = ['一定', '必定', '必然', '絕對', '絕不', '永遠', '完全', '毫無',
            '全部', '所有', '任何', '只有', '只能', '唯一', '無論', '必須']
HEDGE = ['可能', '或許', '通常', '一般而言', '不一定', '視情況', '視需要',
         '無從得知', '難以確定', '不確定', '原則上', '多半', '未必', '不確知']
ABS_RE = re.compile('|'.join(ABSOLUTE))
HEDGE_RE = re.compile('|'.join(HEDGE))

# ---- (4) numeric parsing --------------------------------------------------
CN = {'零': 0, '一': 1, '二': 2, '兩': 2, '三': 3, '四': 4, '五': 5,
      '六': 6, '七': 7, '八': 8, '九': 9}
UNIT_MAP = [('億', 10 ** 8), ('萬', 10 ** 4), ('千', 10 ** 3), ('百', 10 ** 2), ('十', 10)]


def cn_number(s):
    """Parse a pure Chinese numeral like 三十, 一百二十, 二分之一 -> float (None if not one)."""
    if not s or not re.fullmatch(r'[零一二兩三四五六七八九十百千萬億]+', s):
        return None
    total, section, last = 0, 0, 0
    for ch in s:
        if ch in CN:
            last = CN[ch]
        else:
            for u, v in UNIT_MAP:
                if ch == u:
                    if v >= 10 ** 4:
                        total += (section + (last or 0)) * v
                        section, last = 0, 0
                    else:
                        section += (last if last else 1) * v
                        last = 0
                    break
    return float(total + section + last)


def one_number(text):
    """Return (value, unit_signature) if the option contains exactly one magnitude."""
    t = unicodedata.normalize('NFKC', text).replace(',', '').strip()
    ar = re.findall(r'\d+(?:\.\d+)?', t)
    cn = re.findall(r'[零一二兩三四五六七八九十百千萬億]+', t)
    # a Chinese "萬/億" immediately after an Arabic number is a multiplier, not a separate number
    if len(ar) == 1 and not [c for c in cn if c not in ('萬', '億', '千', '百', '十')]:
        v = float(ar[0])
        m = re.search(r'\d+(?:\.\d+)?\s*(億|萬|千|百)', t)
        if m:
            v *= {'億': 10 ** 8, '萬': 10 ** 4, '千': 10 ** 3, '百': 10 ** 2}[m.group(1)]
        unit = re.sub(r'\d+(?:\.\d+)?|[億萬千百]', '', t)
    elif not ar and len(cn) == 1:
        v = cn_number(cn[0])
        if v is None:
            return None
        unit = re.sub(r'[零一二兩三四五六七八九十百千萬億]+', '', t)
    else:
        return None
    return v, re.sub(r'\s+', '', unit)


def wilson(k, n):
    if n == 0:
        return (0.0, 0.0)
    z = 1.96
    p = k / n
    d = 1 + z * z / n
    c = (p + z * z / (2 * n)) / d
    m = z * math.sqrt(p * (1 - p) / n + z * z / (4 * n * n)) / d
    return (max(0, c - m) * 100, min(1, c + m) * 100)


def rate(k, n):
    lo, hi = wilson(k, n)
    return f'{k:>4}/{n:<5} = {k/n*100:>5.1f}%  CI[{lo:>4.0f}%,{hi:>4.0f}%]' if n else '（無樣本）'


def main():
    recs = json.load(open(os.path.join(HERE, 'corpus.json'), encoding='utf-8'))
    pool = [q for q in recs if q['opts'] and len(q['ans']) == 1 and q['ans'] in 'ABCD']
    rep = [f'樣本：{len(pool)} 題（有 A–D 選項且有單一正解）。所有命中率的對照基準是亂猜 25%。']
    dump = {}

    # ================= (2) longest option =================
    clean = [q for q in pool if not any(is_allofabove(b) for _, b in q['opts'])]
    rep.append(f'\n{"="*72}\n【2】最長選項法則')
    rep.append(f'排除含「以上皆是／皆非」的題目後樣本 {len(clean)} 題'
               f'（排除 {len(pool)-len(clean)} 題，因全稱選項又短、命中率又極端，會污染長度統計）')

    def lenrank(qs):
        # rank 1 = longest; ties share the worst rank so "the longest" stays unambiguous
        cnt = Counter()
        hit = Counter()
        uniq = 0
        for q in qs:
            ls = [(len(b), L) for L, b in q['opts']]
            order = sorted(ls, key=lambda x: -x[0])
            if order[0][0] == order[1][0]:
                continue          # no unique longest — skip
            uniq += 1
            for r, (_, L) in enumerate(order, 1):
                cnt[r] += 1
                if q['ans'] == L:
                    hit[r] += 1
        return cnt, hit, uniq

    cnt, hit, uniq = lenrank(clean)
    rep.append(f'（其中 {uniq} 題有唯一的最長選項）')
    rep.append(f'{"長度排名":>8}  {"命中率":<34}')
    for r in range(1, 5):
        lbl = {1: '最長', 2: '第二長', 3: '第三長', 4: '最短'}[r]
        rep.append(f'{lbl:>8}  {rate(hit[r], cnt[r])}')
    dump['length'] = {'cnt': dict(cnt), 'hit': dict(hit), 'n': uniq}

    rep.append('\n分科：')
    for s in SUBJECTS:
        c2, h2, u2 = lenrank([q for q in clean if q['subject'] == s])
        rep.append(f'  {s:<5} 最長 {rate(h2[1], c2[1])}   最短 {rate(h2[4], c2[4])}  (n={u2})')

    # how much longer does it have to be?
    rep.append('\n最長選項「比平均長多少」與命中率的關係：')
    buck = defaultdict(lambda: [0, 0])
    for q in clean:
        ls = {L: len(b) for L, b in q['opts']}
        mx = max(ls.values())
        others = sorted(ls.values())[:-1]
        if list(ls.values()).count(mx) > 1:
            continue
        ratio = mx / (sum(others) / 3) if sum(others) else 1
        b = ('1.0–1.2x' if ratio < 1.2 else '1.2–1.5x' if ratio < 1.5
             else '1.5–2.0x' if ratio < 2.0 else '≥2.0x')
        L = max(ls, key=ls.get)
        buck[b][1] += 1
        if q['ans'] == L:
            buck[b][0] += 1
    for b in ['1.0–1.2x', '1.2–1.5x', '1.5–2.0x', '≥2.0x']:
        k, n = buck[b]
        rep.append(f'  最長/其他平均 {b:<9} {rate(k, n)}')
    dump['length_ratio'] = {k: v for k, v in buck.items()}

    # ================= (3) absolute vs hedge =================
    rep.append(f'\n{"="*72}\n【3】絕對字詞 vs 緩和字詞')
    rep.append('絕對詞表：' + '、'.join(ABSOLUTE))
    rep.append('緩和詞表：' + '、'.join(HEDGE))
    rep.append('「僅」「不得」刻意不列入絕對詞——前者是組合題骨架（僅甲、乙），'
               '後者是法條常態用語（不得超過…），列入只會製造雜訊。')
    cat = defaultdict(lambda: [0, 0])
    per_subj = defaultdict(lambda: defaultdict(lambda: [0, 0]))
    caught = Counter()
    for q in pool:
        for L, b in q['opts']:
            if is_allofabove(b):
                k = '全稱選項（對照）'
            elif ABS_RE.search(b) and not HEDGE_RE.search(b):
                k = '含絕對詞'
                caught[('abs', ABS_RE.search(b).group(0))] += 1
            elif HEDGE_RE.search(b) and not ABS_RE.search(b):
                k = '含緩和詞'
                caught[('hedge', HEDGE_RE.search(b).group(0))] += 1
            elif ABS_RE.search(b) and HEDGE_RE.search(b):
                k = '兩者都含'
            else:
                k = '都不含'
            cat[k][1] += 1
            per_subj[q['subject']][k][1] += 1
            if q['ans'] == L:
                cat[k][0] += 1
                per_subj[q['subject']][k][0] += 1
    rep.append(f'\n{"選項類型":>10}  {"命中率":<34}')
    for k in ['含絕對詞', '含緩和詞', '兩者都含', '都不含', '全稱選項（對照）']:
        v = cat[k]
        rep.append(f'{k:>10}  {rate(v[0], v[1])}')
    rep.append('\n分科（只列絕對詞與緩和詞）：')
    for s in SUBJECTS:
        a, h = per_subj[s]['含絕對詞'], per_subj[s]['含緩和詞']
        rep.append(f'  {s:<5} 絕對 {rate(a[0], a[1])}   緩和 {rate(h[0], h[1])}')
    dump['words'] = {k: v for k, v in cat.items()}
    au = io.StringIO()
    au.write('詞表命中的實際字詞（覆核用）\n')
    for (t, w), n in sorted(caught.items(), key=lambda kv: (kv[0][0], -kv[1])):
        au.write(f'[{t}] {n:>5}  {w}\n')
    open(os.path.join(HERE, '_words_audit.txt'), 'w', encoding='utf-8').write(au.getvalue())

    # ================= (4) numeric middle =================
    rep.append(f'\n{"="*72}\n【4】數字題的中間值法則')
    numq = []
    for q in pool:
        parsed = [one_number(b) for _, b in q['opts']]
        if any(p is None for p in parsed):
            continue
        units = {p[1] for p in parsed}
        if len(units) != 1:
            continue
        vals = [p[0] for p in parsed]
        if len(set(vals)) != 4:
            continue
        numq.append((q, vals))
    rep.append(f'符合「四個選項皆為可比數值、單位一致、彼此相異」的題目：{len(numq)} 題')
    rc, rh = Counter(), Counter()
    for q, vals in numq:
        order = sorted(range(4), key=lambda i: vals[i])   # 0 = smallest
        ansi = 'ABCD'.index(q['ans'])
        pos = order.index(ansi) + 1
        for r in range(1, 5):
            rc[r] += 1
        rh[pos] += 1
    rep.append(f'{"數值排名":>10}  {"命中率":<34}')
    for r, lbl in [(1, '最小'), (2, '第二小'), (3, '第三小'), (4, '最大')]:
        rep.append(f'{lbl:>10}  {rate(rh[r], rc[r])}')
    mid = rh[2] + rh[3]
    ext = rh[1] + rh[4]
    # denominator is the number of QUESTIONS: each question's answer falls in exactly
    # one of the four ranks, so "middle two" covers half the ranks -> 50% is the null
    rep.append(f'\n{"中間兩個合計":>10}  {rate(mid, len(numq))}   '
               f'（四格中佔兩格，若無偏好應為 50%）')
    rep.append(f'{"兩極合計":>10}  {rate(ext, len(numq))}')
    dump['numeric'] = {'n': len(numq), 'hit_by_rank': dict(rh)}
    rep.append('\n分科：')
    for s in SUBJECTS:
        sub = [(q, v) for q, v in numq if q['subject'] == s]
        h = Counter()
        for q, vals in sub:
            order = sorted(range(4), key=lambda i: vals[i])
            h[order.index('ABCD'.index(q['ans'])) + 1] += 1
        m = h[2] + h[3]
        rep.append(f'  {s:<5} n={len(sub):<4} 中間兩個 {rate(m, len(sub))}  '
                   f'最小 {h[1]} 第二小 {h[2]} 第三小 {h[3]} 最大 {h[4]}')

    json.dump(dump, open(os.path.join(HERE, 'optionform.json'), 'w', encoding='utf-8'),
              ensure_ascii=False)
    open(os.path.join(HERE, 'optionform_report.txt'), 'w', encoding='utf-8').write('\n'.join(rep))
    print('\n'.join(rep))


if __name__ == '__main__':
    main()
