"""Follow-ups: (a) same-year vs cross-year at matched gap, (b) edge-bias-corrected
backward recycle rate per exam, (c) empirical hazard table -> forecast for 115-3."""
import json, os, itertools, math
from collections import defaultdict, Counter

HERE = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'out')
os.makedirs(HERE, exist_ok=True)
A = json.load(open(os.path.join(HERE, 'analysis.json'), encoding='utf-8'))
EXAMS = A['exams']
N = len(EXAMS)
YEAR = [int(e.split('-')[0]) for e in EXAMS]
rep = []

# how many exam slots exist at each gap, and how many of those are same-year
slot_gap_total = Counter()
slot_gap_same = Counter()
for a, b in itertools.combinations(range(N), 2):
    g = b - a
    slot_gap_total[g] += 1
    if YEAR[a] == YEAR[b]:
        slot_gap_same[g] += 1

for subj, D in A['subjects'].items():
    clusters = D['clusters']
    rep.append(f'\n{"="*70}\n【{subj}】')

    # ---------- (a) same-year vs cross-year, controlled for gap ----------
    pair_same = Counter()
    pair_cross = Counter()
    for c in clusters:
        for a, b in itertools.combinations(c['slots'], 2):
            g = b - a
            if YEAR[a] == YEAR[b]:
                pair_same[g] += 1
            else:
                pair_cross[g] += 1
    rep.append('\n(a) 同年內 vs 跨年，在「相同間隔」下的每組配對平均共用題數')
    rep.append(f'{"間隔":>4} {"同年配對數":>10} {"同年共用":>8} {"每對":>7} '
               f'{"跨年配對數":>10} {"跨年共用":>8} {"每對":>7} {"同年/跨年":>9}')
    for g in (1, 2, 3):
        ns, nc = slot_gap_same[g], slot_gap_total[g] - slot_gap_same[g]
        ss, sc = pair_same.get(g, 0), pair_cross.get(g, 0)
        rs = ss / ns if ns else 0
        rc = sc / nc if nc else 0
        rep.append(f'{g:>4} {ns:>10} {ss:>8} {rs:>7.2f} {nc:>10} {sc:>8} {rc:>7.2f} '
                   f'{(rs/rc if rc else float("nan")):>9.2f}x')

    # ---------- (b) backward recycle rate ----------
    first_seen = {}
    back = Counter()
    for ci, c in enumerate(clusters):
        s = c['slots']
        for x in s[1:]:
            back[x] += 1          # this exam re-used a question seen earlier
    rep.append('\n(b) 各屆「向前回收率」＝該屆 50 題中曾在更早屆出現過的題數')
    rep.append('    (前幾屆因為可抄的歷史短，數字天生偏低——看 107 年之後的平台期)')
    row = [f'{EXAMS[i]}:{back.get(i,0)}' for i in range(N)]
    for i in range(0, N, 8):
        rep.append('    ' + '  '.join(row[i:i + 8]))
    plateau = [back.get(i, 0) for i in range(N) if i >= 8]
    rep.append(f'    107-1 之後平均回收 {sum(plateau)/len(plateau):.1f} 題/50 '
               f'({sum(plateau)/len(plateau)/50*100:.0f}%)')
    recent = [back.get(i, 0) for i in range(N - 6, N)]
    older = [back.get(i, 0) for i in range(8, N - 6)]
    rep.append(f'    近 6 屆平均 {sum(recent)/len(recent):.1f}  vs  107-1~113-3 平均 {sum(older)/len(older):.1f}')

    # how far back does a recycled question come from?
    srcdist = Counter()
    for c in clusters:
        s = c['slots']
        for i in range(1, len(s)):
            srcdist[s[i] - s[i - 1]] += 1
    tot = sum(srcdist.values())
    cum = 0
    rep.append('\n    回收來源距離（距上一次出現幾屆）累積分布：')
    for g in (1, 2, 3, 4, 6, 8, 12, 16, 24, 36):
        cum = sum(v for k, v in srcdist.items() if k <= g)
        rep.append(f'      ≤{g:>2} 屆: {cum/tot*100:>5.1f}%')

    # ---------- (c) empirical hazard -> forecast ----------
    # state at exam t for a cluster: (k = appearances so far, g = gap since last)
    # outcome: does it appear at exam t?
    hz = defaultdict(lambda: [0, 0])   # (kbin,gbin) -> [hits, trials]

    def kbin(k):
        return min(k, 4)

    def gbin(g):
        return 1 if g <= 1 else 2 if g <= 2 else 3 if g <= 4 else 4 if g <= 8 else 5

    for c in clusters:
        s = set(c['slots'])
        first = min(s)
        k = 0
        last = None
        for t in range(first, N):
            if last is not None:
                key = (kbin(k), gbin(t - last))
                hz[key][1] += 1
                if t in s:
                    hz[key][0] += 1
            if t in s:
                k += 1
                last = t
    rep.append('\n(c) 經驗命中率：一道「已出現 k 次、距上次 g 屆」的題目，在下一屆再出現的機率')
    rep.append(f'{"出現次數":>8} {"距上次":>10} {"樣本":>7} {"再現次數":>8} {"機率":>7}')
    glabel = {1: '1屆', 2: '2屆', 3: '3-4屆', 4: '5-8屆', 5: '>8屆'}
    for kb in range(1, 5):
        for gb in range(1, 6):
            h, n = hz.get((kb, gb), [0, 0])
            if n < 25:
                continue
            rep.append(f'{(str(kb)+"+" if kb==4 else str(kb)):>8} {glabel[gb]:>10} '
                       f'{n:>7} {h:>8} {h/n*100:>6.1f}%')

    # forecast for the next exam (index N, i.e. 115-3)
    scored = []
    for c in clusters:
        s = c['slots']
        k = len(s)
        g = N - max(s)
        h, n = hz.get((kbin(k), gbin(g)), [0, 0])
        p = h / n if n >= 25 else None
        if p:
            scored.append((p, k, g, c['stem'], c['members'][-1], c['ans']))
    scored.sort(key=lambda x: (-x[0], -x[1]))
    exp_hits = sum(x[0] for x in scored)
    rep.append(f'\n    下一屆(115-3) 期望重出題數 = {exp_hits:.1f} / 50 '
               f'（把每個題組的命中機率加總）')
    rep.append('    機率最高的 15 個題組：')
    for p, k, g, stem, last, ans in scored[:15]:
        a = ans[0][0] if ans else '—'
        rep.append(f'      p={p*100:>4.1f}%  已出現{k}次 距上次{g}屆  答{a}  {stem[:70]}')

    D['forecast'] = [{'p': p, 'k': k, 'g': g, 'stem': st, 'last': l,
                      'ans': (an[0][0] if an else '')} for p, k, g, st, l, an in scored[:60]]
    D['hazard'] = {f'{a}|{b}': v for (a, b), v in hz.items()}
    D['back'] = {EXAMS[i]: back.get(i, 0) for i in range(N)}
    D['srcdist'] = dict(srcdist)
    D['pair_same'] = dict(pair_same)
    D['pair_cross'] = dict(pair_cross)

A['slot_gap_total'] = dict(slot_gap_total)
A['slot_gap_same'] = dict(slot_gap_same)
json.dump(A, open(os.path.join(HERE, 'analysis.json'), 'w', encoding='utf-8'), ensure_ascii=False)
open(os.path.join(HERE, 'analysis2_report.txt'), 'w', encoding='utf-8').write('\n'.join(rep))
print('written')
