"""Cluster repeated questions (fuzzy >=90% on the stem) and test the repeat pattern
for regularity against a Monte-Carlo null model. Per subject."""
import json, os, re, random, math, itertools
from collections import defaultdict, Counter
from difflib import SequenceMatcher

HERE = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'out')
os.makedirs(HERE, exist_ok=True)
THRESH = 0.90
SUBJECTS = ['投資學', '財務分析', '法規']

# ---- chronological exam order (37) -------------------------------------
EXAMS = []
for y in range(105, 116):
    n = 4 if y <= 109 else (2 if y == 115 else 3)
    for s in range(1, n + 1):
        EXAMS.append(f'{y}-{s}')
IDX = {e: i for i, e in enumerate(EXAMS)}
N = len(EXAMS)


def cluster(records):
    """Union-find over pairs with SequenceMatcher ratio >= THRESH.
    Length-band + trigram prefilters keep this ~O(n * small)."""
    n = len(records)
    stems = [r['stem'] for r in records]
    tri = [set(s[i:i + 3] for i in range(max(1, len(s) - 2))) for s in stems]
    order = sorted(range(n), key=lambda i: len(stems[i]))
    parent = list(range(n))

    def find(a):
        while parent[a] != a:
            parent[a] = parent[parent[a]]
            a = parent[a]
        return a

    def union(a, b):
        ra, rb = find(a), find(b)
        if ra != rb:
            parent[rb] = ra

    for oi in range(n):
        i = order[oi]
        li = len(stems[i])
        for oj in range(oi + 1, n):
            j = order[oj]
            lj = len(stems[j])
            if lj > li / 0.80:          # ratio>=0.9 impossible beyond this length gap
                break
            if find(i) == find(j):
                continue
            inter = len(tri[i] & tri[j])
            if inter * 2 < 0.72 * (len(tri[i]) + len(tri[j])):
                continue                # trigram Dice prefilter
            sm = SequenceMatcher(None, stems[i], stems[j])
            if sm.real_quick_ratio() < THRESH or sm.quick_ratio() < THRESH:
                continue
            if sm.ratio() >= THRESH:
                union(i, j)
    groups = defaultdict(list)
    for i in range(n):
        groups[find(i)].append(i)
    return [v for v in groups.values()]


def gaps_of(slots):
    s = sorted(slots)
    return [b - a for a, b in zip(s, s[1:])]


def montecarlo(sizes, iters=4000, seed=7):
    """Null: each cluster's k occurrences land on k distinct exams uniformly at random."""
    rng = random.Random(seed)
    gap_hist = Counter()
    cvs = []
    for _ in range(iters):
        for k in sizes:
            slots = rng.sample(range(N), k)
            g = gaps_of(slots)
            for x in g:
                gap_hist[x] += 1
            if len(g) >= 2:
                m = sum(g) / len(g)
                sd = math.sqrt(sum((x - m) ** 2 for x in g) / len(g))
                cvs.append(sd / m if m else 0)
    tot = sum(gap_hist.values()) or 1
    return {k: v / iters for k, v in gap_hist.items()}, (sum(cvs) / len(cvs) if cvs else None), tot / iters


def main():
    recs = json.load(open(os.path.join(HERE, 'corpus.json'), encoding='utf-8'))
    report = []
    dump = {}

    for subj in SUBJECTS:
        rs = [r for r in recs if r['subject'] == subj]
        cl = cluster(rs)
        multi = [c for c in cl if len({rs[i]['exam'] for i in c}) >= 2]
        # slots per cluster (dedupe same-exam duplicates)
        clusters = []
        for c in multi:
            slots = sorted({IDX[rs[i]['exam']] for i in c})
            clusters.append({
                'slots': slots,
                'k': len(slots),
                'stem': min((rs[i]['raw'] for i in c), key=len)[:120],
                'members': sorted({rs[i]['exam'] + ' Q' + str(rs[i]['qno']) for i in c}),
                'ans': Counter(rs[i]['ans'] for i in c if rs[i]['ans']).most_common(1),
            })
        clusters.sort(key=lambda d: -d['k'])

        repeated_q = sum(len(c) for c in multi)
        report.append(f'\n{"="*70}\n【{subj}】 題數={len(rs)}  重複題組={len(clusters)}  '
                      f'涉及題目={repeated_q} ({repeated_q/len(rs)*100:.1f}%)')
        report.append('題組大小分布: ' + ', '.join(
            f'{k}次×{v}組' for k, v in sorted(Counter(c['k'] for c in clusters).items())))

        # --- gap distribution: observed vs null
        obs = Counter()
        obs_cvs = []
        for c in clusters:
            g = gaps_of(c['slots'])
            for x in g:
                obs[x] += 1
            if len(g) >= 2:
                m = sum(g) / len(g)
                sd = math.sqrt(sum((x - m) ** 2 for x in g) / len(g))
                obs_cvs.append(sd / m if m else 0)
        null, null_cv, null_pairs = montecarlo([c['k'] for c in clusters])
        tot_obs = sum(obs.values()) or 1
        report.append(f'\n重出間隔（單位=屆）觀察 vs 隨機虛無模型   總間隔數={tot_obs}')
        report.append(f'{"間隔":>4} {"觀察":>6} {"觀察%":>7} {"隨機期望":>8} {"期望%":>7} {"obs/exp":>8}')
        for gp in range(1, 25):
            o = obs.get(gp, 0)
            e = null.get(gp, 0)
            if o == 0 and e < 0.5:
                continue
            report.append(f'{gp:>4} {o:>6} {o/tot_obs*100:>6.1f}% {e:>8.1f} '
                          f'{e/(null_pairs or 1)*100:>6.1f}% {(o/e if e else float("nan")):>8.2f}')
        obs_cv = sum(obs_cvs) / len(obs_cvs) if obs_cvs else None
        report.append(f'\n間隔規律性（≥3 次題組的間隔變異係數 CV，越小=越規律）：'
                      f'觀察={obs_cv:.3f}  隨機={null_cv:.3f}  '
                      f'（{"更規律" if obs_cv and obs_cv < null_cv else "並未更規律"}）'
                      f'  樣本={len(obs_cvs)} 組')

        # --- exam-pair co-occurrence
        pair = Counter()
        for c in clusters:
            for a, b in itertools.combinations(c['slots'], 2):
                pair[(a, b)] += 1
        # per-exam recycled load
        per_exam_rep = Counter()
        for c in clusters:
            for s in c['slots']:
                per_exam_rep[s] += 1
        report.append('\n各屆「與其他屆重複」的題數（該屆 50 題中）：')
        row = []
        for i, e in enumerate(EXAMS):
            row.append(f'{e}:{per_exam_rep.get(i,0)}')
        for i in range(0, len(row), 8):
            report.append('  ' + '  '.join(row[i:i + 8]))

        # strongest exam pairs vs expectation
        exp_pair = {}
        for (a, b), v in pair.items():
            # expected under: b's repeats spread evenly over all other exams
            exp = (per_exam_rep[a] * per_exam_rep[b]) / max(1, sum(per_exam_rep.values()))
            exp_pair[(a, b)] = exp
        top = sorted(pair.items(), key=lambda kv: -(kv[1] - exp_pair[kv[0]]))[:12]
        report.append('\n最強屆次配對（共用題數 vs 期望；期望=兩屆重複題量的乘積模型）：')
        for (a, b), v in top:
            report.append(f'  {EXAMS[a]} ↔ {EXAMS[b]}  共用 {v} 題  期望 {exp_pair[(a,b)]:.1f}  '
                          f'倍數 {v/exp_pair[(a,b)]:.2f}  間隔 {b-a} 屆')

        dump[subj] = {
            'n': len(rs), 'clusters': clusters, 'obs_gap': dict(obs),
            'null_gap': null, 'obs_cv': obs_cv, 'null_cv': null_cv,
            'pair': {f'{a}|{b}': v for (a, b), v in pair.items()},
            'per_exam_rep': {EXAMS[i]: per_exam_rep.get(i, 0) for i in range(N)},
            'repeated_q': repeated_q,
        }

    json.dump({'exams': EXAMS, 'subjects': dump}, open(os.path.join(HERE, 'analysis.json'), 'w', encoding='utf-8'), ensure_ascii=False)
    open(os.path.join(HERE, 'analysis_report.txt'), 'w', encoding='utf-8').write('\n'.join(report))
    print('written')


if __name__ == '__main__':
    main()
