"""Corrections + validation:
 1. exact-match repeat rate (cross-check against the existing 統整檔's 48.6%)
 2. hazard table over ALL distinct questions incl. singletons (removes upward bias)
 3. backtest: forecast each of the last 6 exams from data strictly before it
 4. fixed 8-屆 window recycle rate (removes the 'longer history' mechanical trend)
 5. final calibrated forecast for 115-3
"""
import json, os, itertools, math
from collections import defaultdict, Counter
from s2_cluster import cluster, EXAMS, IDX, N, SUBJECTS

HERE = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'out')
os.makedirs(HERE, exist_ok=True)
recs = json.load(open(os.path.join(HERE, 'corpus.json'), encoding='utf-8'))
YEAR = [int(e.split('-')[0]) for e in EXAMS]
rep = []
store = {}

# ---------- 1. exact match baseline ----------
rep.append('【1】比對標準敏感度：嚴格一字不差 vs 模糊 ≥90%')
rep.append(f'{"科目":>6} {"總題數":>7} {"嚴格重複題":>10} {"嚴格%":>7} {"模糊重複題":>10} {"模糊%":>7}')
prev = json.load(open(os.path.join(HERE, 'analysis.json'), encoding='utf-8'))
exact_tot = fuzzy_tot = 0
for subj in SUBJECTS:
    rs = [r for r in recs if r['subject'] == subj]
    c = Counter(r['stem'] for r in rs)
    # count questions whose stem appears in >=2 DIFFERENT exams
    bystem = defaultdict(set)
    for r in rs:
        bystem[r['stem']].add(r['exam'])
    ex = sum(1 for r in rs if len(bystem[r['stem']]) >= 2)
    fz = prev['subjects'][subj]['repeated_q']
    exact_tot += ex
    fuzzy_tot += fz
    rep.append(f'{subj:>6} {len(rs):>7} {ex:>10} {ex/len(rs)*100:>6.1f}% {fz:>10} {fz/len(rs)*100:>6.1f}%')
rep.append(f'{"合計":>6} {len(recs):>7} {exact_tot:>10} {exact_tot/len(recs)*100:>6.1f}% '
           f'{fuzzy_tot:>10} {fuzzy_tot/len(recs)*100:>6.1f}%')
rep.append('  → 統整檔用嚴格標準得 48.6%，本次嚴格重算應落在同一量級（互相驗證）')

# ---------- build full clusters incl. singletons ----------
full = {}
for subj in SUBJECTS:
    rs = [r for r in recs if r['subject'] == subj]
    cl = cluster(rs)
    items = []
    for c in cl:
        slots = sorted({IDX[rs[i]['exam']] for i in c})
        items.append({'slots': slots,
                      'stem': min((rs[i]['raw'] for i in c), key=len)[:130],
                      'ans': Counter(rs[i]['ans'] for i in c if rs[i]['ans']).most_common(1)})
    full[subj] = items


def kbin(k):
    return min(k, 4)


def gbin(g):
    return 1 if g <= 1 else 2 if g <= 2 else 3 if g <= 4 else 4 if g <= 8 else 5


GL = {1: '1屆', 2: '2屆', 3: '3-4屆', 4: '5-8屆', 5: '>8屆'}


def build_hazard(items, upto):
    """Trials only over exams < upto, so a forecast for `upto` uses no future info."""
    hz = defaultdict(lambda: [0, 0])
    for c in items:
        s = set(x for x in c['slots'] if x < upto)
        if not s:
            continue
        first = min(s)
        k = 0
        last = None
        for t in range(first, upto):
            if last is not None:
                key = (kbin(k), gbin(t - last))
                hz[key][1] += 1
                if t in s:
                    hz[key][0] += 1
            if t in s:
                k += 1
                last = t
    return hz


def forecast(items, upto, hz, prior):
    """Expected number of questions in exam `upto` that already appeared before it."""
    tot = 0.0
    detail = []
    for c in items:
        s = [x for x in c['slots'] if x < upto]
        if not s:
            continue
        k, g = len(s), upto - max(s)
        h, n = hz.get((kbin(k), gbin(g)), [0, 0])
        p = h / n if n >= 30 else prior
        tot += p
        detail.append((p, k, g, c))
    detail.sort(key=lambda x: (-x[0], -x[1]))
    return tot, detail


# ---------- 4. fixed-window recycle rate ----------
rep.append('\n\n【4】回收率趨勢的混淆校正')
rep.append('  「向前回收率」會隨屆次自然上升，因為越晚的考卷可抄的歷史越長（機械效應）。')
rep.append('  改用固定 8 屆視窗：只算「該題在前 8 屆內出現過」，各屆條件一致（108-4 起可比）。')
for subj in SUBJECTS:
    items = full[subj]
    win8, allback = {}, {}
    for t in range(N):
        w = a = 0
        for c in items:
            s = [x for x in c['slots'] if x < t]
            if t in c['slots'] and s:
                a += 1
                if t - max(s) <= 8:
                    w += 1
        win8[t], allback[t] = w, a
    rep.append(f'\n  【{subj}】 固定8屆視窗回收題數 / 全歷史回收題數（每屆 50 題）')
    row = [f'{EXAMS[i]}:{win8[i]}/{allback[i]}' for i in range(N)]
    for i in range(0, N, 6):
        rep.append('    ' + '   '.join(row[i:i + 6]))
    early = [i for i in range(N) if i >= IDX['108-4'] and i < IDX['112-1']]
    late = [i for i in range(N) if i >= IDX['112-1']]
    rep.append(f'    108-4~111-3 平均: 視窗 {sum(win8[i] for i in early)/len(early):.1f} | '
               f'全歷史 {sum(allback[i] for i in early)/len(early):.1f}')
    rep.append(f'    112-1~115-2 平均: 視窗 {sum(win8[i] for i in late)/len(late):.1f} | '
               f'全歷史 {sum(allback[i] for i in late)/len(late):.1f}')
    store.setdefault(subj, {})['win8'] = {EXAMS[i]: win8[i] for i in range(N)}
    store[subj]['allback'] = {EXAMS[i]: allback[i] for i in range(N)}

# ---------- 2+3. unbiased hazard + backtest ----------
rep.append('\n\n【2】以「全部相異題目」為分母的命中率（含只出現過一次、從未重出的題）')
for subj in SUBJECTS:
    items = full[subj]
    hz = build_hazard(items, N)
    rep.append(f'\n  【{subj}】 相異題目數={len(items)}')
    rep.append(f'  {"已出現次數":>10} {"距上次":>8} {"樣本":>8} {"再現":>6} {"下屆再現機率":>12}')
    for kb in range(1, 5):
        for gb in range(1, 6):
            h, n = hz.get((kb, gb), [0, 0])
            if n < 30:
                continue
            rep.append(f'  {(str(kb)+"+" if kb==4 else str(kb)):>10} {GL[gb]:>8} '
                       f'{n:>8} {h:>6} {h/n*100:>11.1f}%')
    store[subj]['hazard_all'] = {f'{a}|{b}': v for (a, b), v in hz.items()}

rep.append('\n\n【3】回測校準：用「該屆之前」的資料預測該屆，與實際相比')
rep.append(f'{"科目":>6} {"預測屆":>8} {"預測重出題數":>12} {"實際":>6} {"誤差":>7}')
back_rows = []
for subj in SUBJECTS:
    items = full[subj]
    for t in range(N - 6, N):
        hz = build_hazard(items, t)
        base = sum(v[0] for v in hz.values()) / max(1, sum(v[1] for v in hz.values()))
        pred, _ = forecast(items, t, hz, base)
        actual = sum(1 for c in items if t in c['slots'] and any(x < t for x in c['slots']))
        rep.append(f'{subj:>6} {EXAMS[t]:>8} {pred:>12.1f} {actual:>6} {pred-actual:>+7.1f}')
        back_rows.append((subj, EXAMS[t], pred, actual))
for subj in SUBJECTS:
    rs = [r for r in back_rows if r[0] == subj]
    mae = sum(abs(p - a) for _, _, p, a in rs) / len(rs)
    bias = sum(p - a for _, _, p, a in rs) / len(rs)
    rep.append(f'  {subj}: 平均絕對誤差 {mae:.1f} 題，系統偏誤 {bias:+.1f} 題')
    # s5 reads these back — never hard-code them, or a re-run silently
    # applies the previous cycle's correction to new data
    store[subj]['backtest'] = {'mae': mae, 'bias': bias,
                               'rows': [{'exam': e, 'pred': p, 'actual': a}
                                        for _, e, p, a in rs]}

# ---------- 5. final forecast ----------
rep.append('\n\n【5】115-3 預測（模型只看 115-2 為止的資料）')
grand = 0
for subj in SUBJECTS:
    items = full[subj]
    hz = build_hazard(items, N)
    base = sum(v[0] for v in hz.values()) / max(1, sum(v[1] for v in hz.values()))
    pred, detail = forecast(items, N, hz, base)
    store[subj]['forecast_raw'] = pred
    grand += pred
    rep.append(f'\n  【{subj}】 期望重出 {pred:.1f} / 50 題 ({pred/50*100:.0f}%)')
    rep.append('    命中機率最高的 20 題：')
    for p, k, g, c in detail[:20]:
        a = c['ans'][0][0] if c['ans'] else '—'
        rep.append(f'      {p*100:>4.1f}%  出現{k}次 距{g}屆  答{a}  {c["stem"][:78]}')
    store[subj]['forecast'] = [
        {'p': p, 'k': k, 'g': g, 'ans': (c['ans'][0][0] if c['ans'] else ''),
         'stem': c['stem'], 'slots': [EXAMS[i] for i in c['slots']]}
        for p, k, g, c in detail[:80]]
rep.append(f'\n  三科合計期望重出 {grand:.1f} / 150 題 ({grand/150*100:.0f}%)')

json.dump(store, open(os.path.join(HERE, 'analysis3.json'), 'w', encoding='utf-8'), ensure_ascii=False)
open(os.path.join(HERE, 'analysis3_report.txt'), 'w', encoding='utf-8').write('\n'.join(rep))
print('written')
