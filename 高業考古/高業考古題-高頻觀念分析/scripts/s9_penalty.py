"""Is the most severe penalty the right answer?

Severity cannot be ranked by regex — 「3年以下有期徒刑，得併科3萬元罰金」 vs
「2年以下有期徒刑，得併科240萬元罰金」 needs a judgement about which criterion
dominates (imprisonment does). So the ranking below is hand-coded, one row per
question, and every inclusion/exclusion decision is recorded with its reason so
the call can be audited rather than taken on trust.

Candidate pool was generated mechanically (>=3 of 4 options carry penalty language,
or the stem asks about 罰鍰/罰金/徒刑 and all 4 options are comparable numbers);
each candidate was then read individually. See out/_penalty_candidates.txt and
out/_penalty_numeric.txt for the raw candidate dumps.
"""
import json, os, math
from collections import Counter

HERE = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'out')
os.makedirs(HERE, exist_ok=True)

# sev = option letters ordered MOST severe -> LEAST severe
INCLUDED = [
    ('109-3', 19, '純量刑', 'ABCD', '刑度：3~10年+2億 > 5年以下 > 2年以下 > 1年以下'),
    ('111-3', 10, '純量刑', 'DACB', '刑度上限：10 > 7 > 3 > 2 年（罰金為次要判準）'),
    ('111-3', 23, '純量刑', 'DBCA', '刑度上限：10 > 7 > 3 > 1 年'),
    ('111-3', 24, '罰鍰階梯', 'CDBA', '罰鍰上限：60 > 30 > 25 > 20 萬'),
    ('112-1', 10, '純量刑', 'DBCA', '刑度上限：10 > 7 > 5 > 1 年'),
    ('112-2', 23, '純量刑', 'DCBA', '刑度上限：10 > 5 > 2 > 1 年'),
    ('113-1', 23, '純量刑', 'DBCA', '與 112-1 Q10 為同一題（重複題）'),
    ('113-3', 10, '純量刑', 'DCBA', '刑度：10 > 5 > 3 > 1 年'),
    ('113-3', 25, '純量刑', 'ABCD', '刑度：7 > 5 > 2 > 1 年'),
    ('115-1', 28, '處分手段', 'DCAB', '責任範圍：本息+民刑 > 本息+損賠 > 返還本金 > 返還利息'),
    ('115-2', 3, '處分手段', 'DCBA', '民事責任：懲罰性 > 連帶 > 區別 > 單獨'),
]

EXCLUDED = [
    ('109-4', 5, '財務分析', '「處分損失」的「處分」是會計用語，與罰則無關，關鍵字誤抓'),
    ('109-4', 18, '法規', '比的是責任「種類」（刑事／民事／皆有／無）而非嚴重程度，且含全稱選項，已由 s7 涵蓋'),
    ('112-1', 19, '法規', '反向題幹「關於內線交易之處罰，下列何者為非」——選最重必然失效，邏輯不適用'),
    ('113-1', 16, '法規', '董事解任的敘述是非題，選項並非可排序的罰則'),
    ('113-3', 5, '法規', '同上（與 113-1 Q16 為同一題）'),
    ('114-3', 6, '法規', '反向題幹「下列何者為錯誤」，且問的是何者不是合法處分'),
    ('114-3', 22, '法規', '問「得為哪些處分」而非比嚴重度，且正解是全稱選項，已由 s7 涵蓋'),
    ('115-1', 10, '法規', '同上（與 114-3 Q22 為同一題）'),
    ('115-1', 28, '法規', ''),   # placeholder removed below
]
EXCLUDED = [e for e in EXCLUDED if e[3]]


def wilson(k, n):
    if n == 0:
        return (0.0, 0.0)
    z = 1.96
    p = k / n
    d = 1 + z * z / n
    c = (p + z * z / (2 * n)) / d
    m = z * math.sqrt(p * (1 - p) / n + z * z / (4 * n * n)) / d
    return (max(0, c - m) * 100, min(1, c + m) * 100)


def binom_ge(k, n, p):
    """P(X >= k) for Binomial(n, p) — is the result beyond blind guessing?"""
    return sum(math.comb(n, i) * p ** i * (1 - p) ** (n - i) for i in range(k, n + 1))


def main():
    recs = json.load(open(os.path.join(HERE, 'corpus.json'), encoding='utf-8'))
    idx = {(q['exam'], q['qno']): q for q in recs if q['subject'] != '' and q['opts']}

    rows = []
    for exam, qno, kind, sev, why in INCLUDED:
        q = idx.get((exam, qno))
        if q is None:
            print('MISSING', exam, qno)
            continue
        rank = sev.index(q['ans']) + 1
        rows.append({'exam': exam, 'qno': qno, 'kind': kind, 'sev': sev,
                     'ans': q['ans'], 'rank': rank, 'why': why,
                     'head': q['head'], 'subject': q['subject'],
                     'opts': q['opts']})

    rep = []
    rep.append('【判定範圍】')
    rep.append(f'  機械撈出候選 19 題（17 題罰則選項 + 2 題純數字罰則），逐題人工覆核後：')
    rep.append(f'  納入 {len(rows)} 題，排除 {len(EXCLUDED)} 題。')
    rep.append('\n  排除清單與理由：')
    for exam, qno, subj, why in EXCLUDED:
        rep.append(f'    {exam} Q{qno} {subj}｜{why}')

    n = len(rows)
    c = Counter(r['rank'] for r in rows)
    rep.append(f'\n【結果】納入 {n} 題，正解落在第幾重：')
    rep.append(f'{"嚴重度":>8}  {"題數":>4}  {"占比":>7}  {"95% CI":>16}')
    for r, lbl in [(1, '最重'), (2, '第二重'), (3, '第三重'), (4, '最輕')]:
        lo, hi = wilson(c[r], n)
        rep.append(f'{lbl:>8}  {c[r]:>4}  {c[r]/n*100:>6.1f}%  [{lo:>5.1f}%,{hi:>5.1f}%]')

    top2 = c[1] + c[2]
    rep.append(f'\n  最重或第二重合計：{top2}/{n} = {top2/n*100:.1f}%')
    rep.append(f'  第三重或最輕：{c[3]+c[4]}/{n} = {(c[3]+c[4])/n*100:.1f}%')
    rep.append(f'\n  「選最重」相對亂猜 25% 的顯著性：P(X≥{c[1]}|n={n},p=.25) = '
               f'{binom_ge(c[1], n, 0.25):.4f}')
    rep.append(f'  「正解落在前二重」相對亂猜 50% 的顯著性：P(X≥{top2}|n={n},p=.5) = '
               f'{binom_ge(top2, n, 0.5):.5f}')

    # dedupe: 113-1 Q23 is the same question as 112-1 Q10
    seen, ded = set(), []
    for r in rows:
        key = (r['sev'], r['ans'], r['head'][:40])
        if key in seen:
            continue
        seen.add(key)
        ded.append(r)
    cd = Counter(r['rank'] for r in ded)
    rep.append(f'\n【去重後】相異題目 {len(ded)} 題：最重 {cd[1]}、第二重 {cd[2]}、'
               f'第三重 {cd[3]}、最輕 {cd[4]}  →  選最重命中 {cd[1]/len(ded)*100:.1f}%')

    rep.append('\n【逐題明細】')
    rep.append(f'{"屆次":>7} {"題號":>4} {"類型":>6} {"嚴重度序":>9} {"正解":>4} {"落在":>5}  說明')
    for r in sorted(rows, key=lambda r: (r['exam'], r['qno'])):
        lbl = {1: '最重', 2: '第二重', 3: '第三重', 4: '最輕'}[r['rank']]
        rep.append(f'{r["exam"]:>7} Q{r["qno"]:<3} {r["kind"]:>6} {r["sev"]:>9} '
                   f'{r["ans"]:>4} {lbl:>5}  {r["why"]}')

    rep.append('\n【納入題目全文】（供覆核嚴重度排序是否合理）')
    for r in sorted(rows, key=lambda r: (r['exam'], r['qno'])):
        rep.append(f'\n{r["exam"]} Q{r["qno"]}（正解 {r["ans"]}，嚴重度序 {r["sev"]}）')
        rep.append(f'  {r["head"][:110]}')
        for L, b in r['opts']:
            mark = '★' if L == r['ans'] else ' '
            rep.append(f'   {mark}({L}) {b[:84]}')

    json.dump({'rows': [{k: v for k, v in r.items() if k != 'opts'} for r in rows],
               'excluded': EXCLUDED},
              open(os.path.join(HERE, 'penalty.json'), 'w', encoding='utf-8'),
              ensure_ascii=False)
    open(os.path.join(HERE, 'penalty_report.txt'), 'w', encoding='utf-8').write('\n'.join(rep))
    print('\n'.join(rep[:32]))


if __name__ == '__main__':
    main()
