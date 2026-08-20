"""How often is 「以上皆是」／「以上皆非」 actually the right answer?

Counts only *true* all-of-the-above options — ones that refer to the OTHER OPTIONS
(「選項(A)(B)(C)皆是」「以上皆非」). Deliberately excludes 「甲、乙、丙皆是」, which
refers to items listed in the stem, not to the options, and is a combination question.

Sample: the 2,849 questions that have both a parsed A–D option set and a single-letter
official answer (answer cards exist from 109 onward only).
"""
import json, os, re, io
from collections import Counter, defaultdict

HERE = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'out')
os.makedirs(HERE, exist_ok=True)

# must point at the other options, not at 甲乙丙 items in the stem
REFERS = re.compile(r'以上|選項|[（(]\s*[A-D]\s*[)）]')
NEG = re.compile(r'(?:皆|均|都)\s*(?:非|錯誤|錯|不正確|不對|不是|不可)')
POS = re.compile(r'(?:皆|均|都)\s*(?:是|對|正確|可能|可|能|得|屬|有)')


def classify(text):
    if not REFERS.search(text):
        return None
    if NEG.search(text):
        return 'neg'
    if POS.search(text):
        return 'pos'
    return None


def period(exam):
    y = int(exam.split('-')[0])
    return '109-111' if y <= 111 else '112-115'


def main():
    recs = json.load(open(os.path.join(HERE, 'corpus.json'), encoding='utf-8'))
    pool = [q for q in recs if q['opts'] and len(q['ans']) == 1 and q['ans'] in 'ABCD']

    rows = []           # one per (question, all-of-above option)
    caught = Counter()  # for manual review of the detector
    for q in pool:
        for L, b in q['opts']:
            kind = classify(b)
            if kind:
                caught[(kind, b.strip()[:40])] += 1
                rows.append({'q': q, 'letter': L, 'kind': kind, 'text': b,
                             'hit': q['ans'] == L})

    # detector audit dump
    au = io.StringIO()
    au.write(f'偵測到的全稱選項字串（共 {len(rows)} 個，{len(caught)} 種寫法）\n')
    au.write('人工覆核用：任何不該入列的寫法都代表偵測器要修\n\n')
    for (k, t), n in sorted(caught.items(), key=lambda kv: (kv[0][0], -kv[1])):
        au.write(f'[{k}] {n:>4}  {t}\n')
    open(os.path.join(HERE, '_aoa_detector_audit.txt'), 'w', encoding='utf-8').write(au.getvalue())

    rep = []
    rep.append(f'樣本：{len(pool)} 題（有選項且有單一正解）')
    rep.append(f'其中含全稱選項的題目：{len({id(r["q"]) for r in rows})} 題，'
               f'全稱選項共 {len(rows)} 個\n')

    def block(title, subset):
        out = []
        for kind, label in [('pos', '肯定型（以上皆是／皆正確／皆可）'),
                            ('neg', '否定型（以上皆非／皆錯誤）')]:
            s = [r for r in subset if r['kind'] == kind]
            if not s:
                continue
            hit = sum(1 for r in s if r['hit'])
            out.append(f'  {label:<28} 出現 {len(s):>4} 次，是正解 {hit:>4} 次'
                       f'  →  命中率 {hit/len(s)*100:>5.1f}%')
        s = subset
        hit = sum(1 for r in s if r['hit'])
        out.append(f'  {"兩類合計":<28} 出現 {len(s):>4} 次，是正解 {hit:>4} 次'
                   f'  →  命中率 {hit/len(s)*100:>5.1f}%')
        return [title] + out

    rep += block('【全體 三科合計】', rows)
    rep.append('')
    for subj in ['投資學', '財務分析', '法規']:
        rep += block(f'【{subj}】', [r for r in rows if r['q']['subject'] == subj])
        rep.append('')

    rep.append('【年度趨勢】')
    for p in ['109-111', '112-115']:
        rep += block(f'  期間 {p}', [r for r in rows if period(r['q']['exam']) == p])
        rep.append('')

    # position sanity: are these always (D)?
    rep.append('【位置分布】這類選項排在哪一格')
    pc = Counter(r['letter'] for r in rows)
    for L in 'ABCD':
        rep.append(f'  ({L}) {pc.get(L,0):>4} 次')

    # counter-examples: option present but NOT the answer
    miss = [r for r in rows if not r['hit']]
    rep.append(f'\n【反例】有全稱選項、但正解不是它：{len(miss)} 題')
    mm = io.StringIO()
    mm.write(f'反例清單（{len(miss)} 題）——這些題「跟著選以上皆是/皆非」會錯\n')
    mm.write('格式：屆次 題號 科目 ｜ 全稱選項在哪格（類型）｜ 正解 ｜ 題幹\n\n')
    for r in sorted(miss, key=lambda r: (r['q']['exam'], r['q']['qno'])):
        q = r['q']
        mm.write(f'{q["exam"]} Q{q["qno"]:<2} {q["subject"]:<5}｜({r["letter"]}) '
                 f'{"肯定" if r["kind"]=="pos" else "否定"}｜正解 {q["ans"]}｜{q["head"][:78]}\n')
    open(os.path.join(HERE, '_aoa_counterexamples.txt'), 'w', encoding='utf-8').write(mm.getvalue())

    json.dump({'rows': [{'exam': r['q']['exam'], 'qno': r['q']['qno'],
                         'subject': r['q']['subject'], 'letter': r['letter'],
                         'kind': r['kind'], 'ans': r['q']['ans'], 'hit': r['hit'],
                         'head': r['q']['head'], 'opt': r['text']} for r in rows]},
              open(os.path.join(HERE, 'allofabove.json'), 'w', encoding='utf-8'),
              ensure_ascii=False)
    open(os.path.join(HERE, 'allofabove_report.txt'), 'w', encoding='utf-8').write('\n'.join(rep))
    print('\n'.join(rep[:12]))


if __name__ == '__main__':
    main()
