"""Among questions whose STEM repeats verbatim, how often are the OPTIONS rewritten?
That is the quantitative version of the 統整檔's '不要背選項字母' warning.
Also produces the bias-corrected 115-3 forecast."""
import fitz, glob, os, re, json, unicodedata
from collections import defaultdict, Counter
from s1_extract import SUBJ_PAT, HEAD_PAT, subject_of, split_questions, norm_stem, BASE

HERE = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'out')
os.makedirs(HERE, exist_ok=True)
rep = []


def norm_opts(raw):
    t = unicodedata.normalize('NFKC', raw)
    m = re.search(r'[（(]\s*[AＡ]\s*[)）]', t)
    if not m:
        return None
    t = t[m.start():]
    t = re.sub(r'\s+', '', t)
    t = re.sub(r'[，。：；？（）()\[\]「」『』、,.:;?!\'"~*#_\-─—－]', '', t)
    return t


rows = []
for y in range(105, 116):
    for p in sorted(glob.glob(f'{BASE}/{y}/*.pdf')):
        name = os.path.basename(p)
        if re.fullmatch(r'\d{5}a\.pdf', name) or '說明' in name:
            continue
        d = fitz.open(p)
        text = unicodedata.normalize('NFC', ''.join(pg.get_text() for pg in d))
        d.close()
        hm = HEAD_PAT.search(text)
        if not hm:
            continue
        exam = f'{int(hm.group(1))}-{int(hm.group(2))}'
        marks = [(m.start(), subject_of(m.group(1))) for m in SUBJ_PAT.finditer(text)]
        marks = [(s, sub) for s, sub in marks if sub]
        for i, (s, sub) in enumerate(marks):
            e = marks[i + 1][0] if i + 1 < len(marks) else len(text)
            for n, raw in split_questions(text[s:e]):
                st = norm_stem(raw)
                if len(st) < 3:
                    continue
                rows.append({'exam': exam, 'subject': sub, 'qno': n,
                             'stem': st, 'opts': norm_opts(raw)})

rep.append('【題幹重複時，選項是否也一字不差】')
rep.append(f'{"科目":>6} {"題幹重複題數":>12} {"選項也相同":>10} {"選項被改寫":>10} {"改寫比例":>9}')
tot_same = tot_diff = 0
for subj in ['投資學', '財務分析', '法規']:
    rs = [r for r in rows if r['subject'] == subj]
    by = defaultdict(list)
    for r in rs:
        by[r['stem']].append(r)
    same = diff = 0
    examples = []
    for stem, g in by.items():
        exams = {x['exam'] for x in g}
        if len(exams) < 2:
            continue
        # a question counts as "選項也相同" if its (stem, options) pair itself
        # recurs across >=2 exams — matching 統整檔's 「題幹＋選項一字不差」 wording
        byopt = defaultdict(set)
        for x in g:
            byopt[x['opts']].add(x['exam'])
        for x in g:
            if len(byopt[x['opts']]) >= 2:
                same += 1
            else:
                diff += 1
        if len(byopt) > 1 and len(examples) < 3 and len(stem) > 12:
            examples.append((stem, sorted(exams), len(byopt)))
    tot_same += same
    tot_diff += diff
    rep.append(f'{subj:>6} {same+diff:>12} {same:>10} {diff:>10} '
               f'{diff/(same+diff)*100:>8.1f}%')
    for stem, exams, k in examples:
        rep.append(f'    例：{stem[:52]}… 出現於 {"、".join(exams)}，共 {k} 種選項版本')
rep.append(f'{"合計":>6} {tot_same+tot_diff:>12} {tot_same:>10} {tot_diff:>10} '
           f'{tot_diff/(tot_same+tot_diff)*100:>8.1f}%')
rep.append(f'\n  題幹重複題共 {tot_same+tot_diff} 題 = 全部 5550 題的 '
           f'{(tot_same+tot_diff)/5550*100:.1f}%')
rep.append(f'  其中選項完全沒動的 {tot_same} 題 = {tot_same/5550*100:.1f}%'
           f'（這才是統整檔「題幹＋選項一字不差」口徑，其 48.6% 應與此數對照）')

# ---------- bias-corrected forecast ----------
rep.append('\n\n【115-3 預測的偏誤校正】')
A3 = json.load(open(os.path.join(HERE, 'analysis3.json'), encoding='utf-8'))
BIAS = {s: A3[s]['backtest']['bias'] for s in ['投資學', '財務分析', '法規']}
RAW = {s: A3[s]['forecast_raw'] for s in ['投資學', '財務分析', '法規']}
MAE = {s: A3[s]['backtest']['mae'] for s in ['投資學', '財務分析', '法規']}
rep.append(f'{"科目":>6} {"原始預測":>9} {"回測偏誤":>9} {"校正後":>8} {"合理區間":>14}')
tot = 0
for s in ['投資學', '財務分析', '法規']:
    adj = RAW[s] - BIAS[s]
    tot += adj
    rep.append(f'{s:>6} {RAW[s]:>9.1f} {BIAS[s]:>+9.1f} {adj:>8.1f} '
               f'{adj-MAE[s]:>6.0f} ~ {adj+MAE[s]:<5.0f}')
rep.append(f'{"合計":>6} {sum(RAW.values()):>9.1f} {-sum(BIAS.values()):>+9.1f} {tot:>8.1f}'
           f'   / 150 題 = {tot/150*100:.0f}%')
rep.append('  模型系統性高估，原因與固定視窗分析一致：訓練資料含較高回收率的早期年份，')
rep.append('  近年回收強度實際上在下降（法規最明顯），所以未校正的預測偏高。')

open(os.path.join(HERE, 'analysis4_report.txt'), 'w', encoding='utf-8').write('\n'.join(rep))
print('written')
