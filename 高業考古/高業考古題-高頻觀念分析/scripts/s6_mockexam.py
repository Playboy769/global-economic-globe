"""Build a mock paper for the next 屆 by probability-weighted sampling.

    python s6_mockexam.py [seed]

Sampling is weighted by each question's modelled reappearance probability, WITHOUT
replacement (Efraimidis–Spirakis A-Res). High-probability questions are likely but
not guaranteed to appear — that is the point of simulating rather than ranking.

Only clusters with a printable A–D option set AND a single-letter official answer
are eligible, so every question on the paper can actually be graded.
"""
import json, os, random, sys, html
from collections import defaultdict, Counter
from s2_cluster import cluster, EXAMS, IDX, N, SUBJECTS

HERE = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'out')
os.makedirs(HERE, exist_ok=True)
OUT_HTML = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                        '..', '115-3_模擬考卷.html')
SEED = int(sys.argv[1]) if len(sys.argv) > 1 else 20261115
NEXT_EXAM = '115-3'
PER_SUBJECT = 50


def kbin(k):
    return min(k, 4)


def gbin(g):
    return 1 if g <= 1 else 2 if g <= 2 else 3 if g <= 4 else 4 if g <= 8 else 5


def build_hazard(items, upto):
    hz = defaultdict(lambda: [0, 0])
    for c in items:
        s = set(x for x in c['slots'] if x < upto)
        if not s:
            continue
        k, last = 0, None
        for t in range(min(s), upto):
            if last is not None:
                key = (kbin(k), gbin(t - last))
                hz[key][1] += 1
                if t in s:
                    hz[key][0] += 1
            if t in s:
                k += 1
                last = t
    return hz


def optsig(o):
    return '||'.join(b for _, b in o) if o else None


def main():
    rng = random.Random(SEED)
    recs = json.load(open(os.path.join(HERE, 'corpus.json'), encoding='utf-8'))
    paper = {}
    stats = {}

    for subj in SUBJECTS:
        rs = [r for r in recs if r['subject'] == subj]
        groups = cluster(rs)
        items = []
        for g in groups:
            members = [rs[i] for i in g]
            items.append({'slots': sorted({IDX[m['exam']] for m in members}),
                          'members': members})
        hz = build_hazard(items, N)
        base = sum(v[0] for v in hz.values()) / max(1, sum(v[1] for v in hz.values()))

        cands = []
        for c in items:
            k, g = len(c['slots']), N - max(c['slots'])
            h, n = hz.get((kbin(k), gbin(g)), [0, 0])
            p = h / n if n >= 30 else base
            # display version: latest occurrence that is both printable and gradable
            usable = [m for m in c['members']
                      if m['opts'] and len(m['ans']) == 1 and m['ans'] in 'ABCD']
            if not usable:
                continue
            show = max(usable, key=lambda m: (IDX[m['exam']], m['qno']))
            sigs = {optsig(m['opts']) for m in c['members'] if m['opts']}
            cands.append({'p': p, 'k': k, 'g': g,
                          'exams': [EXAMS[i] for i in c['slots']],
                          'head': show['head'], 'opts': show['opts'],
                          'ans': show['ans'], 'from': show['exam'],
                          'variants': len(sigs)})

        # Efraimidis–Spirakis weighted sampling without replacement
        keyed = sorted(cands, key=lambda c: -(rng.random() ** (1.0 / max(c['p'], 1e-9))))
        picked = keyed[:PER_SUBJECT]
        rng.shuffle(picked)
        paper[subj] = picked
        stats[subj] = {
            'pool': len(cands),
            'exp_hits': sum(c['p'] for c in cands),
            'mean_p': sum(c['p'] for c in picked) / len(picked),
            'rewritten': sum(1 for c in picked if c['variants'] > 1),
            'k_dist': Counter(min(c['k'], 5) for c in picked),
        }

    write_html(paper, stats)
    lines = [f'seed={SEED}  next={NEXT_EXAM}']
    for s in SUBJECTS:
        st = stats[s]
        lines.append(f'{s}: 候選池 {st["pool"]} 組 → 抽 {PER_SUBJECT} 題 | '
                     f'平均命中機率 {st["mean_p"]*100:.1f}% | '
                     f'選項曾被改寫 {st["rewritten"]} 題 | '
                     f'出現次數分布 {dict(sorted(st["k_dist"].items()))}')
    open(os.path.join(HERE, 'mockexam_report.txt'), 'w', encoding='utf-8').write('\n'.join(lines))
    print('\n'.join(lines))
    print('->', os.path.normpath(OUT_HTML))


CSS = """
:root{color-scheme:light}
*{box-sizing:border-box}
body{margin:0;background:#fff;color:#1a1a1a;
 font:16px/1.75 "Noto Sans TC","PingFang TC","Microsoft JhengHei",system-ui,sans-serif}
.wrap{max-width:900px;margin:0 auto;padding:28px 20px 80px}
h1{font-size:26px;font-weight:600;margin:0 0 10px;letter-spacing:.02em}
.sub{color:#555;font-size:14px;margin-bottom:18px}
.note{border-left:3px solid #999;background:#fafafa;padding:12px 16px;
 font-size:14px;color:#333;margin-bottom:20px}
.note b{color:#111}
h2{font-size:20px;font-weight:600;margin:34px 0 6px;padding-bottom:8px;
 border-bottom:2px solid #1a1a1a}
.shint{font-size:13px;color:#666;margin-bottom:18px}
.q{border:1px solid #e2e2e2;border-radius:6px;padding:14px 16px;margin-bottom:18px}
.qh{display:flex;gap:10px;align-items:baseline;margin-bottom:10px}
.qn{flex:0 0 auto;font-weight:600;color:#444;font-variant-numeric:tabular-nums}
.qt{flex:1 1 auto}
.opt{display:block;padding:7px 10px;border-radius:4px;cursor:pointer;margin-bottom:2px}
.opt:hover{background:#f4f4f4}
.opt input{margin-right:9px}
.opt .L{font-weight:600;color:#444;margin-right:5px}
.bar{position:sticky;top:0;z-index:9;background:#fff;border-bottom:1px solid #ddd;
 padding:11px 0;margin-bottom:18px;display:flex;gap:14px;align-items:center;flex-wrap:wrap}
.bar .prog{font-size:14px;color:#333;font-variant-numeric:tabular-nums}
button{font:inherit;font-size:15px;padding:8px 20px;border-radius:5px;
 border:1px solid #1a1a1a;background:#1a1a1a;color:#fff;cursor:pointer}
button.ghost{background:#fff;color:#1a1a1a}
button:disabled{opacity:.4;cursor:not-allowed}
.warn{color:#b45309;font-size:14px}
.q.right{border-color:#15803d;border-left:4px solid #15803d}
.q.wrong{border-color:#b91c1c;border-left:4px solid #b91c1c}
.mark{font-size:13px;font-weight:600;margin-left:auto;flex:0 0 auto}
.right .mark{color:#15803d}
.wrong .mark{color:#b91c1c}
.opt.sol{background:#f0fdf4;border:1px solid #86efac}
.opt.pick{background:#fef2f2;border:1px solid #fca5a5}
#result{display:none;border:2px solid #1a1a1a;border-radius:6px;padding:18px 20px;margin-bottom:20px}
#result h3{margin:0 0 12px;font-size:20px}
.score{font-size:34px;font-weight:700;font-variant-numeric:tabular-nums}
.stab{width:100%;border-collapse:collapse;font-size:14px;margin-top:12px}
.stab th,.stab td{border:1px solid #ddd;padding:7px 10px;text-align:left}
.stab th{background:#f5f5f5;font-weight:600}
.stab td.n{text-align:right;font-variant-numeric:tabular-nums}
#answers{display:none}
.a{border:1px solid #e2e2e2;border-left:3px solid #999;border-radius:5px;
 padding:11px 14px;margin-bottom:16px;font-size:14px}
.a .t{font-weight:600;color:#111;margin-bottom:5px}
.a .m{color:#555}
.a .m span{margin-right:16px;white-space:nowrap}
.badge{display:inline-block;font-size:12px;padding:1px 8px;border-radius:10px;
 border:1px solid #b45309;color:#b45309;margin-left:6px}
.correct{color:#15803d;font-weight:600}
.foot{margin-top:40px;padding-top:16px;border-top:1px solid #ddd;
 font-size:13px;color:#666}
@media print{.bar,button{display:none}#answers,#result{display:block!important}}
"""

JS = """
const KEY=__KEY__, TOT=__TOT__, SUBJ=__SUBJ__;
let done=false;
function answered(){return document.querySelectorAll('input[type=radio]:checked').length}
function refresh(){
  const n=answered();
  document.getElementById('prog').textContent=`已作答 ${n} / ${TOT}`;
  document.getElementById('submit').disabled=done;
}
document.addEventListener('change',e=>{if(e.target.type==='radio')refresh()});
function submit(){
  if(done)return;
  const n=answered();
  if(n<TOT){
    const miss=[];
    for(let i=0;i<TOT;i++){if(!document.querySelector(`input[name=q${i}]:checked`))miss.push(i+1)}
    const show=miss.slice(0,12).join('、')+(miss.length>12?` …等 ${miss.length} 題`:'');
    if(!confirm(`還有 ${TOT-n} 題未作答（第 ${show}）。\\n仍要交卷嗎？`))return;
  }
  done=true;
  let right=0; const per={};
  SUBJ.forEach(s=>per[s]=[0,0]);
  for(let i=0;i<TOT;i++){
    const box=document.getElementById('q'+i);
    const sol=KEY[i][0], s=KEY[i][1];
    const sel=document.querySelector(`input[name=q${i}]:checked`);
    const got=sel?sel.value:null;
    per[s][1]++;
    if(got===sol){right++;per[s][0]++;box.classList.add('right');
      box.querySelector('.mark').textContent='✓ 答對';}
    else{box.classList.add('wrong');
      box.querySelector('.mark').textContent=got?`✗ 你選 ${got}，正解 ${sol}`:`— 未作答，正解 ${sol}`;}
    box.querySelector(`.opt[data-v="${sol}"]`).classList.add('sol');
    if(got&&got!==sol)box.querySelector(`.opt[data-v="${got}"]`).classList.add('pick');
    box.querySelectorAll('input').forEach(x=>x.disabled=true);
  }
  document.getElementById('score').textContent=(right/TOT*100).toFixed(1)+' 分';
  document.getElementById('rightn').textContent=`答對 ${right} / ${TOT} 題`;
  const tb=document.getElementById('sbody'); tb.innerHTML='';
  SUBJ.forEach(s=>{const [r,t]=per[s];
    tb.insertAdjacentHTML('beforeend',
      `<tr><td>${s}</td><td class="n">${r} / ${t}</td><td class="n">${(r/t*100).toFixed(0)}%</td>
       <td class="n">${t*2}</td><td class="n">${r*2}</td></tr>`)});
  document.getElementById('result').style.display='block';
  document.getElementById('answers').style.display='block';
  document.getElementById('submit').disabled=true;
  document.getElementById('result').scrollIntoView({behavior:'smooth'});
}
refresh();
"""


def write_html(paper, stats):
    E = html.escape
    key, body, ansblk = [], [], []
    idx = 0
    for subj in SUBJECTS:
        st = stats[subj]
        body.append(f'<h2>{E(subj)}</h2>')
        body.append(f'<div class="shint">共 {PER_SUBJECT} 題，每題 2 分。'
                    f'候選池 {st["pool"]} 組題目，依模型命中機率加權隨機抽出；'
                    f'本卷平均命中機率 {st["mean_p"]*100:.1f}%。</div>')
        ansblk.append(f'<h2>{E(subj)} · 解答</h2>')
        for j, q in enumerate(paper[subj], 1):
            key.append([q['ans'], subj])
            opts = ''.join(
                f'<label class="opt" data-v="{l}">'
                f'<input type="radio" name="q{idx}" value="{l}">'
                f'<span class="L">({l})</span>{E(b)}</label>'
                for l, b in q['opts'])
            body.append(
                f'<div class="q" id="q{idx}"><div class="qh">'
                f'<span class="qn">{j}.</span><span class="qt">{E(q["head"])}</span>'
                f'<span class="mark"></span></div>{opts}</div>')
            warn = ('<span class="badge">選項曾被改寫，勿背字母</span>'
                    if q['variants'] > 1 else '')
            ansblk.append(
                f'<div class="a"><div class="t">{j}. 正解 '
                f'<span class="correct">({q["ans"]})</span>{warn}</div>'
                f'<div class="m"><span>歷次出現：{"、".join(q["exams"])}（共 {q["k"]} 次）</span>'
                f'<span>距上次 {q["g"]} 屆</span>'
                f'<span>命中機率 {q["p"]*100:.1f}%</span>'
                f'<span>題面取自 {q["from"]}</span></div></div>')
            idx += 1

    tot = idx
    js = (JS.replace('__KEY__', json.dumps(key, ensure_ascii=False))
            .replace('__TOT__', str(tot))
            .replace('__SUBJ__', json.dumps(SUBJECTS, ensure_ascii=False)))
    grand = sum(stats[s]['exp_hits'] for s in SUBJECTS)

    doc = f"""<!DOCTYPE html>
<html lang="zh-Hant"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>{NEXT_EXAM} 模擬考卷 · 證券商高級業務員</title>
<style>{CSS}</style></head><body><div class="wrap">
<h1>{NEXT_EXAM} 模擬考卷</h1>
<div class="sub">證券商高級業務員資格測驗 · 投資學／財務分析／證券交易相關法規與實務 · 共 {tot} 題，每題 2 分</div>
<div class="note">
<b>這份考卷是怎麼來的。</b>把 105–115 年 37 屆 5,550 題全部比對出重複題組，
用「已出現 k 次、距上次 g 屆」推估每道題在下一屆的命中機率，再<b>依機率加權隨機抽樣</b>而成
（隨機種子 {SEED}，可重現）。所以它模擬的是「115-3 有可能長什麼樣」，
<b>不是押題</b>——模型估計整屆約 {grand/3:.0f}／50 題會是舊題重出，但單題最高命中機率也不到 10%。<br>
題面與正解取自該題<b>最近一次有答案卡的出現版本</b>（109 年起才有答案卡）。
法規題可能因修法而答案已變動，作答後請以最新法條為準。
</div>
<div class="bar">
  <span class="prog" id="prog"></span>
  <button id="submit" onclick="submit()">交卷</button>
  <span class="warn">交卷後才會顯示分數與解答</span>
</div>
<div id="result">
  <h3>成績</h3>
  <div class="score" id="score"></div>
  <div id="rightn"></div>
  <table class="stab"><thead><tr><th>科目</th><th>答對</th><th>正確率</th>
  <th>滿分</th><th>得分</th></tr></thead><tbody id="sbody"></tbody></table>
</div>
{''.join(body)}
<div id="answers"><h2 style="margin-top:46px">解答與出處</h2>
<div class="shint">「命中機率」是模型估計該題出現在 {NEXT_EXAM} 的機率；
「歷次出現」可回查 <code>高業考古/</code> 底下的原卷。</div>
{''.join(ansblk)}</div>
<div class="foot">由 <code>scripts/s6_mockexam.py</code> 產生（seed {SEED}）。
重跑可換一份不同的抽樣結果。本文僅供個人準備參考，非考試預測，法規內容請以主管機關最新公告為準。</div>
</div><script>{js}</script></body></html>"""
    with open(OUT_HTML, 'w', encoding='utf-8', newline='\n') as f:
        f.write(doc)


if __name__ == '__main__':
    main()
