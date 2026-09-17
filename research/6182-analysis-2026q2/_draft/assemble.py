"""report_filled.html + data/charts/*.html -> 6182_FY2026Q2_Analysis/6182_FY2026Q2_Analysis.html"""
import os, re

D = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(D)
CH = os.path.join(ROOT, 'data', 'charts')
OUT = os.path.join(ROOT, '6182_FY2026Q2_Analysis', '6182_FY2026Q2_Analysis.html')

s = open(os.path.join(D, 'report_filled.html'), encoding='utf-8').read()

FUNDING = '''    <div class="chart-wrap">
      <div class="chart-title">擴產資金來源組成（億元）</div>
      {{CHART:funding_stack}}
    </div>
    <div class="note">⚑ 解讀：股權端（現增 5.56 億＋CB 轉換 11.51 億，合計約 17.1 億）只占三年來長期借款增幅 50.2 億的三分之一，擴產的主力資金仍是銀行聯貸——這也是利息保障倍數在 2026Q1 一度跌到 1.01 倍的背景。三個數字期間不同（現增、CB 為 2026 上半年，借款為 2023Q3→2026Q2 累計），此圖用來看量級而非同期比較。出處：2026Q2 合併財報附註（股本、可轉換公司債、長期借款）。</div>

'''
SUBS = '''
    <div class="chart-wrap">
      <div class="chart-title">大陸轉投資事業 2026 上半年損益（億元）</div>
      {{CHART:subsidiary_pnl}}
    </div>
    <div class="note">⚑ 解讀：陸廠獲利集中在上海晶盟（磊晶片 +2.28 億）與上海合晶（+0.86 億），鄭州合晶 −0.44 億、揚州合晶 −0.12 億仍在虧損；而母公司的 84.4 億背書保證正掛在鄭州等陸廠身上——虧損的廠與擔保曝險重疊，是風險矩陣「背書保證與子公司虧損」一列的量化依據。數字為被投資公司本期損益，非合晶依持股認列金額。出處：2026Q2 合併財報附註十三「大陸投資資訊」。</div>
'''

anchor = '    <div class="card">\n      <div class="card-title">核心疑問'
assert s.count(anchor) == 1
s = s.replace(anchor, FUNDING + anchor)

# 業務部門 tab：bu_mix 解讀之後插子公司損益
i = s.index('id="panel-bu"')
j = s.index('{{CHART:bu_mix}}', i)
k = s.index('</div>', s.index('<div class="note">', j)) + len('</div>')
s = s[:k] + '\n' + SUBS + s[k:]

s = re.sub(r'<!--[^>]*\{\{[^>]*-->\n?', '', s)


def chart(m):
    return open(os.path.join(CH, m.group(1) + '.html'), encoding='utf-8').read()


s = re.sub(r'\{\{CHART:([a-z_]+)\}\}', chart, s)
SRC = [(r'ar2025\.txt\s*~?\d*', '2025 年報'), (r'q2_pdf\.txt', '2026Q2 合併財報'),
       (r'derived?_metrics\.json', 'MOPS 各季合併財報（本報告計算）'), (r'quarterly\.json', 'MOPS 各季合併財報'),
       (r'monthly_revenue\.json', 'MOPS 上櫃月營收'), (r'peers_yahoo\.json|peers_summary\.txt', 'Yahoo Finance'),
       (r'insiders_summary\.txt|insiders_raw\.json', 'MOPS 內部人持股異動月報'),
       (r'qualitative_notes\.md|industry_notes\.md', '本報告研究筆記')]
for pat, rep in SRC:
    s = re.sub(pat, rep, s)
left = re.findall(r'\{\{[A-Z]+:[a-z_]+\}\}', s)
assert not left, left
os.makedirs(os.path.dirname(OUT), exist_ok=True)
open(OUT, 'w', encoding='utf-8').write(s)
print('ok', len(s), OUT)
