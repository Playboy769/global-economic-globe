import re, pathlib
R = pathlib.Path(__file__).resolve().parents[2]
D = R/'3532-analysis-2026q2'/'data'
tpl = (R/'6488-analysis-2026q2'/'6488_FY2026Q2_Analysis'/'6488_FY2026Q2_Analysis.html').read_text(encoding='utf-8')
head = tpl[:tpl.index('</head>')]
head = re.sub(r'<title>.*?</title>', '<title>台勝科（3532）— FY2026 Q2 財報完整分析</title>', head)
qual = (D/'fragments'/'qual.html').read_text(encoding='utf-8')
quant = (D/'fragments'/'quant.html').read_text(encoding='utf-8')
# capex_gantt skipped (no dated intervals disclosed): keep note, drop chart frame
qual = qual.replace('''  <div class="chart-wrap">
    <div class="chart-title">資本支出承諾與折舊費用時程（gantt-chart）</div>
    <!--CHART:capex_gantt-->
    <div class="note">⚑ 解讀：這張圖要呈現的是''', '''  <div class="note">⚑ 時序觀察（財報僅揭露期末承諾餘額、無到期分層，故不繪時程圖）：''')
qual = qual.replace('''（Dep 逐季序列）。</div>
  </div>''', '''（Dep 逐季序列）。</div>''', 1)
# bu_bar belongs to quant panel-bu; qual keeps only the cross-reference note
qual = qual.replace('''  <div class="chart-wrap">
    <div class="chart-title">月營收 YoY 動能（台股必備圖，由量化層損益表 tab 呈現實際圖表）</div>
    <!--CHART:bu_bar-->
''', '  <div>\n')
qual = qual.replace('Q2 單季營收 YoY +20.6%', 'Q2 單季營收 YoY +20.6%')
body = qual + '\n' + quant
def chart(m):
    p = D/'charts'/f'{m.group(1)}.html'
    return p.read_text(encoding='utf-8') if p.exists() else f'<!-- missing chart {m.group(1)} -->'
body = re.sub(r'<!--CHART:([a-z_]+)-->', chart, body)
body = body.replace('<div class="panel" id="panel-event">', '<div class="panel active" id="panel-event">', 1)
tabs = [('event','擴產與資本支出'),('tech','技術與產品線'),('supply','供應鏈'),('risk','風險矩陣'),
        ('is','損益表'),('bu','業務部門'),('bs','資產負債表'),('cf','現金流量'),('val','估值觀察')]
tabhtml = '\n'.join(f'    <div class="tab{" active" if i==0 else ""}" onclick="sw(\'{k}\',this)">{n}</div>' for i,(k,n) in enumerate(tabs))
header = '''<div class="page-header">
  <h1>台灣勝高科技股份有限公司（TWSE: 3532）— FY2026 Q2 財報完整分析</h1>
  <p>擴產與資本支出事件拆解 / 技術與產品線 / 供應鏈（含 SUMCO 與台塑關係人交易）/ 風險矩陣 / 損益表 / 業務部門 / 資產負債表 / 現金流量 / 估值觀察（共 9 個分頁；Fallback 模式省略 Q&amp;A 分頁）
  &nbsp;·&nbsp; 主軸期間 <strong>2026 年第二季（115Q2）</strong>
  &nbsp;·&nbsp; 資料來源：115Q2 合併財務報告暨會計師核閱報告（MOPS t164sb01 iXBRL ＋ doc.twse 全文 PDF）、2023Q3–2026Q2 共 12 季財報、114 年度年報、MOPS 上市月營收（113/7–115/8 共 26 個月，經累計欄逆算交叉驗證）、母公司 SUMCO FY26Q1 分析報告交叉引用、Yahoo Finance 同業估值（2026-09-16）
  &nbsp;·&nbsp; 幣別新台幣；季度數字僅經核閱未經查核
  &nbsp;·&nbsp; <strong>Fallback 模式</strong>：公司未公布法說會逐字稿，質化判斷以書面揭露＋財報數字推導為主。</p>
</div>'''
footer = '''  <div class="footer">本報告依據台灣勝高科技股份有限公司 115 年第二季財務報告暨會計師核閱報告（MOPS t164sb01 iXBRL 與 doc.twse 全文 PDF）、2023Q3–2026Q2 共 12 季財報、114 年度年報、MOPS 上市月營收公告（113/7–115/8，均經累計欄逆算交叉驗證）、SUMCO CORPORATION FY26Q1 分析報告與 Yahoo Finance 同業估值（2026-09-16）整理。本次分析無法說會逐字稿（Fallback 模式）；產能利用率、未開始租賃、承諾到期分層、產品別量價等項目查無公開揭露，已於對應 tab 標註「無此資料」，不代表不存在；內部人持股月報查詢端點僅回傳靜態快照，未能取得逐月異動。Sankey 僅畫出 filing 揭露之進銷貨占比連結。僅供資訊參考，非投資建議。</div>'''
script = tpl[tpl.index('<script>\nfunction sw('):]
out = head + '</head>\n<body>\n' + header + '\n<div class="container">\n  <div class="tabs">\n' + tabhtml + '\n  </div>\n\n' + body + '\n' + footer + '\n</div>\n\n' + script
(R/'3532-analysis-2026q2'/'3532_FY2026Q2_Analysis'/'3532_FY2026Q2_Analysis.html').write_text(out, encoding='utf-8')
print(len(out), body.count('missing chart'), len(re.findall(r'class="panel( active)?"', out)))
