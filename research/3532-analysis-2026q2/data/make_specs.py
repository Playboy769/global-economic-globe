import json, os, sys

D = os.path.dirname(os.path.abspath(__file__)) + '/'
OUT = sys.argv[1] if len(sys.argv) > 1 else D + 'chart_specs'

r = json.load(open(D + 'quarterly.json', encoding='utf-8'))
qs = sorted(r)
L = [q[2:] for q in qs]
B = lambda v: round(v / 1e8, 2)

S = {}

# 1. IS trend (12Q revenue + GM% + OPM%)
S['is_trend_line'] = {
    'title': '台勝科 近 12 季營收與利潤率（2023Q3–2026Q2）',
    'caption': '左軸：單季營收（億元）；右軸：毛利率／營業利益率（%）。來源：各季合併財報（MOPS t164sb01）。2026Q2 毛利率降至 6.9%，為 12 季最低。',
    'leftLabel': '營收（億元）', 'rightLabel': '%', 'categories': L,
    'series': [
        {'name': '營收', 'axis': 'left', 'values': [B(r[q]['Rev']) for q in qs]},
        {'name': '毛利率', 'axis': 'right', 'values': [round(r[q]['GP'] / r[q]['Rev'] * 100, 1) for q in qs]},
        {'name': '營業利益率', 'axis': 'right', 'values': [round(r[q]['OpInc'] / r[q]['Rev'] * 100, 1) for q in qs]},
    ],
    'opts': {'grid': True, 'markers': True, 'values': False, 'leftZero': True, 'rightZero': True}
}

# 2. Monthly revenue + YoY
m_raw = json.load(open(D + 'monthly_revenue.json', encoding='utf-8'))
m_keys = sorted(m_raw, key=lambda k: (m_raw[k]['ad_year'], m_raw[k]['month']))
S['monthly_rev_line'] = {
    'title': '台勝科 26 個月月營收與 YoY（2024/07–2026/08）',
    'caption': '左軸：單月營收（億元）；右軸：YoY（%）。2026/07、08 為 Q2 財報期間之後的最新月份，尚未被任何一期財報涵蓋，YoY 已連續轉正並擴大（21.95%→29.27%）。來源：MOPS 上市月營收彙總表 t21sc03_sii，各月均以累計欄反算交叉驗證。',
    'leftLabel': '月營收（億元）', 'rightLabel': 'YoY %',
    'categories': [k[2:].replace('-', '/') for k in m_keys],
    'series': [
        {'name': '月營收', 'axis': 'left', 'values': [round(m_raw[k]['cur_month_rev'] / 1e5, 2) for k in m_keys]},
        {'name': 'YoY', 'axis': 'right', 'values': [m_raw[k]['yoy_pct'] for k in m_keys]},
    ],
    'opts': {'grid': True, 'markers': True, 'leftZero': True, 'rightZero': False}
}

# 3. BU / segment mix -- only quantified split disclosed is domestic vs export (annual report)
S['bu_bar'] = {
    'title': '內外銷營收占比：114 年度 vs 113 年度',
    'caption': '單位：億元。年報未揭露 8 吋／12 吋個別產品之營收拆分（僅文字敘述占比趨勢），唯一量化的產品線相關指標為內外銷比例。114 年外銷占比降至 15.8%（113 年 19.0%），內銷集中度上升。來源：114 年度股東年報「二、市場及產銷概況（一）市場分析 1.主要產品別內外銷比例及地區」。',
    'unit': '億元', 'categories': ['內銷', '外銷'], 'mode': 'grouped', 'orient': 'v',
    'series': [
        {'name': '113年度', 'values': [round(10088991/1e5, 2), round(2332644/1e5, 2)]},
        {'name': '114年度', 'values': [round(10381565/1e5, 2), round(1952727/1e5, 2)]},
    ],
    'opts': {'grid': True, 'values': True, 'legend': True}
}

# 4. FCF waterfall (2026 H1 YTD)
pbt_ytd = r['2026Q1']['PBT'] + r['2026Q2']['PBT']
dep_amort_ytd = r['2026Q2']['Dep_ytd'] + r['2026Q2']['Amort_ytd']
cfo_ytd = r['2026Q2']['CFO_ytd']
capex_ytd = r['2026Q2']['Capex_ytd']
items_raw = [
    ('稅前淨利（H1 YTD）', pbt_ytd, 'base'),
    ('折舊攤銷', dep_amort_ytd, 'delta'),
    ('存貨增加', r['2026Q2']['dInv_ytd'], 'delta'),
    ('應收帳款增加', r['2026Q2']['dAR_ytd'], 'delta'),
    ('應付帳款增加', r['2026Q2']['dAP_ytd'], 'delta'),
    ('合約負債減少', r['2026Q2']['dContractLiab_ytd'], 'delta'),
    ('支付所得稅', r['2026Q2']['TaxPaid_ytd'], 'delta'),
    ('支付利息', r['2026Q2']['IntPaid_ytd'], 'delta'),
    ('租賃本金償還', r['2026Q2']['Lease_ytd'], 'delta'),
]
plug = cfo_ytd - sum(v for _, v, _ in items_raw)
items = [{'name': n, 'value': round(v / 1e8, 2), 'type': t} for n, v, t in items_raw]
items.append({'name': '其他營運調整項目（含利息收入等）', 'value': round(plug / 1e8, 2), 'type': 'delta'})
items.append({'name': '營運現金流（H1 YTD）', 'value': round(cfo_ytd / 1e8, 2), 'type': 'total'})
items.append({'name': '資本支出', 'value': round(capex_ytd / 1e8, 2), 'type': 'delta'})
items.append({'name': '自由現金流（H1 YTD）', 'value': round((cfo_ytd + capex_ytd) / 1e8, 2), 'type': 'total'})
S['fcf_waterfall'] = {
    'title': '2026 上半年（H1 YTD）自由現金流橋接（億元）',
    'caption': '稅前淨利 H1 僅 −0.2 億（Q1 虧損 Q2 微幅獲利相抵），營運現金流靠折舊攤銷（17.1 億）與營運資金貢獻撐至 9.7 億，扣資本支出 5.8 億後 FCF 約 3.9 億。來源：2026Q2 合併現金流量表（H1 累計，Q1+Q2 相加或直接讀 YTD 欄位）。',
    'unit': '億元', 'items': items, 'opts': {'values': True, 'conn': True, 'grid': True}
}

# 5. Peer valuation scatter (USD-normalized bubble size; fwd P/E except Siltronic EV/EBITDA)
S['peer_scatter'] = {
    'title': '矽晶圓同業：預估本益比 vs 營收成長（2026-09-16）',
    'caption': 'X：最近一季營收 YoY（%）；Y：Forward P/E（倍，Siltronic 因 Forward P/E 為負改用 EV/EBITDA）；氣泡＝市值（十億美元，以 2026-09-16 USD/TWD 31.748、USD/JPY 155.05、EUR/USD 1.1537 換算，僅供跨幣別相對比較）。台勝科 Trailing P/E 高達 338 倍（獲利趨近於零）不具參考性，故採 Forward P/E。來源：Yahoo Finance quoteSummary（peers_yahoo.json）。',
    'xName': '營收 YoY（%）', 'yName': 'Forward P/E（倍，Siltronic 為 EV/EBITDA）', 'refMode': 'median',
    'points': [
        {'name': '台勝科 3532', 'x': 20.6, 'y': 74.58, 'size': 5.38},
        {'name': '環球晶 6488', 'x': -5.0, 'y': 28.45, 'size': 14.11},
        {'name': '合晶 6182', 'x': 7.4, 'y': 38.23, 'size': 2.21},
        {'name': 'SUMCO 3436', 'x': 10.4, 'y': 60.46, 'size': 6.71},
        {'name': '信越化學 4063', 'x': 5.4, 'y': 17.65, 'size': 68.13},
        {'name': 'Siltronic WAF', 'x': -2.3, 'y': 11.88, 'size': 2.65},
    ],
    'opts': {'labels': True, 'grid': True, 'trend': False, 'xZero': False, 'yZero': True}
}

# 6. Interest coverage (12Q)
S['interest_coverage_line'] = {
    'title': '利息保障倍數 12 季（營業利益 ÷ 財務成本）',
    'caption': '倍數＝營業利益÷財務成本（>60 倍者截尾於 60，避免壓縮近期崩落的視覺）。2026Q1 營業利益轉虧為負值（顯示為 0 並標註），2026Q2 倍數僅 1.6x，為 12 季最低，與長期借款維持 205 億高檔、財務成本逐季攀升同向惡化。來源：各季合併財報附註「財務成本」與損益表營業利益。',
    'leftLabel': '倍', 'rightLabel': '', 'categories': L,
    'series': [
        {'name': '營業利益÷財務成本', 'axis': 'left',
         'values': [round(min(r[q]['OpInc'] / r[q]['FinCost'], 60), 2) if r[q]['OpInc'] > 0 else 0 for q in qs]},
    ],
    'opts': {'grid': True, 'markers': True, 'values': True, 'leftZero': True}
}

# 7. DSO / inventory days (12Q)
S['dso_line'] = {
    'title': '應收帳款天數（DSO）與存貨天數 12 季',
    'caption': 'DSO＝期末應收帳款÷單季營收×91；存貨天數＝期末存貨÷單季營業成本×91。2026Q2 DSO 升至 70.9 天（12 季新高），與客戶集中度上升（客戶A占比擴大至 37.6%）方向一致；存貨天數同步走高反映稼動率與去化速度放緩。來源：各季合併資產負債表與綜合損益表。',
    'leftLabel': '天', 'rightLabel': '', 'categories': L,
    'series': [
        {'name': 'DSO', 'axis': 'left', 'values': [round(r[q]['AR'] / r[q]['Rev'] * 91, 1) for q in qs]},
        {'name': '存貨天數', 'axis': 'left', 'values': [round(r[q]['Inv'] / r[q]['COGS'] * 91, 1) for q in qs]},
    ],
    'opts': {'grid': True, 'markers': True, 'leftZero': True}
}

# 8. Contract liabilities (12Q, stacked)
S['contract_liab_line'] = {
    'title': '合約負債餘額 12 季走勢（台股 IFRS 無「預收貨款」科目，對應合約負債）',
    'caption': '單位：億元。流動＋非流動合計由 2023Q3 的 50.9 億持續消化至 2026Q2 的 26.0 億，降幅逾五成，與客戶集中度上升、A公司占比擴大並列觀察：主力客戶提前鎖單力道未見增強。來源：各季合併資產負債表（附註十九合約負債）。',
    'unit': '億元', 'categories': L, 'mode': 'stacked',
    'series': [
        {'name': '合約負債－流動', 'values': [B(r[q]['CL_contract']) for q in qs]},
        {'name': '合約負債－非流動', 'values': [B(r[q]['NCL_contract']) for q in qs]},
    ],
    'opts': {'grid': True, 'bounds': True, 'markers': False, 'bandLabels': False, 'totals': True, 'zero': True}
}

# 9. Risk heatmap
S['risk_heat'] = {
    'title': '風險矩陣：風險項 × 時間窗口（強度 1–5）',
    'caption': '數值為本報告主觀評分：5＝極可能且衝擊大。來源：本報告財報數字推導（Fallback 模式，無逐字稿）。',
    'unit': '',
    'rows': ['毛利率持續低迷', '客戶A集中', '利息保障惡化', '中國8吋產能', '母公司權利金/定價', '12吋需求遞延'],
    'cols': ['近1季', '1年', '3年'],
    'values': [
        [5, 4, 3],
        [3, 4, 5],
        [5, 4, 3],
        [2, 3, 4],
        [2, 2, 3],
        [3, 3, 2],
    ],
    'scale': 'diverging', 'mid': 3, 'opts': {'values': True, 'legend': True, 'border': True}
}

# 10. Supply chain structure (chain-diagram) -- required node labels verbatim
S['supply_chain_chain'] = {
    'title': '台勝科供應鏈結構：上游原料 → 一貫製程 → 下游客戶 → 終端應用',
    'caption': '關係人（SUMCO 體系與台塑）以紫色虛線標示，與一般供應/銷售實線區隔。終端應用層為產品層級之定性對應（8吋主力應用電源管理IC、12吋主力應用記憶體），非逐客戶揭露，故不帶數字、亦不進 Sankey。來源：2026Q2合併財報附註二六關係人交易；114年度股東年報「主要客戶及供應商」「產業上中下游關聯性」章節。',
    'stages': ['上游原料／設備', '台勝科（中游一貫廠）', '下游客戶', '終端應用（依產品定性對應）'],
    'branches': {
        'main': {'label': '實體供應／銷售流', 'color': '#2563eb'},
        'fin': {'label': '關係人（財務／技術授權，非供應流）', 'color': '#9333ea', 'dashed': True},
        'app': {'label': '產品-應用定性對應（無量化拆分）', 'color': '#16a34a', 'dashed': True},
    },
    'nodes': [
        {'id': 'sumco_ingot', 'stage': 0, 'title': 'SUMCO（單晶晶棒）', 'sub': '最終母公司，115年迄前一季進貨占比17.24%', 'note': '關係人；付款條件貨到30-120日', 'emphasis': True, 'branch': 'fin'},
        {'id': 'poly', 'stage': 0, 'title': '多晶矽供應商', 'sub': 'Tokuyama Corp，115年迄前一季進貨占比10.64%', 'note': '非關係人'},
        {'id': 'quartz', 'stage': 0, 'title': '石英坩堝與耗材', 'sub': '未揭露具體占比', 'note': '資料缺口：年報僅列前三大供應商'},
        {'id': 'equip', 'stage': 0, 'title': '設備供應商', 'sub': '未揭露具體占比', 'note': '資料缺口：年報僅列前三大供應商'},
        {'id': 'ftc', 'stage': 1, 'title': '台勝科 3532', 'sub': '8吋／12吋鏡面拋光矽晶圓、磊晶加工一貫廠', 'note': '國內一貫化生產', 'emphasis': True, 'span': True},
        {'id': 'custA', 'stage': 2, 'title': '客戶A', 'sub': '115年Q1銷貨占比37.58%（113年27.71%→114年36.02%持續上升）', 'emphasis': True},
        {'id': 'sumco_buy', 'stage': 2, 'title': 'SUMCO（回銷）', 'sub': '母公司SUMCO TECHXIV，115年Q1銷貨占比10.04%', 'note': '關係人；月結55天', 'branch': 'fin'},
        {'id': 'other_fab', 'stage': 2, 'title': '其他晶圓廠', 'sub': '含M公司12.93%、B公司12.58%等未逐一揭露客戶，合計約52.4%', 'note': '前十大客戶占應收比例未揭露'},
        {'id': 'mem', 'stage': 3, 'title': '記憶體', 'sub': '12吋矽晶圓主力應用（記憶體客戶產能擴建）', 'note': '定性揭露，無量化拆分'},
        {'id': 'pmic', 'stage': 3, 'title': '電源管理IC', 'sub': '8吋矽晶圓主力應用（AI伺服器HVDC、車用/工業電源管理IC）', 'note': '定性揭露，無量化拆分'},
        {'id': 'logic', 'stage': 3, 'title': '邏輯與其他', 'sub': '晶圓代工先進製程等其他應用', 'note': '定性揭露，無量化拆分'},
        {'id': 'ftg', 'stage': 0, 'title': '台塑（29.05% 股東）', 'sub': '具重大影響力之投資者；辦公室租賃＋短期資金融通（已全數清償）', 'note': '非供應鏈成員，財務／租賃關係人', 'branch': 'fin'},
    ],
    'edges': [
        {'from': 'sumco_ingot', 'to': 'ftc', 'label': '單晶晶棒（關係人）', 'branch': 'fin'},
        {'from': 'poly', 'to': 'ftc', 'label': '多晶矽', 'branch': 'main'},
        {'from': 'quartz', 'to': 'ftc', 'branch': 'main'},
        {'from': 'equip', 'to': 'ftc', 'branch': 'main'},
        {'from': 'ftc', 'to': 'custA', 'branch': 'main'},
        {'from': 'ftc', 'to': 'sumco_buy', 'label': '回銷（關係人）', 'branch': 'fin'},
        {'from': 'ftc', 'to': 'other_fab', 'branch': 'main'},
        {'from': 'ftc', 'to': 'pmic', 'label': '8吋主力應用', 'branch': 'app'},
        {'from': 'ftc', 'to': 'mem', 'label': '12吋主力應用', 'branch': 'app'},
        {'from': 'ftc', 'to': 'logic', 'label': '其他應用', 'branch': 'app'},
        {'from': 'ftg', 'to': 'ftc', 'label': '辦公室租賃／短期融通（財務關係）', 'branch': 'fin', 'side': 'right'},
    ],
    'options': {'showStageLabels': True, 'showLegend': True}
}

# 11. Supply chain sankey -- only disclosed flows (upstream purchase share, downstream sales share);
#     terminal-application tier omitted (no quantified split disclosed anywhere in filings)
S['supply_chain_sankey'] = {
    'nodes': [
        {'id': 'SUMCO（單晶晶棒）', 'label': 'SUMCO（單晶晶棒）17.24% 進貨占比', 'color': '#9333ea'},
        {'id': '多晶矽供應商', 'label': '多晶矽供應商（Tokuyama）10.64% 進貨占比', 'color': '#60a5fa'},
        {'id': '台勝科', 'label': '台勝科 3532', 'color': '#0f172a'},
        {'id': '客戶A', 'label': '客戶A 37.58% 銷貨占比', 'color': '#16a34a'},
        {'id': 'SUMCO（回銷）', 'label': 'SUMCO（回銷）10.04% 銷貨占比', 'color': '#a855f7'},
        {'id': '其他晶圓廠', 'label': '其他晶圓廠 52.38%（含M/B公司等）', 'color': '#86efac'},
    ],
    'links': [
        {'source': 'SUMCO（單晶晶棒）', 'target': '台勝科', 'value': 17.24},
        {'source': '多晶矽供應商', 'target': '台勝科', 'value': 10.64},
        {'source': '台勝科', 'target': '客戶A', 'value': 37.58},
        {'source': '台勝科', 'target': 'SUMCO（回銷）', 'value': 10.04},
        {'source': '台勝科', 'target': '其他晶圓廠', 'value': 52.38},
    ]
}

os.makedirs(OUT, exist_ok=True)
for k, v in S.items():
    json.dump(v, open(os.path.join(OUT, k + '.json'), 'w', encoding='utf-8'), ensure_ascii=False)
print(list(S))
