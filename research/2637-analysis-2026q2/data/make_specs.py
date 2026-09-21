import json, os, sys
D = os.path.dirname(os.path.abspath(__file__)) + '/'
OUT = sys.argv[1]
r = json.load(open(D + 'quarterly.json', encoding='utf-8')); qs = sorted(r); L = [q[2:] for q in qs]
M = lambda v: round(v / 1e6, 1)  # NT$ millions (raw values are absolute TWD)
S = {}

# ---------- Tab 1 特殊事件：新造船計畫 ----------
S['newbuild_gantt'] = {
    'title': '慧洋-KY 新造船購建計畫（截至2026-06-30尚未交船之7艘）',
    'caption': '來源：2026Q2合併財報附註九「重大或有負債及未認列之合約承諾」。合約總價US$234,753,200，已付US$55,811,500，餘款分別於安放龍骨／下水／交船時支付，預計2027Q3前付清；尚未辦妥融資。⚠️ 法說會第三方摘要另提及「2026年交付8艘」，與此處財報揭露之「截至6/30尚未交船7艘」口徑不同（前者含揭露時點前已交船者），本圖以財報附註九為準。',
    'scale': 'quarter', 'rowH': 32,
    'items': [
        {'name': '2026年交船批次（3艘，US$93.4M）', 'start': '2026-06-30', 'end': '2026-12-31', 'group': '購建合約（7艘，US$234.75M）'},
        {'name': '2027年交船批次（4艘，US$141.3M）', 'start': '2026-06-30', 'end': '2027-09-30', 'group': '購建合約（7艘，US$234.75M）'},
    ],
    'markers': [
        {'label': '財報揭露基準日 2026-06-30（已付US$55.8M）', 'date': '2026-06-30', 'color': '#111111'},
        {'label': '預計2027Q3前付清全部餘款', 'date': '2027-09-30', 'color': '#b45309'},
    ],
    'opts': {'grid': True, 'alt': True, 'dur': False, 'group': True},
}
S['newbuild_waterfall'] = {
    'title': '新造船購建合約承諾拆解（US$M，截至2026-06-30）',
    'caption': '合約總價US$234.75M（7艘）；已付US$55.81M（安放龍骨等分期款）；尚未支付餘額US$178.94M將於龍骨／下水／交船時分期給付，尚未辦妥融資。下方併列此餘額對應的兩個交船年度之「合約總價」拆分（非按已付/未付切分，僅供對照年度分布）。來源：2026Q2財報附註九。',
    'unit': 'US$M',
    'items': [
        {'name': '合約總價（7艘）', 'value': 234.75, 'type': 'base'},
        {'name': '已付款（截至2026-06-30）', 'value': -55.81, 'type': 'delta'},
        {'name': '尚未支付餘額', 'value': 178.94, 'type': 'total'},
        {'name': '其中：2026年交船部分（3艘）合約值', 'value': -93.43, 'type': 'delta'},
        {'name': '其中：2027年交船部分（4艘）合約值', 'value': -141.32, 'type': 'total'},
    ],
    'opts': {'values': True, 'conn': True, 'grid': True},
}

# ---------- Tab 2 船隊/業務平台 ----------
fleet = json.load(open(D + 'fleet_2026q2.json', encoding='utf-8'))
from collections import defaultdict
byType = defaultdict(lambda: [0, 0])
for x in fleet:
    byType[x['type']][0] += 1
    byType[x['type']][1] += int(x['dwt'])
order = sorted(byType.items(), key=lambda kv: -kv[1][1])
S['fleet_by_type'] = {
    'title': '慧洋船隊結構：依船型分（截至2026-06-30，130艘）',
    'caption': '艘數與總噸位（萬DWT）並列（量綱不同，僅作相對規模比較）。巴拿馬極限型單艘噸位最大但艘數次多；輕便型船艘數最多、以2026年新造節能船為主力交船船型。來源：2026Q2財報附註十二「船隊資訊」。',
    'unit': '', 'categories': [k for k, v in order], 'mode': 'grouped', 'orient': 'h',
    'series': [
        {'name': '艘數', 'values': [v[0] for k, v in order]},
        {'name': '總噸位（萬DWT）', 'values': [round(v[1] / 10000, 1) for k, v in order]},
    ],
    'opts': {'grid': True, 'values': True, 'legend': True},
}
buckets = [('2009年以前', lambda y: y <= 2009), ('2010–2014', lambda y: 2010 <= y <= 2014),
           ('2015–2019', lambda y: 2015 <= y <= 2019), ('2020–2024', lambda y: 2020 <= y <= 2024),
           ('2025–2026（近2年新船）', lambda y: y >= 2025)]
counts = []
for name, f in buckets:
    counts.append(sum(1 for x in fleet if f(int(x['build_year']))))
S['fleet_age'] = {
    'title': '慧洋船隊船齡分布（依建造年份分組，130艘）',
    'caption': '2025–2026新交船6艘（皆輕便型，符合NOx Tier III排放標準），2020–2024批次33艘為近年節能船更新主力；2009年以前老舊船仍有1艘（客輪Pescadores，非散裝主力船隊）。來源：2026Q2財報附註十二「船隊資訊」逐艘建造年份彙總。',
    'unit': '艘', 'categories': [b[0] for b in buckets], 'mode': 'grouped', 'orient': 'v',
    'series': [{'name': '艘數', 'values': counts}],
    'opts': {'grid': True, 'values': True, 'legend': False},
}

# ---------- Tab 3 供應鏈 ----------
S['chain_supply'] = {
    'title': '慧洋-KY 供應鏈結構：上游 → 船隊 → 合約型態 → 客戶',
    'caption': '長約占船隊約88–90%（指數連結型約88艘為主流，據富果直送2026-06-09法說會摘要）；短約約11%；船舶管理服務占比<2%。具名客戶Cargill、Bunge（2025年收購Viterra）、SwissMarine為長期合作之歐洲租家；財報另以匿名代號揭露A/B/C/D四大客戶（2026H1合計US$156.7M，占營收約51%）。滑油／燃油供應商分散，無單一供應商達10%進貨門檻。來源：2025年報肆、二(三)(四)；2026Q2財報附註六(23)。',
    'stages': ['上游供給', '慧洋船隊（依船型）', '合約／收入型態', '客戶'],
    'branches': {'main': {'label': '實體/收入流', 'color': '#2563eb'}, 'fin': {'label': '融資關係', 'color': '#9333ea', 'dashed': True}},
    'nodes': [
        {'id': 'yard', 'stage': 0, 'title': '日本造船廠', 'sub': '2026年新交船6艘皆為輕便型節能船', 'note': '公司列為競爭利基：船廠聲譽佳、未來處分利潤較高'},
        {'id': 'bank', 'stage': 0, 'title': '融資銀行', 'sub': '長期借款NT$199.3億+一年內到期NT$37.5億', 'note': '新造船US$178.9M餘款尚未辦妥融資', 'branch': 'fin'},
        {'id': 'lube', 'stage': 0, 'title': '滑油供應商（4–5家）', 'sub': '分散採購，各家占比未達10%'},
        {'id': 'fuel', 'stage': 0, 'title': '燃油供應商', 'sub': '依航程/港口分散購置'},
        {'id': 'panamax', 'stage': 1, 'title': '巴拿馬極限型', 'sub': '35艘／288.5萬DWT'},
        {'id': 'supra', 'stage': 1, 'title': '超級極限型船', 'sub': '29艘／177.0萬DWT'},
        {'id': 'handy', 'stage': 1, 'title': '輕便型船', 'sub': '63艘／205.3萬DWT', 'note': '2026年新造船6艘皆屬此型', 'emphasis': True},
        {'id': 'other_ship', 'stage': 1, 'title': '其他型（LPG/滾裝/客輪）', 'sub': '3艘'},
        {'id': 'lta_idx', 'stage': 2, 'title': '長約－指數連結', 'sub': '約88艘（主流），與BDI/BSI連動', 'emphasis': True},
        {'id': 'lta_short', 'stage': 2, 'title': '長約其他＋短約', 'sub': '合計約11%；短約租期<半年'},
        {'id': 'mgmt', 'stage': 2, 'title': '船舶管理服務', 'sub': '收入占比<2%'},
        {'id': 'cargill', 'stage': 3, 'title': 'Cargill', 'sub': '具名長約客戶（歐洲租家）'},
        {'id': 'bunge', 'stage': 3, 'title': 'Bunge', 'sub': '2025年收購Viterra，具名長約客戶'},
        {'id': 'swiss', 'stage': 3, 'title': 'SwissMarine', 'sub': '具名長約客戶'},
        {'id': 'spot', 'stage': 3, 'title': '現貨市場客戶', 'sub': '東南亞／印度／中國／中東，逐步拓展非洲、南美'},
    ],
    'edges': [
        {'from': 'yard', 'to': 'handy', 'label': '新造交船', 'branch': 'main'},
        {'from': 'bank', 'to': 'panamax', 'branch': 'fin'}, {'from': 'bank', 'to': 'supra', 'branch': 'fin'}, {'from': 'bank', 'to': 'handy', 'branch': 'fin'},
        {'from': 'lube', 'to': 'panamax', 'branch': 'main'}, {'from': 'lube', 'to': 'supra', 'branch': 'main'}, {'from': 'lube', 'to': 'handy', 'branch': 'main'},
        {'from': 'fuel', 'to': 'panamax', 'branch': 'main'}, {'from': 'fuel', 'to': 'supra', 'branch': 'main'}, {'from': 'fuel', 'to': 'handy', 'branch': 'main'}, {'from': 'fuel', 'to': 'other_ship', 'branch': 'main'},
        {'from': 'panamax', 'to': 'lta_idx', 'branch': 'main'}, {'from': 'supra', 'to': 'lta_idx', 'branch': 'main'}, {'from': 'handy', 'to': 'lta_idx', 'branch': 'main'},
        {'from': 'panamax', 'to': 'lta_short', 'branch': 'main'}, {'from': 'supra', 'to': 'lta_short', 'branch': 'main'}, {'from': 'handy', 'to': 'lta_short', 'branch': 'main'},
        {'from': 'other_ship', 'to': 'mgmt', 'branch': 'main'},
        {'from': 'lta_idx', 'to': 'cargill', 'branch': 'main'}, {'from': 'lta_idx', 'to': 'bunge', 'branch': 'main'}, {'from': 'lta_idx', 'to': 'swiss', 'branch': 'main'},
        {'from': 'lta_short', 'to': 'spot', 'branch': 'main'},
    ],
    'options': {'showStageLabels': True, 'showLegend': True},
}
# Sankey: two genuinely quantifiable, same-unit flows sharing the chain-diagram's node names.
S['sankey_supply'] = {
    'nodes': [
        {'id': '輕便型船', 'label': '輕便型船 205.3萬DWT', 'color': '#2563eb'},
        {'id': '巴拿馬極限型', 'label': '巴拿馬極限型 288.5萬DWT', 'color': '#1d4ed8'},
        {'id': '超級極限型船', 'label': '超級極限型船 177.0萬DWT', 'color': '#3b82f6'},
        {'id': '其他型', 'label': '其他型 1.0萬DWT', 'color': '#93c5fd'},
        {'id': '慧洋合計船隊', 'label': '慧洋合計船隊 130艘／671.9萬DWT', 'color': '#0f172a'},
        {'id': '長約收入', 'label': '長約收入 NT$86.19億', 'color': '#16a34a'},
        {'id': '短約收入', 'label': '短約收入 NT$9.45億', 'color': '#22c55e'},
        {'id': '船舶管理收入', 'label': '船舶管理收入 NT$0.15億', 'color': '#86efac'},
        {'id': '其他收入', 'label': '其他收入 NT$1.01億', 'color': '#bbf7d0'},
        {'id': '2026H1合計營收', 'label': '2026H1合計營收 NT$96.80億', 'color': '#111111'},
    ],
    'links': [
        {'source': '輕便型船', 'target': '慧洋合計船隊', 'value': 205.3},
        {'source': '巴拿馬極限型', 'target': '慧洋合計船隊', 'value': 288.5},
        {'source': '超級極限型船', 'target': '慧洋合計船隊', 'value': 177.0},
        {'source': '其他型', 'target': '慧洋合計船隊', 'value': 1.0},
        {'source': '長約收入', 'target': '2026H1合計營收', 'value': 86.19},
        {'source': '短約收入', 'target': '2026H1合計營收', 'value': 9.45},
        {'source': '船舶管理收入', 'target': '2026H1合計營收', 'value': 0.15},
        {'source': '其他收入', 'target': '2026H1合計營收', 'value': 1.01},
    ],
}

# ---------- Tab 4 風險矩陣 ----------
S['risk_heat'] = {
    'title': '風險矩陣：風險項 × 機率/衝擊/時間窗口（強度1–5）',
    'caption': '數值為本報告主觀評分，5＝極可能且衝擊大／時間窗口內顯著惡化。散裝運價循環反轉與長約到期重議風險隨時間窗口拉長而上升；新造船US$178.9M餘款尚未辦妥融資，2027年到期部位風險高於2026年。來源：本報告風險矩陣tab論證，綜合財報附註九、市場評述整理。',
    'unit': '',
    'rows': ['BDI運價下行', '浮動利率再融資', '新造船付款未融資', '二手船價下跌減損', '美元/新台幣匯兌', '長約到期重議', 'CII碳稅法規'],
    'cols': ['機率', '衝擊', '近期(0–12個月)', '中期(1–3年)'],
    'values': [
        [3, 5, 2, 4],
        [2, 3, 2, 3],
        [3, 3, 2, 4],
        [2, 4, 2, 3],
        [3, 3, 3, 3],
        [3, 4, 2, 4],
        [4, 3, 2, 3],
    ],
    'scale': 'greenred', 'mid': 3,
    'opts': {'values': True, 'legend': True, 'border': True},
}

# ---------- Tab 6 損益表 ----------
S['income_trend'] = {
    'title': '慧洋-KY 近12季營收與利潤率（2023Q3–2026Q2）',
    'caption': '左軸：單季營收/營業利益（NT$百萬元）；右軸：毛利率／營業利益率（%）。2025Q1散裝運價低迷使毛利率降至6.8%，其後隨BDI回升逐季改善，2026Q2毛利率39.2%、營益率38.0%創近12季新高。來源：MOPS合併財報t164sb01，Q4為全年減前三季。',
    'leftLabel': 'NT$百萬元', 'rightLabel': '%', 'categories': L,
    'series': [
        {'name': '營業收入', 'axis': 'left', 'values': [M(r[q]['Rev']) for q in qs]},
        {'name': '營業利益', 'axis': 'left', 'values': [M(r[q]['OpInc']) for q in qs]},
        {'name': '毛利率', 'axis': 'right', 'values': [round(r[q]['GP'] / r[q]['Rev'] * 100, 1) for q in qs]},
        {'name': '營業利益率', 'axis': 'right', 'values': [round(r[q]['OpInc'] / r[q]['Rev'] * 100, 1) for q in qs]},
    ],
    'opts': {'grid': True, 'markers': True, 'values': False, 'leftZero': True, 'rightZero': True},
}
mrev = json.load(open(D + 'monthly_revenue.json', encoding='utf-8'))
S['monthly_revenue'] = {
    'title': '慧洋-KY 近27個月月營收與YoY（2024/06–2026/08）',
    'caption': '左軸：單月營收（NT$百萬元）；右軸：YoY（%）。2026/07、08為2026Q2季報期間之後最新月份，尚未被任何一期財報涵蓋——單月YoY分別達56.2%與35.0%，顯示Q3運價動能延續。來源：MOPS上市月營收彙總表t21sc03_sii，累計欄反算交叉驗證。',
    'leftLabel': 'NT$百萬元', 'rightLabel': 'YoY %', 'categories': [x['ym'][2:] for x in mrev],
    'series': [
        {'name': '月營收', 'axis': 'left', 'values': [round(x['rev_k'] / 1000, 1) for x in mrev]},
        {'name': 'YoY', 'axis': 'right', 'values': [x['yoy_pct'] for x in mrev]},
    ],
    'opts': {'grid': True, 'markers': True, 'leftZero': True, 'rightZero': False},
}

# ---------- Tab 7 業務部門 ----------
S['revenue_breakdown'] = {
    'title': '收入結構：2025H1 vs 2026H1（依收入類別）',
    'caption': '單位：NT$百萬元。長約收入（含指數連結型）占比約90%，為BDI/BSI上漲之主要傳導管道；短約收入年增顯著但占比仍小。船舶管理收入占比不足2%。來源：2026Q2合併財報附註六(廿三)「與客戶合約之收入」，收入細分表。',
    'unit': 'NT$百萬元', 'categories': ['長約收入', '短約收入', '船舶管理收入', '其他'], 'mode': 'grouped', 'orient': 'h',
    'series': [
        {'name': '2025H1', 'values': [6241.1, 969.7, 9.8, 95.1]},
        {'name': '2026H1', 'values': [8619.4, 944.5, 14.6, 101.3]},
    ],
    'opts': {'grid': True, 'values': True, 'legend': True},
}
S['volume_price'] = {
    'title': '船隊規模 vs 隱含單艘營收：2025Q4 vs 2026Q2',
    'caption': '⚠️ 資料限制：財報僅揭露2025年底（126艘）與2026Q2末（130艘）兩個時點的船隊總數，無逐季船隊艘數序列，故僅能取此兩個可比時點做「量」（艘數）與「隱含價」（單季營收÷艘數，NT$百萬元/艘）對比，非連續QoQ序列。單艘隱含營收由40.4微升至41.1，顯示同期營收成長主要來自運價（價），船隊擴張（量）貢獻有限。來源：quarterly.json 2025Q4/2026Q2營收、2025年報與2026Q2財報船隊揭露。',
    'unit': '', 'categories': ['2025Q4（126艘）', '2026Q2（130艘）'], 'mode': 'grouped', 'orient': 'v',
    'series': [
        {'name': '船隊艘數', 'values': [126, 130]},
        {'name': '隱含單艘營收（NT$百萬元/艘）', 'values': [round(r['2025Q4']['Rev'] / 1e6 / 126, 1), round(r['2026Q2']['Rev'] / 1e6 / 130, 1)]},
    ],
    'opts': {'grid': True, 'values': True, 'legend': True},
}

# ---------- Tab 8 資產負債表 ----------
S['dso'] = {
    'title': '應收帳款天數（DSO）12季走勢',
    'caption': 'DSO＝期末應收帳款÷單季營收×91。2637應收帳款占資產比重極低（船舶租賃業以預收租金為主，非應收帳款驅動），DSO長期落在2.5–4天量級，屬正常水準，非核心觀察指標。來源：各季合併資產負債表與綜合損益表。',
    'leftLabel': '天', 'rightLabel': '', 'categories': L,
    'series': [{'name': 'DSO', 'axis': 'left', 'values': [round(r[x]['AR'] / r[x]['Rev'] * 91, 1) for x in qs]}],
    'opts': {'grid': True, 'markers': True, 'leftZero': True},
}
S['prepaid_advances'] = {
    'title': '預收款項（CurrentAdvances）12季餘額走勢',
    'caption': '2637（散裝航運）無ContractLiabilities類科目，以「預收款項」科目取代，金額小、非核心觀察指標（資料缺口：見README，此為誠實揭露之替代科目，非合約負債）。2026Q2達NT$567.6百萬，為12季新高，方向上與運價回升期客戶提前鎖艙的傾向一致，但金額仍遠小於資產負債表規模。來源：各季合併資產負債表。',
    'unit': 'NT$百萬元', 'categories': L,
    'series': [{'name': '預收款項', 'axis': 'left', 'values': [M(r[x]['CurrentAdvances']) for x in qs]}],
    'leftLabel': 'NT$百萬元', 'rightLabel': '',
    'opts': {'grid': True, 'markers': True, 'leftZero': True},
}

# ---------- Tab 9 現金流量 ----------
S['fcf_waterfall'] = {
    'title': '2026H1自由現金流橋接（NT$百萬元）',
    'caption': '營運現金流NT$5,357.1M為近年高點；扣造船/購船資本支出NT$441.6M、加處分舊船價款NT$543.9M後FCF達NT$5,459.4M。股利分配NT$1,669.2M（對2025年度盈餘，每股NT$3.5元）後，淨借款變動（含租賃清償）為正NT$448.2M。⚠️ 此簡化橋接聚焦營運現金生成與資本配置主軸，未納入CFI項下其他金融資產（定存/受限資金等）異動，故末項不等於實際現金總變動（H1實際現金減少NT$516.4M，主因短期理財性資產調整與匯率換算，非核心營運現金流所致）。來源：2026Q2合併現金流量表YTD數字。',
    'unit': 'NT$百萬元',
    'items': [
        {'name': '營運現金流(CFO)', 'value': 5357.1, 'type': 'base'},
        {'name': '造船/購船資本支出', 'value': -441.6, 'type': 'delta'},
        {'name': '處分舊船價款', 'value': 543.9, 'type': 'delta'},
        {'name': '自由現金流(FCF)', 'value': 5459.4, 'type': 'total'},
        {'name': '股利分配', 'value': -1669.2, 'type': 'delta'},
        {'name': '淨借款變動(含租賃清償)', 'value': 448.2, 'type': 'total'},
    ],
    'opts': {'values': True, 'conn': True, 'grid': True},
}
S['cf_components'] = {
    'title': '營運／投資／籌資活動現金流 12季組成',
    'caption': '單位：NT$百萬元。2026Q2投資活動現金流轉為顯著淨流出（-3,129.1），主因資本支出加速與處分船舶價款減少，籌資活動則因借款/租賃淨減少持續流出。來源：各季合併現金流量表（單季，YTD相減）。',
    'unit': 'NT$百萬元', 'categories': L, 'mode': 'overlap',
    'series': [
        {'name': '營業活動', 'values': [M(r[x]['CFO']) for x in qs]},
        {'name': '投資活動', 'values': [M(r[x]['CFI']) for x in qs]},
        {'name': '籌資活動', 'values': [M(r[x]['CFF']) for x in qs]},
    ],
    'opts': {'grid': True, 'bounds': True, 'zero': True},
}

# ---------- Tab 10 估值觀察 ----------
peers = json.load(open(D + 'peers_yahoo.json', encoding='utf-8'))
name_map = {'2605.TW': '正德海運', '2606.TW': '裕民航運', '2612.TW': '中航', '2617.TW': '台航', '2637.TW': '慧洋-KY',
            'DSX': 'Diana Shipping', 'SB': 'Safe Bulkers', 'SBLK': 'Star Bulk'}
pe_points = []
pb_points = []
for k, v in peers.items():
    nm = name_map.get(k, k)
    mc_bn = round(v['marketCap'] / 1e9, 2)
    if v.get('trailingPE') and v.get('revenueGrowth') is not None:
        pe_points.append({'name': nm, 'x': round(v['revenueGrowth'] * 100, 1), 'y': round(v['trailingPE'], 1), 'size': mc_bn})
    if v.get('priceToBook') and v.get('returnOnEquity') is not None:
        pb_points.append({'name': nm, 'x': round(v['returnOnEquity'] * 100, 1), 'y': round(v['priceToBook'], 2), 'size': mc_bn})
S['pe_growth_scatter'] = {
    'title': '散裝航運同業：本益比(P/E) vs 營收成長（2026-09-20）',
    'caption': 'X：最近一期營收YoY（%）；Y：Trailing P/E（倍）；氣泡＝市值（十億，台股TWD/美股USD混列，量綱不同僅供組內相對比較）。此圖為框架要求之必備圖表。來源：Yahoo Finance quoteSummary（fc.yahoo.com cookie+crumb）。',
    'xName': '營收YoY（%）', 'yName': 'Trailing P/E（倍）', 'refMode': 'median',
    'points': pe_points, 'opts': {'labels': True, 'grid': True, 'trend': False, 'xZero': False, 'yZero': True},
}
S['pb_roe_scatter'] = {
    'title': '散裝航運同業：股價淨值比(P/B) vs ROE（2026-09-20）',
    'caption': 'X：ROE（%）；Y：P/B（倍）；氣泡＝市值（十億）。慧洋P/B 1.54倍高於多數台股同業（正德0.76、中航0.90、台航0.77），但ROE 14.3%亦領先，估值溢價有基本面支撐。來源：Yahoo Finance quoteSummary。',
    'xName': 'ROE（%）', 'yName': 'P/B（倍）', 'refMode': 'median',
    'points': pb_points, 'opts': {'labels': True, 'grid': True, 'trend': False, 'xZero': False, 'yZero': True},
}
S['int_cov'] = {
    'title': '利息保障倍數12季走勢（營業利益÷財務成本）',
    'caption': '財務成本自2023Q3高點NT$553.2M逐季降至2026Q2的NT$258.4M（借款餘額下降＋降息環境），加上營業利益隨運價回升，利息保障倍數由2025Q1谷底0.5倍大幅改善至2026Q2的7.9倍，財務體質同步佐證管理層「負債比自2022年49.9%降至2026Q1的40.2%」之說法方向一致。來源：各季合併綜合損益表營業利益與財務成本。',
    'leftLabel': '倍', 'rightLabel': '', 'categories': L,
    'series': [{'name': '營業利益÷財務成本', 'axis': 'left', 'values': [round(r[x]['OpInc'] / r[x]['FinCost'], 1) for x in qs]}],
    'opts': {'grid': True, 'markers': True, 'values': True, 'leftZero': True},
}

os.makedirs(OUT, exist_ok=True)
for k, v in S.items():
    json.dump(v, open(os.path.join(OUT, k + '.json'), 'w', encoding='utf-8'), ensure_ascii=False)
print(list(S))
print('count:', len(S))
