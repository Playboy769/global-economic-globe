import json, os, sys
D = os.path.dirname(os.path.abspath(__file__)) + '/'
OUT = sys.argv[1]
r = json.load(open(D + 'quarterly.json')); qs = sorted(r); L = [q[2:] for q in qs]
dm = json.load(open(D + 'derived_metrics.json'))
B = lambda v: round(v / 1e8, 2)
S = {}

S['income_trend'] = {'title': '合晶科技 近 12 季營收與利潤率（2023Q3–2026Q2）', 'caption': '左軸：單季營收（億元）；右軸：毛利率／營業利益率（%）。來源：MOPS 合併財報 t164sb01，Q1/Q2 為季度直接揭露值。', 'leftLabel': '營收（億元）', 'rightLabel': '%', 'categories': L,
    'series': [{'name': '營收', 'axis': 'left', 'values': [B(r[q]['Rev']) for q in qs]}, {'name': '毛利率', 'axis': 'right', 'values': [round(r[q]['GP'] / r[q]['Rev'] * 100, 1) for q in qs]}, {'name': '營業利益率', 'axis': 'right', 'values': [round(r[q]['OpInc'] / r[q]['Rev'] * 100, 1) for q in qs]}], 'opts': {'grid': True, 'markers': True, 'values': False, 'leftZero': True, 'rightZero': True}}

m = json.load(open(D + 'monthly_revenue.json')); mk = sorted(m)
S['monthly_revenue'] = {'title': '合晶科技 近 32 個月月營收與 YoY（2024/01–2026/08）', 'caption': '左軸：單月營收（億元）；右軸：YoY（%）。2026/07、08 為季報（截至6/30）期間之後的最新月份，尚未被任何一期財報涵蓋，8月營收NT$1,004,466千為近32個月新高、首度單月破十億元。全部月份均以累計欄反算交叉驗證（cross_check_ok=true）。來源：MOPS 上櫃月營收彙總表 t21sc03_otc。', 'leftLabel': '月營收（億元）', 'rightLabel': 'YoY %', 'categories': [k[2:] for k in mk],
    'series': [{'name': '月營收', 'axis': 'left', 'values': [round(m[k]['revenue_thousand'] / 1e5, 2) for k in mk]}, {'name': 'YoY', 'axis': 'right', 'values': [m[k]['yoy_pct'] for k in mk]}], 'opts': {'grid': True, 'markers': True, 'leftZero': True, 'rightZero': False}}

# bu_mix: 6182 揭露單一營運部門（附註十四），季報無產品別/地區別收入細分。
# 改用年報（ar2025.txt）揭露之銷售地區別（FY113=2024 vs FY114=2025，僅有年度資料，非季度）。
S['bu_mix'] = {'title': '銷售地區別營收：FY2024 vs FY2025（年度，非季度）', 'caption': '單位：億元。合晶財報附註十四揭露「單一營運部門」，2026Q2合併財報無產品別／地區別季度收入細分；地區別資料僅見於年報（股東會年報），故本圖以FY2024（民國113年）vs FY2025（民國114年）年度資料呈現，非季度對比，亦無2026年資料可用。來源：股東會年報（ar2025.txt）「銷售地區別」表。', 'unit': '億元', 'categories': ['台灣', '中國大陸(含香港)', '美國', '其他國家'], 'mode': 'grouped', 'orient': 'v',
    'series': [{'name': 'FY2024', 'values': [27.43, 26.44, 9.62, 23.72]}, {'name': 'FY2025', 'values': [27.67, 28.27, 9.05, 33.19]}], 'opts': {'grid': True, 'values': True, 'legend': True}}

S['fcf_waterfall'] = {'title': '2026Q2 自由現金流橋接（億元）', 'caption': 'derive_metrics.json：2026Q2 FCF=-2.32億（CFO 5.38億－Capex 7.70億），D&A 5.46億。中間項目為估算拆解，用以呈現營運現金流→資本支出→FCF的橋接邏輯，非財報逐項揭露；折舊攤銷加回、營運資金變動合計以「其他營運調整」軋平至CFO。來源：quarterly.json（2026Q2 CFO/Capex/Dep，季度直接揭露值）、derive_metrics.json。', 'unit': '億元',
    'items': [{'name': '稅前淨利', 'value': B(r['2026Q2']['PBT']), 'type': 'base'}, {'name': '折舊攤銷加回', 'value': round(r['2026Q2']['Dep'] / 1e8, 2), 'type': 'delta'}, {'name': '其他營運調整', 'value': round((r['2026Q2']['CFO'] - r['2026Q2']['PBT'] - r['2026Q2']['Dep']) / 1e8, 2), 'type': 'delta'}, {'name': '營運現金流', 'value': round(r['2026Q2']['CFO'] / 1e8, 2), 'type': 'total'}, {'name': '資本支出', 'value': round(r['2026Q2']['Capex'] / 1e8, 2), 'type': 'delta'}, {'name': '自由現金流', 'value': round((r['2026Q2']['CFO'] + r['2026Q2']['Capex']) / 1e8, 2), 'type': 'total'}], 'opts': {'values': True, 'conn': True, 'grid': True}}

S['valuation_scatter'] = {'title': '矽晶圓同業：Trailing 本益比 vs 營收成長（2026-09-16）', 'caption': 'X：最近一季營收 YoY（%，Yahoo revGrowth）；Y：Trailing P/E（倍，同業多無 Forward P/E 資料，故全部改用 Trailing 口徑以求一致）；氣泡＝市值（十億美元，TWD/JPY 概略換算 31.6／147.5，僅供氣泡相對大小參考）。Sumco(3436) 無 Trailing／Forward P/E 資料（P/B 1.83），不入圖。合晶(6182) Trailing P/E 高達 466.7x 係因 2026Q2 NI_parent 僅NT$3,119萬、獲利基期極低所致，本益比比較意義有限，宜改參考 P/B 3.91 或 EV/Sales；此圖為框架要求之必備圖表，仍照實揭露。來源：Yahoo Finance quoteSummary（peers_yahoo.json）。', 'xName': '營收 YoY（%）', 'yName': 'Trailing P/E（倍）', 'refMode': 'median',
    'points': [{'name': '合晶 6182', 'x': 7.4, 'y': 466.7, 'size': 2.2}, {'name': '環球晶 6488', 'x': -5.0, 'y': 45.6, 'size': 14.2}, {'name': '台勝科 3532', 'x': 20.6, 'y': 338.5, 'size': 5.4}, {'name': '嘉晶 3016', 'x': 37.1, 'y': 96.2, 'size': 1.06}, {'name': '信越化學 4063', 'x': 5.4, 'y': 23.0, 'size': 71.6}], 'opts': {'labels': True, 'grid': True, 'trend': False, 'xZero': False, 'yZero': True}}

S['dso'] = {'title': '應收帳款天數（DSO）與存貨天數 12 季', 'caption': 'DSO＝期末應收帳款÷單季營收×91；存貨天數＝期末存貨÷單季營業成本×91。2026Q2 DSO 67.0天，較2026Q1 68.7天略降；存貨天數135.1天，2025Q4曾達157.6天。來源：derive_metrics.json（各季合併資產負債表與綜合損益表）。', 'leftLabel': '天', 'rightLabel': '', 'categories': L,
    'series': [{'name': 'DSO', 'axis': 'left', 'values': [round(dm[q]['DSO'], 1) for q in qs]}, {'name': '存貨天數', 'axis': 'left', 'values': [round(dm[q]['InvDays'], 1) for q in qs]}], 'opts': {'grid': True, 'markers': True, 'leftZero': True}}

S['int_cov'] = {'title': '利息保障倍數 12 季（營業利益 ÷ 利息費用）', 'caption': 'derive_metrics.json IntCoverage＝OpInc÷IntExp。2024Q1營業虧損季度倍數為負(-0.28)；隨擴產舉債（長期借款自2697百萬增至7717百萬）利息費用上升，2026Q1降至1.01倍（接近臨界），2026Q2回升至3.15倍，仍低於2025Q4的7.50倍。來源：quarterly.json IS.OpInc／IS.IntExp，derive_metrics.json。', 'leftLabel': '倍', 'rightLabel': '', 'categories': L,
    'series': [{'name': '營業利益÷利息費用', 'axis': 'left', 'values': [round(dm[q]['IntCoverage'], 2) for q in qs]}], 'opts': {'grid': True, 'markers': True, 'values': True, 'leftZero': True}}

S['contract_liab'] = {'title': '合約負債（客戶預付款）餘額 12 季走勢', 'caption': '單位：億元。合晶合約負債全數為流動項目，2026Q2餘額2.58億，較2025Q4的2.71億續降。附註九.5另揭露與下游策略聯盟客戶A/B相關合約負債合計2.22億（111.04–113.12合約期間，可能展延），與此處流動合約負債是否重疊財報未明確交叉說明，信心度中等。來源：quarterly.json BS.CL_contract；q2_pdf.txt（合約負債附註、附註九.5）。', 'unit': '億元', 'categories': L, 'mode': 'stacked',
    'series': [{'name': '合約負債－流動', 'values': [round((r[q]['CL_contract'] or 0) / 1e8, 2) for q in qs]}], 'opts': {'grid': True, 'bounds': True, 'markers': False, 'bandLabels': False, 'totals': True, 'zero': True}}

S['capex_dep'] = {'title': '資本支出 vs 折舊 vs 營運現金流（單季，億元）', 'caption': '2026Q1資本支出達22.76億（近12季新高，對應二林/龍潭擴產），2026Q2回落至7.70億；折舊自2023Q3的12.3億逐季墊高至2026Q2的5.46億（單季）並將隨新產線陸續轉固持續上升。來源：quarterly.json 各季合併現金流量表（2026Q1/Q2為季度直接揭露值）。', 'unit': '億元', 'categories': L, 'mode': 'grouped', 'orient': 'v',
    'series': [{'name': '資本支出', 'values': [round(-(r[q]['Capex'] if r[q]['Capex'] is not None else 0) / 1e8, 2) for q in qs]}, {'name': '折舊', 'values': [round((r[q]['Dep'] if r[q]['Dep'] is not None else 0) / 1e8, 2) for q in qs]}, {'name': '營運現金流', 'values': [round((r[q]['CFO'] if r[q]['CFO'] is not None else 0) / 1e8, 2) for q in qs]}], 'opts': {'grid': True, 'values': False, 'legend': True}}

S['risk_heat'] = {'title': '風險矩陣：風險項 × 時間窗口（強度 1–5）', 'caption': '數值為本報告主觀評分：5＝極可能且衝擊大。來源：qualitative_notes.md 第7節風險清單論證。', 'unit': '',
    'rows': ['獲利基期低、P/E失真', '大股東華榮連續五月減碼', '陸子公司鉅額背書保證', '增資+CB+舉債稀釋槓桿', '利息保障倍數波動', '存貨天數偏高波動', '陸廠加速在地化競爭'], 'cols': ['0–6 個月', '6–12 個月', '1–2 年'],
    'values': [[4, 4, 3], [4, 3, 2], [2, 3, 4], [3, 4, 4], [4, 3, 2], [3, 3, 2], [2, 3, 4]], 'scale': 'diverging', 'mid': 3, 'opts': {'values': True, 'legend': True, 'border': True}}

S['subsidiary_pnl'] = {'title': '大陸轉投資事業 2026 上半年損益（附註十三）', 'caption': '單位：億元。被投資公司本期損益（非本公司認列投資損益）。上海晶盟（磊晶片）與上海合晶（矽晶片）獲利，揚州合晶（矽晶棒）與鄭州合晶虧損；鄭州合晶為本公司對外背書保證NT$84.4億之標的公司之一。來源：q2_pdf.txt附註十三「大陸被投資公司」表。', 'unit': '億元', 'categories': ['上海晶盟（磊晶片）', '上海合晶（矽晶片）', '揚州合晶（矽晶棒）', '鄭州合晶（矽晶片）'], 'mode': 'grouped', 'orient': 'h',
    'series': [{'name': 'H1 2026 被投資公司損益', 'values': [2.28, 0.86, -0.12, -0.44]}], 'opts': {'grid': True, 'values': True, 'legend': False}}

S['funding_stack'] = {'title': '2025–2026 資本結構籌資組成（億元）', 'caption': '單位：億元。現金增資：115.06.05基準日，追補發行20,000仟股×27.80元＝5.56億。CB轉換：國內第七次+第八次無擔保CB於115年上半年申請轉換金額合計11.51億（215,002仟元+935,500仟元，換發普通股共31,743仟股）。長期借款：2023Q3至2026Q2累計增加50.2億（26.97億→77.17億，近三倍）。三者合計反映擴產所需資金來源，並對應風險矩陣「股本稀釋＋財務槓桿」一項。來源：q2_pdf.txt附註（股本／可轉換公司債／長期借款），quarterly.json（LTB 2023Q3 vs 2026Q2）。', 'unit': '億元', 'categories': ['現金增資(115.06)', 'CB轉換(115H1)', '長期借款增加(2023Q3→2026Q2)'], 'mode': 'grouped', 'orient': 'v',
    'series': [{'name': '籌資金額', 'values': [5.56, 11.51, 50.20]}], 'opts': {'grid': True, 'values': True, 'legend': False}}

S['chain_supply'] = {'title': '合晶科技供應鏈結構：上游原料 → 一貫製程 → 銷售通路 → 終端客戶', 'caption': '年報揭露FY113/FY114/115Q1「主要供應商」表皆列「其他100.00%」，無單一供應商達10%門檻（來源：2025 年報主要原料供應狀況）；惟合併財報附註九.3/9.4另揭露甲/乙兩家供應商之長期原料採購合約（q2_pdf.txt附註九，分期預付貨款並依約履行進貨義務）。Helitek Company Ltd為合晶100%持股之美國子公司，主營「半導體矽晶圓材料銷售」，據附表二關係人交易揭露佔合晶科技（母公司單體）總銷貨40.41%（q2_pdf.txt附表二）——這是集團內部銷售通路而非最終客戶，合併報表中此內部交易已沖銷，Helitek之後的終端客戶身份未揭露。客戶A／B為年報揭露之合併層級10%以上銷貨客戶（與Helitek為不同統計基礎，無法確認是否重疊）。虛線＝金融關係（背書保證），非供應流。⚠️ 甲／乙長約供應商無金額揭露，下方 Sankey 將其併入「其他供應商」（年報進貨淨額合計 32.27 億）；地區別營收見業務部門 tab。',
    'stages': ['上游原料／供應商', '合晶科技一貫製程', '銷售通路', '終端客戶／地區'],
    'branches': {'main': {'label': '實體供應流', 'color': '#2563eb'}, 'fin': {'label': '金融關係（非供應）', 'color': '#9333ea', 'dashed': True}},
    'nodes': [{'id': 'jia', 'stage': 0, 'title': '甲供應商（原料）', 'sub': '長期採購合約，分期預付貨款', 'note': '110.08簽訂，展延至117.09'}, {'id': 'yi', 'stage': 0, 'title': '乙供應商（原料）', 'sub': '長期採購合約', 'note': '111.08簽訂，展延至118.12'}, {'id': 'oth_sup', 'stage': 0, 'title': '其他供應商', 'sub': '矽多晶料／研磨粉拋光漿／坩堝；FY2025 進貨 32.3 億', 'note': '無單一供應商達10%（分散採購）'},
        {'id': 'gwc', 'stage': 1, 'title': '合晶科技 6182', 'sub': '長晶→切片→倒角→研磨→拋光→清洗→包裝', 'note': '龍潭廠（既有 12 吋，40K/月）／中科二林廠（新建 12 吋）', 'emphasis': True, 'span': True}, {'id': 'cn_subs', 'stage': 1, 'title': '陸廠：上海合晶／上海晶盟／揚州合晶／鄭州合晶', 'sub': '矽晶片／磊晶片／矽晶棒研發生產', 'note': '持股比例42.63%（透過境外控股公司）'},
        {'id': 'helitek', 'stage': 2, 'title': 'Helitek Company Ltd', 'sub': '100%持股美國子公司，半導體矽晶圓材料銷售', 'note': '母公司單體銷貨佔比40.41%，合併報表內部沖銷', 'emphasis': True},
        {'id': 'A', 'stage': 3, 'title': '客戶A（合併層級10%+）', 'sub': 'FY114銷貨占比16.31%（NT$1,600,791千）', 'emphasis': True}, {'id': 'B', 'stage': 3, 'title': '客戶B（合併層級10%+）', 'sub': 'FY115Q1銷貨占比12.18%（NT$301,661千），115年新出現', 'emphasis': True}, {'id': 'oth_cus', 'stage': 3, 'title': '其他客戶', 'sub': 'FY2025 83.69%；地區：台灣28.2%／中國28.8%／美國9.2%／其他33.8%'},
        {'id': 'sh_gt', 'stage': 3, 'title': '上海合晶／鄭州合晶（背書保證標的）', 'sub': '母公司對其背書保證限額NT$84.4億', 'branch': 'fin'}],
    'edges': [{'from': 'jia', 'to': 'gwc', 'label': '原料', 'branch': 'main'}, {'from': 'yi', 'to': 'gwc', 'label': '原料', 'branch': 'main'}, {'from': 'oth_sup', 'to': 'gwc', 'branch': 'main'},
        {'from': 'gwc', 'to': 'cn_subs', 'label': '技術／資金', 'branch': 'main'}, {'from': 'gwc', 'to': 'helitek', 'label': '母公司單體銷貨40.41%', 'branch': 'main'}, {'from': 'cn_subs', 'to': 'oth_cus', 'label': '陸廠在地銷售', 'branch': 'main'},
        {'from': 'gwc', 'to': 'A', 'label': 'FY2025 16.0 億', 'branch': 'main'}, {'from': 'gwc', 'to': 'B', 'label': '2026Q1 3.0 億', 'branch': 'main'}, {'from': 'gwc', 'to': 'oth_cus', 'label': 'FY2025 82.2 億', 'branch': 'main'}, {'from': 'helitek', 'to': 'oth_cus', 'label': '終端客戶未揭露', 'branch': 'main'},
        {'from': 'gwc', 'to': 'sh_gt', 'label': '背書保證NT$84.4億', 'branch': 'fin', 'side': 'right'}], 'options': {'showStageLabels': True, 'showLegend': True}}

# 注意：甲/乙供應商與Helitek通路之金額（分別為長約性質敘述與母公司單體口徑40.41%）
# 與此處FY2025合併層級地區別營收非同一統計基礎，無法量化並入同一張流量圖（見chain_supply
# caption說明），故Sankey僅呈現唯一具備完整、一致基礎揭露數字的流量：公司→銷售地區。
S['sankey_supply'] = {'nodes': [{'id': '其他供應商', 'label': '其他供應商 32.3億（無單一≥10%）', 'color': '#60a5fa'},
        {'id': '合晶科技 6182', 'label': '合晶科技 6182（FY2025 合併營收 98.2億）', 'color': '#0f172a'},
        {'id': '客戶A', 'label': '客戶A 16.0億，16.31%', 'color': '#dc2626'}, {'id': '其他客戶', 'label': '其他客戶 82.2億，83.69%', 'color': '#f59e0b'}],
    'links': [{'source': '其他供應商', 'target': '合晶科技 6182', 'value': 32.27}, {'source': '合晶科技 6182', 'target': '客戶A', 'value': 16.01}, {'source': '合晶科技 6182', 'target': '其他客戶', 'value': 82.16}]}
# Sankey 與 chain_supply 同一組實體；客戶B 只在 2026Q1 達 10%（FY2025 未達），故 FY2025 流量圖不列；甲／乙長約供應商無金額揭露，併入「其他供應商」。
# 地區別營收見 bu_mix。

os.makedirs(OUT, exist_ok=True)
for k, v in S.items():
    json.dump(v, open(os.path.join(OUT, k + '.json'), 'w', encoding='utf-8'), ensure_ascii=False)
print(list(S))
