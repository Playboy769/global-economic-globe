import json
d=json.load(open('financials_raw.json'))
qs=sorted(d)
def g(q,sec,n): return d[q][sec].get(n)
IS_keys={'Rev':'ifrs-full:Revenue','OtherRev':'ifrs-full:OtherRevenue','COGS':'tifrs-bsci-ci:OperatingCosts','GP':'ifrs-full:GrossProfit','Adm':'ifrs-full:AdministrativeExpense','Opex':'ifrs-full:OperatingExpense','OpInc':'ifrs-full:ProfitLossFromOperatingActivities','FinCost':'ifrs-full:FinanceCosts','ShipGain':'ifrs-full:GainsOnDisposalsOfPropertyPlantAndEquipment','EquityMethod':'ifrs-full:ShareOfProfitLossOfAssociatesAndJointVenturesAccountedForUsingEquityMethod','PBT':'ifrs-full:ProfitLossBeforeTax','Tax':'ifrs-full:IncomeTaxExpenseContinuingOperations','NI':'ifrs-full:ProfitLoss','NI_parent':'ifrs-full:ProfitLossAttributableToOwnersOfParent','EPS':'ifrs-full:BasicEarningsLossPerShare','EPS_diluted':'ifrs-full:DilutedEarningsLossPerShare'}
CF_keys={'CFO':'ifrs-full:CashFlowsFromUsedInOperatingActivities','CFI':'tifrs-SCF:NetCashFlowsFromUsedInInvestingActivities','CFF':'tifrs-SCF:CashFlowsFromUsedInFinancingActivities','Capex':'ifrs-full:PurchaseOfPropertyPlantAndEquipmentClassifiedAsInvestingActivities','ShipDisposalProceeds':'ifrs-full:ProceedsFromSalesOfPropertyPlantAndEquipmentClassifiedAsInvestingActivities','Dep':'ifrs-full:AdjustmentsForDepreciationExpense','Div':'ifrs-full:DividendsPaidClassifiedAsFinancingActivities','Lease':'ifrs-full:PaymentsOfLeaseLiabilitiesClassifiedAsFinancingActivities','TaxPaid':'ifrs-full:IncomeTaxesPaidRefundClassifiedAsOperatingActivities'}
BS_keys={'Cash':'ifrs-full:CashAndCashEquivalents','AR':'tifrs-bsci-ci:AccountsReceivableNet','CA':'ifrs-full:CurrentAssets','PPE':'ifrs-full:PropertyPlantAndEquipment','ROU':'ifrs-full:RightofuseAssets','Assets':'ifrs-full:Assets','STB':'ifrs-full:ShorttermBorrowings','LTB_cur':'ifrs-full:CurrentPortionOfLongtermBorrowings','LeaseLiab_cur':'ifrs-full:CurrentLeaseLiabilities','CL':'ifrs-full:CurrentLiabilities','LTB':'ifrs-full:LongtermBorrowings','Bonds_nc':'ifrs-full:NoncurrentPortionOfNoncurrentBondsIssued','LeaseLiab_nc':'ifrs-full:NoncurrentLeaseLiabilities','NCL':'ifrs-full:NoncurrentLiabilities','Liab':'ifrs-full:Liabilities','Eq_parent':'ifrs-full:EquityAttributableToOwnersOfParent','Equity':'ifrs-full:Equity','Capital':'ifrs-full:IssuedCapital','RE':'ifrs-full:RetainedEarnings','CurrentAdvances':'ifrs-full:CurrentAdvances','CurrentPrepayments':'ifrs-full:CurrentPrepayments','PrepayToRelated':'tifrs-notes:TotalPrepaymentsToRelatedParties'}
rows={}
for i,q in enumerate(qs):
    r={}
    is3=d[q]['IS_3M']; isy=d[q]['IS_YTD']
    for k,n in IS_keys.items():
        if is3:
            r[k]=is3.get(n)
        else:
            prev=qs[i-1]
            a=isy.get(n); b=d[prev]['IS_YTD'].get(n)
            r[k]=(a-b) if (a is not None and b is not None) else None
    for k,n in BS_keys.items(): r[k]=d[q]['BS'].get(n)
    cfy=d[q]['CF_YTD']
    for k,n in CF_keys.items():
        a=cfy.get(n)
        if q.endswith('Q1'):
            r[k]=a
        else:
            b=d[qs[i-1]]['CF_YTD'].get(n)
            r[k]=(a-b) if (a is not None and b is not None) else None
        r[k+'_ytd']=a
    rows[q]=r
json.dump(rows,open('quarterly.json','w'),indent=1)

def M(v): return '' if v is None else f'{v/1e6:,.1f}'
def P(a,b): return '' if (a is None or not b) else f'{a/b*100:.1f}%'

lines=[]
lines.append('# 慧洋-KY (2637) 財報數據 12 季（2023Q3–2026Q2）')
lines.append('')
lines.append('單位：百萬元 NTD（原始 XBRL 為新台幣仟元），除 EPS 外。來源：MOPS 合併財報 t164sb01，Q4 = 全年 YTD − 前三季 YTD。')
lines.append('')
lines.append('## 損益表（單季）')
lines.append('| 項目 | '+' | '.join(qs)+' |')
lines.append('|---|'+'---|'*len(qs))
for k,label in [('Rev','營業收入'),('OtherRev','其他收入'),('COGS','營業成本'),('GP','營業毛利'),('Adm','管理費用'),('Opex','營業費用'),('OpInc','營業利益'),('FinCost','財務成本'),('ShipGain','處分資產利益'),('EquityMethod','權益法認列損益'),('PBT','稅前淨利'),('Tax','所得稅費用'),('NI','本期淨利'),('NI_parent','歸屬母公司淨利')]:
    lines.append(f'| {label} | '+' | '.join(M(rows[q][k]) for q in qs)+' |')
lines.append('| EPS（基本） | '+' | '.join('' if rows[q]['EPS'] is None else f"{rows[q]['EPS']:.2f}" for q in qs)+' |')
lines.append('| GM 毛利率 | '+' | '.join(P(rows[q]['GP'],rows[q]['Rev']) for q in qs)+' |')
lines.append('| OPM 營業利益率 | '+' | '.join(P(rows[q]['OpInc'],rows[q]['Rev']) for q in qs)+' |')
lines.append('| NPM(母公司) | '+' | '.join(P(rows[q]['NI_parent'],rows[q]['Rev']) for q in qs)+' |')
lines.append('')
lines.append('## 資產負債表（期末）')
lines.append('| 項目 | '+' | '.join(qs)+' |')
lines.append('|---|'+'---|'*len(qs))
bs_labels={'Cash':'現金及約當現金','AR':'應收帳款淨額','CA':'流動資產','PPE':'不動產廠房及設備(含船舶)','ROU':'使用權資產','Assets':'資產總額','STB':'短期借款','LTB_cur':'一年內到期長期借款','LeaseLiab_cur':'流動租賃負債','CL':'流動負債','LTB':'長期借款','Bonds_nc':'應付公司債(非流動)','LeaseLiab_nc':'非流動租賃負債','NCL':'非流動負債','Liab':'負債總額','Eq_parent':'歸屬母公司權益','Equity':'權益總額','Capital':'股本','RE':'保留盈餘','CurrentAdvances':'預收款項','CurrentPrepayments':'預付款項(含預付船款)','PrepayToRelated':'對關係人預付款'}
for k,label in bs_labels.items():
    lines.append(f'| {label} | '+' | '.join(M(rows[q][k]) for q in qs)+' |')
lines.append('')
lines.append('## 現金流量表（單季）')
lines.append('| 項目 | '+' | '.join(qs)+' |')
lines.append('|---|'+'---|'*len(qs))
cf_labels={'CFO':'營業活動現金流','CFI':'投資活動現金流','CFF':'籌資活動現金流','Capex':'購置不動產廠房設備(造船/購船)','ShipDisposalProceeds':'處分不動產廠房設備價款','Dep':'折舊費用','Div':'發放股利','Lease':'租賃負債清償','TaxPaid':'支付所得稅'}
for k,label in cf_labels.items():
    lines.append(f'| {label} (單季) | '+' | '.join(M(rows[q][k]) for q in qs)+' |')
lines.append('')
for k,label in cf_labels.items():
    lines.append(f'| {label} (YTD) | '+' | '.join(M(rows[q][k+'_ytd']) for q in qs)+' |')

open('financial_table.md','w',encoding='utf-8').write('\n'.join(lines))
print('\n'.join(lines[:40]))
