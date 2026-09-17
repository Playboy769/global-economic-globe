import json
d=json.load(open('financials_raw.json'))
qs=sorted(d)
def g(q,sec,n): return d[q][sec].get(n)
IS_keys={'Rev':'ifrs-full:Revenue','COGS':'tifrs-bsci-ci:OperatingCosts','GP':'ifrs-full:GrossProfit','Sell':'ifrs-full:SellingExpense','Adm':'ifrs-full:AdministrativeExpense','RD':'ifrs-full:ResearchAndDevelopmentExpense','Opex':'ifrs-full:OperatingExpense','OpInc':'ifrs-full:ProfitLossFromOperatingActivities','NonOp':'tifrs-bsci-ci:NonoperatingIncomeAndExpenses','FinCost':'ifrs-full:FinanceCosts','PBT':'ifrs-full:ProfitLossBeforeTax','Tax':'ifrs-full:IncomeTaxExpenseContinuingOperations','NI':'ifrs-full:ProfitLoss','NI_parent':'ifrs-full:ProfitLossAttributableToOwnersOfParent','EPS':'ifrs-full:BasicEarningsLossPerShare','IntExp':'tifrs-notes:InterestExpense_n','FX':'tifrs-notes:ForeignExchangeGainsLosses_n','IntInc':'tifrs-notes:InterestIncome_n','OtherGL':'ifrs-full:OtherGainsLosses','CI_parent':'ifrs-full:ComprehensiveIncomeAttributableToOwnersOfParent'}
CF_keys={'CFO':'ifrs-full:CashFlowsFromUsedInOperatingActivities','CFI':'tifrs-SCF:NetCashFlowsFromUsedInInvestingActivities','CFF':'tifrs-SCF:CashFlowsFromUsedInFinancingActivities','Capex':'ifrs-full:PurchaseOfPropertyPlantAndEquipmentClassifiedAsInvestingActivities','Dep':'ifrs-full:AdjustmentsForDepreciationExpense','Amort':'ifrs-full:AdjustmentsForAmortisationExpense','Div':'ifrs-full:DividendsPaidClassifiedAsFinancingActivities','dContractLiab':'tifrs-SCF:IncreaseDecreaseInContractLiabilities','LTDproceeds':'tifrs-SCF:ProceedsFromLongTermDebt','LTDrepay':'tifrs-SCF:RepaymentsOfLongTermDebt','Bonds':'tifrs-SCF:ProceedsFromIssuingBonds','BondRepay':'tifrs-SCF:RepaymentsOfBonds','STLup':'tifrs-SCF:IncreaseInShortTermLoans','STLdn':'tifrs-SCF:DecreaseInShortTermLoans','IntPaid':'ifrs-full:InterestPaidClassifiedAsOperatingActivities','TaxPaid':'ifrs-full:IncomeTaxesPaidRefundClassifiedAsOperatingActivities','dAR':'tifrs-SCF:DecreaseIncreaseInAccountsReceivable','dInv':'ifrs-full:AdjustmentsForDecreaseIncreaseInInventories','dAP':'tifrs-SCF:IncreaseDecreaseInAccountsPayable','Lease':'ifrs-full:PaymentsOfLeaseLiabilitiesClassifiedAsFinancingActivities'}
BS_keys={'Cash':'ifrs-full:CashAndCashEquivalents','AR':'tifrs-bsci-ci:AccountsReceivableNet','Inv':'ifrs-full:Inventories','CA':'ifrs-full:CurrentAssets','PPE':'ifrs-full:PropertyPlantAndEquipment','ROU':'ifrs-full:RightofuseAssets','Assets':'ifrs-full:Assets','STB':'ifrs-full:ShorttermBorrowings','CL_contract':'ifrs-full:CurrentContractLiabilities','NCL_contract':'ifrs-full:NoncurrentContractLiabilities','LTB_cur':'ifrs-full:CurrentPortionOfLongtermBorrowings','LTL_cur':'tifrs-bsci-ci:LongtermLiabilitiesCurrentPortion','Bonds_cur':'ifrs-full:CurrentBondsIssuedAndCurrentPortionOfNoncurrentBondsIssued','CL':'ifrs-full:CurrentLiabilities','LTB':'ifrs-full:LongtermBorrowings','BondsPay':'tifrs-bsci-ci:BondsPayable-BondsPayable','Bonds_nc':'ifrs-full:NoncurrentPortionOfNoncurrentBondsIssued','Liab':'ifrs-full:Liabilities','Eq_parent':'ifrs-full:EquityAttributableToOwnersOfParent','NCI':'ifrs-full:NoncontrollingInterests','Equity':'ifrs-full:Equity','Capital':'ifrs-full:IssuedCapital','AP':'ifrs-full:TradeAndOtherCurrentPayablesToTradeSuppliers','OtherCL':'ifrs-full:OtherCurrentLiabilities','OtherNCL':'ifrs-full:OtherNoncurrentLiabilities','FVTPL_cur':'ifrs-full:CurrentFinancialAssetsAtFairValueThroughProfitOrLoss','OtherFA_cur':'ifrs-full:OtherCurrentFinancialAssets','FVOCI_nc':'ifrs-full:NoncurrentFinancialAssetsAtFairValueThroughOtherComprehensiveIncome','Intang':'ifrs-full:IntangibleAssetsAndGoodwill','OtherNCA':'ifrs-full:OtherNoncurrentAssets','NCL':'ifrs-full:NoncurrentLiabilities','RE':'ifrs-full:RetainedEarnings','OtherEq':'ifrs-full:OtherEquityInterest','OtherPay':'ifrs-full:OtherCurrentPayables'}
rows={}
for i,q in enumerate(qs):
    r={}
    is3=d[q]['IS_3M']; isy=d[q]['IS_YTD']
    for k,n in IS_keys.items():
        if is3: r[k]=is3.get(n)
        else:
            prev=qs[i-1]
            a=isy.get(n); b=d[prev]['IS_YTD'].get(n)
            r[k]=(a-b) if (a is not None and b is not None) else None
    for k,n in BS_keys.items(): r[k]=d[q]['BS'].get(n)
    cfy=d[q]['CF_YTD']
    for k,n in CF_keys.items():
        a=cfy.get(n)
        if q.endswith('Q1') or i==0:
            r[k]=a if q.endswith('Q1') else None
        else:
            b=d[qs[i-1]]['CF_YTD'].get(n)
            r[k]=(a-b) if (a is not None and b is not None) else None
        r[k+'_ytd']=a
    rows[q]=r
json.dump(rows,open('quarterly.json','w'),indent=1)
def M(v): return '' if v is None else f'{v/1e6:,.0f}'
def P(a,b): return '' if (a is None or not b) else f'{a/b*100:.1f}%'
print('| 項目 | '+' | '.join(qs)+' |')
for k in ['Rev','COGS','GP','Opex','RD','OpInc','NonOp','FinCost','IntExp','FX','PBT','Tax','NI','NI_parent','CI_parent']:
    print(f'| {k} | '+' | '.join(M(rows[q][k]) for q in qs)+' |')
print('| EPS | '+' | '.join('' if rows[q]['EPS'] is None else f"{rows[q]['EPS']:.2f}" for q in qs)+' |')
print('| GM | '+' | '.join(P(rows[q]['GP'],rows[q]['Rev']) for q in qs)+' |')
print('| OPM | '+' | '.join(P(rows[q]['OpInc'],rows[q]['Rev']) for q in qs)+' |')
print('| NPM(parent) | '+' | '.join(P(rows[q]['NI_parent'],rows[q]['Rev']) for q in qs)+' |')
print()
for k in BS_keys: print(f'| {k} | '+' | '.join(M(rows[q][k]) for q in qs)+' |')
print()
for k in CF_keys: print(f'| {k} (Q) | '+' | '.join(M(rows[q][k]) for q in qs)+' |')
print()
for k in ['CFO','CFI','CFF','Capex','Dep','Div']: print(f'| {k} (YTD) | '+' | '.join(M(rows[q][k+'_ytd']) for q in qs)+' |')
