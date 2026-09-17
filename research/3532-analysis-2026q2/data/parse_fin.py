import re, json, glob, os
IS=['ifrs-full:Revenue','tifrs-bsci-ci:OperatingCosts','ifrs-full:GrossProfit','ifrs-full:SellingExpense','ifrs-full:AdministrativeExpense','ifrs-full:ResearchAndDevelopmentExpense','ifrs-full:OperatingExpense','ifrs-full:ProfitLossFromOperatingActivities','tifrs-bsci-ci:NonoperatingIncomeAndExpenses','ifrs-full:FinanceCosts','ifrs-full:ProfitLossBeforeTax','ifrs-full:IncomeTaxExpenseContinuingOperations','ifrs-full:ProfitLoss','ifrs-full:ProfitLossAttributableToOwnersOfParent','ifrs-full:BasicEarningsLossPerShare','tifrs-notes:InterestExpense_n','tifrs-notes:ForeignExchangeGainsLosses_n','tifrs-notes:InterestIncome_n','ifrs-full:OtherGainsLosses','ifrs-full:ComprehensiveIncomeAttributableToOwnersOfParent']
BS=['ifrs-full:CashAndCashEquivalents','tifrs-bsci-ci:AccountsReceivableNet','ifrs-full:Inventories','ifrs-full:CurrentAssets','ifrs-full:PropertyPlantAndEquipment','ifrs-full:RightofuseAssets','ifrs-full:Assets','ifrs-full:ShorttermBorrowings','ifrs-full:CurrentContractLiabilities','ifrs-full:NoncurrentContractLiabilities','ifrs-full:CurrentPortionOfLongtermBorrowings','tifrs-bsci-ci:LongtermLiabilitiesCurrentPortion','ifrs-full:CurrentBondsIssuedAndCurrentPortionOfNoncurrentBondsIssued','ifrs-full:CurrentLiabilities','ifrs-full:LongtermBorrowings','tifrs-bsci-ci:BondsPayable-BondsPayable','ifrs-full:NoncurrentPortionOfNoncurrentBondsIssued','ifrs-full:Liabilities','ifrs-full:EquityAttributableToOwnersOfParent','ifrs-full:NoncontrollingInterests','ifrs-full:Equity','ifrs-full:IssuedCapital','ifrs-full:TradeAndOtherCurrentPayablesToTradeSuppliers','ifrs-full:OtherCurrentLiabilities','ifrs-full:OtherNoncurrentLiabilities','ifrs-full:CurrentFinancialAssetsAtFairValueThroughProfitOrLoss','ifrs-full:OtherCurrentFinancialAssets','ifrs-full:NoncurrentFinancialAssetsAtFairValueThroughOtherComprehensiveIncome','ifrs-full:IntangibleAssetsAndGoodwill','ifrs-full:OtherNoncurrentAssets','ifrs-full:OtherCurrentAssets','ifrs-full:NoncurrentLiabilities','ifrs-full:RetainedEarnings','ifrs-full:CapitalReserve','ifrs-full:OtherEquityInterest','ifrs-full:DeferredTaxAssets','ifrs-full:DeferredTaxLiabilities','ifrs-full:OtherCurrentPayables','ifrs-full:NoncurrentRecognisedLiabilitiesDefinedBenefitPlan']
CF=['ifrs-full:CashFlowsFromUsedInOperatingActivities','tifrs-SCF:NetCashFlowsFromUsedInInvestingActivities','tifrs-SCF:CashFlowsFromUsedInFinancingActivities','ifrs-full:PurchaseOfPropertyPlantAndEquipmentClassifiedAsInvestingActivities','ifrs-full:AdjustmentsForDepreciationExpense','ifrs-full:AdjustmentsForAmortisationExpense','ifrs-full:DividendsPaidClassifiedAsFinancingActivities','tifrs-SCF:IncreaseDecreaseInContractLiabilities','tifrs-SCF:ProceedsFromLongTermDebt','tifrs-SCF:RepaymentsOfLongTermDebt','tifrs-SCF:ProceedsFromIssuingBonds','tifrs-SCF:RepaymentsOfBonds','tifrs-SCF:IncreaseInShortTermLoans','tifrs-SCF:DecreaseInShortTermLoans','ifrs-full:InterestPaidClassifiedAsOperatingActivities','ifrs-full:IncomeTaxesPaidRefundClassifiedAsOperatingActivities','ifrs-full:IncreaseDecreaseInCashAndCashEquivalents','tifrs-SCF:DecreaseIncreaseInAccountsReceivable','ifrs-full:AdjustmentsForDecreaseIncreaseInInventories','tifrs-SCF:IncreaseDecreaseInAccountsPayable','ifrs-full:PaymentsOfLeaseLiabilitiesClassifiedAsFinancingActivities','tifrs-SCF:AdjustmentsToReconcileProfitLoss','tifrs-SCF:ChangesInOperatingAssetsAndLiabilities','tifrs-SCF:ProfitLossBeforeTax','ifrs-full:CashFlowsFromUsedInOperations']
pat=re.compile(r'<ix:nonFraction([^>]*)>(.*?)</ix:nonFraction>',re.S)
def attrs(s): return dict(re.findall(r'(\w+)="([^"]*)"',s))
def num(txt,a):
    t=re.sub(r'<[^>]+>','',txt).replace(',','').strip()
    if t in ('','-','—'): return 0.0
    try: v=float(t)
    except: return None
    if a.get('sign')=='-': v=-v
    sc=int(a.get('scale','0') or 0)
    return v*(10**sc)
out={}
for f in sorted(glob.glob('raw/fin_*.html')):
    q=os.path.basename(f)[4:10]; y=int(q[:4]); s=int(q[5])
    h=open(f,'rb').read().decode('cp950','replace')
    facts={}
    for a,txt in pat.findall(h):
        d=attrs(a); n=d.get('name'); c=d.get('contextRef','')
        if '_' in c: continue  # skip member contexts
        v=num(txt,d)
        if v is None: continue
        facts.setdefault(n,{})[c]=v
    qend={1:'0331',2:'0630',3:'0930',4:'1231'}[s]
    asof=f'AsOf{y}{qend}'
    qstart={1:'0101',2:'0401',3:'0701',4:'1001'}[s]
    q3m=f'From{y}{qstart}To{y}{qend}'
    ytd=f'From{y}0101To{y}{qend}'
    rec={'IS_3M':{},'IS_YTD':{},'BS':{},'CF_YTD':{}}
    for n in IS:
        if n in facts:
            if q3m in facts[n]: rec['IS_3M'][n]=facts[n][q3m]
            if ytd in facts[n]: rec['IS_YTD'][n]=facts[n][ytd]
    for n in BS:
        if n in facts and asof in facts[n]: rec['BS'][n]=facts[n][asof]
    for n in CF:
        if n in facts and ytd in facts[n]: rec['CF_YTD'][n]=facts[n][ytd]
    rec['contexts']=sorted(set(c for n in facts for c in facts[n]))
    out[q]=rec
    print(q, 'IS3M',len(rec['IS_3M']),'ISYTD',len(rec['IS_YTD']),'BS',len(rec['BS']),'CF',len(rec['CF_YTD']), 'Rev3M', rec['IS_3M'].get('ifrs-full:Revenue'), 'RevYTD', rec['IS_YTD'].get('ifrs-full:Revenue'))
json.dump(out,open('financials_raw.json','w'),indent=1)
