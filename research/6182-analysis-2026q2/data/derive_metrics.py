import json
d = json.load(open('quarterly.json'))
qs = sorted(d)
out = {}
for q in qs:
    r = d[q]
    m = {}
    rev = r.get('Rev')
    ar = r.get('AR')
    inv = r.get('Inv')
    cogs = r.get('COGS')
    # DSO using 91 days (quarter convention per instructions)
    m['DSO'] = (ar/rev*91) if (ar is not None and rev not in (None,0)) else None
    # Inventory days: use COGS (abs, since COGS reported may be negative sign convention) 91 days
    cogs_abs = abs(cogs) if cogs is not None else None
    m['InvDays'] = (inv/cogs_abs*91) if (inv is not None and cogs_abs not in (None,0)) else None
    m['GM'] = (r['GP']/rev*100) if (r.get('GP') is not None and rev) else None
    m['OPM'] = (r['OpInc']/rev*100) if (r.get('OpInc') is not None and rev) else None
    m['NPM_parent'] = (r['NI_parent']/rev*100) if (r.get('NI_parent') is not None and rev) else None
    # Interest coverage = OpInc / IntExp (quarterly, abs)
    intexp = r.get('IntExp')
    intexp_abs = abs(intexp) if intexp is not None else None
    m['IntCoverage'] = (r['OpInc']/intexp_abs) if (r.get('OpInc') is not None and intexp_abs) else None
    # FCF = CFO(Q) - |Capex(Q)|
    cfo = r.get('CFO')
    capex = r.get('Capex')
    m['FCF'] = (cfo - abs(capex)) if (cfo is not None and capex is not None) else None
    dep = r.get('Dep')
    amort = r.get('Amort')
    m['D&A'] = ( (dep or 0) + (amort or 0) ) if (dep is not None or amort is not None) else None
    out[q] = m
json.dump(out, open('derived_metrics.json','w'), indent=1, ensure_ascii=False)
def M(v):
    return '' if v is None else f'{v:,.1f}'
def M2(v):
    return '' if v is None else f'{v/1e6:,.0f}'
print('| 指標 | ' + ' | '.join(qs) + ' |')
for k in ['DSO','InvDays','GM','OPM','NPM_parent','IntCoverage']:
    print(f'| {k} | ' + ' | '.join(M(out[q][k]) for q in qs) + ' |')
print(f"| FCF(M) | " + ' | '.join(M2(out[q]['FCF']) for q in qs) + ' |')
