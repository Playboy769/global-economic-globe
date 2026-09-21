import json, glob, os

out = {}
for f in sorted(glob.glob('raw/peers/*.json')):
    t = os.path.basename(f).replace('.json', '').replace('_', '.')
    d = json.load(open(f))
    try:
        res = d['quoteSummary']['result'][0]
    except Exception:
        out[t] = {'error': d.get('quoteSummary', {}).get('error')}
        continue
    sd = res.get('summaryDetail', {})
    ks = res.get('defaultKeyStatistics', {})
    fd = res.get('financialData', {})
    pr = res.get('price', {})

    def gv(x):
        return x.get('raw') if isinstance(x, dict) else x

    out[t] = {
        'shortName': pr.get('shortName'),
        'currency': pr.get('currency'),
        'marketCap': gv(pr.get('marketCap')),
        'regularMarketPrice': gv(pr.get('regularMarketPrice')),
        'trailingPE': gv(sd.get('trailingPE')),
        'forwardPE': gv(sd.get('forwardPE')),
        'priceToBook': gv(ks.get('priceToBook')),
        'revenueGrowth': gv(fd.get('revenueGrowth')),
        'earningsGrowth': gv(fd.get('earningsGrowth')),
        'earningsQuarterlyGrowth': gv(ks.get('earningsQuarterlyGrowth')),
        'totalRevenue': gv(fd.get('totalRevenue')),
        'grossMargins': gv(fd.get('grossMargins')),
        'returnOnEquity': gv(fd.get('returnOnEquity')),
        'dividendYield': gv(sd.get('dividendYield')),
    }

json.dump(out, open('peers_yahoo.json', 'w'), indent=1)
for k, v in out.items():
    print(k, v.get('shortName'), v.get('marketCap'), v.get('trailingPE'), v.get('priceToBook'))
