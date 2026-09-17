import urllib.request, json, time

headers = {'User-Agent':'Mozilla/5.0'}
def get(url):
    req = urllib.request.Request(url, headers=headers)
    return urllib.request.urlopen(req, timeout=20).read()

url = 'https://query1.finance.yahoo.com/v8/finance/chart/6182.TWO?range=2y&interval=1wk'
data = get(url)
j = json.loads(data)
open('price_weekly.json','w',encoding='utf-8').write(json.dumps(j, indent=1, ensure_ascii=False))
res = j['chart']['result'][0]
ts = res['timestamp']
closes = res['indicators']['quote'][0]['close']
print('points:', len(ts))
print('first:', ts[0], closes[0])
print('last:', ts[-1], closes[-1])
