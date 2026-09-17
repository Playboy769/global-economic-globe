import urllib.request, urllib.parse, http.cookiejar, json, time

cj = http.cookiejar.CookieJar()
opener = urllib.request.build_opener(urllib.request.HTTPCookieProcessor(cj))
headers = {'User-Agent':'Mozilla/5.0'}

def get(url):
    req = urllib.request.Request(url, headers=headers)
    return opener.open(req, timeout=20).read()

# Step 1: get cookie
try:
    get('https://fc.yahoo.com')
except Exception as e:
    print('cookie step err (expected 404 ok):', e)
# Step 2: get crumb
crumb = get('https://query1.finance.yahoo.com/v1/test/getcrumb').decode()
print('crumb:', crumb)

tickers = ['6182.TWO','6488.TWO','3532.TW','3016.TW','4063.T','3436.T']
out = {}
for t in tickers:
    modules = 'price,summaryDetail,defaultKeyStatistics,financialData'
    url = f'https://query1.finance.yahoo.com/v10/finance/quoteSummary/{urllib.parse.quote(t)}?modules={modules}&crumb={urllib.parse.quote(crumb)}'
    try:
        data = get(url)
        j = json.loads(data)
        out[t] = j
        print(t, 'OK', len(data))
    except Exception as e:
        print(t, 'ERR', e)
    time.sleep(1)
json.dump(out, open('peers_yahoo.json','w',encoding='utf-8'), indent=1, ensure_ascii=False)
