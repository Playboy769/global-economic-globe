import re, json, glob, os

def parse_file(path, roc_year, month):
    h = open(path, 'rb').read().decode('big5', 'replace')
    # find the row starting with company code 3532
    m = re.search(r'<tr[^>]*><td align=center>3532</td>.*?</tr>', h, re.S | re.I)
    if not m:
        return None
    row = m.group(0)
    tds = re.findall(r'<td[^>]*>(.*?)</td>', row, re.S | re.I)
    vals = [re.sub(r'<[^>]+>', '', t).strip() for t in tds]
    def num(s):
        s = s.replace(',', '').strip()
        if s in ('', '-', '—'):
            return None
        try:
            return float(s)
        except ValueError:
            return None
    # vals: [code, name, curMonth, lastMonth, lastYearSameMonth, MoM%, YoY%, cumThis, cumLastYear, cumYoY%, note]
    rec = {
        'code': vals[0],
        'name': vals[1],
        'cur_month_rev': num(vals[2]),
        'last_month_rev': num(vals[3]),
        'last_year_same_month_rev': num(vals[4]),
        'mom_pct': num(vals[5]),
        'yoy_pct': num(vals[6]),
        'cum_this_year': num(vals[7]),
        'cum_last_year': num(vals[8]),
        'cum_yoy_pct': num(vals[9]) if len(vals) > 9 else None,
    }
    rec['roc_year'] = roc_year
    rec['month'] = month
    rec['ad_year'] = roc_year + 1911
    return rec

out = {}
for f in sorted(glob.glob('raw/mrev_*.html')):
    base = os.path.basename(f)
    m = re.match(r'mrev_(\d+)_(\d+)\.html', base)
    y, mo = int(m.group(1)), int(m.group(2))
    rec = parse_file(f, y, mo)
    key = f'{y+1911}-{mo:02d}'
    out[key] = rec
    if rec:
        print(key, 'cur', rec['cur_month_rev'], 'cumThis', rec['cum_this_year'], 'YoY%', rec['yoy_pct'])
    else:
        print(key, 'NOT FOUND')

# cross-check: cum_this_year[month] - cum_this_year[month-1] should == cur_month_rev (except Jan)
keys = sorted(out.keys())
print('\n--- cross-check cumulative differences ---')
for i, k in enumerate(keys):
    r = out[k]
    if r is None:
        continue
    mo = r['month']
    if mo == 1 or i == 0:
        continue
    prev_key = keys[i-1]
    prev = out.get(prev_key)
    if prev is None or prev['cum_this_year'] is None or r['cum_this_year'] is None:
        continue
    diff = r['cum_this_year'] - prev['cum_this_year']
    match = abs(diff - r['cur_month_rev']) < 1.0 if r['cur_month_rev'] is not None else None
    if not match:
        print(f'MISMATCH {k}: cur_month_rev={r["cur_month_rev"]} but cum_diff={diff}')

json.dump(out, open('monthly_revenue.json', 'w'), indent=1, ensure_ascii=False)
print('\nSaved monthly_revenue.json,', len(out), 'months')
