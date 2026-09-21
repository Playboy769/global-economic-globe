import re, json, glob, os

files = sorted(glob.glob('raw/monthly/t21_*.html'), key=lambda f: (int(f.split('_')[1]), int(f.split('_')[2].split('.')[0])))
rows = []
for f in files:
    base = os.path.basename(f)[:-5]
    _, y, m = base.split('_')
    y, m = int(y), int(m)
    h = open(f, 'rb').read().decode('cp950', 'replace')
    idx = h.find('>2637<')
    if idx < 0:
        idx = h.find("center'>2637<")
    # find the <tr...>...2637...</tr> block
    tr_start = h.rfind('<tr', 0, idx)
    tr_end = h.find('</tr>', idx)
    block = h[tr_start:tr_end]
    cells = re.findall(r'<td[^>]*>(.*?)</td>', block, re.S | re.I)
    cells = [re.sub(r'<[^>]+>', '', c).strip() for c in cells]
    # cells: [code, name, cur, lastmonth, lastyear_samemonth, mom%, yoy%, cum, cum_lastyear, cum_yoy%, note]
    def num(s):
        s = s.replace(',', '').strip()
        if s in ('', '-', '—'):
            return None
        try:
            return float(s)
        except:
            return None
    rec = {
        'y_roc': y, 'm': m,
        'ym': f'{y+1911}-{m:02d}',
        'code': cells[0], 'name': cells[1],
        'rev_k': num(cells[2]),
        'rev_lastmonth_k': num(cells[3]),
        'rev_lastyear_samemonth_k': num(cells[4]),
        'mom_pct': num(cells[5]),
        'yoy_pct': num(cells[6]),
        'cum_k': num(cells[7]),
        'cum_lastyear_k': num(cells[8]),
        'cum_yoy_pct': num(cells[9]) if len(cells) > 9 else None,
    }
    rows.append(rec)

rows.sort(key=lambda r: r['ym'])

# cross-validate: this month's cum - last month's cum (within same ROC year, Jan resets) should equal this month's rev
by_ym = {r['ym']: r for r in rows}
issues = []
for i, r in enumerate(rows):
    if r['m'] == 1:
        continue  # cum resets in Jan, can't check against prior year's Dec cum this way
    prev = rows[i-1]
    if prev['ym'][:4] != r['ym'][:4]:
        continue
    if r['cum_k'] is not None and prev['cum_k'] is not None:
        implied = r['cum_k'] - prev['cum_k']
        diff = abs(implied - r['rev_k']) if r['rev_k'] is not None else None
        r['implied_single_month_k'] = implied
        if diff is not None and diff > 5:  # tolerance 5 thousand NTD
            issues.append((r['ym'], r['rev_k'], implied, diff))
    else:
        r['implied_single_month_k'] = None

json.dump(rows, open('monthly_revenue.json', 'w', encoding='utf-8'), indent=1, ensure_ascii=False)
print(f'{len(rows)} months parsed, {len(issues)} validation issues')
for it in issues:
    print('ISSUE', it)
for r in rows:
    print(r['ym'], r['rev_k'], 'YoY%', r['yoy_pct'], 'cum', r['cum_k'])
