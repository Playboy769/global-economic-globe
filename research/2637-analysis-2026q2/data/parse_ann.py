# -*- coding: utf-8 -*-
import re, html

rows = []
for fn in ['raw/ann_2026.html', 'raw/ann_2025.html']:
    h = open(fn, 'rb').read().decode('utf-8', 'replace')
    trs = re.findall(r'<tr[^>]*>(.*?)</tr>', h, re.S | re.I)
    for tr in trs:
        cells = re.findall(r'<td[^>]*>(.*?)</td>', tr, re.S | re.I)
        cells = [re.sub(r'<[^>]+>', ' ', c) for c in cells]
        cells = [html.unescape(re.sub(r'\s+', ' ', c)).strip() for c in cells]
        if len(cells) >= 5 and re.match(r'^\d{3}/\d{2}/\d{2}$', cells[2]):
            rows.append(cells)

rows.sort(key=lambda r: r[2] + r[3])
lines = ['# 慧洋-KY (2637) 重大訊息（2025-06 起）\n', '來源：MOPS `ajax_t05st01`，UTF-8 解碼。\n']
for r in rows:
    if r[2] >= '114/06/01':
        lines.append(f'- **{r[2]} {r[3]}**：{r[4]}')

open('announcements_raw.md', 'w', encoding='utf-8').write('\n'.join(lines))
print(len([r for r in rows if r[2] >= '114/06/01']), 'items since 2025/06')
print(len(rows), 'total rows')
