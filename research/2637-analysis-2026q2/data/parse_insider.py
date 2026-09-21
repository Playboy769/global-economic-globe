# -*- coding: utf-8 -*-
import re, glob, os

out_lines = ['# 慧洋-KY (2637) 董監大股東持股變化（2026/03–2026/08）\n',
             '來源：MOPS `ajax_stapap1`（內部人持股月報），cp950 解碼。\n']

for f in sorted(glob.glob('raw/insider_115_*.html'), key=lambda x: int(x.split('_')[-1].split('.')[0])):
    m = int(f.split('_')[-1].split('.')[0])
    h = open(f, 'rb').read().decode('utf-8', 'replace')
    ym = re.search(r'資料年月:(\d+)', h)
    ym = ym.group(1) if ym else f'115{m:02d}'
    rows = re.findall(r"<TR class='(?:odd|even)'>(.*?)</TR>", h, re.S)
    out_lines.append(f'\n## {int(ym[:3])+1911}/{int(ym[3:]):02d}\n')
    out_lines.append('| 職稱 | 姓名 | 期初持股 | 目前持股 | 設質股數 | 設質佔持股% |')
    out_lines.append('|---|---|---|---|---|---|')
    for r in rows:
        cells = re.findall(r"<TD[^>]*>(.*?)</TD>", r, re.S | re.I)
        cells = [re.sub(r'<[^>]+>', '', c).strip() for c in cells]
        if len(cells) >= 6:
            out_lines.append('| ' + ' | '.join(cells[:6]) + ' |')

open('insider_and_announcements.md', 'w', encoding='utf-8').write('\n'.join(out_lines))
print('done, lines:', len(out_lines))
