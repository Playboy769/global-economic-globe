import re, json

lines = [l.strip() for l in open('raw/fleet_raw.txt', encoding='utf-8')]
lines = [l for l in lines if l and l != '[PAGE 99]']
# find start: first line that is a pure integer '1'
start = None
for i, l in enumerate(lines):
    if l == '1':
        start = i
        break
lines = lines[start:]

ships = []
i = 0
n = len(lines)
expect = 1
while i < n:
    if lines[i].isdigit() and int(lines[i]) == expect:
        seq = int(lines[i])
        name = lines[i+1]
        year = lines[i+2]
        dwt = lines[i+3]
        stype = lines[i+4]
        ships.append({'seq': seq, 'name': name, 'build_year': year, 'dwt': dwt.replace(',', ''), 'type': stype})
        i += 5
        expect += 1
    else:
        i += 1

json.dump(ships, open('fleet_2026q2.json', 'w', encoding='utf-8'), indent=1, ensure_ascii=False)
print(len(ships), 'ships parsed')
from collections import Counter
c = Counter(s['type'] for s in ships)
print(c)
print(ships[:3])
print(ships[-3:])
