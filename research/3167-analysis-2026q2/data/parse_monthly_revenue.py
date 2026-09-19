import re, glob, os

months = ["113_8","113_9","113_10","113_11","113_12",
"114_1","114_2","114_3","114_4","114_5","114_6","114_7","114_8","114_9","114_10","114_11","114_12",
"115_1","115_2","115_3","115_4","115_5","115_6","115_7","115_8"]

results = {}
for m in months:
    path = f"t21_{m}.html"
    with open(path, "rb") as f:
        raw = f.read()
    text = raw.decode("big5", errors="ignore")
    # find the row containing 3167
    idx = text.find(">3167<")
    if idx == -1:
        idx = text.find("3167")
    if idx == -1:
        results[m] = None
        continue
    # find enclosing <tr>...</tr>
    tr_start = text.rfind("<tr", 0, idx)
    tr_end = text.find("</tr>", idx)
    row = text[tr_start:tr_end]
    # extract all <td...>...</td> cell contents
    cells = re.findall(r"<td[^>]*>(.*?)</td>", row, re.S)
    cleaned = [re.sub("<[^>]+>", "", c).strip().replace("&nbsp;","").replace(",", "") for c in cells]
    results[m] = cleaned

for m in months:
    print(m, results[m])
