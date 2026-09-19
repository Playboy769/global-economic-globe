import re

months = ["113_8","113_9","113_10","113_11","113_12",
"114_1","114_2","114_3","114_4","114_5","114_6","114_7","114_8","114_9","114_10","114_11","114_12",
"115_1","115_2","115_3","115_4","115_5","115_6","115_7","115_8"]

label = {"113":"2024","114":"2025","115":"2026"}

rows = []
for m in months:
    path = f"t21_{m}.html"
    with open(path, "rb") as f:
        raw = f.read()
    text = raw.decode("big5", errors="ignore")
    idx = text.find(">3167<")
    tr_start = text.rfind("<tr", 0, idx)
    tr_end = text.find("</tr>", idx)
    row = text[tr_start:tr_end]
    cells = re.findall(r"<td[^>]*>(.*?)</td>", row, re.S)
    cleaned = [re.sub("<[^>]+>", "", c).strip().replace("&nbsp;", "").replace(",", "") for c in cells]
    roc_y, mon = m.split("_")
    ad_y = label[roc_y]
    cur = int(cleaned[2]); prev_m = int(cleaned[3]); prev_y = int(cleaned[4])
    cum = int(cleaned[6]); cum_prev = int(cleaned[7])
    yoy = (cur - prev_y) / prev_y * 100
    mom = (cur - prev_m) / prev_m * 100
    rows.append(dict(ad_y=ad_y, mon=int(mon), roc=m, cur=cur, prev_y=prev_y, yoy=yoy, mom=mom, cum=cum, cum_prev=cum_prev))

# cross-check: cur month == cum(this) - cum(prev month in same file's cumulative col isn't available across files directly,
# but consecutive months' cum should satisfy cum[i] - cum[i-1] == cur[i] when same fiscal year)
print(f"{'年月':<10}{'當月營收':>12}{'去年同月':>12}{'YoY%':>9}{'MoM%':>9}{'累計營收':>12}")
for r in rows:
    print(f"{r['ad_y']}/{r['mon']:>2}{'':<4}{r['cur']:>12,}{r['prev_y']:>12,}{r['yoy']:>8.1f}%{r['mom']:>8.1f}%{r['cum']:>12,}")

with open("monthly_revenue_3167.csv", "w", encoding="utf-8") as f:
    f.write("year,month,revenue_ntd_k,yoy_pct,mom_pct,cum_revenue_ntd_k\n")
    for r in rows:
        f.write(f"{r['ad_y']},{r['mon']},{r['cur']},{r['yoy']:.2f},{r['mom']:.2f},{r['cum']}\n")
print("\nWrote monthly_revenue_3167.csv")
