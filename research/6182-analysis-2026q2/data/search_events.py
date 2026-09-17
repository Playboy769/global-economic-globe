h = open('q2_pdf.txt', encoding='utf-8').read()
keywords = ['龍潭','擴產','資本支出','投資計畫','12吋','現金增資','私募','發行新股','特別股','可轉換公司債','公司債']
out = []
for kw in keywords:
    idxs = []
    start = 0
    while True:
        idx = h.find(kw, start)
        if idx == -1:
            break
        idxs.append(idx)
        start = idx + 1
        if len(idxs) > 10:
            break
    out.append(f'{kw}: count={len(idxs)} first_positions={idxs[:6]}')
open('event_keyword_scan.txt','w',encoding='utf-8').write('\n'.join(out))
print('done')
