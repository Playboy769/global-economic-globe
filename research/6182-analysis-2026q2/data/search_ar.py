h = open('ar2025.txt', encoding='utf-8').read()
keywords = ['主要客戶','主要供應商','客戶A','客戶B','供應商甲','產銷量值','生產設備','產能','轉投資',
            '重摻','輕摻','磊晶','8吋','12吋','6吋','市場及產銷','應用','產品組合','營業比重']
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
open('ar_keyword_scan.txt','w',encoding='utf-8').write('\n'.join(out))
print('done')
