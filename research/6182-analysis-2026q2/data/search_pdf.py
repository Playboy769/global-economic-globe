import sys, io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
fname = sys.argv[1]
h = open(fname, encoding='utf-8').read()
keywords = ['部門資訊','營運部門','地區','子公司','承諾及或有事項','合約負債','未開始','採購承諾',
            '客戶集中度','主要客戶','供應商','租賃','減損','期後事項','保證','背書','關係人交易','股權']
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
    out.append(f'{kw}: count={len(idxs)} first_positions={idxs[:5]}')
open('pdf_keyword_scan.txt','w',encoding='utf-8').write('\n'.join(out))
print('done')
