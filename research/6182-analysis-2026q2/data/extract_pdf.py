import sys
fname = sys.argv[1]
outname = sys.argv[2]
spans = eval(sys.argv[3])  # list of (start,end) offsets
h = open(fname, encoding='utf-8').read()
out = []
for (s,e) in spans:
    out.append(f'--- [{s}:{e}] ---')
    out.append(h[s:e])
open(outname,'w',encoding='utf-8').write('\n'.join(out))
print('done', len(out))
