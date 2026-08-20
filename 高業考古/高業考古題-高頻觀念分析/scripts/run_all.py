"""Re-run the whole repeat-pattern pipeline.

    python run_all.py

Drop the new 屆 PDFs into 高業考古/<年>/ first (any filename — the year/session is read
from the paper header, not the filename), then re-run. Add the new 屆 to the EXAMS
list in s2_cluster.py, otherwise it is parsed but never enters the timeline.

Outputs land in scripts/out/ :
    corpus.json          every question, normalized
    analysis.json        clusters, gap distribution, exam-pair matrix
    analysis3.json       hazard tables, backtest, forecast
    *_report.txt         the human-readable versions
"""
import subprocess, sys, os

# the Windows console is cp950 here; report files are UTF-8 regardless
try:
    sys.stdout.reconfigure(encoding='utf-8', errors='replace')
    sys.stderr.reconfigure(encoding='utf-8', errors='replace')
except Exception:
    pass

HERE = os.path.dirname(os.path.abspath(__file__))
STEPS = [
    ('s1_extract.py',   'PDF -> 逐題語料（切題、正規化、對答案卡）'),
    ('s2_cluster.py',   '模糊分群 90% + 間隔分布 vs 隨機虛無模型'),
    ('s3_pattern.py',   '同年/跨年對照、回收來源距離、屆次配對'),
    ('s4_calibrate.py', '偏誤校正、回測、下一屆預測'),
    ('s5_options.py',   '選項改寫比例、預測偏誤校正'),
]

for i, (script, desc) in enumerate(STEPS, 1):
    print(f'[{i}/{len(STEPS)}] {script}  — {desc}', flush=True)
    r = subprocess.run([sys.executable, script], cwd=HERE)
    if r.returncode != 0:
        print(f'  FAILED at {script}', file=sys.stderr)
        sys.exit(r.returncode)
print('\nDone. See scripts/out/*_report.txt')
