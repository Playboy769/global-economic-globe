# -*- coding: utf-8 -*-
"""Inject food_data.csv into the HTML template to produce the final standalone app."""
import json
import math
import pandas as pd
from pathlib import Path

ROOT = Path(__file__).resolve().parent
SRC_CSV       = ROOT / "food_data.csv"
TEMPLATE_PATH = ROOT / "template.html"
OUT_HTML      = ROOT / "食物熱量查詢.html"

df = pd.read_csv(SRC_CSV)

def clean(v):
    if v is None:
        return None
    if isinstance(v, float) and math.isnan(v):
        return None
    if isinstance(v, str):
        s = v.strip()
        return s if s else None
    return v

records = []
for _, row in df.iterrows():
    rec = {k: clean(row[k]) for k in df.columns}
    if rec.get("id") and rec.get("name"):
        records.append(rec)

print(f"Records: {len(records)}")

data_json = json.dumps(records, ensure_ascii=False, separators=(",", ":"))
html = TEMPLATE_PATH.read_text(encoding="utf-8")
html = html.replace("__DATA__", data_json)
OUT_HTML.write_text(html, encoding="utf-8")
print(f"HTML: {OUT_HTML}  ({OUT_HTML.stat().st_size // 1024} KB)")
