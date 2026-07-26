import os, uuid, tempfile
from pathlib import Path
from fastapi import FastAPI, UploadFile, File, HTTPException, Request
from fastapi.responses import FileResponse
import pandas as pd
from core.file_loader import load_file
from core.analyzer import build_chart_recipes
from core.chart_builder import build_plotly_charts

app = FastAPI(title="台美股分析工具")
_store: dict = {}
_HTML = Path(__file__).parent / "templates" / "index.html"


@app.get("/")
async def root():
    return FileResponse(_HTML, media_type="text/html")


@app.post("/api/upload")
async def upload(file: UploadFile = File(...)):
    content = await file.read()
    if len(content) > 30 * 1024 * 1024:
        raise HTTPException(413, "檔案過大（上限 30 MB）")
    fname = file.filename or "upload"
    suffix = Path(fname).suffix.lower()
    if suffix not in (".csv", ".xlsx", ".xlsm"):
        raise HTTPException(400, f"不支援的格式：{suffix}")
    with tempfile.NamedTemporaryFile(delete=False, suffix=suffix) as tmp:
        tmp.write(content)
        tmp_path = tmp.name
    try:
        data = load_file(tmp_path)
        data["meta"]["filename"] = fname
    except Exception as e:
        raise HTTPException(400, f"讀取失敗：{e}")
    finally:
        try:
            os.unlink(tmp_path)
        except Exception:
            pass
    sid = str(uuid.uuid4())
    _store[sid] = data
    sheet_names = list(data["sheets"].keys())
    first = sheet_names[0]
    df = data["sheets"][first]
    return {
        "session_id": sid,
        "meta": data["meta"],
        "sheet_names": sheet_names,
        "current_sheet": first,
        "table": _df_to_table(df),
        "charts": _get_charts(df),
    }


@app.post("/api/sheet")
async def switch_sheet(req: Request):
    body = await req.json()
    sid, sheet_name = body.get("session_id"), body.get("sheet_name")
    data = _store.get(sid)
    if not data:
        raise HTTPException(404, "Session 已過期，請重新上傳")
    df = data["sheets"].get(sheet_name)
    if df is None:
        raise HTTPException(404, f"找不到工作表：{sheet_name}")
    return {"table": _df_to_table(df), "charts": _get_charts(df)}


def _df_to_table(df: pd.DataFrame) -> dict:
    headers = list(df.columns.astype(str))
    rows = []
    for _, row in df.iterrows():
        r = []
        for v in row:
            if pd.isna(v) if not isinstance(v, str) else False:
                r.append("")
            elif isinstance(v, float):
                r.append(round(float(v), 6))
            elif isinstance(v, int):
                r.append(int(v))
            else:
                r.append(str(v))
        rows.append(r)
    return {"headers": headers, "rows": rows}


def _get_charts(df: pd.DataFrame) -> list:
    try:
        recipes = build_chart_recipes(df)
        return build_plotly_charts(df, recipes)
    except Exception as e:
        return [{"title": "錯誤", "json": None, "error": str(e)}]
