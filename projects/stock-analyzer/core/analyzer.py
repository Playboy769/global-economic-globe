"""
analyzer.py — 依語意類型（SchemaType）生成圖表配方清單
每個配方 = dict，由 chart_panel.py 渲染。
"""
import pandas as pd
import numpy as np
from core.schema_detector import SchemaType, detect


# ── 公開入口 ──────────────────────────────────────────────────────────────────

def build_chart_recipes(df: pd.DataFrame) -> list[dict]:
    schema = detect(df)
    dispatch = {
        SchemaType.TW_MOMENTUM: _tw_momentum,
        SchemaType.TW_MARKET:   _tw_market,
        SchemaType.US_STOCK:    _us_stock,
        SchemaType.TW_SIGNAL:   _tw_signal,
        SchemaType.GENERIC:     _generic,
    }
    return dispatch[schema](df)


def describe_df(df: pd.DataFrame) -> pd.DataFrame:
    num = [c for c in df.columns if pd.api.types.is_numeric_dtype(df[c])]
    if not num:
        return pd.DataFrame()
    desc = df[num].describe().T
    desc["median"] = df[num].median()
    desc = desc[["count", "mean", "median", "std", "min", "max"]]
    desc.columns = ["筆數", "平均", "中位數", "標準差", "最小", "最大"]
    return desc.round(4)


def column_info(df: pd.DataFrame) -> pd.DataFrame:
    rows = []
    for col in df.columns:
        nn = int(df[col].notna().sum())
        rows.append({
            "欄位": col, "型別": str(df[col].dtype),
            "非空": nn, "唯一": int(df[col].nunique()),
            "範例": str(df[col].dropna().iloc[0]) if nn else "",
        })
    return pd.DataFrame(rows)


# ── TW_MOMENTUM（台股短波動能 — XQ 工作表1）────────────────────────────────────

def _tw_momentum(df: pd.DataFrame) -> list[dict]:
    name_col   = _col("商品", df)
    change_col = _col("漲幅%", df)
    vol_col    = _col("總量", df)
    bid_col    = _col("內外盤比圖", df)
    roe_col    = _col("ROE%", df)
    gm_col     = _col("毛利率%", df)
    beta_col   = _col("Beta", df)
    pe_col     = _col("PE(市盈率)", df)

    r = []

    # 1. 漲幅排行
    if name_col and change_col:
        r.append({"type": "bar_change", "title": "漲幅% 排行",
                  "x": name_col, "y": change_col})

    # 2. 內外盤 vs 漲幅（核心圖）
    if bid_col and change_col:
        r.append({"type": "scatter_labeled",
                  "title": "內外盤力道 vs 漲幅%（右上角 = 強勢買盤推升）",
                  "x": bid_col, "y": change_col,
                  "label_col": name_col, "color_col": vol_col,
                  "xline": 50, "yline": None})

    # 3. ROE × 毛利率象限圖
    if roe_col and gm_col:
        r.append({"type": "quadrant",
                  "title": "ROE% × 毛利率% 象限（氣泡=漲幅）",
                  "x": gm_col, "y": roe_col,
                  "size_col": change_col, "label_col": name_col,
                  "xlabel": "毛利率%", "ylabel": "ROE%"})

    # 4. Beta vs 漲幅
    if beta_col and change_col:
        r.append({"type": "scatter_labeled",
                  "title": "Beta vs 漲幅%（找低Beta異常大漲）",
                  "x": beta_col, "y": change_col,
                  "label_col": name_col, "color_col": None,
                  "xline": 1.0, "yline": None})

    # 5. 成交量前20排行
    if name_col and vol_col:
        r.append({"type": "bar_top", "title": "總量排行 Top 20（量能確認）",
                  "x": name_col, "y": vol_col, "top": 20})

    # 6. PE 分布
    if pe_col:
        r.append({"type": "hist", "title": "PE 估值分布", "col": pe_col})

    return r


# ── TW_MARKET（台股大盤龍頭 — XQ 工作表2）────────────────────────────────────────

def _tw_market(df: pd.DataFrame) -> list[dict]:
    name_col    = _col("商品", df)
    change_col  = _col("漲幅%", df)
    mktcap_col  = _col("市值", df)
    vol_col     = _col("總量", df)
    industry_col = _col("產業", df)
    sub_col     = _col("細產業", df)
    eps_col     = _col("盈餘", df)
    bid_col     = _col("內外盤比圖", df)

    r = []

    # 1. 市值 × 漲幅 × 成交量氣泡圖
    if mktcap_col and change_col and vol_col:
        r.append({"type": "bubble",
                  "title": "市值 × 漲幅% × 成交量（氣泡大=量大）",
                  "x": change_col, "y": mktcap_col,
                  "size_col": vol_col, "label_col": name_col,
                  "xlabel": "漲幅%", "ylabel": "市值（億）"})

    # 2. 產業分布圓餅
    if industry_col:
        r.append({"type": "pie", "title": "產業分布",
                  "col": industry_col})

    # 3. 各產業漲幅盒鬚圖
    if industry_col and change_col:
        r.append({"type": "boxplot_group",
                  "title": "各產業漲幅% 分布（找今日強勢板塊）",
                  "group_col": industry_col, "val_col": change_col})

    # 4. 細產業主題頻次橫條（核心圖）
    if sub_col:
        r.append({"type": "theme_bar",
                  "title": "細產業主題頻次 Top 25（今日主流題材）",
                  "col": sub_col, "top": 25})

    # 5. 漲幅排行（含產業）
    if name_col and change_col:
        r.append({"type": "bar_change",
                  "title": "漲幅% 排行",
                  "x": name_col, "y": change_col})

    # 6. 市值 vs 盈餘
    if mktcap_col and eps_col:
        r.append({"type": "scatter_labeled",
                  "title": "市值 vs 盈餘（估值對照）",
                  "x": eps_col, "y": mktcap_col,
                  "label_col": name_col, "color_col": change_col,
                  "xline": None, "yline": None,
                  "xlabel": "盈餘（元）", "ylabel": "市值（億）"})

    return r


# ── US_STOCK（美股多期間動能）────────────────────────────────────────────────────

def _us_stock(df: pd.DataFrame) -> list[dict]:
    name_col    = _col("name", df) or _col("symbol", df)
    sym_col     = _col("symbol", df)
    chg_col     = _col("% chg", df)
    beta_col    = _col("beta", df)
    industry_col = _col("industry", df)
    ema30_col   = _col("ema-ema30", df)
    price_col   = _col("price", df)
    period_cols = [c for c in df.columns
                   if any(t in c.lower() for t in ["5d chg","10d chg","20d chg",
                                                    "60d chg","120d chg","250d chg","ytd chg"])]

    r = []

    # 1. 多期間熱力圖（核心圖）
    if period_cols and sym_col:
        r.append({"type": "heatmap_pct",
                  "title": "多期間動能熱力圖（值 = 百分比）",
                  "cols": period_cols,
                  "name_col": sym_col})

    # 2. 動能分數排行（加權合計）
    if period_cols and (sym_col or name_col):
        r.append({"type": "momentum_bar",
                  "title": "綜合動能分數排行（近期權重高）",
                  "period_cols": period_cols,
                  "name_col": sym_col or name_col})

    # 3. 今日漲幅排行
    if (sym_col or name_col) and chg_col:
        r.append({"type": "bar_change",
                  "title": "今日漲幅% 排行",
                  "x": sym_col or name_col, "y": chg_col})

    # 4. Beta vs 今日%Chg
    if beta_col and chg_col:
        r.append({"type": "scatter_labeled",
                  "title": "Beta vs 今日漲幅%（低Beta大漲=突破訊號）",
                  "x": beta_col, "y": chg_col,
                  "label_col": sym_col, "color_col": None,
                  "xline": 1.0, "yline": 0})

    # 5. EMA30 乖離率橫條
    if ema30_col and price_col and sym_col:
        r.append({"type": "ema_deviation",
                  "title": "EMA-30 乖離率（正=超買，負=超賣）",
                  "price_col": price_col, "ema_col": ema30_col,
                  "name_col": sym_col})

    # 6. Industry 漲幅盒鬚
    if industry_col and chg_col:
        r.append({"type": "boxplot_group",
                  "title": "各產業今日漲幅分布",
                  "group_col": industry_col, "val_col": chg_col})

    return r


# ── TW_SIGNAL（台股均線策略 CSV）─────────────────────────────────────────────────

def _tw_signal(df: pd.DataFrame) -> list[dict]:
    name_col   = _col("商品", df)
    change_col = _col("漲跌幅", df) or _col("漲幅", df)
    peg_col    = _col("peg指標", df) or _col("peg", df)
    ind_col    = _col("產業", df)
    ma_cols    = [c for c in df.columns if "均線" in str(c)]
    price_col  = _col("收盤價", df) or _col("成交", df)

    r = []
    if name_col and change_col:
        r.append({"type": "bar_change", "title": "漲跌幅% 排行",
                  "x": name_col, "y": change_col})
    if ma_cols:
        r.append({"type": "boxplot_ma", "title": "均線分布（盒鬚圖）",
                  "cols": ma_cols[:6]})
    if ind_col:
        r.append({"type": "pie", "title": "主力產業分布", "col": ind_col})
    if peg_col and name_col:
        r.append({"type": "bar_top", "title": "PEG 指標排行",
                  "x": name_col, "y": peg_col, "top": 30})
    if price_col and change_col:
        r.append({"type": "scatter_labeled",
                  "title": "股價 vs 漲跌幅%",
                  "x": price_col, "y": change_col,
                  "label_col": name_col, "color_col": None,
                  "xline": None, "yline": 0})
    return r


# ── GENERIC ───────────────────────────────────────────────────────────────────

def _generic(df: pd.DataFrame) -> list[dict]:
    r = []
    num_cols = [c for c in df.columns if pd.api.types.is_numeric_dtype(df[c])]
    cat_cols = [c for c in df.columns
                if not pd.api.types.is_numeric_dtype(df[c]) and 2 <= df[c].nunique() <= 30]
    for col in num_cols[:8]:
        r.append({"type": "hist", "title": f"{col} 分布", "col": col})
    for col in cat_cols[:4]:
        r.append({"type": "bar_count", "title": f"{col} 類別分布", "col": col})
    return r


# ── 工具函式 ──────────────────────────────────────────────────────────────────

def _col(keyword: str, df: pd.DataFrame):
    """找第一個欄位名稱含 keyword（不分大小寫）的欄位，找不到回傳 None。"""
    kw = keyword.lower()
    for c in df.columns:
        if kw in str(c).lower():
            return c
    return None
