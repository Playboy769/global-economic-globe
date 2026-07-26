"""
chart_builder.py — 將 analyzer.py 產生的 recipe dict 轉成 Plotly JSON。
每個 recipe 回傳 { title, json, error }，json 為 fig.to_json() 字串。
"""
import numpy as np
import pandas as pd
import plotly.graph_objects as go

# ── 顏色常數 ──────────────────────────────────────────────────────────────────
POS  = "#2ecc71"
NEG  = "#e74c3c"
BLUE = "#4a7fff"
GRAY = "#555555"

# ── 深色主題基底 ──────────────────────────────────────────────────────────────
_FONT = dict(color="#ffffff", family="Calibri, 'Yu Gothic', sans-serif")

def _dark_layout(**kwargs) -> dict:
    base = dict(
        paper_bgcolor="#0a0a0a",
        plot_bgcolor="#0f0f0f",
        font=_FONT,
    )
    base.update(kwargs)
    return base

def _axis(**kwargs) -> dict:
    base = dict(
        gridcolor="#1a1a1a",
        tickcolor="#ffffff",
        tickfont=_FONT,
        automargin=True,
    )
    base.update(kwargs)
    return base


# ── 公開入口 ──────────────────────────────────────────────────────────────────

def build_plotly_charts(df: pd.DataFrame, recipes: list[dict]) -> list[dict]:
    results = []
    for recipe in recipes:
        title = recipe.get("title", "")
        try:
            fig = _dispatch(df, recipe)
            results.append({"title": title, "json": fig.to_json(), "error": None})
        except Exception as e:
            results.append({"title": title, "json": None, "error": str(e)})
    return results


# ── 分派 ─────────────────────────────────────────────────────────────────────

def _dispatch(df: pd.DataFrame, r: dict) -> go.Figure:
    t = r["type"]
    if t == "bar_change":      return _bar_change(df, r)
    if t == "hist":            return _hist(df, r)
    if t == "pie":             return _pie(df, r)
    if t == "bar_count":       return _bar_count(df, r)
    if t in ("bar_top", "bar_peg"):  return _bar_top(df, r)
    if t == "boxplot_ma":      return _boxplot_ma(df, r)
    if t == "boxplot_group":   return _boxplot_group(df, r)
    if t == "heatmap_pct":     return _heatmap_pct(df, r)
    if t == "scatter_labeled": return _scatter_labeled(df, r)
    if t == "quadrant":        return _quadrant(df, r)
    if t == "bubble":          return _bubble(df, r)
    if t == "theme_bar":       return _theme_bar(df, r)
    if t == "momentum_bar":    return _momentum_bar(df, r)
    if t == "ema_deviation":   return _ema_deviation(df, r)
    raise ValueError(f"未知的 recipe type: {t}")


# ── 各圖表實作 ────────────────────────────────────────────────────────────────

def _bar_change(df: pd.DataFrame, r: dict) -> go.Figure:
    x_col, y_col = r["x"], r["y"]
    sub = df[[x_col, y_col]].dropna()
    sub = sub.sort_values(y_col)
    n = len(sub)
    if n > 40:
        sub = pd.concat([sub.head(20), sub.tail(20)]).drop_duplicates()
    vals = sub[y_col].tolist()
    names = sub[x_col].astype(str).tolist()
    colors = [POS if v >= 0 else NEG for v in vals]
    fig = go.Figure(go.Bar(
        x=vals, y=names, orientation="h",
        marker_color=colors,
        text=[f"{v:+.2f}" for v in vals],
        textposition="outside",
    ))
    h = max(300, len(names) * 26 + 80)
    fig.update_layout(
        **_dark_layout(height=h, margin=dict(l=10, r=60, t=40, b=20)),
        xaxis=_axis(title=y_col),
        yaxis=_axis(),
    )
    return fig


def _hist(df: pd.DataFrame, r: dict) -> go.Figure:
    col = r["col"]
    vals = pd.to_numeric(df[col], errors="coerce").dropna()
    fig = go.Figure(go.Histogram(x=vals, marker_color=BLUE, opacity=0.85))
    fig.update_layout(
        **_dark_layout(height=320, margin=dict(l=10, r=10, t=40, b=20)),
        xaxis=_axis(title=col),
        yaxis=_axis(title="頻次"),
        bargap=0.05,
    )
    return fig


def _pie(df: pd.DataFrame, r: dict) -> go.Figure:
    col = r["col"]
    counts = df[col].dropna().astype(str).value_counts()
    if len(counts) > 7:
        top = counts.head(7)
        other = pd.Series({"其他": counts.iloc[7:].sum()})
        counts = pd.concat([top, other])
    fig = go.Figure(go.Pie(
        labels=counts.index.tolist(),
        values=counts.values.tolist(),
        hole=0.35,
        textfont=_FONT,
    ))
    fig.update_layout(
        **_dark_layout(height=380, margin=dict(l=10, r=10, t=40, b=20)),
    )
    return fig


def _bar_count(df: pd.DataFrame, r: dict) -> go.Figure:
    col = r["col"]
    counts = df[col].dropna().astype(str).value_counts().sort_values()
    n = len(counts)
    fig = go.Figure(go.Bar(
        x=counts.values.tolist(),
        y=counts.index.tolist(),
        orientation="h",
        marker_color=BLUE,
    ))
    h = max(300, n * 26 + 80)
    fig.update_layout(
        **_dark_layout(height=h, margin=dict(l=10, r=40, t=40, b=20)),
        xaxis=_axis(title="數量"),
        yaxis=_axis(),
    )
    return fig


def _bar_top(df: pd.DataFrame, r: dict) -> go.Figure:
    x_col, y_col = r["x"], r["y"]
    top = r.get("top", 20)
    sub = df[[x_col, y_col]].dropna()
    sub = pd.to_numeric(sub[y_col], errors="coerce").pipe(
        lambda s: sub.assign(**{y_col: s})
    ).dropna(subset=[y_col])
    sub = sub.nlargest(top, y_col).sort_values(y_col)
    vals = sub[y_col].tolist()
    names = sub[x_col].astype(str).tolist()
    n = len(names)
    fig = go.Figure(go.Bar(
        x=vals, y=names, orientation="h",
        marker_color=BLUE,
        text=[f"{v:.2f}" for v in vals],
        textposition="outside",
    ))
    h = max(300, n * 26 + 80)
    fig.update_layout(
        **_dark_layout(height=h, margin=dict(l=10, r=60, t=40, b=20)),
        xaxis=_axis(title=y_col),
        yaxis=_axis(),
    )
    return fig


def _boxplot_ma(df: pd.DataFrame, r: dict) -> go.Figure:
    cols = r["cols"]
    fig = go.Figure()
    for col in cols:
        vals = pd.to_numeric(df[col], errors="coerce").dropna()
        fig.add_trace(go.Box(y=vals.tolist(), name=str(col), marker_color=BLUE))
    fig.update_layout(
        **_dark_layout(height=340, margin=dict(l=10, r=10, t=40, b=20)),
        xaxis=_axis(),
        yaxis=_axis(),
        showlegend=False,
    )
    return fig


def _boxplot_group(df: pd.DataFrame, r: dict) -> go.Figure:
    grp_col, val_col = r["group_col"], r["val_col"]
    sub = df[[grp_col, val_col]].copy()
    sub[val_col] = pd.to_numeric(sub[val_col], errors="coerce")
    sub = sub.dropna()
    # top 12 groups by count
    top_groups = sub[grp_col].value_counts().head(12).index.tolist()
    sub = sub[sub[grp_col].isin(top_groups)]
    fig = go.Figure()
    for grp in top_groups:
        vals = sub[sub[grp_col] == grp][val_col].tolist()
        fig.add_trace(go.Box(y=vals, name=str(grp), marker_color=BLUE))
    fig.update_layout(
        **_dark_layout(height=340, margin=dict(l=10, r=10, t=40, b=40)),
        xaxis=_axis(tickangle=-30),
        yaxis=_axis(),
        showlegend=False,
    )
    return fig


def _heatmap_pct(df: pd.DataFrame, r: dict) -> go.Figure:
    cols = r["cols"]
    name_col = r.get("name_col")
    sub = df.copy()
    for c in cols:
        sub[c] = pd.to_numeric(sub[c], errors="coerce")
    sub = sub.dropna(subset=cols)
    if name_col and name_col in sub.columns:
        labels = sub[name_col].astype(str).tolist()
    else:
        labels = [str(i) for i in sub.index]

    mat = sub[cols].values.astype(float)
    # multiply by 100 if values look like decimals (abs mean < 2)
    if np.nanmean(np.abs(mat)) < 2:
        mat = mat * 100

    n_rows = len(labels)
    h = max(300, n_rows * 20 + 100)
    fig = go.Figure(go.Heatmap(
        z=mat,
        x=[str(c) for c in cols],
        y=labels,
        colorscale="RdYlGn",
        zmid=0,
        text=np.round(mat, 1),
        texttemplate="%{text}%",
        textfont=dict(size=10, color="#ffffff"),
    ))
    fig.update_layout(
        **_dark_layout(height=h, margin=dict(l=10, r=10, t=40, b=60)),
        xaxis=_axis(tickangle=-30),
        yaxis=_axis(),
    )
    return fig


def _scatter_labeled(df: pd.DataFrame, r: dict) -> go.Figure:
    x_col, y_col = r["x"], r["y"]
    label_col = r.get("label_col")
    color_col = r.get("color_col")
    xline = r.get("xline")
    yline = r.get("yline")
    xlabel = r.get("xlabel", x_col)
    ylabel = r.get("ylabel", y_col)

    sub = df[[c for c in [x_col, y_col, label_col, color_col] if c]].copy()
    for c in [x_col, y_col]:
        sub[c] = pd.to_numeric(sub[c], errors="coerce")
    sub = sub.dropna(subset=[x_col, y_col])

    marker_opts = dict(size=8, color=BLUE, opacity=0.8)
    if color_col and color_col in sub.columns:
        c_vals = pd.to_numeric(sub[color_col], errors="coerce")
        marker_opts = dict(
            size=8,
            color=c_vals.tolist(),
            colorscale="RdYlGn",
            showscale=True,
            opacity=0.85,
        )

    text_labels = None
    if label_col and label_col in sub.columns:
        text_labels = sub[label_col].astype(str).tolist()

    fig = go.Figure(go.Scatter(
        x=sub[x_col].tolist(),
        y=sub[y_col].tolist(),
        mode="markers+text" if text_labels else "markers",
        text=text_labels,
        textposition="top center",
        textfont=dict(size=9, color="#cccccc"),
        marker=marker_opts,
    ))

    shapes = []
    if xline is not None:
        shapes.append(dict(type="line", x0=xline, x1=xline, y0=0, y1=1,
                           yref="paper", line=dict(color=GRAY, dash="dash", width=1)))
    if yline is not None:
        shapes.append(dict(type="line", x0=0, x1=1, xref="paper",
                           y0=yline, y1=yline, line=dict(color=GRAY, dash="dash", width=1)))

    fig.update_layout(
        **_dark_layout(height=460, margin=dict(l=10, r=10, t=40, b=20)),
        xaxis=_axis(title=xlabel),
        yaxis=_axis(title=ylabel),
        shapes=shapes,
    )
    return fig


def _quadrant(df: pd.DataFrame, r: dict) -> go.Figure:
    x_col, y_col = r["x"], r["y"]
    size_col = r.get("size_col")
    label_col = r.get("label_col")
    xlabel = r.get("xlabel", x_col)
    ylabel = r.get("ylabel", y_col)

    cols_needed = [c for c in [x_col, y_col, size_col, label_col] if c]
    sub = df[cols_needed].copy()
    for c in [x_col, y_col]:
        sub[c] = pd.to_numeric(sub[c], errors="coerce")
    if size_col:
        sub[size_col] = pd.to_numeric(sub[size_col], errors="coerce")
    sub = sub.dropna(subset=[x_col, y_col])

    xmed = sub[x_col].median()
    ymed = sub[y_col].median()

    sizes = None
    if size_col and size_col in sub.columns:
        raw = sub[size_col].fillna(0)
        mn, mx = raw.min(), raw.max()
        if mx > mn:
            sizes = (((raw - mn) / (mx - mn)) * 28 + 6).tolist()
        else:
            sizes = [12] * len(sub)

    text_labels = None
    if label_col and label_col in sub.columns:
        text_labels = sub[label_col].astype(str).tolist()

    fig = go.Figure(go.Scatter(
        x=sub[x_col].tolist(),
        y=sub[y_col].tolist(),
        mode="markers+text" if text_labels else "markers",
        text=text_labels,
        textposition="top center",
        textfont=dict(size=9, color="#cccccc"),
        marker=dict(
            size=sizes if sizes else 10,
            color=BLUE, opacity=0.75,
            line=dict(width=0.5, color="#ffffff"),
        ),
    ))

    shapes = [
        dict(type="line", x0=xmed, x1=xmed, y0=0, y1=1, yref="paper",
             line=dict(color=GRAY, dash="dash", width=1)),
        dict(type="line", x0=0, x1=1, xref="paper", y0=ymed, y1=ymed,
             line=dict(color=GRAY, dash="dash", width=1)),
    ]

    annotations = [
        dict(x=sub[x_col].max(), y=sub[y_col].max(), text="強ROE+高毛利",
             showarrow=False, font=dict(color=POS, size=10), xanchor="right"),
        dict(x=sub[x_col].min(), y=sub[y_col].max(), text="高ROE+低毛利",
             showarrow=False, font=dict(color="#f39c12", size=10), xanchor="left"),
        dict(x=sub[x_col].max(), y=sub[y_col].min(), text="低ROE+高毛利",
             showarrow=False, font=dict(color="#3498db", size=10), xanchor="right"),
        dict(x=sub[x_col].min(), y=sub[y_col].min(), text="弱勢區",
             showarrow=False, font=dict(color=NEG, size=10), xanchor="left"),
    ]

    fig.update_layout(
        **_dark_layout(height=500, margin=dict(l=10, r=10, t=40, b=20)),
        xaxis=_axis(title=xlabel),
        yaxis=_axis(title=ylabel),
        shapes=shapes,
        annotations=annotations,
    )
    return fig


def _bubble(df: pd.DataFrame, r: dict) -> go.Figure:
    x_col, y_col = r["x"], r["y"]
    size_col = r.get("size_col")
    label_col = r.get("label_col")
    xlabel = r.get("xlabel", x_col)
    ylabel = r.get("ylabel", y_col)

    cols_needed = [c for c in [x_col, y_col, size_col, label_col] if c]
    sub = df[cols_needed].copy()
    for c in [x_col, y_col]:
        sub[c] = pd.to_numeric(sub[c], errors="coerce")
    if size_col:
        sub[size_col] = pd.to_numeric(sub[size_col], errors="coerce")
    sub = sub.dropna(subset=[x_col, y_col])

    sizes = [10] * len(sub)
    if size_col and size_col in sub.columns:
        raw = sub[size_col].fillna(0).clip(lower=0)
        mn, mx = raw.min(), raw.max()
        if mx > mn:
            sizes = (((raw - mn) / (mx - mn)) * 36 + 6).tolist()

    text_labels = None
    if label_col and label_col in sub.columns:
        text_labels = sub[label_col].astype(str).tolist()

    y_vals = sub[y_col].tolist()
    use_log = False
    if len(y_vals) > 0:
        pos_vals = [v for v in y_vals if v and v > 0]
        if pos_vals and (max(pos_vals) / (min(pos_vals) + 1e-9)) > 100:
            use_log = True

    fig = go.Figure(go.Scatter(
        x=sub[x_col].tolist(),
        y=y_vals,
        mode="markers+text" if text_labels else "markers",
        text=text_labels,
        textposition="top center",
        textfont=dict(size=9, color="#cccccc"),
        marker=dict(size=sizes, color=BLUE, opacity=0.7,
                    line=dict(width=0.5, color="#ffffff")),
    ))

    fig.update_layout(
        **_dark_layout(height=460, margin=dict(l=10, r=10, t=40, b=20)),
        xaxis=_axis(title=xlabel),
        yaxis=_axis(title=ylabel, type="log" if use_log else "linear"),
    )
    return fig


def _theme_bar(df: pd.DataFrame, r: dict) -> go.Figure:
    col = r["col"]
    top = r.get("top", 25)
    # split by comma, explode
    series = df[col].dropna().astype(str)
    tags = series.str.split(",").explode().str.strip()
    tags = tags[tags != ""]
    counts = tags.value_counts().head(top).sort_values()
    n = len(counts)
    fig = go.Figure(go.Bar(
        x=counts.values.tolist(),
        y=counts.index.tolist(),
        orientation="h",
        marker_color=BLUE,
        text=counts.values.tolist(),
        textposition="outside",
    ))
    h = max(300, n * 26 + 80)
    fig.update_layout(
        **_dark_layout(height=h, margin=dict(l=10, r=50, t=40, b=20)),
        xaxis=_axis(title="頻次"),
        yaxis=_axis(),
    )
    return fig


def _momentum_bar(df: pd.DataFrame, r: dict) -> go.Figure:
    period_cols = r["period_cols"]
    name_col = r["name_col"]

    sub = df.copy()
    for c in period_cols:
        sub[c] = pd.to_numeric(sub[c], errors="coerce")

    # weights: more recent = higher weight
    # order: assume sorted from shorter to longer period
    weights = list(range(len(period_cols), 0, -1))
    sub["_score"] = sum(
        sub[c].fillna(0) * w for c, w in zip(period_cols, weights)
    )
    sub = sub.dropna(subset=[name_col])
    sub = sub.sort_values("_score")

    vals = sub["_score"].tolist()
    names = sub[name_col].astype(str).tolist()
    colors = [POS if v >= 0 else NEG for v in vals]
    n = len(names)
    fig = go.Figure(go.Bar(
        x=vals, y=names, orientation="h",
        marker_color=colors,
        text=[f"{v:+.1f}" for v in vals],
        textposition="outside",
    ))
    h = max(300, n * 26 + 80)
    fig.update_layout(
        **_dark_layout(height=h, margin=dict(l=10, r=70, t=40, b=20)),
        xaxis=_axis(title="加權動能分數"),
        yaxis=_axis(),
    )
    return fig


def _ema_deviation(df: pd.DataFrame, r: dict) -> go.Figure:
    price_col = r["price_col"]
    ema_col   = r["ema_col"]
    name_col  = r["name_col"]

    sub = df[[name_col, price_col, ema_col]].copy()
    sub[price_col] = pd.to_numeric(sub[price_col], errors="coerce")
    sub[ema_col]   = pd.to_numeric(sub[ema_col], errors="coerce")
    sub = sub.dropna()
    sub["_dev"] = (sub[price_col] - sub[ema_col]) / sub[ema_col].replace(0, np.nan) * 100
    sub = sub.dropna(subset=["_dev"]).sort_values("_dev")

    vals = sub["_dev"].tolist()
    names = sub[name_col].astype(str).tolist()
    colors = [POS if v >= 0 else NEG for v in vals]
    n = len(names)
    fig = go.Figure(go.Bar(
        x=vals, y=names, orientation="h",
        marker_color=colors,
        text=[f"{v:+.1f}%" for v in vals],
        textposition="outside",
    ))
    h = max(300, n * 26 + 80)
    fig.update_layout(
        **_dark_layout(height=h, margin=dict(l=10, r=70, t=40, b=20)),
        xaxis=_axis(title="EMA-30 乖離率 (%)"),
        yaxis=_axis(),
    )
    return fig
