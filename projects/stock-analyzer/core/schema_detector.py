"""
schema_detector.py — 根據欄位名稱組合偵測資料語意類型
"""
from enum import Enum
import pandas as pd


class SchemaType(Enum):
    TW_MOMENTUM = "tw_momentum"  # 台股短波動能（XQ 工作表1）
    TW_MARKET   = "tw_market"   # 台股大盤龍頭（XQ 工作表2）
    TW_SIGNAL   = "tw_signal"   # 台股均線策略 CSV
    US_STOCK    = "us_stock"    # 美股多期間動能 xlsm
    GENERIC     = "generic"     # 其他通用


# 各類型的辨識特徵欄位
_SIGNATURES: list[tuple[SchemaType, list[str]]] = [
    # 順序很重要：較具體的放前面
    (SchemaType.TW_MARKET,   ["內外盤比圖", "市值"]),
    (SchemaType.TW_MOMENTUM, ["內外盤比圖", "漲幅%"]),
    (SchemaType.US_STOCK,    ["symbol", "5d chg"]),       # 欄位 lower 比對
    (SchemaType.US_STOCK,    ["% chg", "ema-ema30"]),
    (SchemaType.TW_SIGNAL,   ["短均線", "peg指標"]),
    (SchemaType.TW_SIGNAL,   ["乖離", "peg"]),
]


def detect(df: pd.DataFrame) -> SchemaType:
    """回傳 df 最符合的語意類型。"""
    col_set = {str(c).lower() for c in df.columns}
    col_str = " ".join(col_set)

    for schema_type, keywords in _SIGNATURES:
        if all(_kw_in(kw, col_set, col_str) for kw in keywords):
            return schema_type

    return SchemaType.GENERIC


def _kw_in(keyword: str, col_set: set, col_str: str) -> bool:
    kw = keyword.lower()
    return kw in col_set or kw in col_str


def label(schema: SchemaType) -> str:
    return {
        SchemaType.TW_MOMENTUM: "台股短波動能",
        SchemaType.TW_MARKET:   "台股大盤龍頭",
        SchemaType.TW_SIGNAL:   "台股均線策略",
        SchemaType.US_STOCK:    "美股多期間動能",
        SchemaType.GENERIC:     "通用",
    }[schema]
