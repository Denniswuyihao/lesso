from __future__ import annotations

from dataclasses import dataclass
import pandas as pd


@dataclass
class PriceRule:
    mode: str = "原价"          # 原价 / 加点 / 打折
    percent: float = 0.0        # 加点 10 = +10%; 打折 10 = -10%


def apply_price_rule(price: float, rule: PriceRule) -> float:
    price = float(price or 0)
    pct = float(rule.percent or 0) / 100
    if rule.mode == "加点":
        return round(price * (1 + pct), 4)
    if rule.mode == "打折":
        return round(price * (1 - pct), 4)
    return round(price, 4)


def build_quote_items(selected_df: pd.DataFrame, rule: PriceRule) -> pd.DataFrame:
    if selected_df.empty:
        return pd.DataFrame()

    df = selected_df.copy()
    if "quantity" not in df.columns:
        df["quantity"] = 1
    df["quantity"] = pd.to_numeric(df["quantity"], errors="coerce").fillna(1).astype(int)
    df.loc[df["quantity"] <= 0, "quantity"] = 1

    df["base_price"] = pd.to_numeric(df["price"], errors="coerce").fillna(0)
    df["quote_price"] = df["base_price"].apply(lambda x: apply_price_rule(x, rule))
    df["amount"] = (df["quote_price"] * df["quantity"]).round(2)
    df["total_volume"] = (pd.to_numeric(df["packing_volume"], errors="coerce").fillna(0) * df["quantity"]).round(4)
    return df


def quote_summary(items: pd.DataFrame) -> dict:
    if items.empty:
        return {"sku_count": 0, "total_qty": 0, "total_amount": 0, "total_volume": 0}
    return {
        "sku_count": int(items["sap"].nunique()),
        "total_qty": int(items["quantity"].sum()),
        "total_amount": float(items["amount"].sum().round(2)),
        "total_volume": float(items["total_volume"].sum().round(4)),
    }
