"""Assertions that the cleaned extract and the KPI definitions agree."""

from __future__ import annotations

import math

import pandas as pd

from pipeline.catalog import CATEGORIES, CHANNELS, REGIONS, STATUSES
from pipeline.metrics import in_scope, kpis


def run_checks(raw: pd.DataFrame, clean: pd.DataFrame, rejected: pd.DataFrame, log: dict, params: dict) -> list[dict]:
    checks = []

    def add(name: str, passed: bool, detail: str):
        checks.append({"name": name, "passed": bool(passed), "detail": detail})

    add("row balance", log["balance_ok"] and len(raw) == len(clean) + len(rejected), f"raw {len(raw)}, clean {len(clean)}, rejected {len(rejected)}")
    add("unique order ids", clean["OrderID"].is_unique, f"duplicate ids left: {int(clean['OrderID'].duplicated().sum())}")
    add("regions controlled", set(clean["Region"]) <= set(REGIONS), str(sorted(clean["Region"].unique())))
    add("channels controlled", set(clean["Channel"]) <= set(CHANNELS), str(sorted(clean["Channel"].unique())))
    add("categories controlled", set(clean["Category"]) <= set(CATEGORIES), str(sorted(clean["Category"].unique())))
    add("statuses controlled", set(clean["Status"]) <= set(STATUSES), str(sorted(clean["Status"].unique())))
    add("dates are real", pd.api.types.is_datetime64_any_dtype(clean["OrderDate"]), str(clean["OrderDate"].dtype))
    add("no zero quantity", int((clean["Quantity"] == 0).sum()) == 0, "")
    add("prices positive", bool((clean["UnitPrice"] > 0).all()), "")
    add("discounts in range", bool(((clean["Discount"] >= 0) & (clean["Discount"] <= 1)).all()), "")
    revenue_gap = (clean["NetRevenue"] - clean["Quantity"] * clean["UnitPrice"] * (1 - clean["Discount"])).abs().max()
    add("net revenue identity", revenue_gap < 1e-6, f"max gap {revenue_gap}")
    cost_gap = (clean["TotalCost"] - clean["Quantity"] * clean["UnitCost"]).abs().max()
    add("cost identity", cost_gap < 1e-6, f"max gap {cost_gap}")
    margin_gap = clean.loc[clean["NetRevenue"] != 0, "Margin"] - (
        (clean.loc[clean["NetRevenue"] != 0, "NetRevenue"] - clean.loc[clean["NetRevenue"] != 0, "TotalCost"])
        / clean.loc[clean["NetRevenue"] != 0, "NetRevenue"]
    )
    add("margin identity", float(margin_gap.abs().max()) < 1e-6, "")
    add("exact duplicates removed", log["reason_counts"].get("exact_duplicate", 0) >= params["exact_duplicate_copies"], str(log["reason_counts"]))
    add("unknown region rejected", log["reason_counts"].get("unknown_region", 0) >= 1, "")
    add("bad date rejected", log["reason_counts"].get("bad_date", 0) >= 1, "")
    add("zero qty rejected", log["reason_counts"].get("zero_qty", 0) >= 1, "")
    add("bad price rejected", log["reason_counts"].get("bad_price", 0) >= 1, "")
    add("conflict resolved", log["reason_counts"].get("order_id_conflict", 0) >= 1, "")
    add("negative qty corrected", log["adjustments"]["negative_qty_corrected"] >= 1, "")
    add("ship days cleared", log["adjustments"]["ship_days_cleared"] >= 1, "")

    scoped = in_scope(clean, params)
    old = clean[clean["OrderID"] == "EDGE-OLD"]
    future = clean[clean["OrderID"] == "EDGE-FUTURE"]
    add("old edge kept but outside window", len(old) == 1 and old.iloc[0]["OrderID"] not in set(scoped["OrderID"]), "")
    add("future edge kept but outside window", len(future) == 1 and future.iloc[0]["OrderID"] not in set(scoped["OrderID"]), "")
    cards = kpis(clean, params)
    add("orders match scope", cards["orders"] == len(scoped), f"{cards['orders']} vs {len(scoped)}")
    add("revenue finite", math.isfinite(cards["net_revenue"]) and cards["net_revenue"] > 0, str(cards["net_revenue"]))
    add("margin finite", cards["gross_margin"] is not None and math.isfinite(cards["gross_margin"]), str(cards["gross_margin"]))
    wider = dict(params)
    wider["include_cancelled"] = True
    add(
        "cancelled switch increases revenue",
        kpis(clean, wider)["net_revenue"] > cards["net_revenue"],
        "",
    )
    narrowed = dict(params)
    narrowed["analysis_end"] = "2024-06-30"
    add(
        "shorter window lowers revenue",
        kpis(clean, narrowed)["net_revenue"] < cards["net_revenue"],
        "",
    )
    return checks


def assert_all(checks: list[dict]) -> None:
    failed = [item for item in checks if not item["passed"]]
    if failed:
        detail = "; ".join(f"{item['name']}: {item['detail']}" for item in failed)
        raise AssertionError(detail)
