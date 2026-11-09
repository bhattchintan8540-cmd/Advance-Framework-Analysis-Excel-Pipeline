"""Clean the raw extract and keep an audit trail of every rejected row."""

from __future__ import annotations

import math

import pandas as pd

from pipeline.catalog import (
    CATEGORY_ALIASES,
    CHANNEL_ALIASES,
    RAW_COLUMNS,
    REGION_ALIASES,
    SEGMENT_ALIASES,
    STATUS_ALIASES,
)
from pipeline.numbers import parse_date, parse_number


def _blank(value) -> bool:
    return str(value).strip() == ""


def _lookup(value, mapping: dict[str, str]) -> str | None:
    key = " ".join(str(value).strip().lower().split())
    return mapping.get(key)


def _money(quantity: float, unit_price: float, discount: float, unit_cost: float) -> tuple[float, float, float, float | None]:
    gross = quantity * unit_price
    net = gross * (1 - discount)
    cost = quantity * unit_cost
    margin = None if net == 0 else (net - cost) / net
    return gross, net, cost, margin


def clean_raw(raw: pd.DataFrame) -> tuple[pd.DataFrame, pd.DataFrame, dict]:
    if list(raw.columns) != RAW_COLUMNS:
        raise ValueError(f"Unexpected raw columns: {list(raw.columns)}")

    working = raw.copy()
    duplicate_mask = working.duplicated(keep="first")
    rejected = []
    for position in working.index[duplicate_mask]:
        rejected.append(_reject(working.loc[position], "exact_duplicate", int(position)))

    kept_rows = []
    notes = {
        "negative_qty_corrected": 0,
        "discount_percent_scaled": 0,
        "ship_days_cleared": 0,
        "blank_segment_imputed": 0,
        "unknown_segment_imputed": 0,
    }

    for position, record in working.loc[~duplicate_mask].iterrows():
        cleaned, reason, row_notes = _clean_row(record)
        for key in row_notes:
            notes[key] += row_notes[key]
        if cleaned is None:
            rejected.append(_reject(record, reason or "rejected", int(position)))
            continue
        cleaned["SourceRow"] = int(position) + 2  # header is row 1 in the workbook
        kept_rows.append(cleaned)

    preliminary = pd.DataFrame(kept_rows)
    if preliminary.empty:
        raise RuntimeError("Cleaning removed every row")

    conflict_losers = []
    winners = []
    for order_id, group in preliminary.groupby("OrderID", sort=False):
        if len(group) == 1:
            winners.append(group.iloc[0])
            continue
        ranked = group.sort_values(["NetRevenue", "OrderDate"], ascending=[False, False])
        winners.append(ranked.iloc[0])
        for _, loser in ranked.iloc[1:].iterrows():
            conflict_losers.append(order_id)
            rejected.append(
                {
                    "OrderID": order_id,
                    "Reason": "order_id_conflict",
                    "Detail": "A higher net-revenue row with the same OrderID was kept",
                    "OrderDate": str(loser["OrderDate"].date()),
                    "Region": loser["Region"],
                    "Status": loser["Status"],
                }
            )
    notes["order_id_conflict"] = len(conflict_losers)

    clean = pd.DataFrame(winners).reset_index(drop=True)
    clean["OrderDate"] = pd.to_datetime(clean["OrderDate"])
    rejected_frame = pd.DataFrame(rejected)
    reason_counts = rejected_frame["Reason"].value_counts().to_dict() if not rejected_frame.empty else {}
    exact_duplicates = int(duplicate_mask.sum())
    failed = int(len(rejected_frame) - exact_duplicates - len(conflict_losers))
    balance_ok = int(len(raw)) == exact_duplicates + failed + len(conflict_losers) + int(len(clean))
    log = {
        "raw_rows": int(len(raw)),
        "clean_rows": int(len(clean)),
        "rejected_rows": int(len(rejected_frame)),
        "exact_duplicates": exact_duplicates,
        "reason_counts": {key: int(value) for key, value in reason_counts.items()},
        "adjustments": notes,
        "balance_ok": balance_ok,
    }
    if not balance_ok:
        raise RuntimeError(f"Row accounting failed: {log}")
    return clean, rejected_frame, log


def _reject(record, reason: str, source_index: int) -> dict:
    return {
        "OrderID": record.get("OrderID", ""),
        "Reason": reason,
        "Detail": reason,
        "OrderDate": record.get("OrderDate", ""),
        "Region": record.get("Region", ""),
        "Status": record.get("Status", ""),
        "SourceRow": source_index + 2,
    }


def _clean_row(record: pd.Series) -> tuple[dict | None, str | None, dict]:
    notes = {
        "negative_qty_corrected": 0,
        "discount_percent_scaled": 0,
        "ship_days_cleared": 0,
        "blank_segment_imputed": 0,
        "unknown_segment_imputed": 0,
    }
    region = _lookup(record["Region"], REGION_ALIASES)
    if region is None:
        return None, "unknown_region", notes
    channel = _lookup(record["Channel"], CHANNEL_ALIASES)
    if channel is None:
        return None, "unknown_channel", notes
    category = _lookup(record["Category"], CATEGORY_ALIASES)
    if category is None:
        return None, "unknown_category", notes
    status = _lookup(record["Status"], STATUS_ALIASES)
    if status is None:
        return None, "unknown_status", notes
    order_date = parse_date(record["OrderDate"])
    if order_date is None:
        return None, "bad_date", notes

    if _blank(record["Segment"]):
        segment = "Unknown"
        notes["blank_segment_imputed"] = 1
    else:
        segment = _lookup(record["Segment"], SEGMENT_ALIASES)
        if segment is None:
            segment = "Unknown"
            notes["unknown_segment_imputed"] = 1

    quantity = parse_number(record["Quantity"])
    if quantity is None or not math.isfinite(quantity):
        return None, "bad_qty", notes
    if abs(quantity - round(quantity)) > 1e-6:
        return None, "bad_qty", notes
    quantity = int(round(quantity))
    qty_adjusted = "No"
    if quantity < 0:
        quantity = abs(quantity)
        qty_adjusted = "Yes"
        notes["negative_qty_corrected"] = 1
    if quantity == 0:
        return None, "zero_qty", notes

    unit_price = parse_number(record["UnitPrice"])
    if unit_price is None or unit_price <= 0:
        return None, "bad_price", notes
    unit_cost = parse_number(record["UnitCost"])
    if unit_cost is None or unit_cost < 0:
        return None, "bad_cost", notes
    discount = parse_number(record["Discount"])
    if discount is None:
        return None, "bad_discount", notes
    if 1 < discount <= 100:
        discount = discount / 100.0
        notes["discount_percent_scaled"] = 1
    if discount < 0 or discount > 1:
        return None, "bad_discount", notes

    ship_days = parse_number(record["ShipDays"])
    if ship_days is None:
        ship_value = None
    elif ship_days < 0 or ship_days > 120:
        ship_value = None
        notes["ship_days_cleared"] = 1
    else:
        ship_value = int(round(ship_days))

    gross, net, cost, margin = _money(quantity, unit_price, discount, unit_cost)
    return (
        {
            "OrderID": str(record["OrderID"]).strip(),
            "OrderDate": order_date,
            "Region": region,
            "Channel": channel,
            "Category": category,
            "Product": str(record["Product"]).strip(),
            "Segment": segment,
            "Quantity": quantity,
            "UnitPrice": float(unit_price),
            "Discount": float(discount),
            "UnitCost": float(unit_cost),
            "ShipDays": ship_value,
            "Status": status,
            "SalesRep": str(record["SalesRep"]).strip(),
            "QtyAdjusted": qty_adjusted,
            "GrossRevenue": gross,
            "NetRevenue": net,
            "TotalCost": cost,
            "Margin": margin,
        },
        None,
        notes,
    )
