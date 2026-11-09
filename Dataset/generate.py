"""Build a deterministic dirty sales extract plus known edge cases."""

from __future__ import annotations

import calendar
from datetime import datetime

import numpy as np
import pandas as pd

from pipeline.catalog import (
    CATALOG,
    CATEGORIES,
    CHANNEL_DIRT,
    CHANNELS,
    RAW_COLUMNS,
    REGION_DIRT,
    REGIONS,
    REPS,
    SEGMENTS,
    STATUS_DIRT,
)
from pipeline.numbers import camel_words, int_to_words


def _month_weights():
    months = []
    for year in (2024, 2025):
        last_month = 12 if year == 2024 else 9
        for month in range(1, last_month + 1):
            weight = 1.0
            if month in (11, 12):
                weight = 1.85
            elif month in (1, 2):
                weight = 0.72
            elif month in (6, 7):
                weight = 1.18
            months.append((year, month, weight))
    return months


def _pick_month(rng: np.random.Generator):
    months = _month_weights()
    weights = np.array([item[2] for item in months], dtype=float)
    weights /= weights.sum()
    index = int(rng.choice(len(months), p=weights))
    return months[index][0], months[index][1]


def _format_date(moment: datetime, style: int) -> str:
    if style == 0:
        return moment.strftime("%m/%d/%Y")
    if style == 1:
        return moment.strftime("%d-%b-%y")
    if style == 2:
        return moment.strftime("%Y/%m/%d")
    return moment.strftime("%d.%m.%Y")


def _base_record(rng: np.random.Generator, order_id: str) -> dict:
    year, month = _pick_month(rng)
    day = int(rng.integers(1, calendar.monthrange(year, month)[1] + 1))
    region = str(rng.choice(REGIONS, p=[0.28, 0.26, 0.16, 0.30]))
    channel = str(rng.choice(CHANNELS, p=[0.46, 0.39, 0.15]))
    category = str(rng.choice(CATEGORIES, p=[0.34, 0.18, 0.27, 0.21]))
    product, base_price, cost_ratio = CATALOG[category][int(rng.integers(0, len(CATALOG[category])))]
    if channel == "Wholesale":
        quantity = int(rng.integers(8, 36))
        discount = float(rng.uniform(0.12, 0.32))
    elif channel == "Online":
        quantity = int(rng.integers(1, 4))
        discount = float(rng.uniform(0.02, 0.18))
    else:
        quantity = int(rng.integers(1, 5))
        discount = float(rng.uniform(0.0, 0.12))
    price = base_price * float(rng.uniform(0.92, 1.08))
    if region == "East":
        price *= 0.90
    if month in (11, 12):
        quantity += 1 if channel != "Wholesale" else 4
    if category == "Textiles":
        status = str(rng.choice(["Completed", "Returned", "Cancelled"], p=[0.83, 0.14, 0.03]))
    else:
        status = str(rng.choice(["Completed", "Returned", "Cancelled"], p=[0.93, 0.04, 0.03]))
    ship = int(rng.normal(6 if region != "East" else 9, 2))
    ship = max(1, min(ship, 21))
    segment = str(rng.choice(SEGMENTS, p=[0.62, 0.24, 0.14]))
    return {
        "OrderID": order_id,
        "OrderDate": datetime(year, month, day),
        "Region": region,
        "Channel": channel,
        "Category": category,
        "Product": product,
        "Segment": segment,
        "Quantity": quantity,
        "UnitPrice": round(price, 2),
        "Discount": round(discount, 2),
        "UnitCost": round(price * cost_ratio, 2),
        "ShipDays": ship,
        "Status": status,
        "SalesRep": str(rng.choice(REPS)),
    }


def _dirty_record(record: dict, rng: np.random.Generator, index: int) -> dict:
    region = record["Region"]
    channel = record["Channel"]
    status = record["Status"]
    if rng.random() < 0.30:
        region = str(rng.choice(REGION_DIRT[region]))
    if rng.random() < 0.22:
        channel = str(rng.choice(CHANNEL_DIRT[channel]))
    if rng.random() < 0.20:
        status = str(rng.choice(STATUS_DIRT[status]))
    segment = record["Segment"]
    if rng.random() < 0.04:
        segment = ""
    elif rng.random() < 0.08:
        segment = segment.lower()
    quantity = str(record["Quantity"])
    if record["Quantity"] <= 12 and rng.random() < 0.03:
        quantity = camel_words(int(record["Quantity"])) if rng.random() < 0.5 else int_to_words(int(record["Quantity"]))
    price = f"{record['UnitPrice']:.2f}"
    if rng.random() < 0.08:
        price = f"${record['UnitPrice']:,.2f}"
    discount = f"{record['Discount']:.2f}"
    if rng.random() < 0.06:
        discount = f"{int(round(record['Discount'] * 100))}%"
    ship = str(record["ShipDays"])
    if rng.random() < 0.03:
        ship = "N/A"
    category = record["Category"] if rng.random() > 0.1 else record["Category"].upper()
    return {
        "OrderID": record["OrderID"],
        "OrderDate": _format_date(record["OrderDate"], index % 4),
        "Region": region,
        "Channel": channel,
        "Category": category,
        "Product": record["Product"],
        "Segment": segment,
        "Quantity": quantity,
        "UnitPrice": price,
        "Discount": discount,
        "UnitCost": f"{record['UnitCost']:.2f}",
        "ShipDays": ship,
        "Status": status,
        "SalesRep": record["SalesRep"],
    }


def _edge_rows() -> list[dict]:
    """Hand-built dirty rows. Tests lock their expected cleaning outcome."""
    common = {
        "Channel": "Store",
        "Category": "Lighting",
        "Product": "Arc Lamp",
        "Segment": "Consumer",
        "Quantity": "2",
        "UnitPrice": "310.00",
        "Discount": "0.05",
        "UnitCost": "148.80",
        "ShipDays": "4",
        "Status": "Completed",
        "SalesRep": "A. Shah",
        "Region": "North",
        "OrderDate": "01/15/2024",
    }

    def row(order_id: str, **overrides) -> dict:
        record = dict(common)
        record["OrderID"] = order_id
        record.update(overrides)
        return record

    duplicate = row("EDGE-DUP", OrderDate="02/02/2024", Region="South", Quantity="3", UnitPrice="180.00")
    conflict_low = row("EDGE-CONF", OrderDate="03/03/2024", UnitPrice="100.00", UnitCost="40.00", Quantity="1")
    conflict_high = row("EDGE-CONF", OrderDate="03/04/2024", UnitPrice="400.00", UnitCost="120.00", Quantity="2", Region="West")
    return [
        duplicate,
        dict(duplicate),
        row("EDGE-WORD", Quantity="Five", UnitPrice="TwoHundredForty", OrderDate="04/04/2024"),
        row("EDGE-CURRENCY", UnitPrice="$1,250.50", Quantity="1", UnitCost="700.00", OrderDate="05/05/2024", Category="Furniture", Product="Linen Sofa"),
        row("EDGE-NEGQ", Quantity="-3", OrderDate="06/06/2024"),
        row("EDGE-BADREG", Region="Atlantis", OrderDate="01/20/2024"),
        row("EDGE-BADDATE", OrderDate="yesterday"),
        row("EDGE-CANCEL", Status="cancled", OrderDate="07/07/2024"),
        row("EDGE-COMPLETE", Status="Compleated", OrderDate="07/08/2024"),
        row("EDGE-REGION", Region="  nort ", OrderDate="15-Aug-24"),
        row("EDGE-ZERO", Quantity="0", OrderDate="08/01/2024"),
        conflict_low,
        conflict_high,
        row("EDGE-PCT", Discount="20%", OrderDate="09/09/2024"),
        row("EDGE-LEAP", OrderDate="2024/02/29", Quantity="1"),
        row("EDGE-SEG", Segment="", OrderDate="10/10/2024"),
        row("EDGE-SHIP", ShipDays="N/A", OrderDate="11/11/2024"),
        row("EDGE-SHIPNEG", ShipDays="-4", OrderDate="11/12/2024"),
        row("EDGE-PRICE", UnitPrice="(50.00)", OrderDate="12/01/2024"),
        row("EDGE-OLD", OrderDate="2023-12-01", Region="East"),
        row("EDGE-FUTURE", OrderDate="10/15/2025", Region="West"),
        row("EDGE-SPACE", Region=" west ", OrderDate="12.12.2024"),
        row("EDGE-RETURN", Status="RETURNED", OrderDate="01/18/2025", Category="Textiles", Product="Wool Throw"),
        row(
            "EDGE-CLEARANCE",
            OrderDate="02/14/2025",
            Category="Furniture",
            Product="Lounge Chair",
            Quantity="1",
            UnitPrice="200.00",
            UnitCost="190.00",
            Discount="15%",
            Region="East",
        ),
    ]


def generate_raw(params: dict) -> pd.DataFrame:
    rng = np.random.default_rng(params["random_seed"])
    records = []
    for index in range(params["raw_row_target"]):
        clean = _base_record(rng, f"HC-{100000 + index}")
        records.append(_dirty_record(clean, rng, index))
    copy_count = min(params["exact_duplicate_copies"], len(records))
    copied = [dict(records[int(i)]) for i in rng.choice(len(records), size=copy_count, replace=False)]
    records.extend(copied)
    records.extend(_edge_rows())
    frame = pd.DataFrame(records, columns=RAW_COLUMNS)
    return frame.astype(str).replace({"nan": ""})
