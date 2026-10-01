"""Checks for cleaning rules, KPI definitions, and the Excel workbook structure."""

from pathlib import Path

import json

import pandas as pd
import pytest
from openpyxl import load_workbook

from pipeline.clean import clean_raw
from pipeline.generate import generate_raw
from pipeline.metrics import high_value_cutoff, in_scope, kpis, summarize
from pipeline.numbers import camel_words, int_to_words, parse_date, parse_number, words_to_number
from pipeline.parameters import load_parameters
from pipeline.qa import run_checks
from pipeline.workbook import build_workbook

ROOT = Path(__file__).resolve().parents[1]


@pytest.fixture(scope="module")
def params():
    return load_parameters(ROOT / "config" / "parameters.json")


@pytest.fixture(scope="module")
def built(params):
    raw = generate_raw(params)
    clean, rejected, log = clean_raw(raw)
    summary = summarize(clean, params)
    summary["narrative"] = "test narrative"
    summary["include_cancelled"] = False
    return raw, clean, rejected, log, summary


def test_word_numbers_roundtrip_through_9999():
    for value in list(range(0, 200)) + [240, 896, 1000, 1250, 23896, 99999]:
        assert words_to_number(int_to_words(value)) == value
        assert words_to_number(camel_words(value)) == value


def test_number_and_date_parsers():
    assert parse_number("$1,250.50") == 1250.50
    assert parse_number("20%") == 0.20
    assert parse_number("(50.00)") == -50
    assert parse_number("N/A") is None
    assert parse_number("Five") == 5
    assert parse_date("15-Aug-24") == pd.Timestamp("2024-08-15")
    assert parse_date("2024/02/29") == pd.Timestamp("2024-02-29")
    assert parse_date("12.12.2024") == pd.Timestamp("2024-12-12")
    assert parse_date("yesterday") is None


def test_parameters_reject_bad_values(tmp_path):
    source = tmp_path / "parameters.json"
    original = load_parameters(ROOT / "config" / "parameters.json")
    original["high_value_percentile"] = 1.4
    source.write_text(json.dumps(original), encoding="utf-8")
    with pytest.raises(ValueError):
        load_parameters(source)


def test_cleaning_edges(built):
    _raw, clean, rejected, log, _summary = built
    checks = {item["name"]: item["passed"] for item in run_checks(_raw, clean, rejected, log, load_parameters(ROOT / "config" / "parameters.json"))}
    assert all(checks.values()), checks
    word = clean.set_index("OrderID").loc["EDGE-WORD"]
    assert word["Quantity"] == 5
    assert word["UnitPrice"] == 240
    currency = clean.set_index("OrderID").loc["EDGE-CURRENCY"]
    assert currency["UnitPrice"] == 1250.50
    assert clean.set_index("OrderID").loc["EDGE-NEGQ"]["Quantity"] == 3
    assert clean.set_index("OrderID").loc["EDGE-NEGQ"]["QtyAdjusted"] == "Yes"
    assert clean.set_index("OrderID").loc["EDGE-CANCEL"]["Status"] == "Cancelled"
    assert clean.set_index("OrderID").loc["EDGE-COMPLETE"]["Status"] == "Completed"
    assert clean.set_index("OrderID").loc["EDGE-REGION"]["Region"] == "North"
    assert clean.set_index("OrderID").loc["EDGE-REGION"]["OrderDate"] == pd.Timestamp("2024-08-15")
    assert clean.set_index("OrderID").loc["EDGE-PCT"]["Discount"] == pytest.approx(0.20)
    assert clean.set_index("OrderID").loc["EDGE-LEAP"]["OrderDate"] == pd.Timestamp("2024-02-29")
    assert clean.set_index("OrderID").loc["EDGE-SEG"]["Segment"] == "Unknown"
    assert pd.isna(clean.set_index("OrderID").loc["EDGE-SHIP"]["ShipDays"])
    assert pd.isna(clean.set_index("OrderID").loc["EDGE-SHIPNEG"]["ShipDays"])
    assert clean.set_index("OrderID").loc["EDGE-SPACE"]["Region"] == "West"
    assert clean.set_index("OrderID").loc["EDGE-RETURN"]["Status"] == "Returned"
    assert clean.set_index("OrderID").loc["EDGE-CONF"]["Region"] == "West"
    assert (clean["OrderID"] == "EDGE-DUP").sum() == 1
    reasons = set(rejected["Reason"])
    for required in ("exact_duplicate", "unknown_region", "bad_date", "zero_qty", "bad_price", "order_id_conflict"):
        assert required in reasons
    assert "EDGE-BADREG" not in set(clean["OrderID"])
    assert "EDGE-ZERO" not in set(clean["OrderID"])
    clearance = clean.set_index("OrderID").loc["EDGE-CLEARANCE"]
    assert clearance["NetRevenue"] == pytest.approx(170)
    assert clearance["Margin"] < 0


def test_kpi_definitions(built, params):
    _raw, clean, _rejected, _log, summary = built
    cards = summary["kpis"]
    scoped = in_scope(clean, params)
    assert cards["orders"] == len(scoped)
    assert cards["net_revenue"] == pytest.approx(scoped["NetRevenue"].sum())
    assert cards["gross_margin"] == pytest.approx((scoped["NetRevenue"].sum() - scoped["TotalCost"].sum()) / scoped["NetRevenue"].sum())
    cutoff = high_value_cutoff(clean, params)
    assert cards["high_value_orders"] == int((scoped["NetRevenue"] >= cutoff).sum())
    non_cancelled = scoped[scoped["Status"] != "Cancelled"]
    shipped = non_cancelled[non_cancelled["ShipDays"].notna()]
    assert cards["on_time_rate"] == pytest.approx((shipped["ShipDays"] <= params["late_shipment_days"]).mean())
    assert cards["return_rate"] == pytest.approx((non_cancelled["Status"] == "Returned").mean())
    east = summary["regions"].set_index("Name").loc["East", "NetRevenue"]
    filtered = kpis(clean, params, {"region": "East"})
    assert filtered["net_revenue"] == pytest.approx(east)
    # The region comparison keeps every region when the dashboard region is East,
    # because that chart ignores the region filter. The KPI card does not.
    assert filtered["orders"] < cards["orders"]


def test_workbook_structure(built, params, tmp_path):
    raw, clean, rejected, log, summary = built
    path = tmp_path / "pipeline.xlsx"
    build_workbook(raw, clean, rejected, log, summary, params, path)
    wb = load_workbook(path)
    expected = {
        "Cover", "Parameters", "Raw_Data", "Lookup_Maps", "Cleaned_Data", "Chart_Data",
        "Analysis", "Dashboard", "Insights", "QA_Tests", "Rejected_Rows", "Data_Dictionary", "Cleaning_Log", "Lists",
    }
    assert expected <= set(wb.sheetnames)
    assert len(wb["Dashboard"]._charts) == 10
    assert "CleanData" in wb["Cleaned_Data"].tables
    assert "RawData" in wb["Raw_Data"].tables
    kpi = wb["Dashboard"]["A9"].value
    assert kpi.startswith("=SUMIFS(CleanData[NetRevenue]")
    assert "pRegion" in kpi and "pIncludeCancelled" in kpi
    assert wb["Cleaned_Data"]["Q2"].value.startswith("=P2*(1-J2)")
    assert wb["Cleaned_Data"]["W2"].value.startswith("=IF(L2=")
    names = set(wb.defined_names)
    for required in ("pRegion", "pChannel", "pCategory", "pStart", "pEnd", "pIncludeCancelled", "pLateDays", "pLowMargin", "pHighCutoff"):
        assert required in names
    validations = list(wb["Dashboard"].data_validations.dataValidation)
    assert len(validations) == 4
    assert wb["QA_Tests"]["E5"].value.startswith('=IF(')
    assert wb["Parameters"]["C8"].value == pytest.approx(summary["cutoff"], abs=0.01)
    assert wb["Dashboard"]["I5"].value == "No"
    assert wb["Chart_Data"]["B4"].value.startswith("=SUMIFS")
    assert "honor" not in wb["Chart_Data"]["B4"].value
    # Region chart formula must pin the region to the row label and still respect the other filters.
    assert "CleanData[Region],A4" in wb["Chart_Data"]["B4"].value
    assert "pChannel" in wb["Chart_Data"]["B4"].value
