"""Run the Harbor & Co. Excel analytics pipeline from raw extract to workbook."""

from __future__ import annotations

import json
from pathlib import Path

import pandas as pd

from pipeline.charts import render_dashboards
from pipeline.clean import clean_raw
from pipeline.generate import generate_raw
from pipeline.metrics import build_insights, summarize
from pipeline.parameters import ROOT, load_parameters
from pipeline.qa import assert_all, run_checks
from pipeline.workbook import build_workbook


def _write_clean_csv(frame: pd.DataFrame, path: Path) -> None:
    export = frame.copy()
    export["OrderDate"] = export["OrderDate"].dt.strftime("%Y-%m-%d")
    export.to_csv(path, index=False)


def run(params: dict | None = None, root: Path | None = None) -> dict:
    params = params or load_parameters()
    root = root or ROOT
    raw_path = root / "data" / "raw" / "harbor_sales_raw.csv"
    clean_path = root / "data" / "processed" / "harbor_sales_clean.csv"
    rejected_path = root / "data" / "processed" / "rejected_rows.csv"
    log_path = root / "data" / "processed" / "cleaning_log.json"
    workbook_path = root / "excel" / "Harbor_and_Co_Excel_Pipeline.xlsx"
    insights_path = root / "outputs" / "insights.md"
    qa_path = root / "outputs" / "qa_report.json"
    summary_path = root / "outputs" / "summary.json"

    for folder in (raw_path.parent, clean_path.parent, workbook_path.parent, insights_path.parent, root / "dashboards"):
        folder.mkdir(parents=True, exist_ok=True)

    raw = generate_raw(params)
    raw.to_csv(raw_path, index=False)
    clean, rejected, cleaning_log = clean_raw(raw)
    checks = run_checks(raw, clean, rejected, cleaning_log, params)
    assert_all(checks)
    _write_clean_csv(clean, clean_path)
    rejected.to_csv(rejected_path, index=False)
    log_path.write_text(json.dumps(cleaning_log, indent=2), encoding="utf-8")

    summary = summarize(clean, params)
    summary["narrative"] = build_insights(clean, params, summary, cleaning_log)
    summary["include_cancelled"] = params["include_cancelled"]
    build_workbook(raw, clean, rejected, cleaning_log, summary, params, workbook_path)
    images = render_dashboards(summary, params, cleaning_log, root / "dashboards")
    insights_path.write_text(summary["narrative"], encoding="utf-8")

    cards = summary["kpis"]
    payload = {
        "raw_rows": cleaning_log["raw_rows"],
        "clean_rows": cleaning_log["clean_rows"],
        "rejected_rows": cleaning_log["rejected_rows"],
        "net_revenue": cards["net_revenue"],
        "orders": cards["orders"],
        "aov": cards["aov"],
        "gross_margin": cards["gross_margin"],
        "return_rate": cards["return_rate"],
        "on_time_rate": cards["on_time_rate"],
        "high_value_orders": cards["high_value_orders"],
        "high_value_cutoff": cards["high_value_cutoff"],
        "low_margin_orders": cards["low_margin_orders"],
        "checks_passed": len(checks),
        "workbook": str(workbook_path.relative_to(root)),
        "dashboards": [str(path.relative_to(root)) for path in images],
    }
    summary_path.write_text(json.dumps(payload, indent=2), encoding="utf-8")
    qa_path.write_text(json.dumps({"passed": True, "checks": checks}, indent=2), encoding="utf-8")
    return payload


def main() -> int:
    payload = run()
    print(json.dumps(payload, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
