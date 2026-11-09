"""Load and validate pipeline parameters."""

from __future__ import annotations

import json
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
DEFAULT_PATH = ROOT / "config" / "parameters.json"

REQUIRED = (
    "analysis_start",
    "analysis_end",
    "high_value_percentile",
    "late_shipment_days",
    "low_margin_threshold",
    "include_cancelled",
    "random_seed",
    "raw_row_target",
    "exact_duplicate_copies",
)


def load_parameters(path: Path | None = None) -> dict:
    source = Path(path) if path else DEFAULT_PATH
    with source.open(encoding="utf-8") as handle:
        params = json.load(handle)
    missing = [key for key in REQUIRED if key not in params]
    if missing:
        raise ValueError(f"Missing parameters: {', '.join(missing)}")
    percentile = float(params["high_value_percentile"])
    if not 0 < percentile < 1:
        raise ValueError("high_value_percentile must be between 0 and 1")
    if params["analysis_start"] > params["analysis_end"]:
        raise ValueError("analysis_start is after analysis_end")
    late_days = int(params["late_shipment_days"])
    if late_days < 0:
        raise ValueError("late_shipment_days cannot be negative")
    margin = float(params["low_margin_threshold"])
    if not 0 <= margin <= 1:
        raise ValueError("low_margin_threshold must be between 0 and 1")
    if int(params["raw_row_target"]) < 100:
        raise ValueError("raw_row_target must be at least 100")
    params["high_value_percentile"] = percentile
    params["late_shipment_days"] = late_days
    params["low_margin_threshold"] = margin
    params["include_cancelled"] = bool(params["include_cancelled"])
    params["random_seed"] = int(params["random_seed"])
    params["raw_row_target"] = int(params["raw_row_target"])
    params["exact_duplicate_copies"] = int(params["exact_duplicate_copies"])
    filters = params.setdefault("dashboard_filters", {})
    filters.setdefault("region", "All")
    filters.setdefault("channel", "All")
    filters.setdefault("category", "All")
    return params
