"""KPI definitions shared by the Excel formulas, the tests, and the dashboards.

Filter rules
- KPI cards honor region, channel, category, date window, and the cancelled switch.
- Comparison charts drop only the filter that would collapse the comparison
  (a region chart ignores the region filter, a channel chart ignores the channel filter).
- Return rate always uses non-cancelled orders in the denominator.
- On-time fulfillment always ignores cancelled orders and rows with no ship days.
- Status mix always shows Completed, Returned, and Cancelled.
- The high-value cutoff is the linear percentile of in-scope net revenue for the
  saved parameter set. Dashboard dropdowns do not recompute that cutoff.
"""

from __future__ import annotations

import pandas as pd

REGIONS = ["North", "South", "East", "West"]
CHANNELS = ["Store", "Online", "Wholesale"]
CATEGORIES = ["Furniture", "Lighting", "Textiles", "Kitchen"]
STATUSES = ["Completed", "Returned", "Cancelled"]


def _as_timestamp(value) -> pd.Timestamp:
    return pd.Timestamp(value)


def active_filters(params: dict, overrides: dict | None = None) -> dict:
    filters = dict(params.get("dashboard_filters", {}))
    filters.setdefault("region", "All")
    filters.setdefault("channel", "All")
    filters.setdefault("category", "All")
    if overrides:
        filters.update(overrides)
    return filters


def dimension_mask(
    frame: pd.DataFrame,
    params: dict,
    *,
    region: str = "All",
    channel: str = "All",
    category: str = "All",
    respect_cancelled: bool = True,
) -> pd.Series:
    start = _as_timestamp(params["analysis_start"])
    end = _as_timestamp(params["analysis_end"])
    mask = (frame["OrderDate"] >= start) & (frame["OrderDate"] <= end)
    if region != "All":
        mask &= frame["Region"] == region
    if channel != "All":
        mask &= frame["Channel"] == channel
    if category != "All":
        mask &= frame["Category"] == category
    if respect_cancelled and not params["include_cancelled"]:
        mask &= frame["Status"] != "Cancelled"
    return mask


def in_scope(frame: pd.DataFrame, params: dict, filters: dict | None = None) -> pd.DataFrame:
    chosen = active_filters(params, filters)
    mask = dimension_mask(
        frame,
        params,
        region=chosen["region"],
        channel=chosen["channel"],
        category=chosen["category"],
        respect_cancelled=True,
    )
    return frame.loc[mask].copy()


def high_value_cutoff(frame: pd.DataFrame, params: dict) -> float:
    scoped = in_scope(frame, params)
    if scoped.empty:
        return 0.0
    return float(scoped["NetRevenue"].quantile(params["high_value_percentile"], interpolation="linear"))


def kpis(frame: pd.DataFrame, params: dict, filters: dict | None = None, cutoff: float | None = None) -> dict:
    scoped = in_scope(frame, params, filters)
    net = float(scoped["NetRevenue"].sum()) if not scoped.empty else 0.0
    cost = float(scoped["TotalCost"].sum()) if not scoped.empty else 0.0
    orders = int(len(scoped))
    chosen = active_filters(params, filters)
    non_cancelled = frame.loc[
        dimension_mask(
            frame,
            params,
            region=chosen["region"],
            channel=chosen["channel"],
            category=chosen["category"],
            respect_cancelled=False,
        )
        & (frame["Status"] != "Cancelled")
    ]
    returned = int((non_cancelled["Status"] == "Returned").sum())
    return_base = int(len(non_cancelled))
    shipped = non_cancelled[non_cancelled["ShipDays"].notna()]
    on_time = int((shipped["ShipDays"] <= params["late_shipment_days"]).sum()) if len(shipped) else 0
    threshold = high_value_cutoff(frame, params) if cutoff is None else float(cutoff)
    high_value = int((scoped["NetRevenue"] >= threshold).sum()) if orders else 0
    low_margin = int((scoped["Margin"].notna() & (scoped["Margin"] < params["low_margin_threshold"])).sum()) if orders else 0
    return {
        "net_revenue": net,
        "total_cost": cost,
        "orders": orders,
        "aov": (net / orders) if orders else None,
        "gross_margin": ((net - cost) / net) if net else None,
        "return_rate": (returned / return_base) if return_base else None,
        "returned_orders": returned,
        "return_base": return_base,
        "on_time_rate": (on_time / len(shipped)) if len(shipped) else None,
        "on_time_orders": on_time,
        "shipped_orders": int(len(shipped)),
        "high_value_orders": high_value,
        "high_value_cutoff": threshold,
        "low_margin_orders": low_margin,
        "units": int(scoped["Quantity"].sum()) if orders else 0,
    }


def _group_sum(frame: pd.DataFrame, column: str, values: list[str]) -> pd.DataFrame:
    grouped = frame.groupby(column, dropna=False).agg(
        NetRevenue=("NetRevenue", "sum"),
        Orders=("OrderID", "size"),
        TotalCost=("TotalCost", "sum"),
        Units=("Quantity", "sum"),
    )
    grouped = grouped.reindex(values, fill_value=0)
    grouped["AOV"] = grouped.apply(lambda row: (row["NetRevenue"] / row["Orders"]) if row["Orders"] else None, axis=1)
    grouped["GrossMargin"] = grouped.apply(
        lambda row: ((row["NetRevenue"] - row["TotalCost"]) / row["NetRevenue"]) if row["NetRevenue"] else None,
        axis=1,
    )
    return grouped.reset_index().rename(columns={column: "Name"})


def region_summary(frame, params, filters=None) -> pd.DataFrame:
    chosen = active_filters(params, filters)
    scoped = in_scope(frame, params, {**chosen, "region": "All"})
    summary = _group_sum(scoped, "Region", REGIONS)
    margins = []
    for region in REGIONS:
        part = scoped.loc[scoped["Region"] == region, "Margin"].dropna()
        margins.append(float(part.mean()) if len(part) else None)
    summary["AvgMargin"] = margins
    return summary


def channel_summary(frame, params, filters=None) -> pd.DataFrame:
    chosen = active_filters(params, filters)
    scoped = in_scope(frame, params, {**chosen, "channel": "All"})
    return _group_sum(scoped, "Channel", CHANNELS)


def category_summary(frame, params, filters=None) -> pd.DataFrame:
    chosen = active_filters(params, filters)
    scoped = in_scope(frame, params, {**chosen, "category": "All"})
    summary = _group_sum(scoped, "Category", CATEGORIES)
    rates = []
    for category in CATEGORIES:
        part = scoped[scoped["Category"] == category]
        base = part[part["Status"] != "Cancelled"]
        rates.append(float((base["Status"] == "Returned").mean()) if len(base) else None)
    summary["ReturnRate"] = rates
    return summary


def month_index(params: dict) -> pd.PeriodIndex:
    start = _as_timestamp(params["analysis_start"]).to_period("M")
    end = _as_timestamp(params["analysis_end"]).to_period("M")
    return pd.period_range(start, end, freq="M")


def monthly_trend(frame, params, filters=None) -> pd.DataFrame:
    scoped = in_scope(frame, params, filters)
    if scoped.empty:
        totals = pd.Series(0.0, index=month_index(params))
        orders = pd.Series(0, index=month_index(params))
    else:
        periods = scoped["OrderDate"].dt.to_period("M")
        totals = scoped.groupby(periods)["NetRevenue"].sum().reindex(month_index(params), fill_value=0.0)
        orders = scoped.groupby(periods)["OrderID"].size().reindex(month_index(params), fill_value=0)
    result = pd.DataFrame({"YearMonth": totals.index.astype(str), "NetRevenue": totals.values, "Orders": orders.values})
    return result


def quarterly_trend(frame, params, filters=None) -> pd.DataFrame:
    scoped = in_scope(frame, params, filters)
    start = _as_timestamp(params["analysis_start"]).to_period("Q")
    end = _as_timestamp(params["analysis_end"]).to_period("Q")
    index = pd.period_range(start, end, freq="Q")
    if scoped.empty:
        totals = pd.Series(0.0, index=index)
    else:
        totals = scoped.groupby(scoped["OrderDate"].dt.to_period("Q"))["NetRevenue"].sum().reindex(index, fill_value=0.0)
    return pd.DataFrame({"Quarter": [f"{period.year}-Q{period.quarter}" for period in totals.index], "NetRevenue": totals.values})


def category_region_matrix(frame, params, filters=None) -> pd.DataFrame:
    chosen = active_filters(params, filters)
    scoped = in_scope(frame, params, {**chosen, "region": "All", "category": "All"})
    pivot = pd.pivot_table(
        scoped,
        values="NetRevenue",
        index="Category",
        columns="Region",
        aggfunc="sum",
        fill_value=0.0,
    )
    return pivot.reindex(index=CATEGORIES, columns=REGIONS, fill_value=0.0)


def monthly_orders_by_region(frame, params, filters=None) -> pd.DataFrame:
    chosen = active_filters(params, filters)
    scoped = in_scope(frame, params, {**chosen, "region": "All"})
    months = month_index(params)
    data = {"YearMonth": months.astype(str)}
    for region in REGIONS:
        part = scoped[scoped["Region"] == region]
        if part.empty:
            counts = pd.Series(0, index=months)
        else:
            counts = part.groupby(part["OrderDate"].dt.to_period("M"))["OrderID"].size().reindex(months, fill_value=0)
        data[region] = counts.values
    return pd.DataFrame(data)


def aov_matrix(frame, params, filters=None) -> pd.DataFrame:
    chosen = active_filters(params, filters)
    scoped = in_scope(frame, params, {**chosen, "region": "All", "channel": "All"})
    rows = []
    for region in REGIONS:
        row = {"Region": region}
        for channel in CHANNELS:
            part = scoped[(scoped["Region"] == region) & (scoped["Channel"] == channel)]
            row[channel] = float(part["NetRevenue"].sum() / len(part)) if len(part) else None
        rows.append(row)
    return pd.DataFrame(rows)


def status_mix(frame, params, filters=None) -> pd.DataFrame:
    chosen = active_filters(params, filters)
    mask = dimension_mask(
        frame,
        params,
        region=chosen["region"],
        channel=chosen["channel"],
        category=chosen["category"],
        respect_cancelled=False,
    )
    scoped = frame.loc[mask]
    counts = scoped["Status"].value_counts()
    return pd.DataFrame({"Status": STATUSES, "Orders": [int(counts.get(status, 0)) for status in STATUSES]})


def top_orders(frame, params, filters=None, limit: int = 10) -> pd.DataFrame:
    scoped = in_scope(frame, params, filters).sort_values(["NetRevenue", "OrderID"], ascending=[False, True])
    columns = ["OrderID", "OrderDate", "Region", "Category", "Product", "Status", "NetRevenue", "Margin"]
    return scoped[columns].head(limit).reset_index(drop=True)


def build_insights(frame: pd.DataFrame, params: dict, summary: dict, cleaning_log: dict) -> str:
    cards = summary["kpis"]
    regions = summary["regions"]
    categories = summary["categories"]
    channels = summary["channels"]
    top_region = regions.sort_values("NetRevenue", ascending=False).iloc[0]
    weak_region = regions.sort_values("NetRevenue", ascending=True).iloc[0]
    margin_leader = regions.sort_values("GrossMargin", ascending=False).iloc[0]
    margin_laggard = regions.sort_values("GrossMargin", ascending=True).iloc[0]
    top_category = categories.sort_values("NetRevenue", ascending=False).iloc[0]
    return_leader = categories.sort_values("ReturnRate", ascending=False).iloc[0]
    top_channel = channels.sort_values("NetRevenue", ascending=False).iloc[0]
    best_month = summary["monthly"].sort_values("NetRevenue", ascending=False).iloc[0]
    slow_month = summary["monthly"].sort_values("NetRevenue", ascending=True).iloc[0]
    top = summary["top_orders"].iloc[0]
    reasons = cleaning_log["reason_counts"]
    adjustments = cleaning_log["adjustments"]

    def money(value: float) -> str:
        return f"${value:,.0f}"

    def pct(value: float | None) -> str:
        return "n/a" if value is None else f"{value * 100:.1f}%"

    lines = [
        "# Key insights",
        "",
        (
            f"Harbor & Co. recorded {money(cards['net_revenue'])} of in-scope net revenue "
            f"across {cards['orders']:,} orders from {params['analysis_start']} to {params['analysis_end']}. "
            f"Average order value was {money(cards['aov'] or 0)} and weighted gross margin was {pct(cards['gross_margin'])}."
        ),
        "",
        (
            f"{top_region['Name']} is the largest region at {money(top_region['NetRevenue'])} "
            f"({top_region['Orders']:,} orders). {weak_region['Name']} is the smallest at "
            f"{money(weak_region['NetRevenue'])}. {margin_leader['Name']} leads weighted margin at "
            f"{pct(margin_leader['GrossMargin'])}; {margin_laggard['Name']} is lowest at {pct(margin_laggard['GrossMargin'])}."
        ),
        "",
        (
            f"{top_category['Name']} contributes the most category revenue ({money(top_category['NetRevenue'])}). "
            f"{return_leader['Name']} has the highest return rate at {pct(return_leader['ReturnRate'])}. "
            f"Overall return rate is {pct(cards['return_rate'])} "
            f"({cards['returned_orders']:,} returns out of {cards['return_base']:,} non-cancelled orders)."
        ),
        "",
        (
            f"{top_channel['Name']} is the largest channel by net revenue ({money(top_channel['NetRevenue'])}). "
            f"The strongest month is {best_month['YearMonth']} at {money(best_month['NetRevenue'])}; "
            f"the softest month inside the window is {slow_month['YearMonth']} at {money(slow_month['NetRevenue'])}."
        ),
        "",
        (
            f"On-time fulfillment (shipped within {params['late_shipment_days']} days, cancelled orders excluded) "
            f"is {pct(cards['on_time_rate'])}. "
            f"{cards['high_value_orders']:,} orders meet or exceed the {params['high_value_percentile']:.0%} "
            f"net-revenue cutoff of {money(cards['high_value_cutoff'])}. "
            f"{cards['low_margin_orders']:,} in-scope orders sit below the {pct(params['low_margin_threshold'])} margin floor. "
            f"The largest single order is {top['OrderID']} ({top['Region']}, {top['Product']}) at {money(top['NetRevenue'])}."
        ),
        "",
        (
            f"The extract started with {cleaning_log['raw_rows']:,} rows and finished with "
            f"{cleaning_log['clean_rows']:,} analysis-ready orders. "
            f"Exact duplicate rows removed: {reasons.get('exact_duplicate', 0):,}. "
            f"Order-ID conflicts resolved by keeping the higher net revenue: {reasons.get('order_id_conflict', 0):,}. "
            f"Rows rejected for invalid region, date, quantity, or price: "
            f"{reasons.get('unknown_region', 0) + reasons.get('bad_date', 0) + reasons.get('zero_qty', 0) + reasons.get('bad_price', 0):,}. "
            f"Negative quantities corrected: {adjustments['negative_qty_corrected']:,}. "
            f"Blank segments imputed to Unknown: {adjustments['blank_segment_imputed']:,}. "
            f"Implausible ship days cleared: {adjustments['ship_days_cleared']:,}."
        ),
        "",
        "## Actions suggested by the numbers",
        "",
        (
            f"- Review fulfillment in the region with the weakest margin ({margin_laggard['Name']}) before adding promotional discount."
        ),
        (
            f"- {return_leader['Name']} returns are high enough to justify a quality or sizing check before the next peak season."
        ),
        (
            "- Hold cancelled orders out of sales KPIs (the workbook default). Turn the dashboard switch to Yes only when the question is about demand originally written, not demand kept."
        ),
        (
            f"- Treat orders at or above {money(cards['high_value_cutoff'])} as a named follow-up list; there are {cards['high_value_orders']:,} of them in the current window."
        ),
    ]
    return "\n".join(lines) + "\n"


def summarize(frame: pd.DataFrame, params: dict) -> dict:
    cutoff = high_value_cutoff(frame, params)
    return {
        "kpis": kpis(frame, params, cutoff=cutoff),
        "regions": region_summary(frame, params),
        "channels": channel_summary(frame, params),
        "categories": category_summary(frame, params),
        "monthly": monthly_trend(frame, params),
        "quarterly": quarterly_trend(frame, params),
        "matrix": category_region_matrix(frame, params),
        "orders_by_region": monthly_orders_by_region(frame, params),
        "aov": aov_matrix(frame, params),
        "status": status_mix(frame, params),
        "top_orders": top_orders(frame, params),
        "cutoff": cutoff,
    }
