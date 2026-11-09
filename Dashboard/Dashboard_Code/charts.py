"""Render dashboard screenshots from the same KPI definitions the workbook uses."""

from __future__ import annotations

from pathlib import Path

import matplotlib

matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib.patches import FancyBboxPatch

INK = "#1C2430"
MUTED = "#5E6A75"
NAVY = "#1E3A5F"
TEAL = "#1F6F64"
GOLD = "#B0893E"
CORAL = "#A33B32"
CREAM = "#F3F0E8"
WHITE = "#FFFFFF"
SERIES = [NAVY, TEAL, GOLD, CORAL, "#6E7F8D"]


def _money(value: float) -> str:
    return f"${value:,.0f}"


def _pct(value: float | None) -> str:
    return "n/a" if value is None else f"{value * 100:.1f}%"


def _style_ax(ax, title: str):
    ax.set_title(title, loc="left", fontsize=12, color=INK, pad=10, fontweight="bold")
    ax.title.set_clip_on(False)
    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)
    ax.spines["left"].set_color("#D5D0C6")
    ax.spines["bottom"].set_color("#D5D0C6")
    ax.tick_params(colors=MUTED, labelsize=8)
    ax.set_facecolor(WHITE)
    ax.grid(axis="y", color="#E6E1D6", linewidth=0.6)
    ax.set_axisbelow(True)


def _card(ax, label, value, tone):
    ax.set_xlim(0, 1)
    ax.set_ylim(0, 1)
    ax.axis("off")
    ax.add_patch(
        FancyBboxPatch(
            (0.02, 0.08),
            0.96,
            0.84,
            boxstyle="round,pad=0.02,rounding_size=0.08",
            facecolor=WHITE,
            edgecolor="#E4DFD4",
            linewidth=0.8,
            mutation_aspect=0.6,
        )
    )
    ax.add_patch(plt.Rectangle((0.02, 0.78), 0.96, 0.06, color=tone, zorder=2))
    ax.text(0.08, 0.62, label.upper(), color=MUTED, fontsize=8, fontweight="bold")
    ax.text(0.08, 0.28, value, color=INK, fontsize=16, fontweight="bold")


def render_dashboards(summary: dict, params: dict, cleaning_log: dict, output_dir: Path) -> list[Path]:
    output_dir.mkdir(parents=True, exist_ok=True)
    executive = output_dir / "01_executive_dashboard.png"
    regional = output_dir / "02_regional_category_dashboard.png"
    _executive(summary, params, executive)
    _regional(summary, params, cleaning_log, regional)
    return [executive, regional]


def _executive(summary, params, path: Path):
    cards = summary["kpis"]
    monthly = summary["monthly"]
    regions = summary["regions"]
    channels = summary["channels"]
    status = summary["status"]
    fig = plt.figure(figsize=(16, 9), dpi=140, facecolor=CREAM)
    fig.suptitle("Harbor & Co.  ·  Sales performance", fontsize=18, color=INK, x=0.04, ha="left", fontweight="bold")
    fig.text(
        0.04,
        0.935,
        f"{params['analysis_start']} to {params['analysis_end']}   ·   cancelled orders excluded   ·   filters: All regions, channels, categories",
        color=MUTED,
        fontsize=9,
    )
    grid = fig.add_gridspec(4, 8, left=0.04, right=0.98, top=0.88, bottom=0.08, hspace=0.55, wspace=0.45)
    values = [
        ("Net revenue", _money(cards["net_revenue"]), TEAL),
        ("Orders", f"{cards['orders']:,}", NAVY),
        ("Average order value", _money(cards["aov"] or 0), TEAL),
        ("Gross margin", _pct(cards["gross_margin"]), GOLD),
        ("Return rate", _pct(cards["return_rate"]), CORAL),
        ("On-time fulfillment", _pct(cards["on_time_rate"]), TEAL),
        ("High-value orders", f"{cards['high_value_orders']:,}", GOLD),
        ("Low-margin orders", f"{cards['low_margin_orders']:,}", CORAL),
    ]
    for index, (label, value, tone) in enumerate(values):
        _card(fig.add_subplot(grid[0, index]), label, value, tone)

    ax = fig.add_subplot(grid[1:3, 0:5])
    ax.plot(monthly["YearMonth"], monthly["NetRevenue"], color=NAVY, linewidth=2.2)
    ax.fill_between(range(len(monthly)), monthly["NetRevenue"], color=NAVY, alpha=0.08)
    ax.set_xticks(range(0, len(monthly), 2))
    ax.set_xticklabels(monthly["YearMonth"].iloc[::2], rotation=45, ha="right")
    _style_ax(ax, "Monthly net revenue")
    ax.yaxis.set_major_formatter(plt.FuncFormatter(lambda value, _pos: f"${value/1000:.0f}k"))

    ax = fig.add_subplot(grid[1:3, 5:8])
    colors = [NAVY, TEAL, GOLD, CORAL]
    ax.barh(regions["Name"][::-1], regions["NetRevenue"][::-1], color=colors[::-1])
    _style_ax(ax, "Revenue by region")
    ax.xaxis.set_major_formatter(plt.FuncFormatter(lambda value, _pos: f"${value/1000:.0f}k"))

    ax = fig.add_subplot(grid[3, 0:3])
    ax.bar(channels["Name"], channels["Orders"], color=[NAVY, TEAL, GOLD])
    _style_ax(ax, "Orders by channel")

    ax = fig.add_subplot(grid[3, 3:5])
    mix_colors = [TEAL, GOLD, CORAL]
    wedges, _texts = ax.pie(
        status["Orders"],
        colors=mix_colors,
        startangle=90,
        wedgeprops={"width": 0.42, "edgecolor": WHITE},
    )
    ax.legend(wedges, status["Status"], loc="center left", bbox_to_anchor=(0.95, 0.5), frameon=False, fontsize=8)
    ax.set_title("Order status mix", loc="left", fontsize=11, color=INK, pad=8)
    ax.set_facecolor(CREAM)

    ax = fig.add_subplot(grid[3, 5:8])
    quarters = summary["quarterly"]
    ax.fill_between(range(len(quarters)), quarters["NetRevenue"], color=TEAL, alpha=0.25)
    ax.plot(range(len(quarters)), quarters["NetRevenue"], color=TEAL, linewidth=2)
    ax.set_xticks(range(len(quarters)))
    ax.set_xticklabels(quarters["Quarter"], rotation=45, ha="right")
    _style_ax(ax, "Quarterly net revenue")
    ax.yaxis.set_major_formatter(plt.FuncFormatter(lambda value, _pos: f"${value/1000:.0f}k"))

    fig.text(0.04, 0.02, "Synthetic extract, seed 42. Screenshots match the default Excel dashboard before any filter is changed.", color=MUTED, fontsize=8)
    fig.savefig(path, dpi=140, bbox_inches="tight", pad_inches=0.35)
    plt.close(fig)


def _regional(summary, params, cleaning_log, path: Path):
    categories = summary["categories"]
    matrix = summary["matrix"]
    orders = summary["orders_by_region"]
    regions = summary["regions"]
    fig = plt.figure(figsize=(16, 9), dpi=140, facecolor=CREAM)
    fig.suptitle("Harbor & Co.  ·  Region, category, and data quality", fontsize=18, color=INK, x=0.04, ha="left", fontweight="bold")
    fig.text(0.04, 0.935, "Same default window as the Excel workbook. Margin bars are unweighted averages of order margins.", color=MUTED, fontsize=9)
    grid = fig.add_gridspec(3, 4, left=0.05, right=0.98, top=0.88, bottom=0.08, hspace=0.5, wspace=0.4)

    ax = fig.add_subplot(grid[0, 0:2])
    bottom = None
    for index, region in enumerate(matrix.columns):
        values = matrix[region].to_numpy()
        ax.bar(matrix.index, values, bottom=bottom, color=SERIES[index], label=region)
        bottom = values if bottom is None else bottom + values
    _style_ax(ax, "Net revenue by category and region")
    ax.legend(frameon=False, ncol=4, fontsize=8)
    ax.yaxis.set_major_formatter(plt.FuncFormatter(lambda value, _pos: f"${value/1000:.0f}k"))
    ax.tick_params(axis="x", labelrotation=0)

    ax = fig.add_subplot(grid[0, 2])
    ax.barh(categories["Name"][::-1], (categories["ReturnRate"][::-1] * 100), color=CORAL)
    _style_ax(ax, "Return rate by category")
    ax.set_xlabel("%", color=MUTED, fontsize=8)

    ax = fig.add_subplot(grid[0, 3])
    margins = regions["AvgMargin"].fillna(0) * 100
    ax.bar(regions["Name"], margins, color=TEAL)
    ax.axhline(params["low_margin_threshold"] * 100, color=CORAL, linewidth=1, linestyle="--")
    _style_ax(ax, "Unweighted margin (%)")

    ax = fig.add_subplot(grid[1, 0:3])
    for index, region in enumerate(["North", "South", "East", "West"]):
        ax.plot(orders["YearMonth"], orders[region], color=SERIES[index], linewidth=1.8, label=region)
    ax.set_xticks(range(0, len(orders), 2))
    ax.set_xticklabels(orders["YearMonth"].iloc[::2], rotation=45, ha="right")
    _style_ax(ax, "Monthly orders by region")
    ax.legend(frameon=False, ncol=4, fontsize=8)

    ax = fig.add_subplot(grid[1, 3])
    aov = summary["aov"].set_index("Region")
    aov.plot(kind="bar", ax=ax, color=[NAVY, TEAL, GOLD], width=0.8, legend=True)
    _style_ax(ax, "AOV by region and channel")
    ax.legend(frameon=False, fontsize=7)
    ax.tick_params(axis="x", labelrotation=0)
    ax.yaxis.set_major_formatter(plt.FuncFormatter(lambda value, _pos: f"${value:,.0f}"))

    ax = fig.add_subplot(grid[2, 0:2])
    reason_labels = {
        "exact_duplicate": "Duplicate row",
        "unknown_region": "Unknown region",
        "bad_date": "Bad date",
        "zero_qty": "Zero quantity",
        "bad_price": "Bad price",
        "order_id_conflict": "Order ID conflict",
        "bad_qty": "Bad quantity",
        "bad_cost": "Bad cost",
        "bad_discount": "Bad discount",
        "unknown_channel": "Unknown channel",
        "unknown_category": "Unknown category",
        "unknown_status": "Unknown status",
    }
    reasons = cleaning_log["reason_counts"]
    labels = [reason_labels.get(key, key) for key in reasons]
    values = [reasons[key] for key in reasons]
    ax.barh(labels[::-1], values[::-1], color=NAVY)
    _style_ax(ax, "Rows removed, by reason")

    ax = fig.add_subplot(grid[2, 2:4])
    ax.axis("off")
    ax.set_xlim(0, 1)
    ax.set_ylim(0, 1)
    lines = [
        f"Raw rows          {cleaning_log['raw_rows']:,}",
        f"Clean orders      {cleaning_log['clean_rows']:,}",
        f"Rejected rows     {cleaning_log['rejected_rows']:,}",
        f"Duplicates        {cleaning_log['exact_duplicates']:,}",
        f"Qty sign fixes    {cleaning_log['adjustments']['negative_qty_corrected']:,}",
        f"Blank segments    {cleaning_log['adjustments']['blank_segment_imputed']:,}",
        f"Ship days cleared {cleaning_log['adjustments']['ship_days_cleared']:,}",
        f"High-value cutoff {_money(summary['cutoff'])}",
    ]
    ax.text(0.02, 0.95, "Pipeline quality", fontsize=12, color=INK, fontweight="bold", va="top")
    ax.text(0.02, 0.78, "\n".join(lines), fontsize=11, color=INK, va="top", family="monospace")
    fig.text(0.05, 0.02, "Rejected rows are kept in data/processed/rejected_rows.csv and on the Rejected_Rows sheet.", color=MUTED, fontsize=8)
    fig.savefig(path, dpi=140, bbox_inches="tight", pad_inches=0.35)
    plt.close(fig)
