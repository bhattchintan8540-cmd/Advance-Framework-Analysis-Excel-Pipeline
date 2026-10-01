# Harbor & Co. — Advanced Data Analysis (Excel Pipeline)

An end-to-end sales analysis you can rerun. A dirty order extract is cleaned, checked, and loaded into an Excel workbook whose KPI cards and ten charts are formulas. Python rebuilds the extract, proves the cleaning rules, and writes the dashboard screenshots.

The workbook is synthetic. Seed `42` rebuilds the same file every time. It is not a client ledger.

## Objective

Give a merchandising lead one place to answer:

- How much net revenue is in the window, and what is a typical order worth?
- Which region, category, and channel carry the business?
- Where are returns, late shipments, and thin margins concentrated?
- What did the cleaning step throw away, and why?

## Tools

| Piece | Role |
| --- | --- |
| Python 3, pandas, numpy | Generate the dirty extract, clean it, and calculate the KPIs used as the test oracle |
| openpyxl | Write the `.xlsx` model: tables, named ranges, data-validation filters, conditional formatting, and charts |
| Excel formulas | `SUMIFS`, `COUNTIFS`, `AVERAGEIFS`, and row-level revenue formulas so the dashboard recalculates when filters change |
| matplotlib | Render the two dashboard screenshots from the same KPI definitions |
| pytest | Lock the cleaning rules, the KPI math, and the workbook structure |

Excel desktop (Microsoft 365) is what recalculates the workbook. The pipeline itself does not need Excel installed.

## Process

1. **Extract.** `pipeline/generate.py` writes 1,600 synthetic orders plus planted defects: exact duplicate rows, mixed date formats (`01/15/2024`, `15-Aug-24`, `2024/02/29`, `12.12.2024`), region aliases (`Nort`, ` west `), status typos (`cancled`, `Compleated`), currency text (`$1,250.50`), spelled numbers (`Five`, `TwoHundredForty`), a negative quantity, a zero quantity, a negative price, a blank segment, `N/A` and negative ship days, an unknown region, an unparseable date, and two rows that share an order id.
2. **Clean.** `pipeline/clean.py` applies the alias maps, parses dates and numbers, drops exact duplicates, fixes negative quantities, clears impossible ship days, imputes a blank segment to `Unknown`, and keeps the higher net-revenue row when an order id collides. Every rejected row is written out with a reason. Raw count = clean count + rejected count.
3. **Model.** Surviving orders land on `Cleaned_Data`. Gross revenue, net revenue, cost, margin, year, quarter, year-month, late flag, high-value flag, and return flag are Excel formulas, not pasted values.
4. **Analyze and present.** `Chart_Data` and `Dashboard` aggregate with `SUMIFS` / `COUNTIFS` / `AVERAGEIFS` against named filter cells. Ten charts read those formula ranges.
5. **Test.** `pytest` and `outputs/qa_report.json` check identities (net revenue, cost, margin), edge-case outcomes, parameter behavior, and workbook structure. The `QA_Tests` sheet repeats the headline numbers as Excel formulas so a desktop recalculation can confirm them.

## How to run

```bash
python3 -m pip install -r requirements.txt
python3 -m pipeline
python3 -m pytest -q
```

`python3 -m pipeline` rewrites the dataset, the workbook, the screenshots, and `outputs/insights.md`.

Open `excel/Harbor_and_Co_Excel_Pipeline.xlsx` in Excel. On **Dashboard**, change Region, Channel, Category, Include cancelled, Start, or End. Excel recalculates the cards and charts. Late-shipment days and the low-margin floor are on **Parameters**.

## Parameters

Defaults live in `config/parameters.json`.

| Parameter | Default | Effect |
| --- | --- | --- |
| `analysis_start` / `analysis_end` | 2024-01-01 to 2025-09-30 | Orders outside the window stay on `Cleaned_Data` and drop out of every aggregate |
| `include_cancelled` | false | Sales KPIs skip cancelled orders. Return rate and on-time rate always skip them |
| `dashboard_filters` | All / All / All | Region, channel, and category controls on the dashboard |
| `late_shipment_days` | 7 | A shipped order is late when ship days are greater than this |
| `low_margin_threshold` | 0.15 | Orders whose own margin is below 15% |
| `high_value_percentile` | 0.90 | Percentile used to set the high-value cutoff at pipeline run time |
| `random_seed` | 42 | Rebuilds the same extract |
| `exact_duplicate_copies` | 25 | Exact copies appended before the planted edge rows |

### Which filters each view honors

KPI cards honor every control.

Comparison charts drop only the filter that would erase the comparison: the region chart ignores the region dropdown, the channel chart ignores the channel dropdown, the category-by-region chart ignores region and category, and average order value by region and channel ignores region and channel. Date window and the cancelled switch still apply. Status mix always shows Completed, Returned, and Cancelled.

## Key insights

From the default window, cancelled orders excluded:

- **$2,220,660** net revenue across **1,558** orders. Average order value **$1,425**. Weighted gross margin **29.4%**.
- **West** is the largest region ($732,437, 487 orders). **East** is the smallest ($425,069) and the weakest on weighted margin (26.1%). **South** has the best weighted margin (31.3%).
- **Furniture** is most of the revenue ($1,698,590). **Textiles** has the highest return rate (13.5%). Overall returns are 7.1% (111 of 1,558 non-cancelled orders).
- **Wholesale** is the largest channel by net revenue ($1,286,845) because those tickets are large, not because there are more of them. Store still has the most orders.
- The strongest month is **2024-11** ($220,964). The softest month in the window is **2024-02** ($40,088).
- On-time fulfillment is **75.5%**. **156** orders are at or above the 90th percentile cutoff of **$3,065**. **21** orders are under the 15% margin floor. The largest order is **HC-100483** (West, Linen Sofa) at $29,394.
- Cleaning started with **1,649** rows and kept **1,618**. Removed: 26 exact duplicates, 1 order-id conflict, and 4 invalid rows (unknown region, bad date, zero quantity, bad price). One negative quantity was flipped. 77 blank segments became `Unknown`. One impossible ship-day value was cleared.

The same narrative is written to `outputs/insights.md` and the **Insights** sheet on every run.

### What to do with that

- Look at East margin before adding another promotion there.
- Textiles returns are high enough to check quality or sizing before the next November peak.
- Leave cancelled orders out of the sales cards unless the question is about demand that was written, not demand that was kept.
- The 156 orders at or above $3,065 are the follow-up list.

## Workbook map

| Sheet | What it is |
| --- | --- |
| Cover | Objective, pipeline, and the headline result |
| Dashboard | Eight KPI cards, six filter controls, ten charts |
| Parameters | Late days, margin floor, high-value cutoff, and the limits of each one |
| Raw_Data | Dirty extract as received |
| Lookup_Maps | Alias tables the cleaner applies |
| Cleaned_Data | Analysis table. Money and flags are formulas |
| Chart_Data | Formula ranges the charts read |
| Analysis | Regional summary (live) and a top-10 snapshot |
| Insights | Narrative from the latest run |
| QA_Tests | Expected value vs Excel formula, with PASS/FAIL |
| Rejected_Rows | Rows that did not survive, with a reason |
| Cleaning_Log | Counts from the transform |
| Data_Dictionary | Field definitions |
| Lists | Hidden lists behind the dropdowns |

### Dashboard charts

1. Net revenue by region
2. Orders by channel
3. Monthly net revenue
4. Net revenue by category and region (stacked)
5. Unweighted average margin by region
6. Quarterly net revenue
7. Order status mix
8. Return rate by category
9. Monthly orders by region
10. Average order value by region and channel

## Project files

```
config/parameters.json          model parameters
data/raw/harbor_sales_raw.csv   dirty extract
data/processed/                 clean orders, rejected rows, cleaning log
excel/Harbor_and_Co_Excel_Pipeline.xlsx
dashboards/                     screenshots of the default view
outputs/insights.md             narrative
outputs/qa_report.json          machine-readable test record
outputs/summary.json            headline numbers
pipeline/                       generate, clean, metrics, workbook, charts
tests/test_pipeline.py          cleaning, KPI, and workbook tests
```

## Definitions that are easy to mix up

- **Net revenue** = quantity × unit price × (1 − discount).
- **Weighted gross margin** = (net revenue − cost) / net revenue. This is the KPI card.
- **Unweighted margin** = the average of order-level margins. This is the regional margin chart. A few large wholesale orders can move the weighted figure without moving the average.
- **Return rate** = returned orders / (completed + returned). Cancelled orders are never in that denominator.
- **On-time** = shipped within `late_shipment_days`, among non-cancelled orders that have a ship-day value.
- **High-value cutoff** = the linear 90th percentile of in-scope net revenue, computed when the pipeline runs and stored in `Parameters!C8`. Dashboard dropdowns do not rebuild it. Change the dates or the percentile in `config/parameters.json` and rerun if you want a new cutoff.

## Limitations

- The extract is synthetic. Do not treat the dollars as a real business result.
- Spelled-out numbers, currency text, and the alias maps are resolved in Python, then the cleaned fields are loaded into Excel. The maps are visible on `Lookup_Maps`. Excel formulas take over from gross revenue onward.
- The high-value cutoff is a stored number, not `PERCENTILE.INC` of whatever the dropdowns currently show.
- The top-10 list on `Analysis` is a snapshot of the default filters from the last pipeline run. The regional summary next to it is live.
- Interactive PivotTable slicers are not embedded. Dropdowns plus `SUMIFS` recalculate without a pivot cache, including in tools that do not refresh Power Pivot. That is a deliberate tradeoff.
- `QA_Tests` shows PASS/FAIL after Excel calculates. The Python tests already compare the same definitions without Excel.

## Debugging

- A rejected row you did not expect is on `Rejected_Rows` and in `data/processed/rejected_rows.csv`, with `Reason`.
- `outputs/qa_report.json` lists every Python check from the last run. This run passed 28 checks.
- If a dashboard card looks wrong after a filter change, confirm **Include cancelled**, **Start**, and **End**, then check the matching formula on `Chart_Data`.
- If `python3 -m pipeline` fails the row balance, cleaning dropped or double-counted a row. The command raises before it writes the workbook.
- Planted order ids (`EDGE-WORD`, `EDGE-CONF`, `EDGE-OLD`, and the others in `pipeline/generate.py`) are the fixtures the tests assert against. Do not rename them without updating `tests/test_pipeline.py`.
