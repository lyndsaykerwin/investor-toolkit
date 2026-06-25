# arr-to-bookings

An agent skill that turns customer-level revenue into an **investor-grade quarterly ACV bookings analysis**. Works in any agent harness that can run a skill.

Give it customer-level recurring revenue in whatever shape you have it:

- **Wide** — one row per customer, one column per month (the typical exported spreadsheet),
- **Long / "tidy"** — one row per customer per month (columns like `customer_id, month, mrr`), or
- **Transaction log** — many rows per customer that get aggregated up,

as an **Excel file or a CSV**, holding either MRR (monthly recurring revenue — what each customer pays per month) or ARR (annual recurring revenue — the same thing annualized). The skill returns an Excel deliverable: quarterly **new-logo vs. upsell ACV** (annual contract value), logo counts, the largest deal per period, year-over-year growth, and annual summaries. Every number is a live formula traceable back to the source data — no hardcoded values, full audit trail.

## Read this first — what the number means

> This is a bookings **estimate** for use when customer-level MRR/ARR is the only data available. It is **not valid for usage-based models**, and it is **not a substitute for a CRM-sourced bookings read** — where a CRM (the sales system that records actual signed deals) is available, that is the most accurate measure of bookings growth. The skill measures bookings as the month-over-month *change* in recurring revenue, which approximates — but is not identical to — deals booked in a sales system.

## What it's for / what it's NOT for

**For:**
- Estimating quarterly/annual bookings (new business vs. expansion) from customer-level recurring revenue
- Spotting "mega deals" — the largest new logo or upsell in each period

**Not for:**
- Usage-based / consumption revenue models
- Forecasting future bookings
- LTV/CAC (lifetime value vs. customer acquisition cost) analysis
- The retention corkscrew itself (use the sibling **retention-analysis** skill)

## How it measures bookings

Bookings are classified **one step at a time** — each month versus the immediately prior month — then **summed into the quarter** (and into the year). The first month of the dataset is the opening installed base and books nothing (there is no prior month to compare it to). This is **gross** bookings: a sign-and-expand inside one quarter is split into new-logo + upsell rather than lumped together, and a sign-and-churn inside one quarter is still captured. Every quarter reconciles: Beginning + New + Upsell − Downsell − Churn = Ending ARR, with an external check that must tie to $0.

## Dependencies

**Required:**
- Python 3.9 or newer
- `openpyxl` (the only non-stdlib Python package) — install with `pip install openpyxl`

**Required for the final recalc step:**
- LibreOffice — provides the `soffice` command used in headless mode (no GUI) to recompute and cache every formula, so the file opens showing values instead of empty formula cells.
  - macOS: `brew install --cask libreoffice`
  - Ubuntu/Debian: `sudo apt install libreoffice`
  - Windows: download installer from libreoffice.org

## Quickstart — run against a bundled fixture

```
cd skills/arr-to-bookings
python3 scripts/arr_to_bookings.py --source scripts/fixtures/sample_corkscrew.xlsx --out /tmp/bookings.xlsx
```

It prints a shape report, self-verifies (independent recompute + LibreOffice recalc; external check = 0), and writes a clickable link to the output. The `scripts/fixtures/quirks/` folder holds 11 messy-input fixtures (text dates, reversed dates, gaps, multi-row, long/tidy, quarterly-native, …) used as the regression suite.

## Using it inside an agent

Point the agent at a folder containing a real revenue workbook and ask something like:

> "Turn this into a quarterly bookings analysis" — or — "new business vs. expansion ACV by quarter"

The agent triggers the skill, auto-detects the file's shape, builds the deliverable, and self-verifies before returning it.

## What's in the repo

```
SKILL.md                              # the spec the agent follows (decisions + workflow)
reference/
  method.md                           # definitions, normalization details, known gaps
scripts/
  arr_to_bookings.py                  # build the Excel deliverable + self-verify
  normalize.py                        # detect & reshape any input into a clean grid
  fixtures/
    sample_corkscrew.xlsx             # synthetic sample input
    quirks/                           # 11 messy-input regression fixtures + MANIFEST
README.md
.gitignore
```
