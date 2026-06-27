# pipeline-analysis

Turn a deal-level pipeline export (open opportunities, closed history, or both)
into a banker-formatted Excel pipeline analysis. Built to **degrade, never break**:
it runs every analysis the data supports and clearly lists what's missing — and
when a file has every supported field, it runs every analysis with nothing blocked.

## What it produces
A single `.xlsx` with an **Analysis tab first** and the **original source sheet(s)
preserved byte-for-byte** behind it. The Analysis tab includes, where the data
allows:

- KPI strip: open pipeline $, deal count, (and win rates if closed history present)
- Pipeline by **Stage × Product**, **Stage × Type**, **Owner × Stage** cross-tabs
- **Top-deal concentration** (top 10, % of pool, top-5 share)
- **Open-pipeline aging** (age buckets + stale >90 days) — needs create date or age
- **Bookings by year** (Won ARR) — needs close dates on won deals
- **Conversion context** (win rate two ways, competitive win rate, sales-cycle
  length, % new-vs-upsell) — only if historical closed deals are included
- "What this data can't tell you" + a ready-to-send data request

Out of scope: win/loss **reasons** (loss-reason coding, competitor/incumbent
diagnosis) — a separate study with different data.

## Run
```
python3 scripts/analyze_deals.py --input <file.xlsx|csv> --output <out.xlsx> \
    [--target <bookings_goal>] [--sheet <name>] [--header-row <n>] [--map <json>]
```
- `--target` enables pipeline coverage (open ÷ target, as a multiple).
- Multi-sheet workbooks (e.g. one tab per year) are combined automatically.
- `--map '{"arr":"G","outcome":"N","account":"D",...}'` handles **headerless**
  exports by mapping canonical fields to column letters.

## Inputs (auto-mapped from common header names)
Minimum: a deal $ value (ARR/amount) **plus** a stage, probability, or close date.
Recommended for the full analysis: probability, create date, close date, deal
type, product, owner, and a bookings target. See
`references/data-request-template.md` for the field list to request from a company.

Requires Python 3 + `openpyxl`. Fixtures in `fixtures/` are fully synthetic.
