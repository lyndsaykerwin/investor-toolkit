---
name: pipeline-analysis
description: >-
  Turn a deal-level pipeline export into a formula-clean Excel pipeline analysis:
  total & weighted (expected) pipeline, coverage vs target, and the pipeline
  split by stage × product / type / owner, plus deal-size concentration and
  aging. When historical CLOSED deals are also supplied, adds conversion context
  — win rate, competitive win rate, sales-cycle length, new-vs-upsell mix. Flags
  what's missing with a ready-to-send data request. Triggers on "analyze this
  pipeline", "pipeline coverage", "pipeline by stage", "open/weighted pipeline",
  "sales pipeline", "CRM export". NOT for win/loss *reasons* analysis (why deals
  are won/lost, competitor/loss-reason coding) — that's a separate study.
---

# Pipeline Analysis

Produce a formula-clean Excel pipeline analysis from one deal-level export, and
name what the data can't support.

**Core principle — degrade, never break.** The skill NEVER errors out because a
field is missing: it runs every analysis the available data supports and lists
the rest under "What this data can't tell you." Conversely, when a file contains
**every** supported field, it runs **every** analysis (stage×product/type/owner
cross-tabs, weighted pipeline, coverage, aging, concentration, bookings-by-period,
and full conversion context) with zero blocked items. The only stop is a file
with no usable deal data at all — and even then it returns a clean data request,
not a crash.

## Steps
1. **Map columns.** Locate the header row; map to canonical fields via
   `references/metrics-and-definitions.md`. Minimum to run: a deal $ value (ARR/
   amount) plus a stage/status, probability, or close date. If absent, STOP and
   output the data request (`references/data-request-template.md`).
2. **Run the engine:** `python3 scripts/analyze_deals.py --input <file>
   --output <Company>_Pipeline_v1.xlsx [--target <bookings_goal>] [--sheet <name>]`.
   It builds an **Analysis tab first**, then preserves the original source
   sheet(s) byte-for-byte. Multi-sheet workbooks are combined.
3. **Report from the Analysis tab:** headline pipeline metrics, the stage cross-
   tabs, blocked metrics, and the follow-up data request.

## What it computes
- **Pipeline:** total open $, weighted/expected $ (needs probability), coverage
  vs target (×), Stage × Product / Type / Owner cross-tabs, top-deal
  concentration, and open-pipeline aging (needs create date or an age column).
- **Bookings by period:** Won ARR by year (needs close dates on won deals).
- **Conversion context (only if historical closed deals are included):** win rate
  (shown two ways), competitive win rate, sales-cycle length, % new-vs-upsell.

Each is emitted only when its inputs exist; otherwise it's listed as blocked.

## Robustness (handled automatically)
- Header row need not be row 1 (auto-detected under title/filter blocks).
- Subtotal/Total/Count rows are stripped, not counted as deals.
- Multi-sheet workbooks (e.g. one sheet per year) are combined; only deal-like
  sheets are included. Pass `--sheet` to force one.
- Anonymized/garbled headers → refuses with the data request, never garbage.
- **Headerless exports** (data present, no column names) → pass
  `--map '{"arr":"G","outcome":"N","account":"D",...}'` (column letters) to map
  positionally; rows with a numeric ARR are treated as data.

## Guardrails
- Coverage = open pipeline ÷ the period's bookings target (a multiple, e.g. 3×).
- Weighted pipeline needs an explicit probability column — never infer from stage.
- **Win/loss *reasons*** (loss-reason coding, competitor/incumbent diagnosis) are
  OUT of scope — different data; route to a separate win-loss-reasons study.
- Output is **Excel only**. Never put real client data in this skill folder —
  write deliverables to the caller's working directory.
