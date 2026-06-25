# ARR-to-bookings — method & definitions

## Intake: the normalization front door (`normalize.py`)
Any messy file is first converted to ONE canonical grid (customers × continuous period axis,
clean numbers) so the bookings math never sees the mess. It handles, generally:
- **Locating data:** scores every sheet; finds the period-header row by *distinct* periods (so a
  `2021,2021,…` year super-header loses to the real month row); skips title/summary/pivot rows;
  finds the customer column anywhere (by cardinality, not position).
- **Dates:** real dates; formula dates (cached values, or recalculated via LibreOffice if a
  freshly-written file has no cache); text in any format (`Jan 21`, `2021-01`, `January 2021`,
  `Q1 2021`, `FY21`, year-only); reversed order (sorted); gaps (placed on a continuous axis);
  junk suffixes (`Jan-21 Or Prior`). Repeated date blocks (actuals + adjusted) are de-duped to
  the first block.
- **Orientation:** wide grid, long/tidy, or transaction log — long is detected when periods
  repeat down many rows (a date *column*) rather than across one header row.
- **Identity:** many rows per customer (per product/persona) are group-summed to one row;
  Total/Subtotal/section rows are dropped.
- **Values:** text-numbers (`9,900`, `$1,542`), accounting negatives `(500)`, blanks, `NA`.
- **Frequency & scale:** monthly/quarterly/annual from cadence; MRR/ARR/annual-revenue from
  magnitude. Trailing all-zero (forecast) columns are trimmed; `--actuals-through` forces a cutoff.

It **refuses with a specific reason** when bookings are impossible: annual-only granularity, a
current-state snapshot (one row per customer, no history), or a sheet too large to model safely
(>60k rows / >5k customers — pre-aggregate first).

## What it produces (consistent output structure)
A clean output workbook (the source file on disk is never modified), ordered:
1. **Bookings analysis** (Quarterly or Annual) — new/upsell ACV, counts, largest deal, YoY,
   annual block, reconciliation.
2. **Normalized** helper — **only when a reshape is needed.** Cells are simple formulas back to
   Raw Data: `SUMIF` (multi-row→one-row), `SUMIFS` (long→wide), `=MAX(Raw Data,0)` (negative
   policy = zero), or cleaned values when the source values are non-numeric text Excel can't sum.
   A clean wide numeric source needs **no** helper — the analysis references Raw Data directly.
3. **Bookings Detail** (visible) — per-customer snapshots `=ROUND(<ref>×factor,2)` and the
   new/upsell/down/churn classification.
4. **Raw Data** — the **final tab, a verbatim copy of the source data, untouched.**

Every figure traces to Raw Data through auditable formulas; the reconciliation ties the rolled-up
bookings to the customer snapshots (which are themselves simple links to Raw Data).

- **Quarterly Bookings** (visible) — by quarter, then an annual block to the right:
  - NEW BUSINESS: New-logo ACV ($), YoY %, New logos (#), Largest new deal ($) + customer
  - EXPANSION: Upsell ACV ($), YoY %, Accounts upsold (#), Largest upsell ($) + customer
  - TOTAL BOOKINGS: Total ACV ($), YoY %, accounts booked (#)
  - RECONCILIATION: Beginning + New + Upsell + Downsell + Churn = Ending ARR; External check = 0
  - An **ARR factor** input cell (B2, blue = hardcoded) the whole model references.
- **Bookings Detail** (visible) — one row per customer; ARR snapshots at each
  period-end (`=ROUND(<ref> × ARR factor, 2)`), then the **per-step** classification columns
  (`mNew/mUp/mDn/mCh` = each month vs the immediately prior month), and finally the quarter
  rollup columns, each `=SUM(...)` of the steps that quarter owns.

## Core definitions
ARR per customer at a period-end = that month-end MRR × ARR factor (12 for monthly MRR).
Bookings are classified **one step at a time** — each month vs the immediately prior month —
then **summed into the quarter** (and into the year / apples-to-apples stub). This is a single,
uniform mechanism: every comparison is a one-step delta, so a quarter is just the sum of its
months and needs no special "compare quarter-end to quarter-end" logic.

Per step, comparing each customer's ARR to the **prior month**:

| case | condition | booking |
|---|---|---|
| New-logo | prev ≈ 0 and curr > 0 | New ACV = curr |
| Upsell | prev > 0 and curr > prev | Upsell ACV = curr − prev |
| Downsell | prev > 0 and 0 < curr < prev | (reconciliation only) |
| Churn | prev > 0 and curr ≈ 0 | (reconciliation only) |

Quarter New ACV = Σ (monthly New over the quarter's months); Quarter Upsell ACV = Σ (monthly
Upsell). Bookings = New + Upsell. Because this is **gross** (each month's rise is booked, not
netted against a later dip), it captures activity an end-to-end quarter snapshot would miss —
e.g. a customer that signs in month 1 and churns in month 3 still books its month-1 new-logo,
and a customer that signs and expands in the same quarter is split into new-logo + upsell rather
than lumped entirely into new-logo.

## Three things that make it correct
1. **Left-censoring.** Only the **very first month of the dataset** is the opening installed base
   (no prior month to compare against, so it books nothing) — customers already live in month 1
   signed before the data starts. Every period after that, including the first quarter, is just
   the sum of its one-month steps. (A customer first appearing in a *later* month is a genuine
   new-logo in that month.)
2. **Epsilon ($0.01).** Sub-cent float noise (e.g. MRR×12 rounding) must not register as a
   booking. All four cases use a 0.01 threshold, so a $0.00 change is never an "upsell."
3. **Reconcile to the existing-customer subtotal, not the grand total.** Many corkscrews carry
   a "Pipeline / unidentified" line that isn't a customer. Reconcile against the customer
   subtotal row (`--recon-row`) so the external check ties to exactly 0.

## Periods & the partial-year stub
Quarter-ends are months 3/6/9/12 up to the actuals cutoff. If the last actual month is **not**
a quarter-end, a trailing partial quarter is added (label suffixed `*`). When the final year is
partial, two amber **apples-to-apples** columns are added — `(Y-1) Jan–<mon>` and `Y Jan–<mon>`,
each measuring Dec→same-month — so the stub year is compared to the identical window a year
earlier, with YoY on the later column.

## Counts: a subtlety
"New logos (#)" per quarter = customers with a positive new-logo rollup that quarter; annual =
sum of quarterly counts = **distinct** new logos (a customer is new only once). "Accounts
upsold (#)" annual = sum of quarterly counts = **upsell events** (a customer can expand in
multiple quarters) — labeled as events, not accounts, on purpose.

## Verification (built into the script)
- Independent Python recompute of every figure from the raw cells.
- LibreOffice headless recalc, then assert cached cell values == recompute; external check < $1;
  no `#REF!/#VALUE!/#NAME?/#DIV0!` cells. Exits non-zero on any failure.

## Known gaps / not handled
- **Gross-bookings sensitivity to month-to-month noise:** because each month's rise is booked
  (not netted), genuinely non-deal wobble — a one-month credit, usage-based billing, a data
  artifact — registers as gross upsell even if it reverses next month (the reversal shows as
  downsell in the reconciliation). For contractual SaaS MRR this is negligible; the per-month
  `mUp`/`mDn` columns in the Detail tab make any such wobble fully auditable. No smoothing
  threshold is applied (it would bury numbers behind a hidden rule).
- **Reactivation = new-logo:** a customer that fully churns to $0 and later returns books a *new*
  new-logo on return (it's a genuine re-acquisition). Intentional, but worth knowing if you treat
  logos as strictly once-ever.
- **Annual-only sources** → "Annual Bookings" mode: each year is a period, YoY = vs prior year
  (lookback 1), factor 1 (annual amounts), no quarter-of-year block or stubs. (Quarterly bookings
  still can't be manufactured from annual totals — this is a yearly new-vs-upsell view instead.)
- **Very large sources** (>60k rows / >5k customers) are refused rather than materialize a giant
  per-customer helper; pre-aggregate or sample first.
- **Snapshots** (one row per customer, current MRR + a single date) have no history → refused.
- **Bookings = gross new + gross upsell**; contraction/churn appear only in the reconciliation.
- **Negatives** (credits/refunds): detected and reported; choose `--neg-policy zero|contraction|
  leave`. With `leave`, the reconciliation will surface them as unexplained movement (by design).
- LibreOffice absent → formulas are written but not numerically confirmed (warns); also needed to
  recalc un-cached formula headers.
