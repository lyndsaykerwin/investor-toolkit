---
name: ARR-to-bookings
description: Turn ANY customer-revenue file (ARR/retention corkscrew, customer×month grid, transaction log, or tidy CSV) into a quarterly ACV bookings analysis in Excel — new-logo vs upsell ACV, logo counts, largest deal per period, YoY growth, and annual summaries. Use when a user has customer-level MRR/ARR/revenue and wants bookings, new-business vs expansion, ACV by quarter or year, growth rates, or to spot "mega deals."
---

> **Use as a bookings *estimate* when customer-level MRR/ARR is the only data available — not for
> usage-based models, and not a substitute for a CRM-sourced bookings read, which (where available)
> is the most accurate measure of bookings growth.**

Goal: produce a clean, consistent workbook. The source file on disk is never touched; the
**output** is structured as:
- **Bookings analysis** tab(s) first (Quarterly or Annual), all driven by simple formulas.
- a **Normalized** helper tab *only when a reshape is needed* (long→wide via `SUMIFS`,
  multi-row→one-row via `SUMIF`, or messy text → cleaned values); a clean wide source needs no
  helper and links straight to Raw Data.
- a **Bookings Detail** working tab (visible, so the classification is auditable).
- **Raw Data** as the **final tab — a verbatim copy of the source data, zero changes.**

Every number traces back to Raw Data through auditable formulas (`=Raw Data!cell`, `SUMIF`,
`SUMIFS`, `ROUND`); no complex unreadable formulas, no baked-in numbers where a formula works.

## Workflow

1. **Just point it at the file.** `scripts/arr_to_bookings.py --source FILE --out OUT.xlsx`.
   A normalization front door (`normalize.py`) auto-detects the sheet, orientation (wide /
   long / transaction), date headers (real, formula, or text — any format, even reversed or
   gapped), frequency (monthly/quarterly), customer column, multi-row-per-customer aggregation,
   text-numbers, and MRR-vs-ARR scale. No manual row/column flags needed.

2. **Read the shape report it prints**, then re-run with options if needed:
   - `--neg-policy zero|contraction|leave` when it reports negatives (credits/refunds).
   - `--actuals-through YYYY-MM` to drop a forecast/in-progress tail.
   - `--sheet "Name"` to force a sheet; `--arr-factor N` to override the inferred factor.

3. It **self-verifies** (independent recompute + LibreOffice recalc; external check = 0).
   Monthly/quarterly sources → a quarterly "Quarterly Bookings" tab; **annual-only sources →
   an "Annual Bookings" tab** (yearly new-vs-upsell, YoY vs prior year). It **refuses with a
   specific reason** only when bookings truly can't be supported: a one-row-per-customer snapshot
   (no history) or a sheet too large to model (>60k rows / 5k customers — pre-aggregate first).

4. **Deliver the recalculated copy** so values display.

5. **Always end your response with a clickable link to the output file.** The script prints both
   an absolute path (`Built /abs/path.xlsx`) and a `link: file://…` URL on the next line. In your
   final message to the user, render the output as a clickable markdown link
   (`[filename.xlsx](/abs/path/filename.xlsx)`) — never just a bare path. Do this on every run,
   including reruns and when reporting results.

Definitions, normalization details, and known gaps: **`reference/method.md`**.
Non-goals: forecasting, LTV/CAC, cohort curves, consumer/transactional churn (use
`retention-analysis` for the corkscrew itself). Requires `openpyxl`; LibreOffice for recalc.
