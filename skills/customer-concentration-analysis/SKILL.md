---
name: customer-concentration-analysis
description: Use when the user wants a customer concentration analysis from customer-level recurring-revenue data — ranking customers by current ARR, showing the Top 10 with each customer's ARR and % of total, a Top-10 subtotal %, an all-remaining-customers %, and a grand total that ties out to full current ARR. Produces a banker-grade Excel deliverable where every figure is a live formula linking back to a verbatim copy of the source. Triggers on "customer concentration", "top 10 customers", "revenue concentration", "customer concentration analysis", "how concentrated is the customer base", or a customer-level ARR/MRR file plus a concentration ask. Not for retention/churn (use retention-analysis) or bookings (use ARR-to-bookings).
---

# customer-concentration-analysis

## What it produces

One banker-grade `.xlsx` with **two sheets**:

1. **Customer Concentration** — customers ranked largest→smallest by current run-rate. Top 10 shown individually (rank, customer, value $, % of total); a **Top 10 Subtotal** row (+ % of total); an **All Remaining Customers** row (combined $ + %); a **TOTAL — All Customers** row that ties out; and a visible **tie-out check** (Top 10 + Remaining − Total = 0).
2. **Raw Data** — a verbatim copy of the source sheet.

Every figure is a live formula referencing the Raw Data tab — **zero hardcoded numbers** — so the tie-out is provably sourced.

## The three rules this skill exists to enforce

These are hard-won from a real build that shipped wrong. Do not relax them.

### 1. NEVER transform the values — label honestly instead
The skill **never annualizes** (no ×12) and never rescales. Each customer's value is shown exactly as the source reports it. This is deliberate: concentration is all *ratios* — ranking and every percentage are identical whether the data is monthly or annual (×12 the parts and the whole cancels). The only thing the unit affects is the *absolute dollar labels*, and that's precisely where a transformation step goes wrong. So instead of transforming, the skill just **labels** the value column honestly:
- Source unit **certainly ARR** → "Current ARR ($)"
- Source unit **certainly MRR** → "Current MRR ($)"
- **Not certain** → "Current Run-Rate ($)" (neutral; never guess, never annualize)

`detect_unit.py` reads the workbook's own words (sheet name → title/header text), reports the verdict + evidence + a suggested `column_label`, and only commits to "ARR"/"MRR" when a label is **decisive**. Magnitude is never used to transform — at most it's a flagged tiebreaker. Surface the determination to the user; pass the chosen label as `--unit`. (The real miss this prevents: a tab named "…ARR" with a ~$13.5K median was once annualized ×12 by a magnitude guess, overstating the total 12×. The concentration % was right both times — only the dollars lied.)

### 2. Preserve the raw data verbatim
The Raw Data tab is an exact copy of the source: zero edits, no reformatting, no reordering, no rounding. It exists so the user can see nothing was touched. `build_concentration.py` copies it for you.

### 3. Every analysis cell links back
The ranked rows, subtotal, remaining, and total are all formulas pointing at the Raw Data tab (e.g. `='Raw Data'!AQ44`, total = `SUM('Raw Data'!AQ5:AQ162)`). No hardcoded values anywhere on the analysis sheet. When the source column is currency-formatted *text* (e.g. `"$8,333.33 "`), the formulas coerce it numerically (strip `$`/`,`/spaces) so `SUM` doesn't silently drop it — while Raw Data stays verbatim.

## Workflow

1. **Detect.** Run `python3 scripts/detect_unit.py <source> [--sheet NAME] --json`. It returns the sheet, header row, customer column, latest-period column, the **unit verdict + evidence**, and a suggested `column_label` (ARR / MRR / Run-Rate). Prefer dispatching a subagent for big files (read-only). If the verdict is `TRANSACTIONAL` (invoice/transaction-level, not normalized recurring) or `AMBIGUOUS`, stop and confirm scope — concentration needs one current value per customer.
2. **Confirm with the user** in one message: the column label (use the detector's `column_label` — "Run-Rate" unless ARR/MRR is certain), the current period, the customer column, and the customer-row range — with 2-3 sample values so they can verify without opening the file. The values are shown verbatim either way, so this is about the *label*, not a transformation.
3. **Build.** `python3 scripts/build_concentration.py <source> <out.xlsx> --sheet "<sheet>" --customer-col <L> --value-col <L> --first-row <N> --last-row <N> --unit <ARR|MRR|Run-Rate> --period-label "<label>" --company "<name>"`. Version the filename (`_v1`, `_v2`, …).
4. **Verify (recalc gate — do NOT ship the recalc copy).** Recalculate a *throwaway* copy with LibreOffice to confirm the math: `soffice --headless --calc --convert-to xlsx --outdir <tmp> <out.xlsx>`, then open the tmp copy `data_only=True` and check the **tie-out cell = 0**, Top-10 % + Remaining % = 100%, and no `#REF!/#VALUE!/#DIV0!`. The shipped file is the original build (clean `centerContinuous`, no merges, `fullCalcOnLoad` set so Excel recalcs on open). LibreOffice converts `centerContinuous` into merged cells on save — that is why the recalc copy is verification-only and never shipped.
5. **Banker-QA.** Run the `banker-formatting-qa` skill on the shipped file and fix anything it flags.
6. **Report** the file path, the total (in whatever unit was labeled), Top-10 %, Remaining %, and the unit determination + evidence (and that values were shown verbatim, not annualized).

## Scripts

- `scripts/detect_unit.py` — unit detector (label-first) → verdict + evidence + suggested `column_label`. `--self-test` checks the fixture.
- `scripts/build_concentration.py` — builds the two-sheet banker workbook. `--allow-large` required above 20,000 customer rows (verbatim copy gets heavy — confirm scope first).
- Layout, formatting palette, and the full cell map: `references/layout-and-formatting.md`.
- Synthetic fixture: `fixtures/example_saas_arr.xlsx` (regenerate with `fixtures/make_fixture.py`).

## Final checklist

- [ ] Values shown verbatim — **no annualization / ×12 anywhere**
- [ ] Column labeled ARR/MRR only if certain, else "Run-Rate"; evidence shown to user
- [ ] Two sheets; Raw Data is a verbatim copy
- [ ] Every analysis figure links back to Raw Data; no hardcoded values
- [ ] Tie-out cell = 0; Top-10 % + Remaining % = 100%; no formula errors
- [ ] No merged cells on the analysis sheet; `banker-formatting-qa` passes
- [ ] Filename versioned; prior version moved to `Archive/`
