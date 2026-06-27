# Fixture expected values

`example_saas_arr.xlsx` — synthetic, fully fictional (Example SaaS Co.). Regenerate
with `python3 fixtures/make_fixture.py`.

## Shape
- One sheet named **"Annual Recurring Revenue"** (the label is decisive: ARR).
- Title row 1 says "Annual Recurring Revenue"; header row 4 = `Customer | FY2023 | FY2024 | FY2025`.
- 15 customers, columns B–D are years; current period = FY2025 (col D).
- Values are deliberately small ($8k–$95k) so a magnitude-only guess would
  wrongly say MRR — the fixture proves the detector trusts the **label**.

## detect_unit.py
- Verdict: **ARR**, confidence **decisive** (from the "annual recurring" label,
  not magnitude). `column_label` = **ARR**. This is the self-test assertion.

## build_concentration.py (values verbatim, no transformation)
Built with `--customer-col A --value-col D --first-row 5 --last-row 19 --unit ARR`:
- Total current ARR = **$469,000** (values shown as reported, never annualized)
- Top 10 subtotal = $414,000 → **88.3%** of total
- All remaining (5 customers) = $55,000 → **11.7%**
- Tie-out check cell = **0**; no formula errors; no merged cells on the analysis sheet.
