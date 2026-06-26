---
name: standardize-pnl
description: Turn any profit-and-loss / income statement — PDF, messy .xls/.xlsx, CSV, one file or one-per-year — into a clean, audited Excel P&L. Use when the user wants to consolidate, standardize, transpose, or "make Excel out of" a P&L; combine several annual P&Ls into one; or clean up a messy income-statement file. Produces a verbatim Source/Audit sheet with check-to-zero plus a formatted Output sheet (margins, growth, follow-ups). NOT for balance sheets, cap tables, retention, or forecasting.
---

# Standardize P&L

Two stages with a JSON handoff. **Stage 1 varies by input; Stage 2 is deterministic.**

## Stage 1 — Extract & normalize (judgment)
Read the source(s). Build ONE normalized JSON per `references/normalized_schema.md`:
- Transpose **verbatim** (as printed) — copy numbers exactly, never round or "fix" them.
- Map each line across all periods; carry account codes/labels as-is.
- Record every subtotal/total/computed line with its source-`stated` value (the audit key) and its member keys.
- Detect period structure; mark any partial/stub period `"full_year": false`.
- No COGS in the source? Set `output.cogs`/`output.gross_profit` to `null` (Gross Profit = Revenue). Map taxes/other-expense to their own `output` keys — never fold them into other income.
- Fill the `output` map and bucketed `followups`.
- For messy sources — monthly data, full financial statements, multi-tab/multi-cut sheets, trap "Total" columns, restated comparatives, multi-entity — **read `references/extraction_playbook.md` first.**
Save the JSON and the built workbook in the **caller's working directory**, NOT inside this skill — real company data must never land in the skill folder.

## Stage 2 — Build & audit (deterministic)
```
python3 scripts/build_pnl_workbook.py <input.json> <output.xlsx>
```
Then VERIFY (formulas must compute): recalc and confirm every check = 0 — see `references/formatting_rules.md`. Never ship without the master check reading 0.

## Locked rules (enforced by the script — do not override)
1. Never merge & center — Center Across Selection only.
2. One highlight color for all major lines (Revenue / Gross Profit / EBITDA / Net Income).
3. Whole-number percentages; margins/growth match the normal line style; only Revenue growth shown; COGS collapsed into an expandable group.
4. Follow-ups live on their own tab (classification / functional-split / comparability), never methodology prose.

Never edit raw source files. This is a public skill — keep all real data (inputs, JSON, outputs) out of the skill folder; fixtures here are synthetic only.
