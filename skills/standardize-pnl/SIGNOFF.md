# standardize-pnl — sign-off receipt

**Result:** 8/8 representative cases PASS (master audit check = 0, every individual check cell = 0, Output reconciliation = 0, 0 merged cells) — verified by LibreOffice headless recalc.

> Test inputs were real founder/accountant P&Ls used privately for validation. Per skill policy they are **not** part of this skill and no real company data, names, or business details appear here. Cases are described by shape only.

## Cases tested (by shape)
| # | Form | Stressor exercised |
|---|------|--------------------|
| 1 | 4 annual PDFs | annual columns; partial final year |
| 2 | synthetic fixture (`regression_services.json`) | no-COGS, nested subtotal, income taxes, other expense |
| 3 | 5 annual PDFs (full financial statements) | isolate income statement from full package; restated comparatives; income taxes |
| 4 | xlsx, 3 tabs, monthly | pick the P&L tab; 24 months → 2 years; parent account with own postings |
| 5 | xlsx, monthly ~4yr | months→annual; trailing "Total" trap column; partial year |
| 6 | xlsx, multi-entity monthly | several entities combined; coarse P&L with no COGS |
| 7 | 3 near-duplicate xlsx | single partial period; duplicate-file disambiguation; no COGS |
| 8 | xlsx, large sheet + recurring-revenue tab | mixed/large sheet; stale annual columns; non-footing subtotals; other-expense bridge |

## Bugs / gaps found and fixed (from a 6-agent parallel test run)
1. **No-COGS crash** (hit by 4/6 inputs) — builder required `cogs`/`gross_profit`; now nullable, Gross Profit computes Revenue − COGS. `link()` is null-safe.
2. **Nested subtotals silently dropped** (caused a large miss on one input) — builder now resolves a subtotal whose member is itself a subtotal (`member_present`). Regression fixture #2 locks this in.
3. **No tax / other-expense slot in the bridge** — added `output.income_taxes` and `output.other_expense`; Net Income bridge = EBIT − interest − tax + other_income − other_expense, built dynamically.
4. **verify_workbook only checked the master total** — now scans every individual check cell (offsetting errors can't hide).
5. **Hardcoded "Consolidated" title** — now meta-driven.
6. **Single-period emitted blank growth rows** — now suppressed.
7. **Stage-1 judgment undocumented** — added `references/extraction_playbook.md` (monthly→annual aggregation + two audit modes; isolating the income statement; trap/stale columns; parent-with-postings; non-footing subtotals; restated comparatives; multi-entity; indentation encoding). Schema updated for nullable COGS, nested members, the bridge identity.

## Formatting (locked, enforced by the build script)
Never merge & center (Center Across Selection); one highlight color for major lines; whole-number percentages by default; margins/growth match the normal line style; only Revenue growth shown; COGS collapsed into an expandable group; follow-ups on their own tab.

## Portability
- `build_pnl_workbook.py` — pure openpyxl (Claude Code / Cowork / Codex OK).
- `verify_workbook.py` — uses LibreOffice for recalc; falls back to a Python subtotal recompute (with a printed warning) when LibreOffice is absent.

## Known gaps / future hardening
- Stage-1 extraction is model-judgment; the playbook covers the patterns seen, but a novel layout may need a new note.
- No automated check yet that `output.opex` covers every operating leaf exactly once (currently a Stage-1 responsibility).
