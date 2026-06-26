# Formatting & verification

The build script enforces all formatting. Don't restyle the output by hand — change the JSON or the script.

## Locked formatting rules
1. **Never merge & center.** Titles/spanned headers use *Center Across Selection* (`Alignment(horizontal="centerContinuous")` applied across the span; value only in the leftmost cell). Merging breaks cell addressing and audit readers.
2. **One highlight color.** All major lines (Revenue, Gross Profit, EBITDA, Net Income) use the single fill `HILITE = D6E0F0`. No multi-color banding.
3. **Actionable follow-ups only, on their own tab.** The follow-up list lives on a separate **Follow-ups** sheet (keeps the P&L overview clean), grouped into Classification / Functional-split / Comparability. No methodology paragraphs.
4. **Margins & growth match the normal line style** — same color (black) and size (10) as the P&L lines, not grey/italic.
5. **Percentages have no decimals by default** (`PCT = '0%;(0%)'`). `n/m` is right-aligned to match the percentage cells.
6. **Only Revenue growth is shown** — Gross Profit / EBITDA / Net Income growth rows are omitted. All margins (gross / EBITDA / net) are shown.
7. **COGS is collapsed** into an expandable row group (hidden by default; the user can expand it), with a blank spacer between Revenue growth and Gross Profit.

Other conventions: negatives in parentheses; Source sheet keeps 2 decimals (verbatim), Output shows whole dollars; gridlines off on Output & Follow-ups; partial periods flagged amber and growth into them = `n/m`.

## No real data
This is a public skill. Fixtures and doc examples must be **fully synthetic** (`fixtures/example_pnl.json`, `fixtures/regression_services.json`). Never commit real company names, product names, business specifics, or test-run outputs into the skill folder. Build real deliverables into the caller's own folder, not under the skill.

## Three tabs
- **Output** — clean P&L overview. Links to Source_Audit; computes margins, Revenue growth, EBITDA/EBIT; COGS grouped/hidden.
- **Follow-ups** — the bucketed action list.
- **Source_Audit** — verbatim transpose. Details hard-typed; subtotals/computed are live formulas. Right-side block: `per source (stated)` then `check = computed − stated` (must be 0). Bottom: green **MASTER AUDIT CHECK** = sum of all checks = 0.

## VERIFY before shipping (required)
Formulas are written by openpyxl but not evaluated. Recalc with LibreOffice headless and confirm checks are zero:
```bash
SOFFICE="/Applications/LibreOffice.app/Contents/MacOS/soffice"   # macOS path; or `libreoffice`/`soffice` on PATH
"$SOFFICE" --headless --calc --convert-to xlsx --outdir <run>/recalc <run>/out.xlsx
```
Then load `<run>/recalc/out.xlsx` with `openpyxl(data_only=True)` and assert:
- **MASTER AUDIT CHECK = 0** (Source_Audit), and no individual check cell ≠ 0.
- **Reconciliation row = 0** for every period (Output).
- **0 merged cells** on all sheets (`len(ws.merged_cells.ranges) == 0`).
- Headline values tie to the source-stated numbers.

If LibreOffice isn't installed, recompute the subtotals independently in Python from the JSON `detail` values and compare to `stated`; report that the live-formula recalc was skipped.

## Portability
- The build script is pure `openpyxl` (portable across Claude Code / Cowork / Codex).
- Only the VERIFY step shells out to LibreOffice. On hosts without it, use the Python fallback above and say so in the receipt.
