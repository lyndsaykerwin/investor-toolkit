# Layout & formatting spec

`build_concentration.py` produces this automatically. Read this only when
hand-building or repairing the workbook.

## Sheet order

1. **Customer Concentration** (analysis)
2. **Raw Data** (verbatim copy of the source sheet)

## Customer Concentration — cell map

`<unit>` below is the label only — "ARR", "MRR", or "Run-Rate". Values are never
transformed; the unit never changes the math.

| Row | Content |
|---|---|
| 1 | Title — `<Company> — Customer Concentration Analysis`, navy bold 14pt, `centerContinuous` across A:D (never merged) |
| 2 | Subtitle — `Current <unit> (<period>) — values shown as reported, not annualized` + customer count, grey 9pt, `centerContinuous` |
| 3–4 | spacer rows |
| 5 | Column headers — `Rank`, `Customer`, `Current <unit> ($)`, `% of Total <unit>` — white bold on navy `1F4E79`, medium top border |
| 6–15 | Top 10 customer rows (rank, customer link-back, value link-back, % formula) |
| 16 | **Top 10 Subtotal** — soft-blue `D9E1F2`, bold, `=SUM(C6:C15)` and `=C16/C18` |
| 17 | **All Remaining Customers (N)** — `=C18-C16`, `=C17/C18` |
| 18 | **TOTAL — All Customers** — navy fill white bold, `=SUM('Raw Data'!<col><first>:<col><last>)`, % = `=C16/C18+C17/C18` (definitionally 100%) |
| 20 | Tie-out check — `=C16+C17-C18` (must read 0), grey 9pt |
| 21 | Source note — grey 8pt, names the Raw Data tab + source file/sheet, notes values are verbatim |

Row positions shift only if Top-10 has fewer than 10 rows (small base);
the build computes subtotal/remaining/total/tie-out rows relative to the data
block, so always locate them by label, not by fixed row number.

## Formula conventions

- **Customer name:** `='Raw Data'!<custcol><srcrow>` — green font (cross-sheet ref).
- **Current value:** `='Raw Data'!<valcol><srcrow>` — verbatim, no factor ever. Green font. Number format `$#,##0`. When the source column is currency-formatted *text*, wrap in the coercion `IFERROR(VALUE(SUBSTITUTE(SUBSTITUTE(TRIM(ref),"$",""),",","")),0)` so it sums (use `VALUE`, not `NUMBERVALUE` — the latter is `#NAME?` in headless LibreOffice).
- **% of total:** `=C<r>/C<total>` — black font (in-sheet formula). Format `0.0%`.
- **Grand total:** an independent `SUM` over the source value column (not a sum of the displayed rows) so it is a true external check on the ranked rows. For text columns use `SUMPRODUCT(<coercion over the range>)` (array-aware, no Ctrl-Shift-Enter).
- **Ranking** is done in Python (values read once to order the rows); the displayed numbers stay link-back formulas. Re-sorting by formula is not possible in plain Excel, so order is fixed at build time — if the raw data changes materially, rebuild.

## Color / font palette (banker convention)

| Element | Hex | Note |
|---|---|---|
| Title font | `1F4E79` navy | bold 14pt |
| Header fill | `1F4E79` navy | white bold text |
| Subtotal fill | `D9E1F2` soft blue | black bold |
| Total fill | `1F4E79` navy | white bold |
| Hardcoded input font | `0000FF` blue | reserved for any manually-keyed input |
| Cross-sheet ref font | `008000` green | customer + value link-backs (classic banker convention: green = link to another sheet) |
| In-sheet formula font | `000000` black | % cells, subtotal, tie-out |
| Metadata/footer font | `595959` grey | subtitle, tie-out, source note |

## Hard rules

- **Never merge cells** on the analysis sheet — use `centerContinuous`. The Raw
  Data tab may contain merges only because they were copied verbatim from the
  source; you never create one.
- **`fullCalcOnLoad` is set** so any viewer recalcs on open. The shipped file is
  the original build, never the LibreOffice recalc copy (LibreOffice rewrites
  `centerContinuous` as merged cells on save).
- **No value is ever transformed.** There is no ARR-factor cell and no ×12. The
  unit is a label only (`--unit ARR|MRR|Run-Rate`).
- Currency `$#,##0`; percentages one decimal `0.0%`.
