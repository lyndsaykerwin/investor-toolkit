# Formulas, layout & formatting — implementation spec

**You usually do NOT need this file.** `scripts/deliver.py` builds the entire
workbook — layout, formulas, and formatting — automatically. Read this only when
you are **hand-building or repairing** the workbook cell-by-cell (e.g. the script
doesn't fit an unusual source and you're patching formulas by hand).

`deliver.py` is the source of truth. If this file and the script ever disagree,
the script wins.

A note used throughout: formulas reference the **helper sheet ("Raw Data with
Analysis") if one was built, otherwise Raw Data directly**. The helper is
optional — see SKILL.md "When the helper sheet is needed."

Three referencing modes (`write_corkscrew_sheet_aggregating`):
- **Helper mode** (`raw_geometry=None`) — rollforward ranges point at the helper's
  canonical grid (customer rows from row 12, months from col B); customer counts
  come via `HLOOKUP` into the helper's summary rows 2/3.
- **Two-sheet mode** (`raw_geometry={sheet,first_row,last_row,first_date_col_idx}`,
  the `--two-sheet` path) — no helper; rollforward ranges and the external check
  reference **Raw Data directly** at the *source* block geometry, and customer
  counts are computed **inline** (`COUNTIF` for active, one `SUMPRODUCT` for
  retained) instead of `HLOOKUP`. See "Two-sheet (no-helper) referencing" below.
- **Blended mode** (`write_blended_corkscrew`) — sums per-segment Corkscrew cells
  and reconciles independently against each segment's Raw Data. See "Blended
  Corkscrew" below.

---

## Corkscrew sheet layout

Define all row positions before writing any formula — a formula written before
the layout is locked points to the wrong row when a later header insertion shifts
everything down.

```
Row 1     Title: "<Company> ARR Corkscrew — Retention Analysis"
          centerContinuous, navy fill (#1F4E79), white bold text
Row 2     "Generated:" | timestamp
Row 3     "ARR Factor (MRR × N):" | factor value (BLUE — hardcode)
Row 5     Date headers across columns (Jan-22, Feb-22, …)
Row 6     Optional prior-period reference label ("vs Jan-21", …) for YoY layouts

Rollforward block
Row 8     Beginning ARR              [formula → helper/Raw Data × $ARR_factor]
Row 9       + New customer ARR       [formula]
Row 10      + Expansion (Upsell)     [formula]
Row 11      − Contraction (Downsell) [formula, stored negative]
Row 12      − Churn                  [formula, stored negative]
Row 13    Ending ARR                 [= rows 8+9+10+11+12]
Row 14    External Check (= 0)       [= row13 − independent-sum × $ARR_factor]

Customer count block
Row 16    SECTION BANNER "CUSTOMER COUNTS"
Row 17    # Active (prior period)    [HLOOKUP into helper]
Row 18    # Active (current)         [HLOOKUP into helper]
Row 19    # Churned                  [prior active − retained]
Row 20    # New                      [current active − retained]

Retention metrics block
Row 22    SECTION BANNER "RETENTION RATES"
Row 23    Gross Dollar Retention (GRR)   [= (Beg + Downsell + Churn) / Beg]
Row 24    Net Dollar Retention (NRR)     [= (Beg + Upsell + Downsell + Churn) / Beg]
Row 25    Logo Retention                 [= (Active prior − Churned) / Active prior]

Per-customer metrics
Row 27    SECTION BANNER "PER-CUSTOMER METRICS"
Row 28    Avg ARR per Active Customer    [= Ending ARR / # Active current]
Row 29    Avg ARR per New Customer       [= New ARR / # New]

Decomposed reconciliation (ONLY when multiple in-scope revenue types)
Row 31    SECTION BANNER "RECONCILIATION CHECKS"
Row 32    Recurring ARR              [= helper row 6 × $ARR_factor]
Row 33    Re-occurring ARR           [= helper row 7 × $ARR_factor]
Row 35    Sum customer ARR           [= row 32 + row 33]
Row 36    Variance vs Ending ARR     [= row 35 − row 13]   must = 0
```

Optional LTM corkscrew block (rows 38–48) with the same shape, comparison T-12,
when LTM is part of the methodology.

---

## Helper sheet layout (Raw Data with Analysis)

Built only when the raw data needs transformation. Summary block on top (only
when self-validation is needed), customer data below.

```
Row 1   Month headers           "2021-M1" … "2024-M3". Col A label = "Customer ID"
Row 2   # Active customers      = COUNTIF(<col>$12:<col>$<last>, ">0")
Row 3   # Retained vs prior     For first N columns (N = lookback) the value is "n/a"
                                — no prior period yet. From col N+1 onward, array formula:
                                = SUMPRODUCT((<curr>$12:<curr>$<last>>0) *
                                             (<prior>$12:<prior>$<last>>0))
                                Only place SUMPRODUCT is needed.
Row 4   Check # Active vs Raw   Independent recount directly from Raw Data, must = 0
Row 5   blank divider
Row 6   Recurring MRR total     SUMIFS(INDEX(Raw Data block, 0, MATCH(col$1, header_row, 0)),
                                       type_col, "Recurring")
Row 7   Re-occurring MRR total  Same pattern, "Re-occurring"
Row 8   Non-recurring MRR total Same pattern, "Non-recurring"
                                (Keep even when out of scope — needed for full-coverage recon)
Row 9   Total MRR (all types)   = <col>6 + <col>7 + <col>8
Row 10  Check vs Raw Data       = <col>9 − SUM('Raw Data'!<month_col>)   must = 0
Row 11  Check (Rec + Re-occ)    = (<col>6 + <col>7) − SUM(customer rows)  must = 0
Row 12+ Customer data           Col A = Customer ID. Each month cell uses DIRECT column
                                reference (not INDEX/MATCH) — thousands of rows × dozens
                                of columns, recalc speed matters:
                                = SUMIFS('Raw Data'!$<month>$<first>:$<month>$<last>,
                                         'Raw Data'!$<cust>$<first>:$<cust>$<last>, $A<row>,
                                         'Raw Data'!$<type>$<first>:$<type>$<last>, "<filter>")
```

**Formula style summary:**
- Summary block rows 6-8: `INDEX/MATCH` dynamic column lookup → copies horizontally without hand-edits.
- Customer rows 12+: direct column reference → recalc speed.
- Check rows 4, 10, 11: direct column reference → one-off, not a copyable pattern.

**Freeze panes** at `B12` so labels and summary stay visible while scrolling.

---

## Corkscrew formula patterns

All Corkscrew formulas reference the helper sheet if one was built, otherwise Raw
Data directly. The helper has month headers in row 1, customer data starting at
row 12. For each Corkscrew comparison-period column you need the **current
period** and the **prior period** helper columns.

**Column mapping — work this out before writing any formulas.** For YoY
(12-month lookback) over a 39-month source dataset:

| Corkscrew col | Period label | Helper current col | Helper prior col |
|---|---|---|---|
| C | 2022-M1 (idx 12) | N (idx 12) | B (idx 0) |
| D | 2022-M2 (idx 13) | O (idx 13) | C (idx 1) |
| AC | 2024-M3 (idx 38) | AN (idx 38) | AB (idx 26) |

Rule: for Corkscrew column at offset `i` from the first comparison-period column,
helper current is at month-index `lookback + i`, helper prior is at month-index
`i`. The helper's first month column is B; corresponding helper column letter is
`get_column_letter(2 + month_index)`.

**Rollforward formulas** (`<curr>` and `<prior>` are helper column letters; replace
`'Raw Data with Analysis'` with `'Raw Data'` in the two-sheet case):

```
Beginning ARR  =SUMPRODUCT(('Raw Data with Analysis'!<prior>$12:<prior>$<last>>0)*
                          'Raw Data with Analysis'!<prior>$12:<prior>$<last>)*$C$3

New ARR        =SUMPRODUCT(('Raw Data with Analysis'!<prior>$12:<prior>$<last>=0)*
                          ('Raw Data with Analysis'!<curr>$12:<curr>$<last>>0)*
                          'Raw Data with Analysis'!<curr>$12:<curr>$<last>)*$C$3

Upsell         =SUMPRODUCT(('Raw Data with Analysis'!<prior>$12:<prior>$<last>>0)*
                          ('Raw Data with Analysis'!<curr>$12:<curr>$<last>>'Raw Data with Analysis'!<prior>$12:<prior>$<last>)*
                          ('Raw Data with Analysis'!<curr>$12:<curr>$<last>-'Raw Data with Analysis'!<prior>$12:<prior>$<last>))*$C$3

Downsell       =SUMPRODUCT(('Raw Data with Analysis'!<prior>$12:<prior>$<last>>0)*
                          ('Raw Data with Analysis'!<curr>$12:<curr>$<last>>0)*
                          ('Raw Data with Analysis'!<curr>$12:<curr>$<last><'Raw Data with Analysis'!<prior>$12:<prior>$<last>)*
                          ('Raw Data with Analysis'!<curr>$12:<curr>$<last>-'Raw Data with Analysis'!<prior>$12:<prior>$<last>))*$C$3

Churn          =SUMPRODUCT(('Raw Data with Analysis'!<prior>$12:<prior>$<last>>0)*
                          ('Raw Data with Analysis'!<curr>$12:<curr>$<last>=0)*
                          (-'Raw Data with Analysis'!<prior>$12:<prior>$<last>))*$C$3

Ending         =<col>8+<col>9+<col>10+<col>11+<col>12

External Check =<col>13 - (independent_sum_path × $C$3)
               // independent_sum_path = SUMIFS on Raw Data when no transformation,
               // OR SUM of helper rows 6+7 when type-filter & per-customer aggregation needed
```

**Customer-count formulas** (HLOOKUP — simple, deterministic, easy to audit):

```
# Active prior    =HLOOKUP(SUBSTITUTE(<col>$6,"vs ",""),'Raw Data with Analysis'!$B$1:$<last>$2, 2, FALSE)
# Active current  =HLOOKUP(<col>$5, 'Raw Data with Analysis'!$B$1:$<last>$2, 2, FALSE)
# Churned         =<col>17 - HLOOKUP(<col>$5, 'Raw Data with Analysis'!$B$1:$<last>$3, 3, FALSE)
# New             =<col>18 - HLOOKUP(<col>$5, 'Raw Data with Analysis'!$B$1:$<last>$3, 3, FALSE)
```

**Retention metrics** (all use `IFERROR` so empty-prior periods don't error):

```
GRR    =IFERROR((<col>8 + <col>11 + <col>12) / <col>8, 0)
NRR    =IFERROR((<col>8 + <col>10 + <col>11 + <col>12) / <col>8, 0)
Logo   =IFERROR((<col>17 - <col>19) / <col>17, 0)
```

---

## Two-sheet (no-helper) referencing — `--two-sheet`

In the two-sheet path there is no helper, so the Corkscrew references **Raw Data
directly** at the *source* block geometry passed in `raw_geometry`:

- `first_row` / `last_row` — the source customer block (e.g. rows 5–162 for
  Metazoa's Core Enterprise sheet).
- `first_date_col_idx` — 1-based column index of the source's first month column
  (e.g. column B = 2). Month index `i` → `get_column_letter(first_date_col_idx + i)`.

The rollforward ranges become `'Raw Data'!<curr>$<first_row>:<curr>$<last_row>`
and `'Raw Data'!<prior>...` (same `SUMPRODUCT` shapes as helper mode — only the
sheet and range bounds change). The customer counts, which in helper mode are
`HLOOKUP`s into the helper's summary rows, are computed **inline** against Raw
Data:

```
# Active prior    =COUNTIF('Raw Data'!<prior>$<first>:<prior>$<last>,">0")
# Active current  =COUNTIF('Raw Data'!<curr>$<first>:<curr>$<last>,">0")
# Retained        SUMPRODUCT(('Raw Data'!<curr>...>0)*('Raw Data'!<prior>...>0))
                  (the one legitimate SUMPRODUCT — differential across two periods)
# Churned         = # Active prior  − Retained
# New             = # Active current − Retained
```

The single-type external check (row 14) sums the current-month column straight
from Raw Data: `=Ending − SUM('Raw Data'!<curr>$<first>:<curr>$<last>) × $ARR_factor`.
This path is valid **only** for a clean contiguous single-type block — survey-gated
and opt-in (see SKILL.md). It produces identical Ending / GRR / NRR / Logo to the
three-sheet pass-through for the same clean source.

---

## Blended Corkscrew (multi-segment) — `write_blended_corkscrew`

For a multi-segment workbook, each segment gets its own two-sheet Corkscrew + Raw
Data, and a leftmost **Blended Corkscrew** ties them together with a *real*
reconciliation (never by summing the segments' own check rows):

```
Blended Beginning(t)  = Σ_seg  '<seg> Corkscrew'!<col>8
Blended New/Up/Down/Churn(t) = Σ_seg  '<seg> Corkscrew'!<col>{9,10,11,12}
Blended Ending(t)     = Σ_seg  '<seg> Corkscrew'!<col>13
Blended counts(t)     = Σ_seg  '<seg> Corkscrew'!<col>{17,18,19,20}

Variance(t) [row 14, must = 0 every period]
    = Blended Ending(t)
      − ( Σ_seg SUM('<seg> Raw Data'!<curr month col>$<first>:$<last>) ) × $ARR_factor
```

The Variance right-hand side is an **independent path** straight from each
segment's Raw Data — it never references the per-segment `Beginning+moves=Ending`
identity — so `Variance == 0` every period proves the blend ties back to source.
Blended GRR/NRR/Logo are computed from the blended dollar rows and Σ counts. All
segments must share the same comparison-period count (same month range) so columns
line up; `deliver_segments()` raises otherwise.

---

## Formatting standards

### Cell text colors (standard finance convention)
- **Blue text (RGB 0,0,255)** — hardcoded inputs (ARR factor, raw data values when displayed, date headers, methodology label values).
- **Green text (#006100)** — references to another sheet (cells that pull from Raw Data or the helper).
- **Black text** — formulas computed within the current sheet.

### Fill colors
- Section header banner rows — dark blue `#1F4E79`, white bold text.
- Sub-headers / row labels — light blue `#D9E1F2`, black bold.
- Rollforward anchor rows (Beginning ARR, Ending ARR) — medium blue `#BDD7EE`, black bold. Apply to all data cells in those two rows as bookends. Do NOT apply to retention rates / per-customer metrics — sibling metrics stay visually uniform.
- Check rows: green text on white when passing, red text when failing. Never ship a red check.

No greens / yellows / oranges in the model body. Reserve red for failed checks (never ship) and green text only for passing checks and cross-sheet references.

### Number formats — `$` on top and bottom of a block, not every cell
- **Top of block (Beginning ARR) and bottom (Ending ARR):** `"$"#,##0;("$"#,##0);"-"`
- **Middle rows (New, Upsell, Downsell, Churn, helper customer cells):** `#,##0;(#,##0);"-"`
- **Percentages:** `0.0%;(0.0%);"-"`
- **Customer counts:** `#,##0;(#,##0);"-"`
- **Dates:** `mmm-yy`

### Headers & labels
- Title row: navy fill (`#1F4E79`), white bold, `centerContinuous`.
- Date row: same navy fill, white bold, center-aligned.
- Sub-header rows ("vs prior year"): light-blue fill (`#D9E1F2`), no bold.
- ARR factor label reads `"ARR Factor (MRR × N):"` not `"ARR Factor"`. Title says `"YoY ARR Corkscrew"` or `"Monthly ARR Corkscrew"` — comparison period in the title. Currency unit somewhere visible: `"All figures in $USD"`.

### Other
- **Never merge cells.** Use `Alignment(horizontal="centerContinuous", vertical="center")` on every cell in the span, text written only to the leftmost. Merged cells break selection, sorting, filtering, copy/paste.
- **Borders:** thick (1.5pt) around the rollforward, retention, and reconciliation blocks; thin (0.5pt) on data tables.
- **Column widths:** label column ~38, data columns ~13.
- **Freeze panes** on the Corkscrew at the first data column / first data row (typically `C7`).
