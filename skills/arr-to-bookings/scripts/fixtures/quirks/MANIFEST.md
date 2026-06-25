# Quirk Fixtures — MANIFEST

Synthetic customer-revenue (MRR) workbooks for testing a normalization/parsing layer.
Each file embeds exactly ONE layout "quirk". The underlying numbers are reused across
files (10 customers, monthly MRR 500–15,000, some late starts = new logos, some
step-ups = upsell); only the LAYOUT differs.

**Shared dataset (unless a fixture says otherwise):**
- Customers: `Acme 01`, `Beacon 02`, `Cobalt 03`, `Delta 04`, `Ember 05`, `Falcon 06`, `Granite 07`, `Harbor 08`, `Ionix 09`, `Juniper 10` (10 customers).
- Periods: 36 consecutive months, Jan-2021 through Dec-2023.
- New logos (late starts): Cobalt 03 (Apr-2021), Falcon 06 (Jul-2021), Harbor 08 (Jan-2022), Juniper 10 (Jul-2022).
- Upsells (mid-life step-up): Acme 01, Cobalt 03, Delta 04, Falcon 06, Granite 07, Ionix 09.

For every fixture the data sheet is the only/active sheet, named **`Sheet1`**.

Legend: `header row` = 1-based row holding period labels · `customer col` = column letter
holding customer names · `first data col` = first column holding revenue values ·
`data row range` = 1-based rows holding customer records.

---

## q01_text_dates_mon_yy.xlsx
- **Quirk:** Month headers are TEXT strings like `"Jan 21"`, `"Feb 21"` (not real dates).
- Sheet: `Sheet1` · Orientation: WIDE (customer × month)
- Header row: 1 · Customer col: A · First data col: B · Data rows: 2–11
- Periods: 36 monthly, B1..AK1 · Date format: TEXT `"%b %y"` (e.g. `Jan 21`)

## q02_text_dates_iso.xlsx
- **Quirk:** Month headers are TEXT strings like `"2021-01"`, `"2021-02"`.
- Sheet: `Sheet1` · Orientation: WIDE
- Header row: 1 · Customer col: A · First data col: B · Data rows: 2–11
- Periods: 36 monthly, B1..AK1 · Date format: TEXT `"%Y-%m"` (e.g. `2021-01`)

## q03_formula_dates_eomonth.xlsx
- **Quirk:** First header `B1` is a REAL date `2021-01-31`; every later header is a FORMULA
  `=EOMONTH(B1,1)`, `=EOMONTH(C1,1)`, … In formula-mode (openpyxl default `data_only=False`)
  only B1 resolves to a date; C1.. read back as formula strings. Open with
  `data_only=True` (after Excel has cached values) OR evaluate EOMONTH to recover dates.
- Sheet: `Sheet1` · Orientation: WIDE
- Header row: 1 · Customer col: A · First data col: B · Data rows: 2–11
- Periods: 36 monthly, B1..AK1 · Date format: B1 datetime (number_format `yyyy-mm-dd`); C1..AK1 = `=EOMONTH(prev,1)` formulas

## q04_reversed_dates.xlsx
- **Quirk:** Real month-end dates in REVERSE chronological order — newest (`2023-12-31`) in
  column B, oldest (`2021-01-31`) at the far right (AK1).
- Sheet: `Sheet1` · Orientation: WIDE
- Header row: 1 · Customer col: A · First data col: B · Data rows: 2–11
- Periods: 36 monthly, B1..AK1 DESCENDING · Date format: real datetime, `yyyy-mm-dd`

## q05_gaps_missing_months.xlsx
- **Quirk:** Real month-end dates but TWO months are missing from the sequence: Apr-2021
  and Sep-2021 are skipped entirely (no column for them).
- Sheet: `Sheet1` · Orientation: WIDE
- Header row: 1 · Customer col: A · First data col: B · Data rows: 2–11
- Periods: 34 monthly (36 minus Apr-2021 & Sep-2021), B1..AI1 · Date format: real datetime, `yyyy-mm-dd`

## q06_customer_not_col_a.xlsx
- **Quirk:** Columns A and B are junk/label columns. A = `Type` (value `"Recurring"`),
  B = `Region` (`US`/`EMEA`/`APAC`). The CUSTOMER name is in column **C**; months start in **D**.
- Sheet: `Sheet1` · Orientation: WIDE
- Header row: 1 · Customer col: **C** · First data col: **D** · Data rows: 2–11
- Periods: 36 monthly, D1..AL1 · Date format: real datetime, `yyyy-mm-dd`

## q07_data_below_title_rows.xlsx
- **Quirk:** Rows 1–4 are title/blurb/blank (row1 title, row2–3 blurb, row4 blank). The real
  header row is **row 5**; customers start at row 6.
- Sheet: `Sheet1` · Orientation: WIDE
- Header row: **5** · Customer col: A · First data col: B · Data rows: **6–15**
- Periods: 36 monthly, B5..AK5 · Date format: real datetime, `yyyy-mm-dd`

## q08_multi_row_per_customer.xlsx
- **Quirk:** Each customer spans MULTIPLE rows — one per Product (`Core`, `Add-on`).
  Customer name in col A, Product in col B, months from col C. The correct per-customer
  value for any month is the SUM of that customer's product rows (Core ≈ 70%, Add-on ≈ 30%).
- Sheet: `Sheet1` · Orientation: WIDE, multi-row-per-entity
- Header row: 1 · Customer col: A · Product col: **B** · First data col: **C** · Data rows: 2–21 (10 customers × 2 product rows)
- Periods: 36 monthly, C1..AL1 · Date format: real datetime, `yyyy-mm-dd`
- Aggregation: GROUP BY col A, SUM across product rows.

## q09_text_numbers_and_parens.xlsx
- **Quirk:** Normal real-date headers, but VALUE cells are messy/mixed-type: thousands-text
  like `"9,900"`, currency-text like `"$1,542"`, blanks `""`, `"NA"`, and accounting
  negatives in parentheses like `"(500)"`. Some cells are plain numbers.
- Sheet: `Sheet1` · Orientation: WIDE
- Header row: 1 · Customer col: A · First data col: B · Data rows: 2–11
- Periods: 36 monthly, B1..AK1 · Date format: real datetime, `yyyy-mm-dd`
- Value parsing needed: strip `$` and `,`; treat `( )` as negative; map `""`/`NA` → blank/0.

## q10_long_tidy.xlsx
- **Quirk:** LONG / tidy (NOT wide). Three columns: `Customer`, `Date`, `MRR`. One row per
  customer per month.
- Sheet: `Sheet1` · Orientation: **LONG / tidy**
- Header row: 1 · Columns: A=`Customer`, B=`Date` (real datetime, `yyyy-mm-dd`), C=`MRR`
- Periods: 24 monthly per customer (Jan-2021..Dec-2022) · Data rows: 2–241 (10 customers × 24 months)
- Pivot needed: spread (Customer × Date) to recover the wide grid.

## q11_quarterly_native.xlsx
- **Quirk:** WIDE but columns are QUARTERS, not months. Headers are real quarter-end dates
  (`2021-03-31`, `2021-06-30`, `2021-09-30`, `2021-12-31`, …).
- Sheet: `Sheet1` · Orientation: WIDE (quarterly)
- Header row: 1 · Customer col: A · First data col: B · Data rows: 2–11
- Periods: 12 quarterly, B1..M1 (Q1-2021 .. Q4-2023) · Date format: real datetime quarter-ends, `yyyy-mm-dd`
- Granularity: QUARTERLY — do not assume monthly spacing.
