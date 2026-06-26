# Stage 1 extraction playbook — handling messy real-world sources

Stage 1 (reading the source into the normalized JSON) is where all the judgment lives. These are the patterns real founder/accountant files throw at you, learned from a multi-file sign-off. The deterministic builder downstream is reliable; correctness depends on getting this stage right.

## Pick the right thing to read
- **Multi-tab workbook**: use the P&L / "Profit and Loss" / "Statement of Income" tab. Ignore Balance Sheet, Cash Flows, and recurring-revenue schedules (CMRR/ARR/MRR tabs are not the P&L).
- **Full financial-statement PDF**: find the income-statement page. The P&L **stops at Net Income** — exclude any retained-earnings / "income and deficit" rollforward (beginning RE, dividends, ending RE) printed below it. Skip cover, compilation/auditor report, balance sheet, cash flows, and notes.
- **One tab with several presentations stacked** (GAAP block, then recategorized revenue, a reclassified-opex cut, an EBITDA bridge, an Adjusted/pro-forma block): transpose the **canonical GAAP block** — the one whose detail lines foot to their own total rows. List the alternate cuts as a `comparability` follow-up; don't transpose them.
- **Several files that might be duplicates**: diff them first (cell-by-cell with `openpyxl(data_only=True)`). If identical, use any one and note the duplication in `meta.source`. A file *named* "Consolidated" is not necessarily consolidated — verify.

## Aggregating sub-annual data (monthly → annual)
Very common: the source is monthly but you want annual columns.
- **Sum the months** for each line into each calendar-year column. The **verbatim rule then applies to the line mapping, not the cell value** — your annual number is a sum you computed. Add a `comparability` follow-up disclosing the annual figures are your own monthly sums.
- **Audit has two modes.** If the source prints an annual total (a year-total column or row), use it as `stated` — a true tie-out. If it doesn't, your `stated` = the source's total *row* summed across the year, so the check confirms **internal consistency** (subtotal = sum of members), a weaker guarantee. Say which mode in the follow-up.
- **Year detection**: read the year off each month header ("Jan 2024" → 2024; "Nov 1-19, 2024" → 2024, partial).

## Trap columns / stale annual columns
- A trailing column headed **"Total"** is usually a row-sum across *all* months (every year), **not a period** — never ingest it as a year column. Restrict your period map to the dated month columns.
- If the sheet has **annual-total columns**, do **not** trust them blindly: confirm each foots to the sum of that year's monthly columns. Real files have annual columns that silently drop lines (stale). If they don't foot, sum the months instead and flag the discrepancy.

## Label / structure gotchas
- **Indentation** is often encoded as **leading spaces in the label text**, not `cell.alignment.indent` (which is frequently 0). Detect hierarchy from both.
- **Account codes may be absent** — keys can be any stable slug (e.g. `"rev_serv"`), not just GL codes.
- **Parent-with-postings** and **non-footing subtotals**: see the rules in `normalized_schema.md`.

## Multi-year PDFs with restated comparatives
Each annual statement usually shows current year + a prior-year comparative, and a later file may restate the earlier year differently. **For year Y, transpose the statement whose period IS year Y** (its own primary column), not a later year's comparative. When the two disagree, keep year Y's own figure and record the restatement as a `comparability` follow-up.

## Multi-entity P&Ls
A combined file may stack several entities (and sometimes a pre-summed combined block). Decide the unit of analysis: usually a **combined** annual P&L (sum the in-scope entities). If the source already has a combined block, use it as an independent cross-check of your sum. Note in a `classification` follow-up that per-entity detail was collapsed, and watch for out-of-scope or grand-total blocks you must exclude.
