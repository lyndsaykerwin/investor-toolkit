---
name: finance-formatting
description: Use whenever an .xlsx file needs to be brought up to investment-banker dealbook formatting standards — before a spreadsheet goes to a counterparty (LP, founder, banker, MD, board) or is filed as a workspace deliverable. Triggers on "fix the formatting", "make this banker-grade", "format-check and clean this xlsx", "polish this file", "make this presentable", "format audit". This skill finds deviations AND corrects them.
---

# Finance Formatting (xlsx)

Bring an .xlsx up to the standard an investment-banking analyst applies before a deal book reaches an MD. This skill **fixes**, not just flags.

**Run as a subagent-driven workflow.** A fixer subagent applies corrections via openpyxl; a separate QC subagent verifies the saved file against every norm below. Nothing returns to the user until QC passes — the core session orchestrates and never hands back an unverified file. Re-render the operator Excel after fixes so the deliverable reflects them.

## The non-negotiable rule: never merge cells

Merged cells break filtering/sorting/copy-paste, blow up on row insert, and hide from openpyxl style edits. If `ws.merged_cells.ranges` is non-empty, unmerge and replace:
- **Centering a label across columns** → `Alignment(horizontal="centerContinuous")` on the leftmost cell.
- **One long string in one cell** → widen the column, `wrap_text=False`.

## Norms to enforce

- **Typography** — one family throughout (Calibri/Aptos Narrow). Title ~20pt bold, section headers ~11pt bold, column headers ~10pt bold, data ~10pt, footer ~8pt italic.
- **Color** — confirm a house palette and apply it consistently (classic convention: hardcoded inputs blue font, formulas black). Drift without reason → correct it.
- **Borders** — thin black all four sides on data cells; medium border below section-header rows; no missing-side ragged edges; X-mark columns bordered like data.
- **Alignment** — text left, numbers right, categorical/X-mark center, all vertical-center; headers match their column.
- **Number format** — one currency style per column (`$X.XM`/`$X.XB`/full); years `0`; counts `#,##0`; convert numeric-stored-as-text.
- **Widths/overflow** — long-text columns wide enough to read; numeric ~10–15 chars; X-mark ~3–5; no clipped or bleeding content (widen or wrap with matching row height, never merge).
- **Row heights** — consistent within a section (~32pt data); wrap rows tall enough; ~17pt spacers between sections.
- **File hygiene** — descriptive sheet name (not "Sheet1"); no phantom data or stray formatting past the populated range; frozen panes on header row + key-identifier column; landscape, `fitToWidth=1`, print area covering data + footer.

## Output

Report fixes in three tiers — **BLOCKING** (merged cells, crushed/truncated content, missing borders, wrong palette), **FIXED** (alignment, frozen panes, sheet name, string-typed numerics, print setup), **NIT** (minor polish) — each with cell/range + what changed. After QC, end with one line: `FORMATTING: CLEAN — verified.` or, if a fix couldn't be safely applied, name it.
