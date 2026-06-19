---
name: excel-to-jsonl
description: >-
  Convert any Excel (.xlsx/.xlsm) or CSV file into a JSONL "sidecar" — a plain-text
  file with one JSON record per data row, plus a leading _meta record describing the
  sheet and its columns. Use this whenever you want an agent-readable companion to a
  spreadsheet so an AI can read the data WITHOUT opening the .xlsx and exploring it
  cell-by-cell (saves tokens and avoids guesswork). Trigger this skill whenever the
  user says things like "make a JSON/JSONL version of this spreadsheet", "save a JSON
  sidecar next to this Excel", "turn this xlsx into JSONL", "give the agent a readable
  copy of this workbook", "I want a JSON twin of my template", or whenever a spreadsheet
  is about to be handed to another agent/skill as a data source. Generalizes the
  market-map-template JSONL pattern so it works for ANY spreadsheet, not just market maps.
---

# Excel → JSONL sidecar

## What this does and why

A spreadsheet is easy for people but slow for agents: to understand it, an agent has
to open the file, hunt for the header row, and infer what each column means — every
single time. A **JSONL sidecar** fixes that. It's one small text file saved next to the
spreadsheet, where:

- the **first line** is a `_meta` record: the source filename, which sheet, when it was
  generated, where the header row is, the row count, and every column with its inferred
  type and a sample value;
- **every line after** is one data row, turned into a JSON object keyed by the column
  headers.

Because it's "JSON Lines" (one record per line, not one giant array), an agent can read
the `_meta` line to understand the shape, then stream rows — and you can append new rows
later without rewriting the file. This is the same idea behind the market-map template's
`.jsonl`, generalized so it works for any workbook.

## How to run it

The conversion is deterministic, so use the bundled script rather than transcribing
cells by hand (hand-copying is slow and error-prone):

```bash
python3 scripts/excel_to_jsonl.py <path-to-file.xlsx>
```

The script needs the `openpyxl` library to read `.xlsx`/`.xlsm` files. If you get a
`ModuleNotFoundError: openpyxl`, install it once with `pip install openpyxl` (CSV files
need no extra library).

It writes `<same-name>.jsonl` right next to the source and prints the row count. That's
the whole happy path. Reach for the options only when the defaults need a nudge:

| Option | When you need it |
|---|---|
| `--sheet "Name"` | The workbook has several sheets and you want a specific one |
| `--all-sheets` | Convert every sheet; each becomes `<name>.<sheet>.jsonl` |
| `--header-row N` | Auto-detection picked the wrong row; force the header (1-based) |
| `--out PATH` | Write somewhere other than next to the source |
| `--formulas` | Keep formula text (`=A1+B1`) instead of the last-saved value |
| `--max-rows N` | Sample just the first N rows of a huge file |

By default it reads the **last-saved values** (not the formulas), since the sidecar is
meant to describe the *data*. If a formula-driven file has never been opened/saved in
Excel, those cached values can be blank — open and save the file once, or pass
`--formulas` to capture the formula text instead.

## What the output looks like

**Input** (a sheet with a title banner, then a header row, then data):

| Customer | Region | ARR |
|---|---|---|
| Acme Co | West | 120000 |
| Beta LLC | East | 85000 |

**Output** (`file.jsonl`):

```
{"_meta": true, "source_file": "file.xlsx", "sheet": "Sheet1", "generated": "2026-06-17", "header_row": 1, "row_count": 2, "columns": [{"name": "Customer", "type": "string", "sample": "Acme Co"}, {"name": "Region", "type": "string", "sample": "West"}, {"name": "ARR", "type": "number", "sample": 120000}]}
{"Customer": "Acme Co", "Region": "West", "ARR": 120000}
{"Customer": "Beta LLC", "Region": "East", "ARR": 85000}
```

## Good to know (the edge cases it already handles)

- **Title/banner rows above the real headers** — common in finance models — are skipped
  automatically; the script finds the row that's wide and mostly text. If it guesses
  wrong, override with `--header-row N`.
- **Dates** become ISO strings (`2026-01-31`) so they sort and parse cleanly.
- **Whole numbers** like `3.0` are written as `3`; real decimals are preserved.
- **Blank rows** are dropped; **empty cells** become `null`.
- **Blank or duplicate column headers** get safe names (`column_3`, `Region_2`).
- **Multi-sheet workbooks**: with no flag it picks the most data-rich sheet and tells you
  which; use `--sheet` or `--all-sheets` to control it.

## Scope

This produces a flat, one-record-per-row sidecar — the faithful, generalizable twin of
any spreadsheet. It does **not** invent nested structure or domain fields (like a market
map's per-company citation objects); that kind of enrichment is specific to each dataset
and belongs to the skill that curates it. This skill's job is the reliable, boring,
correct conversion that those richer workflows can build on top of.
