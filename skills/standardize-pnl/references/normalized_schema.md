# Normalized P&L JSON schema

Stage 1 produces exactly this. `scripts/build_pnl_workbook.py` consumes it. See `fixtures/example_pnl.json` for a complete (synthetic) worked example. All examples below are fictional — never embed real company data in this skill.

```jsonc
{
  "meta": { "company": "...", "statement": "Profit & Loss", "basis": "Accrual", "source": "..." },

  "periods": [                       // left → right = earliest → latest
    { "id": "FY2023", "label": "FY2023", "full_year": true },
    { "id": "FY2024YTD", "label": "2024 YTD\n(Jan 1–Jun 30)", "full_year": false }
  ],

  "lines": [                         // in display order, top to bottom
    { "role": "section", "label": "INCOME" },
    { "role": "group_header", "label": "Sales & Marketing" },

    { "role": "detail", "key": "rev_sub", "label": "Subscription revenue",
      "indent": 2, "values": { "FY2023": 1000000, "FY2024YTD": 600000 } },

    { "role": "subtotal", "key": "income_total", "label": "Total Income",
      "members": ["rev_sub","rev_serv"], "level": 1,
      "stated": { "FY2023": 1060000 } },

    { "role": "computed", "key": "gross_profit", "label": "GROSS PROFIT",
      "formula": ["income_total","-","cogs_total"], "level": 0, "major": true,
      "stated": { "FY2023": 915000 } },

    { "role": "blank" }
  ],

  "output": {                        // how the clean Output sheet is assembled
    "revenue": "income_total",       // REQUIRED — must resolve to a line
    "cogs": "cogs_total",            // key or null — NULL when the source has no COGS (Gross Profit = Revenue)
    "gross_profit": "gross_profit",  // key or null — NULL → builder computes Revenue − COGS
    "opex": [                        // summary cost lines, in display order; EXCLUDE interest & D&A
      { "label": "Sales & Marketing", "key": "sm_total" },
      { "label": "Other operating", "keys": ["ga_total","other_total"] }  // sum several
    ],
    "interest": "interest",          // key or null
    "da": "da_amort",                // D&A key or null
    "income_taxes": "income_tax",    // key or null — row shown only if present
    "other_income": "other_inc",     // key or null
    "other_expense": "other_exp",    // key or null — non-interest/non-D&A below-the-line costs
    "net_income": "net_income"       // REQUIRED
  },

  "followups": [
    { "bucket": "classification",   "text": "..." },
    { "bucket": "functional_split", "text": "..." },
    { "bucket": "comparability",    "text": "..." }
  ]
}
```

## Roles
| role | meaning | builder behavior |
|---|---|---|
| `section` | band header (INCOME, OPERATING EXPENSES) | grey label row |
| `group_header` | account group label (no values) | bold sub-label |
| `detail` | a real line with numbers | **hard-typed** values (the verbatim transpose) |
| `subtotal` | sum of `members` (detail or other subtotal keys) | live `=SUM` formula + audit check vs `stated` |
| `computed` | `formula` = list of keys and `+`/`-` tokens | live formula + audit check vs `stated` |
| `blank` | spacer | empty row |

## Rules for Stage 1
- **Verbatim = as printed** (the statement's displayed precision, normally 2 dp), not the raw stored float. Copy values exactly; the audit catches typos: if a detail is wrong, its subtotal formula won't match `stated`.
- **Members may be detail OR other subtotal/computed keys** — nesting is supported (the builder resolves a subtotal whose member is itself a subtotal). This is the normal QuickBooks/NetSuite shape.
- **Every subtotal/total/computed line needs a `stated`** for each period it exists — that is the check. Omit a period key if the line is absent that period. A **blank** value means "not reported that period," not zero.
- **`major: true`** only on Gross Profit and Net Income in `lines` (Revenue is mapped via `output`; EBITDA is built by the script).
- **Partial periods**: set `full_year: false`. The script labels them, flags amber, and marks YoY growth into/out of them `n/m`. A single period emits no growth rows.
- **EBITDA** is computed as Gross Profit − (`output.opex` lines), which by construction exclude interest and D&A. Keep interest, D&A, income taxes, and other income/expense OUT of `output.opex`; expose them via `output.interest` / `output.da` / `output.income_taxes` / `output.other_income` / `output.other_expense`.
- **Net Income bridge** the script enforces: Net Income = EBIT − interest − income_taxes + other_income − other_expense. If your statement has taxes or stray below-the-line expenses, map them to those keys — do NOT fold them into `other_income` as a negative (that misleads a reader).
- **No COGS?** Set `output.cogs` and `output.gross_profit` to `null`. Gross Profit then equals Revenue — honest, not invented. Flag in a `classification` follow-up whether any opex line (hosting, delivery contractors, payment fees) is really cost-of-revenue.
- **Parent account carrying its own postings** (a group row that has both a direct amount AND child accounts): model the parent's direct amount as an extra `detail` member of its subtotal, so the subtotal still foots. Don't drop it (loses money) or treat the parent row as the subtotal (double-counts).
- **Source subtotal that doesn't foot to its own children** (real source error): keep the leaf details (which foot), drop the broken presentation subtotal, and capture any orphan parent amount as its own verbatim line. Never plug a number.

See `references/extraction_playbook.md` for handling messy real-world sources (monthly→annual aggregation, isolating the income statement from full financials, trap columns, restated comparatives).

## Follow-up buckets (be specific and actionable — not methodology)
- **classification** — lines that may sit in the wrong place (e.g., delivery/implementation consulting or hosting that belongs in COGS; processing fees in opex). Name the account and the $ impact.
- **functional_split** — what's needed to split S&M / R&D / G&A: usually a headcount/department map to allocate pooled payroll; flag if there's no R&D account at all. Never invent an allocation.
- **comparability** — anything blocking a clean YoY: a stub period needing full-year actuals, a basis change, a missing year.
