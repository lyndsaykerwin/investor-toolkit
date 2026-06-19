# Screening Criteria
<!-- TEMPLATE — replace via sourcing-setup -->

> **Fill this in (or run the `sourcing-setup` skill).** The filters below are **examples** — replace
> the values with your own. Every filter here must be checkable from the public web before any
> outreach (company site, LinkedIn, news). Things you can only learn from a call or data room — like
> revenue — belong in post-outreach diligence, not this screen.

Last updated: [DATE]

## Section 1 — Hard filters (a candidate must pass ALL of these)

Run as a checklist. Any "no" means the company is out.

- [ ] **Business type** — [e.g. proprietary software/data product, not consulting or managed services]
- [ ] **Headcount** — [e.g. ≤200 employees (FTE)]
- [ ] **Funding** — [e.g. under $20M total raised. Confirm with an authoritative single-record lookup
      (any connected firmographic provider, or time-aware web search), never a bulk/cached one — see
      Rule 5 in `research_rules.md`]
- [ ] **Independence** — [e.g. independent — not a subsidiary, not already acquired]
- [ ] **Company age** — [e.g. founded 2023 or earlier — or leave open if you don't screen on age]
- [ ] **Category fit** — Operates in one of your Active themes (see `Investment_Themes.md`).

## Section 2 — Feature flags (scored signals, not gating)

Optional thesis-specific signals that make a company *more* interesting but don't disqualify it if
absent. Add 0–3 of your own — for example:

| Flag | What it signals | How to detect it |
|---|---|---|
| `[example-flag]` | [why this matters to your thesis] | [where to check — product/integrations page, marketplace listing] |

## Section 3 — Dedup against the live outreach tracker

Before surfacing a candidate as net-new, check it against your outreach pipeline.

- The outreach list is `outreach/tracker.jsonl` — one JSON record per company you've engaged.
- For each candidate: `grep -i '"name": "<candidate>"' outreach/tracker.jsonl` and check `aliases` arrays.
- Found → already in the pipeline; note it as a known name. Not found → net-new and eligible to surface.

*(This dedup checks the master outreach tracker — distinct from the per-theme landscape dedup in Rule 6
of `research_rules.md`, which checks the theme's Market Map JSONL.)*
