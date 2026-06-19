---
name: market-landscape
description: Expand a market landscape — find lookalike companies from the seed(s) and thesis areas in the theme's Market Brief if it exists, or from what the user gives directly, and append them to the landscape JSONL  with how-found provenance.
---

> **Part of a human-in-the-loop chain — run in order: `theme-first-pass` → `market-landscape` → `enrich-companies`. Stop at each handoff for the operator's judgment; blindly accepting every recommendation produces slop, not a market map.**

# Expand the market landscape

Broaden a theme's landscape with more companies in the same space. Take the **seed companies and thesis areas from the theme's Market Brief** (Sections 6-7) if it exists, otherwise from what the user gives you — those are your anchors. Run the fanout below against them and append what you find to the landscape JSONL. The human picks the seeds and decides what to pursue; you do the mechanical reach, aiming for the **non-obvious** additions a generic "[category] tools" search misses — adjacent-vertical players, foreign-language entrants, younger companies at the same conferences. The obvious names are already in the landscape from first pass.

**Run this as a subagent-driven workflow:** subagents do the fanout, the core session QCs every result, and no company enters the landscape until it's verified as still operating under that name and still in the category — catching pivots, renames, acquisitions, and recent funding rounds before they land.

## The four techniques

Run as many as the ~5 WebSearch/WebFetch budget allows, against each seed and the thesis lens:

1. **Keyword bidders** — who buys ads on the seed's product terms.
2. **Shared conferences** — who sponsors or exhibits where the seed does (use the conference's own sponsor list when an exhibitor list is gated).
3. **"vs" / "alternatives to" pages** — who the seed is compared against (treat vendor-authored pages as biased).
4. **Bucket co-occurrence** — who shares the seed's landscape buckets; these usually update existing records with a new signal.

## Core rules (full text in `shared/research_rules.md`)

- **Provenance.** Each candidate carries a `discovered_via` entry: technique, signal, source URL + snippet. Entries stack across techniques for a stronger signal.
- **Missing data stays visible.** A name that surfaces but can't be confirmed as a real company is appended with `confidence: "low"` and `operator_call_flag: "needs-disambiguation"`.
- **Dedup against the live landscape.** `grep -i '"name": "<candidate>"'` the latest JSONL and its aliases: a match updates that record with a new `discovered_via`; a miss appends a new one.

## File flow

Read the latest `themes/<slug>/<Theme>_Market_Map_v*.jsonl`, write `v<N+1>` with all prior records preserved plus your additions, and `git mv` the old version into `Archive/`. Conform records to `templates/market_map_template.jsonl`, setting `non_obvious: true` with a one-line rationale where it fits.

**Deliverable:** the JSONL is the source of truth — render it to the operator-facing Excel via the `excel-to-jsonl` flow (`render-market-map`), then tell the user that after their review and sign-off, the next step is `/enrich-companies`.

## Done when

Three or more techniques ran (or the coverage note explains why not), every new or updated record has a sourced `discovered_via`, dedup has run, and the new version is written with the prior archived. A few strong non-obvious finds beat a long list of known names.
