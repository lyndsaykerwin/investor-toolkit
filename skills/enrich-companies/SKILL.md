---
name: enrich-companies
description: Enrich companies in a landscape JSONL with the investment criteria from shared/Screening_Criteria.md — founded year, FTE, total funding, ownership/independence, and the feature flags that file defines. Use whenever sourcing surfaces names that need firmographics + feature verification, or when the operator says "enrich these names", "get firmographics for X", "check funding on Y", "fill in the screening fields", "verify the feature flags". Works with whatever firmographic source is connected (Apollo, Grata, SourceScrub, Crunchbase, PitchBook…) and defaults to time-aware web search when none is. Funding-gated fields always come from an authoritative single-record lookup, never a bulk/cached one (research_rules.md Rule 5). Every criterion is confirmed with a source or marked not_found/unknown after the fallback ladder. JSONL is the source of truth; renders the operator-facing Excel market landscape each run so it stays current at every stage.
---

# Enrich companies against the screening criteria

Fill the screening fields on partial landscape records. Third step in the human-in-the-loop chain: `theme-first-pass` → `market-landscape` → `enrich-companies`. Stop at each handoff for operator judgment.

**Run this as a subagent-driven workflow:** subagents enrich companies against the criteria, the core session QCs every result, and no field returns to the operator until it's verified with a source or honestly marked `not_found`/`unknown`.

**Editing:** screening filters → `shared/Screening_Criteria.md`; cross-skill rules → `shared/research_rules.md`; record schema → `templates/market_map_template.jsonl`.

## Rules

Applies Rules 1–8 of `shared/research_rules.md`. The one that bites hardest here: **Rule 5 — funding-gated fields need an authoritative single-record lookup, never a bulk/cached one.** Bulk endpoints (e.g. Apollo's) silently return null on funding; recording that null as `not_found` can promote an over-threshold company to outreach. Funding from a query you didn't run is `status: "unknown"`, never `not_found`.

## Data sources — use what's connected, default to web

No provider is required. Probe for a connected firmographic tool — Apollo, Grata, SourceScrub, Crunchbase, PitchBook, or similar — and use it; if two are connected, cross-check funding. If none is connected, default to web search. Provenance records the source actually used: `<provider>:single`, `<provider>:bulk`, `websearch:<query>`, or a URL.

**The web-search default is time-aware.** Funding rounds, headcount, and ownership go stale; today's date is in your environment. Bias to the most recent sources, stamp `citation_date`, and flag any funding or headcount figure older than ~18 months as possibly stale in the `coverage_note`.

## File flow

1. Find the live landscape: `ls themes/<slug>/*_Market_Map_v*.jsonl | sort -V | tail -1` → `vN`. Read all records.
2. Read `templates/market_map_template.jsonl` for field names + provenance shape. Don't invent fields.
3. Write `<Theme>_Market_Map_v<N+1>.jsonl` (all records, preserved + enriched); `git mv` `vN` into `Archive/`.
4. **Render the Excel:** run `render-market-map` on the new version. The xlsx is the operator-facing market landscape — it refreshes every run so the operator can review the enriched landscape at this stage, not only at the end.
5. Leave the brief to the operator + `theme-first-pass`.

Preserve `added_in`; add `enriched_at: "YYYY-MM-DD"`.

## If the screening filters aren't set yet

This is the first skill that applies the hard filters, so it's where they're worth capturing. If `shared/Screening_Criteria.md` still shows the `<!-- TEMPLATE -->` marker or lists no real filters, ask the user **one** question before screening:

> *"What are your must-pass filters for a company to be worth contacting? e.g. funding under $X, under N employees, software not services, independent (not a subsidiary), plus any must-have product features."*

Write their answer into `Screening_Criteria.md` as checklist rows, set `Last updated: <today>`, then continue. If they'd rather skip, screen on whatever filters the file already has and state the gap in the `coverage_note`.

## Fields to fill

Read the hard filters and feature flags from `shared/Screening_Criteria.md` — don't paraphrase from memory, and add nothing the file doesn't list (no invented age caps or vintage windows).

| Criterion type | Field(s) |
|---|---|
| Independence | `ownership_status`, `independence_note` |
| Founded | `founded_year`, `founded_year_citation` |
| Headcount | `fte`, `headcount_citation` |
| Funding (single-record only) | `total_funding_usd`, `last_round`, `funding_citation` |
| Each feature flag the criteria define | per-flag field(s), e.g. `multi_source` |

## Enrichment procedure

**founded_year / fte / hq:** connected provider (single-record or bulk) → LinkedIn company page → company About page → time-aware web search → else `not_found`.

**total_funding_usd / last_round (authoritative single-record):** the connected provider's single-record funding call → time-aware web search (Tracxn, Crunchbase, PitchBook public pages, recent press releases). Verified "no external funding" → `total_funding_usd: 0` with a sourced `funding_note`. Nothing after both → `not_found` with note. If only a bulk/cached lookup ran and funding came back null → `status: "unknown"`, never `not_found`.

**ownership_status:** the provider's ownership flag, then verify against a second source (About / /legal / LinkedIn) — provider ownership graphs carry errors. Two sources agree → `confidence: high`. Disagree → `status: "unconfirmed"`, `operator_call_flag: "ownership-conflict"`. No signal → default `independent`, `confidence: medium`.

**Feature flags (5-step ladder):** /integrations page → docs site → web search → partner marketplace / app-store listing → GitHub repo. All fail → `status: "unknown"` (not `none` — `none` means verified absent). Site blocks bots (e.g. intermittent HTTP 405) → drop to web-search snippets and note the block; never infer absence from your own fetch failure.

**multi_source** is a list, not a boolean: `{sources: [...], citation_url, citation_date, citation_snippet, confidence}`. Pull warehouse names verbatim from the integrations page; match existing records.

## Done when

For every input company, each screening field is confirmed (`_source` + `citation_snippet`) or honestly `not_found`/`unknown` with a note; funding came from an authoritative single-record source (or is `unknown`); any conflict with an operator-supplied value is surfaced via `operator_call_flag`, not overwritten; the output carries a `coverage_note` stating what enriched, what missed, where bot-blocks or stale data hit. A budget-exhausted run is done if the unknowns are honest. The refreshed xlsx is the operator's review surface for this stage.

## Smoke test

Ask the orchestrator: *"enrich <company> in the <theme> market map."* Expect: reads `Screening_Criteria.md` and the live landscape; uses the connected firmographic source (or time-aware web search) and pulls funding from an authoritative single-record lookup, never a bulk one; verifies ownership with a second source; writes `v<N+1>`, moves the prior to `Archive/`, and re-renders the xlsx.

Red flags: a bulk/cached lookup's null funding recorded `not_found` (Rule 5); overwriting an `operator-supplied` field (Rule 7); writing markdown instead of JSONL (Rule 1); substituting a closest-spelling company (Rule 4).
