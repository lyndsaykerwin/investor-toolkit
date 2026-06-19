---
name: theme-first-pass
description: Broad first-pass scan of a market category that builds thesis vocabulary and surfaces the landscape into a Market Brief (markdown) plus a Market Map data file (JSONL). Use when the user asks to start a market landscape or a deep dive on a category.
---

> **Part of a human-in-the-loop chain — run in order: `theme-first-pass` → `market-landscape` → `enrich-companies`. Stop at each handoff for the operator's judgment; blindly accepting every recommendation produces slop, not a market map.**

# Theme first pass

A wide-net sweep of a category, optimized for recall. The seed pick, enrichment, and Excel render come in later steps.

**Run this as a subagent-driven workflow:** subagents do the discovery-surface sweeps, the core session QCs every result, and no company enters the landscape until it's verified as a real, still-operating company in the category.

**Deliverable returned to the user:** a Market Brief (markdown they read, ending in a readable market-leaders/seed-candidate table to start `market-landscape` from) and a Market Map — written as JSONL (the working file the chain edits) and rendered to a banker-grade Excel via the `render-market-map` skill so the user has a readable view. All saved under `themes/<slug>/`.

## Before you start

1. **Criteria.** If `shared/Firm_Profile.md` still shows the marker `<!-- TEMPLATE — replace via sourcing-setup -->`, run the `sourcing-setup` skill first — it's a **single short question** about what the firm invests in and any in/out industries, then it hands straight back here. The detailed screening filters are captured later, when `enrich-companies` first needs them. Most users land here first, so expect that one question on the first run.
2. **Theme — confirm scope before researching.** If the user named no category, or named one that's broad or vague, ask them for one line on what the category is — including who its users are and a sample company if they have one. A sharp one-line scope produces far better results than a broad label, so clarify first whenever the ask is high-level.
3. **Templates.** Output is built from `templates/Market_Brief_Template.md` and `templates/market_map_template.jsonl`, which ship with the bundle. If either is missing, restore it from the repo's `templates/` folder.

## Output — two files

1. **Market Brief** → `themes/<slug>/<Theme>_Market_Brief.md`. A human deliverable: plain language, jargon defined inline as if explaining to a grandparent, with specific examples from the last 12 months.
2. **Market Map** → `themes/<slug>/<Theme>_Market_Map_v1.jsonl`. The data file the rest of the pipeline reads. First record is `_theme_meta`; then one record per company.

Once the Market Map JSONL is written, run the `render-market-map` skill to produce the readable Excel alongside it.

## Core rules (full text in `shared/research_rules.md`)

- **Provenance.** Every fact carries `{status, _source, citation_date, citation_snippet, confidence}`.
- **Missing data stays visible.** When no exact match is found, keep the name with `confidence: "low"` and `operator_call_flag: "needs-disambiguation"`, plus a note on what was tried.
- **Source quality.** Note it inline — flag vendor-published or promotional pages as biased, and individual blogs or clickbait as low-credibility.
- **Coverage.** Finish on honest coverage: state what you reached and what you left for a deeper pass.

## File flow (version up, keep history)

For each output file: if a prior version exists, write the next version number and `git mv` the prior into `themes/<slug>/Archive/`; if none exists, create v1 from the template. For the JSONL, take only the schema from the template and begin with your own records.

## The `_theme_meta` record (first JSONL record)

Feeds the Market Brief's narrative sections (the xlsx renders only the company tables). Fill what research surfaces; leave a field `null` with a `_meta_citations` note when there's no signal.

```json
{
  "_theme_meta": true,
  "theme": "<theme name>",
  "theme_slug": "<filesystem-safe slug>",
  "definition": "<detailed but simple — explain it as you would to a grandparent, 1-2 sentences>",
  "use_case": "<who buys this and the problem it solves>",
  "strategic_value_for_firm": "<draft fit; leave null if no signal>",
  "incumbent_native_offering": "<what a major incumbent ships natively here, or 'N/A'>",
  "market_size_growth": "<one concrete number + source>",
  "key_trends_drivers": "<2-3 things that changed in the last 12 months>",
  "_meta_citations": [{"field": "<name>", "_source": "<url>", "citation_date": "YYYY-MM-DD", "citation_snippet": "<text>"}],
  "first_pass_at": "YYYY-MM-DD"
}
```

## The brief — fill what research surfaces

Fill a section when research gives you real content; otherwise leave the template's placeholder for a later pass. A short brief with three real sections beats a long one with fabricated filler.

- **1 — What it is.** Simple, grandparent-level definition. Fill when research surfaces enough.
- **2 — Why now.** A pressure point in the last 12 months (regulation, a major vendor move, a benchmark), cited.
- **3 — Commoditized vs. defensible.** Fill when you have at least 2 items each for "absorbed into OSS/platforms" and "still paid for."
- **4 — Unsolved problems.** Fill from practitioner threads or analyst reports, each bullet with an evidence cite.
- **5 — Platform / incumbent competition.** Two or three named platforms absorbing the category, one sentence each.
- **6 — Firm thesis fit.** Fill when Sections 3-5 have content and a real wedge surfaces. Synergy with your firm's other products is optional — include it only if relevant. Otherwise leave `> No clear thesis surfaced — revisit after enrichment`. After drafting this section, **ask the user which thesis areas resonate**, and record their steer here as operator-confirmed — that steer becomes the lens `market-landscape` and `enrich-companies` use to prioritize names and answer specific questions about companies.
- **Market leaders & seed candidates** (bottom of the brief). Close the brief with a short, readable table of the notable names surfaced this pass — market leaders and strong-looking fits — with columns: name, one-line what-they-do, why notable. This is what the user reads to pick a seed for `market-landscape`; it puts the JSONL records into human-readable form.
- **Header.** Theme status, today's date, and links to the JSONL (and the xlsx once it exists).

## Research — a light sweep

- **Budget:** about 5 WebSearch calls. If a niche category needs more, note it for the deeper pass and keep this run a sweep.
- **Stay current:** search the last 12 months and weight recent signal; rely on live search rather than prior knowledge, since the field moves faster than any training memory.
- **Discovery surfaces, in priority order:**
  1. Category education ("what is [category]") — feeds brief Sections 1-2 and the `_theme_meta` definition.
  2. Comparison / "alternatives" pages — leaders plus the long tail.
  3. Analyst reports (Gartner, Forrester, GigaOm, IDC).
  4. Standards bodies, associations, and conferences with public sponsor/exhibitor lists.
  5. Practitioner forums where buyers compare and complain — non-obvious names and Section 4 pain points.
- **Operator seeds win:** real-world signal the user brings (URLs, ad observations) gets `_source: "operator-supplied"`, `confidence: "high"`.

## Company records

Conform to `templates/market_map_template.jsonl`. At this stage fill `name`, `aliases`, `website`, `type`, `buckets`, `tags`, `what_they_do`, `source_note`, `confidence`, `added_in`; leave firmographics for `enrich-companies`. Flag clearly-too-big names (public ticker, 1000+ FTE, $1B+ valuation) as `firm_fit: "Pass-OverFunded"` with a rationale.

## Done when

1. 3+ discovery surfaces returned signal (or 5 searches ran with an honest coverage note).
2. The Market Map JSONL exists with `_theme_meta` first, then company records.
3. The Market Brief exists — real sections filled, the rest left as template placeholders, and a readable **Market leaders & seed candidates** table at the bottom.
4. Every unconfirmed name carries `confidence: "low"` and a `source_note` on what was tried.
5. The user has been asked which thesis areas resonate and which companies they'd like to use to seed /market-landscape, and their steer is recorded in Section 6.
