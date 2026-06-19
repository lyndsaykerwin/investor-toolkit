# Research Rules — single source of truth for the sourcing skills

> **Editing:** change a rule here (one place); screening filters → `shared/Screening_Criteria.md`;
> themes → `shared/Investment_Themes.md`; source hierarchy + citations → `shared/research_methodology.md`.

Each SKILL.md states its rules briefly and points here for the full text. These are guardrails, not
style — each prevents a specific bug that bites automated sourcing.

| Skill | Rules that apply |
|---|---|
| `theme-first-pass` | 1, 2, 3, 4, 6, 7, 8 |
| `market-landscape` | 1, 2, 3, 4, 6, 7, 8 |
| `enrich-companies` | 1, 2, 3, 4, 5, 6, 7, 8 |

**File architecture.** The per-theme deliverable set (Market Map JSONL + rendered xlsx + Market Brief,
with `_theme_meta` as the first JSONL record) is described in `shared/research_methodology.md` and
`skills/theme-first-pass/SKILL.md`. The schema lives in `templates/market_map_template.jsonl`. The
JSONL is source-of-truth; the xlsx is rendered from it, never authored directly.

---

## Rule 1 — Sourcing produces exactly two writes: JSONL records and brief updates

Records (appended/updated) in `<Theme>_Market_Map_vN.jsonl`, first record always `_theme_meta`; and
markdown in `<Theme>_Market_Brief.md`. `theme-first-pass` fills the brief sections research supports;
other skills feed the brief by adding facts to the JSONL, not by editing it. Coverage and
downstream-handoff notes go in each record's `coverage_note` (Rule 8) — not a sidecar markdown, not a
hand-authored xlsx (it's rendered), not a monolithic JSON dump.

## Rule 2 — Version up the customized file if it exists; copy the template if not

Before writing a per-theme file: if a version exists in `themes/<slug>/`, write the next version
(`v6`→`v7`) and `git mv` the prior into `Archive/`. If none exists, copy the canonical `templates/`
file, rename to the `<Theme>_<Artifact>` convention, and start the JSONL empty — schema only, no
template example records. Prevents both blowing away an operator's file and hand-rolling a new shape.

## Rule 3 — Every fact carries provenance and confidence

Write each agent-filled fact as `{status, value, _source, citation_date, citation_snippet,
confidence}`, and pair top-level scalars (`founded_year`, `fte`, `total_funding_usd`) with their
`*_citation` sibling. Match field names to `templates/market_map_template.jsonl` — don't invent new
ones. Inferred is not confirmed: "team of 10–20" means `fte: null` with the range in the snippet at
`confidence: partial`.

## Rule 4 — Missing data is reported as missing, never guessed

If no exact match: search name → operator-supplied website → known aliases. Still nothing → emit
`status: not_found` with `operator_call_flag: "needs-disambiguation"` and a note on what was tried.
Never substitute the closest-spelling company; never read "I couldn't find it" as "it doesn't exist."

## Rule 5 — Funding: single-record lookup only, never bulk/cached

Bulk firmographic endpoints silently return `null` on funding even when funding exists, so a bulk null
can wrongly pass an over-funded company. For any funding-gated field (`total_funding_usd`,
`last_round`), use the provider's single-record call or time-aware web search; bulk/cached is fine for
`founded_year`, `fte`, `hq`. Bulk-null funding is `status: unknown`, not `not_found`. Provenance:
`<provider>:single`, `websearch:<query>`, or a URL — bias to sources under ~18 months old.

## Rule 6 — Dedup against the live landscape before adding

Find the live JSONL (`ls themes/<slug>/*_Market_Map_v*.jsonl | sort -V | tail -1`), then grep each
candidate `name` and check `aliases`. Found → return as an *update* (new tags, `discovered_via`
signals, citations; keep the original `added_in`). Not found → new record with the current `added_in`.

## Rule 7 — Skills add, they don't overwrite operator-confirmed values

If a field is `_source: "operator-supplied"` or better-cited than what you'd write, leave it and flag
the conflict in `operator_call_flag` (e.g. `"enrichment-conflict:funding"`) — operator override is the
highest-trust source. For the brief, `theme-first-pass` fills agent sections (1, 2, 5) and leaves the
synthesis sections (3, 4, 6) to the operator; versioning up diffs and merges, never wipes prior edits.

## Rule 8 — Done means accuracy reached, not budget exhausted

Done when every requested field is sourced or honestly marked `not_found`/`unknown`, and the record's
`coverage_note` states what was attempted, what's missing, surfaces worth a deeper pass, and any MCP
that fell back to web search. For research skills, exhaust the search budget before calling a thin
result done — and say so. Leave brief synthesis sections as a visible `> TODO` rather than fabricating.

---

## Tools

Skills use whatever firmographic-data or search MCP is connected, and fall back to WebSearch/WebFetch
when none is — the skill still runs, output quality just drops. Record any fallback in the record's
`coverage_note` so the operator knows. Rule 5's single-record discipline holds for funding regardless
of which provider is connected.
