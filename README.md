# Investor Toolkit

Tools for investors: market mapping, target identification, building retention analytics, formatting files.

## Install

**Claude Cowork:** Customize → Personal plugins **+** → Create plugin → Add marketplace → paste `https://github.com/lyndsaykerwin/investor-toolkit` → Sync, then install the **investor-toolkit** plugin.

**Claude Code** — paste this to your agent:

```
Fork & clone github.com/lyndsaykerwin/investor-toolkit (e.g. gh repo fork --clone), then install skills/ where my agent loads skills (Claude Code: ~/.claude/skills/), keeping shared/ and templates/ alongside.
```

## The Sourcing Workflow

**theme-first-pass** — Activates on "do a first pass on [category]." Broadly scans one category and **creates the Market Map (JSONL data) and Market Brief (narrative)** that the rest of the workflow builds on.

**market-landscape** — Activates on "find competitors of [seed company]." **Takes the map from the first pass** and adds lookalike companies to it, each with how-it-was-found provenance.

**enrich-companies** — Activates on "enrich these names." **Takes the names now in the map** and fills founding year, headcount, funding, and feature flags via any connected firmographic provider or web search.

**render-market-map** — Activates on "render the market map." **Takes the enriched JSONL** and turns it into a banker-grade Excel landscape, regenerated fresh from the data each time.

The agent checks for relevant skills before any task.

## What's Inside

**Sourcing & market mapping**

- **sourcing-setup** — Capture firm profile + screening criteria
- **theme-first-pass** — Category scan → brief + market map
- **market-landscape** — Find lookalikes from a seed company
- **enrich-companies** — Fill firmographics (size, funding, age, feature flags)

**Rendering & formatting**

- **render-market-map** — JSONL → banker-grade Excel landscape (ships a script)
- **market-map-template-reader** — Read the Excel template's layout cheaply
- **finance-formatting** — Bring an Excel file up to dealbook standards
- **standardize-pnl** — Any P&L (PDF or messy Excel) → clean, audited Excel P&L with check-to-zero, margins, and follow-ups
- **excel-to-jsonl** — Turn any spreadsheet into agent-readable JSONL

**Analytics**

- **retention-analysis** — Gross/Net/Logo retention + ARR corkscrew from customer-level revenue
- **arr-to-bookings** — Quarterly new-logo vs. upsell ACV bookings from customer-level MRR/ARR (estimate; not for usage-based models)
- **pipeline-analysis** — Open-pipeline analysis (coverage, weighted pipeline, stage × product/type/owner, aging, concentration) from a deal-level export; adds conversion context when closed history is included

## Principles

**Delegation & verification** — Skills dispatch subagents to do the heavy lifting and verify the work before returning it.

**Personalized** — Setup flows capture your firm's investment criteria, so output is screened against your thesis, not a generic one.

**Human-in-the-loop** — You confirm direction at each stage. Nothing runs end-to-end unchecked, which keeps output deliberate instead of slop.
