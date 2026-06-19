---
name: sourcing-setup
description: First-run setup for the investment-sourcing skill bundle. Asks ONE short question about the firm's investment criteria and any in/out industries, then writes shared/Firm_Profile.md, shared/Screening_Criteria.md, and shared/Investment_Themes.md so the other sourcing skills read real criteria instead of placeholders. Use when the user first installs the bundle, says "set up sourcing" or "configure my investment criteria", or when a sourcing skill reports that shared/Firm_Profile.md is still the template. Re-run anytime to update criteria.
---

# Sourcing setup (first run)

The bundle ships **template** versions of three shared files every sourcing skill reads:
`Firm_Profile.md`, `Screening_Criteria.md`, `Investment_Themes.md`. This skill fills them from a
**single short question** so the user can start sourcing immediately. Finer detail is requested
later, the first time a downstream skill actually needs it — not all up front.

Run it when the user installs the bundle, says *"set up sourcing"* / *"configure my criteria"*, or a
skill reports *"Firm_Profile.md is still the template"* (a file is still a template if it has the
line `<!-- TEMPLATE — replace via sourcing-setup -->` or bracketed placeholders like `[Your firm]`).

## The one question

Ask exactly one, in plain English:

> *"In a few sentences — what does your firm invest in (the kinds of companies and any must-pass
> criteria you care about), and are there industries or categories you specifically want in or out
> of scope?"*

Accept whatever they give, short or long. Infer nothing they didn't say.

## Writing the files

From that single answer, fill only what the user actually stated:

- **What they buy / their thesis** → `Firm_Profile.md` (prose).
- **Any must-pass filters they named** (funding ceiling, headcount range, software vs. services,
  independence, must-have features) → `Screening_Criteria.md` as checklist rows. Skip revenue — it
  isn't reliably findable for private companies pre-outreach.
- **Any in/out industries or categories** → `Investment_Themes.md` (in → **Active**, out → **Skip**),
  one line each.

Read each template first to preserve section order; replace the `<!-- TEMPLATE -->` marker and
placeholders; set `Last updated: <today>`. For anything the user didn't mention, leave a short
`> (not yet specified)` note rather than inventing a thesis, filter, or theme. Don't overwrite an
already-filled file without showing its current content and confirming.

## After setup

Tell the user, one line each: what you wrote, and that the **detailed screening filters** get
captured later — the first time `enrich-companies` runs and needs them. Next step: *"run
theme-first-pass to deep-dive a category."*

This skill only captures criteria — it does no research, and never edits the skills or `templates/`.
