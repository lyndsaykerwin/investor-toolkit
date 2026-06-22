---
name: sourcing-setup
description: First-run setup for the investment-sourcing skill bundle. Asks a short structured intake — investor type, target company size (employee bands), maximum prior funding raised, in/out industries, and an optional free-text catch-all — then writes shared/Firm_Profile.md, shared/Screening_Criteria.md, and shared/Investment_Themes.md so the other sourcing skills read real criteria instead of placeholders. Use when the user first installs the bundle, says "set up sourcing" or "configure my investment criteria", or when a sourcing skill reports that shared/Firm_Profile.md is still the template. Re-run anytime to update criteria.
---

# Sourcing setup (first run)

The bundle ships **template** versions of three shared files every sourcing skill reads:
`Firm_Profile.md`, `Screening_Criteria.md`, `Investment_Themes.md`. This skill fills them from a
short, mostly pick-list intake so the user can start sourcing immediately.

Run it when the user installs the bundle, says *"set up sourcing"* / *"configure my criteria"*, or a
skill reports a file *"is still the template"* (the marker `<!-- TEMPLATE — replace via sourcing-setup -->`
or bracketed placeholders are still present).

## The intake

Ask these in order. For every pick-list the user may **choose more than one**, and each offers
**"doesn't matter"** as a valid answer — record that literally as `Any` so downstream skills never
invent a filter the user didn't set.

1. **Investor type** — pick all that fit:
   - Growth equity
   - Private equity (control / buyout)
   - Venture
   - Strategic acquirer — you already own operating companies and want targets that fit alongside them

   If they pick **strategic acquirer**, ask one follow-up: *"Which companies or platforms you already
   own should a target complement?"* Capture the names — for a strategic buyer, "fit" means
   *complementary to those*, not just in-category.

2. **Target company size** — by employee headcount; pick all bands that fit, or "doesn't matter":
   `1–10` · `11–50` · `51–200` · `201–500` · `500+`

3. **Maximum prior funding raised** — *"surface companies that have raised less than…"*; pick the
   ceiling, or "doesn't matter":
   `under $1M` · `under $5M` · `under $10M` · `under $25M` · `under $50M` · `over $50M is fine`
   (If several are picked, use the largest as the ceiling.)

4. **Industries / categories** — which are in scope, and which are explicitly out? "No restriction"
   is a valid answer.

5. **Anything else that matters to your investment criteria?** — optional free text; skip if nothing.

Accept short answers. Infer nothing the user didn't say.

## Writing the files

- **Investor type + portfolio/platforms + the optional free text** → `Firm_Profile.md`. Set the
  **Investor type** field and the **Fit definition**: for a strategic acquirer, fit = complementary
  to the named portfolio; for a financial buyer (growth/PE/venture), fit = standalone attractiveness.
  Put the free-text answer under **Other criteria**.
- **Size bands + funding ceiling** → `Screening_Criteria.md` hard filters. A "doesn't matter" answer
  becomes the row value `Any`, so no candidate is screened out on it.
- **In / out industries** → `Investment_Themes.md` (in → **Active**, out → **Skip**), one line each.

Read each template first to preserve section order; replace the `<!-- TEMPLATE -->` marker and
placeholders; set `Last updated: <today>`. For anything answered "doesn't matter" or skipped, write
`Any` (or `> (not specified)`), never an invented value. Confirm before overwriting an already-filled
file.

## After setup

Tell the user in one line each: what you wrote, and that finer product-feature filters get captured
later, the first time `enrich-companies` runs. Next step: *"run theme-first-pass to deep-dive a
category."*

This skill only captures criteria — it does no research, and never edits the skills or `templates/`.
