# Pipeline Data Request (synthetic template)

Send this to a company to enable a full pipeline analysis. Ask for **one row per
open opportunity**. Example values are fictional.

## Minimum (analysis cannot run without these)
| Field | One-line definition | Example |
|---|---|---|
| Opportunity / Deal ID | Unique row identifier | OPP-10231 |
| Account / Customer | Who the deal is with | Example Retailer Inc. |
| Amount / ARR ($) | Annual value of the deal | 48000 |
| Stage | Current pipeline stage | Proposal |

## Strongly recommended (unlocks the real analysis)
| Field | Definition | Example |
|---|---|---|
| % Probability | Likelihood of closing (for weighted/expected pipeline) | 0.40 |
| Expected Close Date | When it's forecast to close | 2026-09-30 |
| Create Date | When the deal opened (for aging) | 2026-02-01 |
| Deal Type | New Logo / Upsell / Cross-Sell | New Logo |
| Product / Line | Which product the deal is for | Core Platform |
| Deal Owner | Rep responsible (for by-rep view) | Rep One |
| Bookings Target | The period's quota/goal (one number, not per-row) | 2,000,000 |

## For conversion context (optional but valuable)
Provide the **historical CLOSED deals** too — one row each, with outcome
(Won / Lost / No-Decision) plus create & close dates — so win rate, competitive
win rate, and sales-cycle length can be computed alongside the live pipeline.

> Separate study: win/loss **reasons** (why deals were won or lost — loss-reason
> codes, named competitor, incumbent vendor) are a different analysis with
> different data; request those only if running a dedicated win-loss-reasons review.

> Note: don't request derived columns (year, LTM flag, type rollups) — the
> analysis generates those automatically.
