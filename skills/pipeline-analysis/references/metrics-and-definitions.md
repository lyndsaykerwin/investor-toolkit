# Metrics & Definitions

All examples below are synthetic (e.g. "Example SaaS Co."). Never store real
client data in this folder.

## Canonical fields & header synonyms
The engine fuzzy-matches incoming headers (case/space/punctuation-insensitive)
to these canonical fields. Add synonyms in `scripts/analyze_deals.py` → `SYNONYMS`.

| Canonical | Used for | Common header synonyms |
|---|---|---|
| `deal_id` | row id / counts | deal id, opportunity id, opp id, record id, account id |
| `account` | logo / concentration | account, customer, company, account name, logo |
| `arr` | every $ figure | arr, booking arr, acv, amount, deal value, $arr, value, tcv, mrr |
| `outcome` | stage / status | stage, deal stage, status, stage name, forecast category, sales stage |
| `probability` | weighted pipeline | probability, % probability, win %, likelihood, confidence |
| `close_date` | timing / aging | close date, closed date, expected close, won date, close period |
| `create_date` | aging, sales cycle | create date, created, open date, start date |
| `deal_type` | new-vs-upsell split | deal type, type, opportunity type, business type |
| `product` | product split | product, product family, product line, sku, module |
| `owner` | by-rep split | owner, deal owner, opportunity owner, rep, ae, sales rep |

> Win/loss-**reasons** fields (loss reason, competitor, incumbent) are NOT part
> of this skill — they belong to a separate win-loss-reasons study.

## Outcome classification
Raw stage/status text is bucketed:
- **open** — any live stage (discovery, demo, proposal, negotiation, …), or a row
  with probability strictly between 0 and 1, or a blank close date. *(The focus
  of this skill.)*
- **won / lost / no_decision** — closed outcomes, used only for conversion context
  when historical closed deals are included (won="won/closed won"; lost="lost";
  no_decision="no business awarded/no decision/abandoned").

## Tiers (what's unlocked)
- **Bronze** = value + stage/probability/close → total pipeline $, deal counts,
  pipeline by stage. (If closed deals present: win rate both ways.)
- **Silver** = + product / type / owner / create_date / probability → stage cross-
  tabs, weighted pipeline, concentration, aging; and (if closed) sales cycle and
  new-vs-upsell.

## Metric formulas
**Pipeline (primary):**
- **Open pipeline $** = Σ ARR of open deals.
- **Weighted / expected pipeline** = Σ(open ARR × probability). Needs an explicit
  probability column; do not infer from stage name.
- **Pipeline coverage** = open pipeline ÷ bookings target for the period, as a
  multiple (e.g. "3.2×"). 3–4× is a common healthy rule of thumb; state the rule,
  don't assert pass/fail without the company's own target.
- **Concentration** = top-N deals and their % of total pipeline.
- **Aging** = days open (today − create_date) or an Age column; flag stale deals.

**Conversion context (only if historical closed deals supplied):**
Let W,L,N = ARR (or count) of won, lost, no_decision.
- **Win rate** — show BOTH: competitive `W/(W+L)` and all-in `W/(W+L+N)`. Never
  show only the flattering one.
- **Competitive win rate** = `W/(W+L)` (head-to-head conversion; outcome-based).
- **Sales-cycle length** = mean & median of (close_date − create_date), won deals.
- **% New vs. Upsell** = won ARR where deal_type≈"new logo" ÷ total won ARR.

## Acceptance criteria (Analysis tab must emit each, value OR "BLOCKED — needs X")
1. Open pipeline $ and deal count
2. Weighted/expected pipeline (needs probability)
3. Pipeline coverage (needs target)
4. Pipeline by Stage × Product and Stage × Type (and Owner if present)
5. Top-deal concentration
6. Conversion context (win rate both ways / competitive win rate / sales cycle /
   % new-vs-upsell) — only if closed history is present, else flagged as needing it
7. "What this data can't tell you" + follow-up data request

## Guardrail notes
- If the source pre-computed a metric, recompute from raw rows and reconcile.
- Bookings ARR ≠ recognized revenue; label outputs as pipeline/bookings.
