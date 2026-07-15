# AIO LBO

A tool that generates leveraged buyout (LBO) analysis for public companies — pulling real financial data by ticker, building a formula-driven Excel model, and producing a narrative report. Supports scenario comparison: edit assumptions, generate a second model, and get a deterministic diff between the two.

## Scope (v1)

- **Public companies only**, ticker-based. No manual entry for private companies in v1.
- Primary deliverable priority: **~65% the downloadable Excel model, ~35% the narrative report.**
- The Excel file is the interactive core of the product — formula-driven, editable by the user, recalculates live. The report and comparison tool are generated *from* spreadsheet snapshots, not the other way around.

## Core pipeline

1. User enters a ticker.
2. **Data retrieval layer** fetches financials from SEC EDGAR (primary) and Twelve Data (current price only).
3. **Validator** checks data completeness (hard requirements / soft requirements / disqualifying conditions) and surfaces missing fields to the user directly (e.g. "these key fields are missing").
4. User can gap-fill missing fields or accept defaults; user can preview the pre-filled Assumptions tab before generating.
5. **Excel model** is generated (openpyxl) — five tabs, formula-driven, base case.
6. **Report generator** produces a narrative summary from that spreadsheet snapshot.
7. User can duplicate the spreadsheet, edit assumptions (deal terms and/or fundamentals) to create a scenario, and generate a second report from that snapshot.
8. **Comparison tool** takes both spreadsheets as input, deterministically computes deltas on headline metrics, and generates minimal AI commentary on what changed (no speculative "why," per current design).

## Data layer

### Primary source: SEC EDGAR (`data.sec.gov`)

**Why SEC EDGAR over FMP:** Financial Modeling Prep (FMP) was initially tested but abandoned — their free tier paywalls financial statement data (HTTP 402 errors) for anything below large-cap, which broke the "any public company" scope. SEC EDGAR has no paywall tiers — every filer's data is equally accessible, free, no API key required.

**What SEC EDGAR provides:**
- Revenue, operating income, D&A, total debt (current + noncurrent), cash, capex, shares outstanding, SIC code
- Via the CompanyFacts endpoint: `data.sec.gov/api/xbrl/companyfacts/CIK{cik}.json`
- Ticker → CIK resolution via SEC's `company_tickers.json` (cached locally)
- Rate limit: 10 requests/second
- Requires a descriptive User-Agent header (app name + contact email, read from `SEC_CONTACT_EMAIL` env var)

**EBITDA is calculated, not sourced as a single field** — `Operating Income + D&A`, for consistency across companies regardless of how they report.

### Secondary source: Twelve Data (current price only)

SEC EDGAR does **not** provide current share price or market cap (not filed data). Twelve Data fills this single gap:
- Free tier (~800 calls/day), API key from `TWELVE_DATA_API_KEY` env var
- Used **only** for current share price — not a full second data provider

### XBRL tag inconsistency (known, tested)

Unlike FMP's standardized fields, SEC EDGAR uses raw XBRL tags that vary across filers — different companies use different GAAP concept names for the same line item. The fetch layer uses **fallback tag lists** per concept, confirmed via live testing across 20 tickers spanning large/mid/small-cap, debt-heavy, excluded-sector, recent-IPO, and edge-case companies:

| Concept | Fallback tags (tried in order) |
|---------|-------------------------------|
| Revenue | `Revenues`, `RevenueFromContractWithCustomerExcludingAssessedTax` (some companies report both; merged) |
| D&A | `DepreciationDepletionAndAmortization`, `DepreciationAndAmortization`, or — if neither present — sum of `Depreciation` + `AmortizationOfIntangibleAssets` |
| Total debt | `LongTermDebtNoncurrent` + `LongTermDebtCurrent` (summed), falling back to `LongTermDebt` or `DebtLongtermAndShorttermCombinedAmount` |
| Cash | `CashAndCashEquivalentsAtCarryingValue` |
| Capex | `PaymentsToAcquirePropertyPlantAndEquipment` |
| Shares outstanding | `dei:EntityCommonStockSharesOutstanding` |

### Sector exclusion (SIC-based)

Exclusion is based on SIC code (from SEC submissions endpoint), confirmed via testing to correctly flag:
- Banks (SIC 6000-6199) — tested: JPM
- Insurance (SIC 6300-6411) — tested: AIG
- REITs (SIC 6798) — tested: O
- Utilities (SIC 4900-4999) — tested: DUK

These sectors also show missing `operating_income`/`capex`/`D&A` fields because they genuinely don't report an EBITDA-style operating structure, reinforcing why they're excluded rather than just defaulted.

### Edge cases (tested)

**No 10-K filings:** Companies with no 10-K filings (e.g., recent foreign private issuers that file Form 20-F under IFRS, like ARM) return zero fiscal years of data. This is treated as a clean disqualifying condition in the validator, not a partial-data case — 20-F/IFRS support is out of scope for v1.

**Total debt missing:** Treated as a soft requirement — if missing after all fallback tags, default to $0 with a visible flag rather than disqualifying. Genuinely low/no-debt companies are a normal, valid case. Testing confirmed debt tags reliably return plausible non-zero values for debt-heavy companies (CCL: $27B, AAL: $5B, IRM: $14B), so a missing result elsewhere is more likely a true zero than a tag gap.

**Ticker normalization:** Tickers with a "." (e.g., `BRK.B`) must be converted to "-" format (`BRK-B`) before SEC lookup, with fallback to the original format if that fails.

### Security note

API keys (`SEC_CONTACT_EMAIL`, `TWELVE_DATA_API_KEY`) must live server-side only (environment variables), never in client-side code or committed to the repo.

## Validator

Runs immediately after data fetch, before modeling starts. Returns a structured verdict (`pass` / `degraded` / `fail`) with a list of missing fields, not just a boolean.

- **Hard requirements** (can't build a model without): revenue, EBITDA components, share count, share price, total debt, cash.
- **Soft requirements** (falls back to defaults if missing): debt tranche detail, capex/D&A history, sector classification.
- **Disqualifying conditions** (should hard-flag or refuse): negative/near-zero EBITDA, 2+ missing hard requirements, stale financials, non-EBITDA-based sectors (financials, insurance, REITs, utilities — excluded from v1 or need separate ruleset).
- Missing fields are shown directly to the user, who can gap-fill them manually.

## Excel model structure

Five tabs, strict separation between **independent inputs** (editable) and **formulas** (never directly editable). This is the single most important design constraint in the project — it's what keeps the model internally consistent when a user edits assumptions, and it's why EBITDA, debt balances, and returns are never directly overridable, only their true underlying drivers are.

### The 14 independent inputs (everything else is a formula)
Revenue Growth Rate, EBITDA Margin %, Entry EV/EBITDA Multiple, Offer Premium %, Leverage Multiple, Interest Rate, Tax Rate, Capex %, D&A %, Change in NWC %, Exit Year, Exit EV/EBITDA Multiple, Transaction Fee %, Mandatory Amortization %.

### Tabs
1. **Assumptions** — all 14 inputs, plus API-fetched historicals (revenue, EBITDA, debt, cash, price, shares) pre-filled as reference/base data. Only tab a user should need to touch.
2. **Sources & Uses** — pure formulas. Purchase EV, fees, total uses vs. new debt raised + sponsor equity (the plug). Sources must equal Uses as an internal consistency check.
3. **Operating Model** — pure formulas. Revenue → EBITDA → EBIT → Net Income → Free Cash Flow, year by year through the exit year.
4. **Debt Schedule** — mostly formulas; Mandatory Amortization % is the only true input. Handles the cash sweep (excess FCF pays down debt automatically) and the beginning/ending balance chain. Known complexity: interest-on-debt-balance creates circularity (interest depends on balance, balance depends on cash left after interest) — v1 approach is a simplified non-circular approximation (interest on beginning balance only) before considering Excel's iterative calculation mode.
5. **Returns** — pure formulas. Exit EV → exit equity value → **IRR** (via `XIRR`) and **MOIC**.
6. **Sensitivity** — Excel Data Tables varying the deal-structure inputs (entry multiple, leverage, exit year, exit multiple) against IRR/MOIC.

Color convention (standard IB/PE modeling practice): blue = hardcoded input, black = formula. A third visual state is needed for fields where fetched real data has been knowingly overridden by the user, so the report generator and any reviewer can distinguish real data from hypothetical scenario data at a glance.

## Report generator

Takes a single spreadsheet snapshot (base case or a scenario) and produces a narrative: deal summary, Sources & Uses, debt paydown trajectory, projected FCF, exit scenario, headline IRR/MOIC, sensitivity highlights, and a plain-English verdict. Any user-overridden fundamental (e.g. EBITDA adjusted -10% from actual) must be disclosed explicitly in the narrative, not presented as fact.

## Comparison tool

Input: two spreadsheet snapshots (+ their generated reports).
Process: **deterministic** delta computation on headline metrics (IRR, MOIC, leverage path) — never inferred from prose.
Output: what changed, stated plainly. Commentary is intentionally minimal — states the delta, does not speculate on causal "why" beyond what's directly computable, per current product decision (the target user is PE-literate enough that the "what" carries the value; "why" is often self-evident or genuinely uncertain in multi-variable changes).

## Architecture (not yet finalized)

- **Backend:** Python, openpyxl for Excel generation.
- **Frontend:** shareable web tool (thin UI — ticker in, generate button, download out — rather than duplicating spreadsheet math in JS).
- **Hosting:** small always-on host (Render/Railway/Fly.io) vs. serverless (Vercel/Lambda) — undecided, serverless likely preferred for a low-traffic portfolio project given cost.
- **State:** stateless by default (ticker in, file out) — no accounts/database planned for v1.
- **Rate limiting:** needed given live API + real compute per request on a public link; not yet designed.

## Open items / not yet decided

- **Validator not yet built** — this is the next step, informed by the tested missing-field patterns from the 20-ticker SEC EDGAR test.
- **POOL still missing `ltd_current` field** after all fallbacks — needs manual inspection of raw CompanyFacts JSON (low priority, soft requirement).
- Backend hosting choice.
- Rate limiting implementation.
- Whether debt schedule circularity is handled via Excel iterative calc or a simplified approximation in v1. 
