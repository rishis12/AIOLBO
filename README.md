# AIO LBO

A tool that generates leveraged buyout (LBO) analysis for public companies — pulling real financial data by ticker, building a formula-driven Excel model, and producing a narrative report. Supports scenario comparison: edit assumptions, generate a second model, and get a deterministic diff between the two.

## Scope (v1)

- **Public companies only**, ticker-based. No manual entry for private companies in v1.
- Primary deliverable priority: **~65% the downloadable Excel model, ~35% the narrative report.**
- The Excel file is the interactive core of the product — formula-driven, editable by the user, recalculates live. The report and comparison tool are generated *from* spreadsheet snapshots, not the other way around.

## Core Pipeline

1. User enters a ticker.
2. **Data retrieval layer** fetches financials from SEC EDGAR (primary) and Twelve Data (current price only).
3. **Validator** checks data completeness (hard requirements / soft requirements / disqualifying conditions) and surfaces missing fields to the user directly.
4. User can gap-fill missing fields or accept defaults; user can preview the pre-filled Assumptions tab before generating.
5. **Excel model** is generated (openpyxl) — formula-driven, base case.
6. **Report generator** produces a narrative summary from that spreadsheet snapshot.
7. User can duplicate the spreadsheet, edit assumptions to create a scenario, and generate a second report.
8. **Comparison tool** takes both spreadsheets as input, deterministically computes deltas on headline metrics, and generates minimal AI commentary on what changed.

## Data Layer

### Primary Source: SEC EDGAR

**Why SEC EDGAR over FMP:** Financial Modeling Prep (FMP) was initially tested but abandoned — their free tier paywalls financial statement data (HTTP 402 errors) for anything below large-cap, which broke the "any public company" scope. SEC EDGAR has no paywall tiers — every filer's data is equally accessible, free, no API key required.

**What SEC EDGAR provides:**

- Revenue, operating income, D&A, total debt (current + noncurrent), cash, capex, shares outstanding, SIC code
- Via the CompanyFacts endpoint: `data.sec.gov/api/xbrl/companyfacts/CIK{cik}.json`
- Ticker to CIK resolution via SEC's `company_tickers.json` (cached locally)
- Rate limit: 10 requests/second
- Requires a descriptive User-Agent header (app name + contact email, read from `SEC_CONTACT_EMAIL` env var)

**EBITDA is calculated, not sourced as a single field** — `Operating Income + D&A`, for consistency across companies regardless of how they report.

### Secondary Source: Twelve Data (current price only)

SEC EDGAR does **not** provide current share price or market cap (not filed data). Twelve Data fills this single gap:

- Free tier (~800 calls/day), API key from `TWELVE_DATA_API_KEY` env var
- Used **only** for current share price — not a full second data provider

### XBRL Tag Inconsistency

Unlike FMP's standardized fields, SEC EDGAR uses raw XBRL tags that vary across filers — different companies use different GAAP concept names for the same line item. The fetch layer uses **fallback tag lists** per concept, confirmed via live testing across 20 tickers:

| Concept | Fallback Tags (tried in order) |
| ------- | ------------------------------ |
| Revenue | `Revenues`, `RevenueFromContractWithCustomerExcludingAssessedTax` |
| D&A | `DepreciationDepletionAndAmortization`, `DepreciationAndAmortization`, or sum of `Depreciation` + `AmortizationOfIntangibleAssets` |
| Total Debt | `LongTermDebtNoncurrent` + `LongTermDebtCurrent` (summed), or `LongTermDebt`, or `DebtLongtermAndShorttermCombinedAmount` |
| Cash | `CashAndCashEquivalentsAtCarryingValue` |
| Capex | `PaymentsToAcquirePropertyPlantAndEquipment` |
| Shares Outstanding | `dei:EntityCommonStockSharesOutstanding` |

### Sector Exclusion (SIC-based)

Exclusion is based on SIC code (from SEC submissions endpoint), confirmed via testing:

- Banks (SIC 6000-6199) — tested: JPM
- Insurance (SIC 6300-6411) — tested: AIG
- REITs (SIC 6798) — tested: O
- Utilities (SIC 4900-4999) — tested: DUK

These sectors also show missing `operating_income`/`capex`/`D&A` fields because they genuinely don't report an EBITDA-style operating structure.

### Edge Cases

**No 10-K filings:** Companies with no 10-K filings (e.g., recent foreign private issuers that file Form 20-F under IFRS, like ARM) return zero fiscal years of data. This is treated as a clean disqualifying condition in the validator — 20-F/IFRS support is out of scope for v1.

**Total debt missing:** Treated as a soft requirement — if missing after all fallback tags, default to $0 with a visible flag rather than disqualifying. Genuinely low/no-debt companies are a normal, valid case.

**Ticker normalization:** Tickers with a "." (e.g., `BRK.B`) must be converted to "-" format (`BRK-B`) before SEC lookup, with fallback to the original format if that fails.

### Security Note

API keys (`SEC_CONTACT_EMAIL`, `TWELVE_DATA_API_KEY`) must live server-side only (environment variables), never in client-side code or committed to the repo.

## Validator

Runs immediately after data fetch, before modeling starts. Returns a structured verdict (`pass` / `degraded` / `fail`) with a list of missing fields.

- **Hard requirements** (can't build a model without): revenue, EBITDA components, share count, share price, total debt, cash
- **Soft requirements** (falls back to defaults if missing): debt tranche detail, capex/D&A history, sector classification
- **Disqualifying conditions**: negative/near-zero EBITDA, 2+ missing hard requirements, stale financials, non-EBITDA-based sectors

Missing fields are shown directly to the user, who can gap-fill them manually.

## Excel Model Structure

Five tabs, strict separation between **independent inputs** (editable) and **formulas** (never directly editable). This is the single most important design constraint — it keeps the model internally consistent when a user edits assumptions.

### The 14 Independent Inputs

Everything else is a formula:

1. Revenue Growth Rate
2. EBITDA Margin %
3. Entry EV/EBITDA Multiple
4. Offer Premium %
5. Leverage Multiple
6. Interest Rate
7. Tax Rate
8. Capex %
9. D&A %
10. Change in NWC %
11. Exit Year
12. Exit EV/EBITDA Multiple
13. Transaction Fee %
14. Mandatory Amortization %

### Tabs

1. **Assumptions** — All 14 inputs, plus API-fetched historicals (revenue, EBITDA, debt, cash, price, shares) pre-filled as reference. Only tab a user should need to touch.

2. **Sources & Uses** — Pure formulas. Purchase EV, fees, total uses vs. new debt raised + sponsor equity (the plug). Sources must equal Uses as an internal consistency check.

3. **Operating Model** — Pure formulas. Revenue → EBITDA → EBIT → Net Income → Free Cash Flow, year by year through the exit year.

4. **Debt Schedule** — Mostly formulas; Mandatory Amortization % is the only true input. Handles the cash sweep (excess FCF pays down debt automatically). Interest calculated on beginning balance only to avoid circularity.

5. **Returns** — Pure formulas. Exit EV → exit equity value → IRR (via `XIRR`) and MOIC.

6. **Sensitivity** — Excel Data Tables varying deal-structure inputs (entry multiple, leverage, exit year, exit multiple) against IRR/MOIC.

### Color Convention

Standard IB/PE modeling practice:

- **Blue** = hardcoded input
- **Black** = formula
- **Orange** = placeholder (needs wiring to real data)

## Report Generator

Takes a single spreadsheet snapshot (base case or a scenario) and produces a narrative: deal summary, Sources & Uses, debt paydown trajectory, projected FCF, exit scenario, headline IRR/MOIC, sensitivity highlights, and a plain-English verdict.

Any user-overridden fundamental (e.g. EBITDA adjusted -10% from actual) must be disclosed explicitly in the narrative, not presented as fact.

## Comparison Tool

- **Input:** Two spreadsheet snapshots (+ their generated reports)
- **Process:** Deterministic delta computation on headline metrics (IRR, MOIC, leverage path) — never inferred from prose
- **Output:** What changed, stated plainly. Commentary is intentionally minimal — states the delta, does not speculate on causal "why" beyond what's directly computable

## Architecture

- **Backend:** Python, openpyxl for Excel generation
- **Frontend:** Shareable web tool (thin UI — ticker in, generate button, download out)
- **Hosting:** TBD (Render/Railway/Fly.io vs. serverless)
- **State:** Stateless by default (ticker in, file out) — no accounts/database planned for v1
- **Rate limiting:** Needed given live API + real compute per request

## Current Status

### Implemented

- SEC EDGAR data fetching with XBRL tag fallbacks
- Twelve Data integration for current price
- Validator with pass/degraded/fail status
- Excel generator: Assumptions, Sources & Uses, Operating Model, Debt Schedule tabs
- Mock data for testing without API credentials

### Not Yet Implemented

- Returns tab
- Sensitivity tab
- Report generator
- Comparison tool
- Web frontend
