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

Six tabs, strict separation between **independent inputs** (editable) and **formulas** (never directly editable). This is the single most important design constraint — it keeps the model internally consistent when a user edits assumptions.

### Status: Complete and Verified

All six tabs built and tested end-to-end against three tickers chosen specifically to stress different failure modes:

- **AAPL** — large-cap, clean data, validates happy path
- **CCL** — debt-heavy, weak free cash flow (negative FCF in projections), validates debt schedule edge cases
- **FIZZ** — small-cap, zero existing debt (defaulted), validates soft requirement handling

**Verification method:** Every tab was checked using real recalculation (LibreOffice headless conversion and/or actual Excel/Google Sheets), reading back computed values with openpyxl in `data_only=True` mode — not trusting formula strings or self-reported console output alone. This caught a real bug (see Key Engineering Decisions below) that would have shipped invisibly otherwise.

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

6. **Sensitivity** — 5×5 grid of Entry Multiple × Exit Multiple, showing MOIC and IRR for each combination. Uses fully self-contained per-cell formulas (not Excel Data Tables).

### Color Convention

Standard IB/PE modeling practice:

- **Blue** = hardcoded input
- **Black** = formula
- **Orange** = flagged value (implausible calculation replaced with fallback, e.g., COVID-distorted growth rate)
- **Yellow** = defaulted value (soft requirement missing, using system default)
- **Green** = base case highlight (Sensitivity tab center cell)

### Key Engineering Decisions

**Cross-sheet formulas must use direct cell references, not named ranges.** A significant bug was found and fixed: named ranges (e.g., `=NewDebtRaised`) resolved correctly when used on their home sheet, but silently evaluated to `$0` when referenced from a different sheet, in both Google Sheets and LibreOffice. This went undetected initially because an internal verification method was computing "expected" values via a separate parallel Python calculation rather than reading back real values from the actual generated file. **Fix:** Every cross-sheet formula reference now uses a direct `'Sheet Name'!$C$##` reference instead of a named range. Named ranges are still defined for documentation/readability but are no longer relied upon for any cross-sheet formula logic.

**Debt Schedule circularity solved via beginning-balance-only interest calculation** (not average of beginning/ending balance), avoiding Excel's iterative calculation setting entirely. This is a deliberate, disclosed simplification — it slightly overstates interest expense in years with heavy paydown, but produces a stable, one-directional formula chain that works reliably across spreadsheet engines, rather than depending on iterative calculation settings that behave inconsistently across Excel versions and other spreadsheet applications.

**Sensitivity tab scope limited to Entry Multiple × Exit Multiple only**, not Leverage Multiple or Growth Rate. Entry/Exit Multiple only affect deal pricing and exit valuation, which can be computed as self-contained per-cell formulas. Leverage and Growth Rate drive the entire year-by-year Debt Schedule chain, which can't be cleanly re-derived inside flat grid formulas without rebuilding the full 5-year paydown logic per scenario — documented as a v2 candidate.

**Does NOT use Excel's built-in Data Table (What-If Analysis) feature** for the same reason named ranges were avoided — that feature's `{=TABLE()}` mechanism is tightly coupled to Excel's calculation engine and unreliable across Google Sheets/LibreOffice. The Sensitivity grid instead uses fully independent, self-contained formulas per cell.

**IRR simplification in Sensitivity tab:** Since this model has no interim cash distributions to the sponsor (all free cash flow stays in the company for debt paydown until exit), IRR can be computed as `MOIC^(1/ExitYear) - 1` rather than requiring XIRR with a full cash flow array. The Returns tab still uses true XIRR with actual dates for the headline number; the Sensitivity grid uses the simplified formula per cell since XIRR isn't practical to replicate 25 times across a grid. The two methods produce results that match to within a fraction of a percent (confirmed via testing) — the gap is just XIRR's actual day-count vs. the simplified formula's whole-year assumption.

**Revenue Growth Rate default uses median (not mean)** of historical YoY % changes, with a sanity band (−15% to +25%) — if the calculated default falls outside that range, it's rejected and replaced with a flat 3% fallback, visibly flagged (orange fill + comment). This was necessary because simple-mean calculation produced a 140.8% default growth rate for CCL, driven by COVID-era revenue collapse/recovery distorting the historical average — confirmed via testing, not theoretical.

**Column widths are explicitly set (not auto-fit)** across all tabs, sized to comfortably fit large-cap currency figures (into the trillions) without displaying as `####`.

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
- **Excel generator: All six tabs complete** (Assumptions, Sources & Uses, Operating Model, Debt Schedule, Returns, Sensitivity)
- Mock data for testing without API credentials

### Not Yet Implemented

- Report generator (next planned step)
- Comparison tool
- Web frontend
- Web UI/backend architecture decisions (hosting, rate limiting)

### Known Limitations (v2 Candidates)

- Sensitivity tab only varies Entry/Exit Multiple, not Leverage Multiple or Growth Rate (would require rebuilding full debt paydown logic per grid cell)
- Beginning-balance-only interest calculation slightly overstates interest in heavy-paydown years (deliberate simplification for cross-platform reliability)
