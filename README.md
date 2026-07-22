# AIOLBO — All-In-One LBO

Generates a leveraged buyout (LBO) analysis for any (eligible) US public company from just a ticker: pulls real financials from SEC EDGAR, builds a formula-driven six-tab Excel model, produces an LLM-written narrative report with a deterministic feasibility score, and can diff two models (scenario vs. base case, or two companies).

Full design documentation, engineering decisions, and verification notes live in **[`DESIGN.md`](DESIGN.md)**.

## Repo Layout

```
backend/                 # FastAPI service (api.py) wrapping the pipeline modules
├── sec_edgar_test.py    #   Data retrieval: SEC EDGAR (financials) + Twelve Data (share price)
├── validator.py         #   Pass/degraded/fail verdict on fetched data before modeling
├── excel_generator.py   #   Six-tab, formula-driven .xlsx LBO model (openpyxl)
├── report_generator.py  #   Recalculates the workbook, scores feasibility, LLM narrative (BYOK)
└── comparison_tool.py   #   Deterministic diff of two workbooks + minimal AI commentary
frontend/                # Vite + React UI (ticker in → editable assumptions → model/report out)
docker-compose.yml       # Local dev: backend on :8000, frontend on :5173
```

## Pipeline

1. **Fetch** — SEC EDGAR CompanyFacts (revenue, operating income, D&A, debt, cash, capex, shares) with XBRL tag fallbacks; Twelve Data for current share price only. Free-tier friendly, no paywalled provider.
2. **Validate** — hard/soft requirement checks and disqualifying conditions (e.g., banks/insurance/REITs/utilities excluded by SIC code, negative EBITDA, no 10-K filings).
3. **Generate Excel** — six tabs (Assumptions, Sources & Uses, Operating Model, Debt Schedule, Returns, Sensitivity) with 14 editable inputs; everything else is a live formula using standard IB/PE color conventions.
4. **Report** — the workbook is genuinely recalculated (LibreOffice headless, or Excel COM on Windows dev), values read back with provenance tags, a 0–100 feasibility score computed in Python, and an LLM narrates the already-computed numbers.
5. **Compare** — two workbooks in, deterministic deltas out (scenario diff for same ticker, side-by-side for different tickers, auto-detected).

## Quick Start

### Docker (recommended)

```bash
# Set SEC_CONTACT_EMAIL and TWELVE_DATA_API_KEY in your environment or a .env file
docker-compose up
```

Backend on `http://localhost:8000`, frontend on `http://localhost:5173`.

### Manual

```bash
# Backend (Python 3.10+; LibreOffice required for report/comparison recalculation)
cd backend
pip install -r requirements.txt
uvicorn api:app --port 8000

# Frontend
cd frontend
npm install
npm run dev
```

The pipeline modules also run standalone from `backend/`:

```bash
python sec_edgar_test.py                                          # fetch + validate a ticker
python report_generator.py model.xlsx --provider anthropic --api-key sk-...
python comparison_tool.py base.xlsx scenario.xlsx
```

LLM providers supported: Anthropic, OpenAI, Gemini — user-supplied key per request, never stored.

## Status

All pipeline stages — data layer, validator, Excel generator, report generator, comparison tool — are complete and verified end-to-end against AAPL, CCL, and FIZZ, wrapped by a FastAPI backend and React frontend, deployed with the frontend on Vercel and the backend on Render. See [`DESIGN.md`](DESIGN.md) for known limitations and v2 candidates.
