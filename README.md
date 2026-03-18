# Stock Valuation Program

A desktop application for estimating the intrinsic value of public equities using fundamental valuation models. Built with Python and CustomTkinter, the tool automates data collection from MOEX ISS API and Yahoo Finance, extracts financial statement data via LLM-powered PDF parsing, and produces fair-value estimates through multiple methodologies.

Valuation is the core analytical skill in M&A investment banking — every deal hinges on whether the target is fairly priced. This project implements the same models that analysts use on the desk: DCF, trading comparables (EV/EBITDA, P/E), and intrinsic value approaches (DDM, Residual Income).

## Valuation Models

### Discounted Cash Flow (DCF)

Projects free cash flows and discounts them at the weighted average cost of capital (WACC). WACC is computed from the Capital Asset Pricing Model (CAPM) using the equity risk premium approach:

$$WACC = \frac{E}{V} \cdot r_e + \frac{D}{V} \cdot r_d \cdot (1 - t)$$

where the cost of equity $r_e$ is derived from CAPM: $r_e = r_f + \beta \cdot ERP$.

### EV/EBITDA (Trading Comparables)

Applies sector-median EV/EBITDA multiples to the target company’s EBITDA, subtracts net debt, and divides by shares outstanding to arrive at an implied share price. Includes outlier filtering to remove distorted comps.

### Dividend Discount Model (Gordon Growth)

Estimates fair value as the present value of a perpetually growing dividend stream:

$$P_0 = \frac{D_1}{r_e - g}$$

where $D_1$ is the expected next-year dividend, $r_e$ is the required return, and $g$ is the sustainable growth rate.

### P/E Comparative

Benchmarks the stock against sector peers using the price-to-earnings ratio. Calculates implied price from sector-median P/E applied to the company’s trailing EPS.

### Residual Income Valuation (RIV)

Values equity as book value plus the present value of future residual income (earnings above the cost of equity):

$$V_0 = BVPS_0 + \sum_{t=1}^{T} \frac{(ROE_t - r_e) \cdot BVPS_{t-1}}{(1 + r_e)^t}$$

## Data Pipeline

The application supports both Russian (MOEX) and international equities with automatic data routing:

| Data Point | Russian Stocks | International Stocks |
|---|---|---|
| Price & dividends | MOEX ISS API | Yahoo Finance |
| Financials (EPS, BVPS, ROE) | Claude API (PDF parsing of IFRS reports) | Yahoo Finance |
| Beta, market data | MOEX ISS API | Yahoo Finance |
| Fallback | Manual input | Manual input |

The Claude API integration parses uploaded IFRS/МСФО annual reports in PDF format, extracting EPS, book value per share, ROE, and growth rates directly from financial statements — eliminating manual data entry for Russian stocks where structured API data is limited.

## Architecture

```
Stock-valuation-programm/
├── app.py                 # Desktop GUI (CustomTkinter, 2795 lines)
│                          # - Ticker input and PDF upload
│                          # - Tabbed results view (DCF, multiples, DDM, RIV)
│                          # - Interactive matplotlib charts
│                          # - Portfolio tracking with Excel export
├── stock_valuation.py     # Valuation engine (1357 lines)
│                          # - All 5 valuation models
│                          # - MOEX ISS API / yfinance data fetching
│                          # - Claude API financial statement parsing
│                          # - WACC and cost of equity calculations
├── config.json            # Model parameters and sector multiples
├── portfolio.xlsx         # Portfolio tracking spreadsheet
└── .gitignore
```

## Quick Start

```bash
# Clone
git clone https://github.com/S0uLL6/Stock-valuation-programm.git
cd Stock-valuation-programm

# Install dependencies
pip install customtkinter matplotlib requests beautifulsoup4 openpyxl numpy yfinance anthropic

# Set up API key for PDF parsing (optional — manual input works without it)
echo "ANTHROPIC_API_KEY=sk-ant-..." > .env

# Run
python app.py
```

### Usage

1. Enter a ticker symbol (e.g., `SBER` for Sberbank, `AAPL` for Apple)
2. Optionally upload an IFRS annual report PDF for automated financial data extraction
3. The app runs all applicable valuation models and displays results in a tabbed interface
4. Compare fair value estimates across models to identify mispricing
5. Track positions in the built-in portfolio manager (exports to Excel)

## Why This Matters for M&A

In investment banking, valuation is not a single number — it is a range derived from multiple approaches. A typical M&A pitch book includes a “football field” chart showing valuation ranges from DCF, trading comps, and precedent transactions. This project implements the first two pillars:

- **DCF** — the gold standard for intrinsic valuation in M&A, used to model the standalone value of acquisition targets
- **Trading comps (EV/EBITDA, P/E)** — the primary relative valuation method, used to benchmark offer prices against market consensus
- **DDM and RIV** — supplementary models particularly relevant for financial institutions (banks, insurance) where cash flow-based DCF is less applicable

The tool demonstrates practical competence in the analytical workflow that junior IB analysts perform daily: gathering financial data, building valuation models, and synthesizing multiple approaches into an investment thesis.

## Technologies

| Library | Purpose |
|---|---|
| CustomTkinter | Modern desktop GUI framework |
| matplotlib | Valuation charts and visualizations |
| requests / BeautifulSoup | MOEX ISS API data fetching |
| yfinance | International stock data |
| anthropic | LLM-powered PDF report parsing |
| openpyxl | Excel portfolio export |
| numpy | Numerical calculations |

## License

MIT License
