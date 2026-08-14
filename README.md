# fund-factsheet-generator
VBA fund factsheet automation: pulls NAV data from an Access database, computes return, volatility and max drawdown through encapsulated classes, populates a bookmarked Word template, exports to PDF, prepares the client email in Outlook and logs every run.

## What it does

One click on the Control sheet runs the full chain:

Access → Excel/VBA → business classes → Word → PDF → Outlook → back to Access for logging

The system works for any fund in the database. Fund ID, benchmark ID, email recipient and output folder are all read from cells, so nothing is hardcoded.

## Metrics

- **Period returns:** 1-month, 3-month and 1-year simple returns
- **Annualized volatility:** standard deviation of daily returns, scaled by √252
- **Maximum drawdown:** largest peak-to-trough decline over the period
- **Relative return:** fund performance against its benchmark
- **Insufficient history** returns `-999`, displayed as `N/A` rather than a misleading figure

Sample output — AGG over six months: −2.22% return, 4.45% volatility, −3.72% max drawdown.

## Data sources

Daily closing prices from January to July 2026, cleaned and loaded into Access:

- [SPY — SPDR S&P 500 ETF Trust](https://www.nasdaq.com/market-activity/etf/spy/historical) — Nasdaq
- [AGG — iShares Core U.S. Aggregate Bond ETF](https://www.nasdaq.com/market-activity/etf/agg/historical) — Nasdaq
- [S&P 500 Index](https://www.wsj.com/market-data/quotes/index/SPX/historical-prices) — Wall Street Journal

Each file contains 123 daily observations per instrument.

## Repository structure

- `src/classes/` — `clsTimeSeries`, `ClsFund`, `clsBenchmark`
- `src/modules/` — `modMain` (orchestration), `modDatabase` (Access), `ModReporting` (chart, Word, Outlook)
- `deliverables/` — workbook, database, Word template, project report
- `samples/` — generated factsheets for SPY and AGG
- `docs/` — database schema, source price data, test screenshots

The VBA source is exported as plain text so it is readable and diffable in Git. The workbook remains the executable version.

## Design

Three classes, each with private data and public members only:

| Class | Responsibility |
|---|---|
| `clsTimeSeries` | Stores validated date/value pairs, returns the value at a given date or the closest available one, computes cumulative return, volatility and drawdown |
| `ClsFund` | Fund identity and NAV history, period returns, delegates risk metrics to its time series |
| `clsBenchmark` | Index identity and history, 1-year return, relative return against a fund |

`ClsFund` holds a `clsTimeSeries` rather than reimplementing the maths, and a `clsBenchmark` rather than duplicating index logic. Adding a metric means touching one class.

Database access goes through parameterized ADO commands. Every generation writes to a log table with timestamp, fund, status and PDF path — failures included, which is what makes the log useful.

## Usage

1. Open `deliverables/fundFactsheet.xlsm` and enable macros
2. On the Control sheet, fill in fund ID, benchmark ID, recipient address and output folder
3. Click the generate button

The PDF and chart PNG are written to the output folder, and the email opens in Outlook with the PDF attached.

## Requirements

Windows, with Excel, Access, Word and Outlook. VBA references: ADO, Word and Outlook object libraries.

## Known limitations

- Windows only — the automation relies on Component Object Model
- The Outlook email is prepared and displayed, not sent automatically: a client-facing document should not leave without human review
- The output folder must already exist
- One factsheet per run, no batch processing
