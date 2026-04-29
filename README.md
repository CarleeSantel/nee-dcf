# NextEra Energy (NEE) — DCF Model

A Python-based discounted cash flow model for NextEra Energy, the largest regulated utility and renewable energy operator in North America. The script pulls live financial and market data with yfinance, builds a 5-year projection, calculates WACC from first principles, and exports a formatted, print-ready Excel workbook with a sensitivity table and full assumptions log.

---

## Methodology

### Free Cash Flow Proxy

Standard yfinance FCF (Operating CF − Capex) materially understates NEE's economic earnings because the company deploys $9B+ per year in growth capex to build regulated-return assets. That capex is not a drain on value — it creates rate base that generates future regulated revenue. Using reported FCF would penalize the company for its own growth.

Instead, the model uses **NOPAT as the FCF proxy**:

```
NOPAT = EBIT × (1 − Effective Tax Rate)
```

This normalizes for growth capex and is internally consistent with the revenue growth projections — new capital deployed generates new revenue, and the EBIT margin assumptions capture the earnings from those assets.

### WACC

WACC is built from market and financial data sourced at runtime:

```
Ke = Risk-Free Rate + β × Equity Risk Premium   (CAPM)
Kd = Avg. Interest Expense / Total Debt          (implied cost)
WACC = Ke × We + Kd × (1 − t) × Wd
```

| Parameter | Value | Source |
|---|---|---|
| Risk-Free Rate | 4.3% | 10-year US Treasury |
| Equity Risk Premium | 5.5% | Damodaran US ERP |
| Beta | Live | Yahoo Finance (5-yr monthly) |
| Effective Tax Rate | Live | 2-year avg from income statement |

NEE's low beta (~0.5–0.7) reflects its regulated utility earnings base, compressing the cost of equity relative to pure-play renewables.

### DCF & Terminal Value

Cash flows are discounted at WACC using a 5-year explicit forecast period. Terminal value uses the **Gordon Growth Model**:

```
TV = NOPAT₅ × (1 + g) / (WACC − g)
```

With a terminal growth rate of **3.0%**, anchored to long-run US GDP growth — consistent with a regulated utility whose revenue is tied to rate case outcomes and load growth rather than cyclical earnings.

### Revenue Growth Assumptions

| Year | Growth Rate | Rationale |
|---|---|---|
| 1–2 | 8% | Consistent with NEE's 2024–2027 capital deployment plan |
| 3–4 | 7% | Step-down as rate base additions normalize |
| 5 | 6% | Converging toward long-run growth |

EBIT margins are projected to expand modestly toward **34%**, within NEE's historical 30–36% range, reflecting operating leverage from its growing regulated asset base.

---

## Output

The script produces a three-sheet Excel workbook:

- **DCF Summary** — historical financials (3 years), 5-year projections, valuation bridge (EV → equity value → implied share price vs. current price), and full WACC build
- **Sensitivity Analysis** — implied share price grid across 5 WACC scenarios (±200 bps) and 7 terminal growth rates (1.0%–4.0%), with red-yellow-green conditional formatting
- **Assumptions** — all model parameters with methodology notes, suitable for documentation or review

---

## Key Observations

- **NOPAT consistently exceeds reported FCF** for NEE because growth capex (new wind, solar, and transmission assets) is economically productive rather than value-destructive — reported FCF is not a reliable valuation anchor for regulated utilities with active capital programs.
- **Terminal value dominates the DCF**, as it does for most stable, capital-intensive businesses. The sensitivity table makes this explicit: WACC and TGR assumptions matter far more than near-term revenue projections.
- **NEE's low beta compresses WACC** relative to its peers in the renewables universe, a direct consequence of its regulated Florida Power & Light earnings base, which provides stable, recession-resistant cash flows that reduce equity risk.
- **ITC/PTC tax credits** keep NEE's effective tax rate below the statutory rate, which flows through to a lower WACC (via higher after-tax NOPAT) and higher implied equity value — a structural advantage for regulated utilities with large renewables pipelines.

---

## Data Source

All financial and market data is sourced from **Yahoo Finance via the yfinance library**. Historical income statement and balance sheet data is pulled for the three most recent fiscal years. Select figures (shares outstanding, current price, beta) are sourced from the yfinance `info` object and reflect the most recent available values at runtime.

---

## Tech Stack

- Python 3
- [yfinance](https://github.com/ranaroussi/yfinance) — market & financial data
- [pandas](https://pandas.pydata.org/) — data handling
- [numpy](https://numpy.org/) — numerical calculations
- [openpyxl](https://openpyxl.readthedocs.io/) — Excel workbook generation

---

## How to Run

```bash
pip3 install -r ../requirements.txt
python3 nee_dcf.py
```

The script will print a console summary of key inputs and outputs (price, WACC, implied value, upside/downside) and save `NEE_DCF_Model.xlsx` to the working directory.
