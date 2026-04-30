# NextEra Energy (NEE) — DCF Model

I created a Python-based DCF model for NextEra Energy (NEE), which is currently the largest regulated renewable energy corporation in North America. This again goes along with the common theme in the financials of renewable energy and why corporations and governments are so hesistant to move towards it, as upfront costs are significant but long term investments are typically successful. Script pulls data with yfinance, builds a 5-year projection, calculates WACC from scratch, and exports a neat Excel workbook with separate sheets for sensitivity and full assumptions.

---

## Methodology

### Free Cash Flow Note

NOTE: Used Standard yfinance for free cash flow (Operating CF − Capex), however it understates NEE's economic earnings because the company spends $9B+ per year in capital expidentures to build regulated return assets.

Instead, model uses **NOPAT as the FCF proxy**:

```
NOPAT = EBIT × (1 − Effective Tax Rate)
```

It normalizes for capital expedentures annd is much more consistent with the revenue growth projections. the new capital generates new revenue and the EBIT margin assumptions account for those assets instead.

### WACC

WACC is built from market data using (Weight of Equity × Cost of Equity) + (Weight of Debt × Cost of Debt × (1 - Tax Rate))

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

NEE's low beta (~0.5–0.7) reflects its regulated utility earnings base; compresseswha cost of equity relative to pure-play renewables.

### DCF & Terminal Value

Cash flows are discounted with WACC formula using a 5-year forecast period, using **Gordon Growth Model**:

```
TV = NOPAT₅ × (1 + g) / (WACC − g)
```

Uses terminal growth rate of **3.0%**, (anchored to US long term GDP growth)

### Revenue Growth Assumptions

| Year | Growth Rate | Rationale |
|---|---|---|
| 1–2 | 8% | Consistent with NEE's 2024–2027 capital deployment plan |
| 3–4 | 7% | Step-down as rate base additions normalize |
| 5 | 6% | Converging toward long-run growth |

EBIT margins projected to expand toward **34%**, within NEE's historical 30–36% range, showing operating leverage from its growing regulated asset base.s

---

## Output

The script produces a three-sheet Excel workbook:

- **DCF Summary** — financial histories for 3 years,, 5-year projections (valuation from enterprise value-subtract debt-add back cash-get equity value-divide by shares outstanding-compare to current market price
- **Sensitivity Analysis** — implied share price grid across 5 WACC scenarios (±200 basis points) and 7 terminal growth rates (1.0%–4.0%), with red-yellow-green indicators
- **Assumptions** — all model parameters with methodology notes, suitable for documentation or review

---

## Key Observations

- **NOPAT consistently exceeds reported FCF** for NEE because growth capital expedentures (new wind, solar, etc) is productive rather than destructive to company value — reported FCF not a reliable anchor for regulated utilities.
- **Terminal value dominates the DCF**, as it does for most stable/capital-heavy businesses. The sensitivity table makes this explicit: WACC and TGR assumptions matter far more than near-term projections
- **NEE's low beta compresses WACC** relative to peers in the renewables environment, a direct consequence of its regulated Florida Power & Light earnings base, which provides stable, recession-resistant cash flows that reduce equity risk.
- **ITC/PTC tax credits** keep NEE's effective tax rate below the statutory rate, which flows through to a lower WACC (via higher after-tax NOPAT) and higher implied equity value — a structural advantage for regulated utilities with large renewables pipelines.

---

## Data Source

All data (financial/market) is from **Yahoo Finance via the yfinance library**. Historical statement and balance sheet data is pulled for the past 3 fiscal years. Select figures (shares outstanding, current price, beta) are sourced from yfinance `info` object and reflect most recent values.

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

The script will print a console summary of inputs and outputs (price, WACC, implied value, upside/downside) and save `NEE_DCF_Model.xlsx` to the working directory.
