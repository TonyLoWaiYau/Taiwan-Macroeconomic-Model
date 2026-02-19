# Taiwan Macroeconomic Model (BVAR) — Excel-Based Forecasting & Scenario Analysis

A Windows desktop forecasting tool built on a **Bayesian VAR (BVAR)** model for Taiwan’s key macroeconomic indicators.  
It is designed for **non-technical researchers** to conduct **forecasting and scenario/conditional analysis** via **Excel-based conditioning**—**no coding** and **no R installation** required.

---

## Key Features

- **BVAR model** for Taiwan macro forecasting
- **Unconditional (baseline) forecast** generation
- **Conditional/scenario forecasts** by editing an Excel sheet (`forecast_condition`)
- Exports results to:
  - `result_raw.xlsx` (baseline + conditioning + conditional forecast)
  - `result_summary.xlsx` (level, growth, annual summaries)
- Users can **re-run Part 2 repeatedly** after adjusting scenarios in Excel

---

## Installation (Windows)

1. Go to the **Releases** page and download the latest installer:
   - https://github.com/TonyLoWaiYau/Taiwan-Macroeconomic-Model/releases/latest
2. Run the installer and launch the application from the Start Menu (or desktop shortcut if created).

> Notes:
> - Microsoft Excel (or compatible `.xlsx` editor) is required for scenario editing.
> - The installer bundles everything needed to run (users do **not** need R or RStudio).

---

## How to Use (Workflow)

### Part 1 — Baseline forecast + create conditioning template
1. Open the app.
2. Upload your `data.xlsx` (see required format below).
3. Click **Run Part 1 (Create result_raw.xlsx)**.
4. Click **Open result_raw.xlsx in Excel**.

This creates `result_raw.xlsx` containing:
- `result_raw` (history + baseline forecast)
- `forecast_condition` (editable sheet for conditional/scenario paths)

### Part 2 — Conditional/scenario forecast (repeatable)
1. In Excel, edit **ONLY** the `forecast_condition` sheet (future rows).
2. Save and **close Excel** (the file must not be locked).
3. In the app, click **Run Part 2 (Conditional Forecast + result_summary.xlsx)**.

If the scenario is problematic (e.g., non-numeric values), the app will show an error message.  
You can then **edit `forecast_condition` again** and re-run Part 2 until satisfied.

---

## Input Data Format: `data.xlsx`

- The workbook must contain a sheet named **`clean`** (default; configurable in the app).
- Data must be **quarterly** and include the following variables.

### Variables

| Variable | Description |
|---|---|
| `t` | Time (quarterly) |
| `gdp` | Real GDP in million TWD (2021 prices) |
| `pce` | Real Private Consumption Expenditure in million TWD (2021 prices) |
| `gce` | Real Government Consumption Expenditure in million TWD (2021 prices) |
| `gcf` | Real Gross Capital Formation in million TWD (2021 prices) |
| `exports` | Real Exports in million TWD (2021 prices) |
| `imports` | Real Imports in million TWD (2021 prices) |
| `g_gdp` | Advance estimate of real GDP growth (% YoY) |
| `g_pce` | Advance estimate of real PCE growth (% YoY) |
| `g_gce` | Advance estimate of real GCE growth (% YoY) |
| `g_gcf` | Advance estimate of real GCF growth (% YoY) |
| `g_exports` | Advance estimate of real Exports growth (% YoY) |
| `g_imports` | Advance estimate of real Imports growth (% YoY) |
| `ip` | Industrial production index (2021 = 100, quarterly average) |
| `cpi` | CPI (2021 = 100, quarterly average) |
| `unemploy` | Unemployment rate (quarterly average) |
| `fx` | USD/NTD (quarterly average) |

### Notes on “advance estimate” variables (`g_*`)
- The model computes YoY growth internally from level variables.
- The `g_*` columns are used to optionally provide **advance estimates** for the latest quarter (if the last row of the input data has missing values).

---

## Output Files

### `result_raw.xlsx`
Contains:
- `result_raw`: historical data + baseline forecast
- `forecast_condition`: scenario input sheet (editable)
- `conditional forecast`: written/updated after running Part 2

### `result_summary.xlsx`
Contains:
- `result_level`: quarterly levels (including forecast horizon)
- `result_growth`: quarterly YoY growth rates (levels converted to growth)
- `result_annual_level`: annual aggregates (sum or average depending on variable)
- `result_annual_growth`: annual YoY growth rates

---

## Troubleshooting

- **Part 2 fails / shows an error**:  
  Usually caused by non-numeric cells in the future rows of `forecast_condition`, or by unrealistic conditioning paths that break the conditional forecast solver.  
  Fix the sheet, **save and close Excel**, then try Part 2 again.

- **Cannot overwrite `result_raw.xlsx`**:  
  Ensure Excel is closed (file not locked).


---

## Author

Tony Lo (Feb 2026)
