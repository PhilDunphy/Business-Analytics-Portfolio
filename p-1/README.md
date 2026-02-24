# College Mess Profitability Optimization Model

> **Consulting-grade financial analysis** of an Indian college hostel mess operation.

---

## 📋 Project Objective

Analyze the profitability, cost structure, and optimization levers of a college mess serving **650 hostel students**. Build a bottom-up financial model, run scenario simulations, and deliver actionable recommendations with quantified financial impact.

## 🔬 Methodology

| Phase | Description |
|-------|-------------|
| **Data Collection** | Built a realistic dataset of 13 cost/revenue line items based on Indian college mess benchmarks |
| **Financial Modeling** | Constructed a monthly P&L with variable costs, fixed costs, and margin analysis |
| **Scenario Analysis** | Simulated 11 optimization scenarios across 3 levers — wastage reduction, operational efficiency, pricing |
| **Insight Generation** | Identified root causes of inefficiency and quantified improvement potential |

## 🛠 Tools Used

- **Python** — Data modeling, scenario simulation, automation
- **openpyxl** — Excel workbook generation with formatted sheets and charts
- **Chart.js** — Interactive browser-based data visualizations
- **Financial Modeling** — Bottom-up P&L, break-even analysis, sensitivity testing

## 📊 Key Findings

| Metric | Current | Optimized | Improvement |
|--------|---------|-----------|-------------|
| Monthly Revenue | ₹29.25L | ₹30.55L | +₹1.30L |
| Monthly Profit | ₹5.14L | ₹8.28L | +₹3.14L |
| Profit Margin | 17.59% | 27.09% | +9.50pp |
| Food Wastage Cost | ₹2.11L/mo | ₹0.35L/mo | −₹1.76L |
| Annualized Improvement | — | — | ₹37.68L/yr |

## 📁 Deliverables

| File | Description |
|------|-------------|
| `financial_model.py` | Python model engine — generates all outputs |
| `College_Mess_Profitability_Model.xlsx` | 6-sheet Excel workbook (Dataset, P&L, Scenarios, Dashboard, Insights, Summary) |
| `dashboard.html` | Interactive browser dashboard with charts and insights |

## ▶️ How to Run

```bash
# Ensure Python 3.8+ is installed
cd /path/to/Bain
python3 financial_model.py
# → Generates College_Mess_Profitability_Model.xlsx

# Open the dashboard
open dashboard.html
```

## 🎯 Resume Summary

> **College Mess Profitability Optimization Model** — Built a consulting-grade financial model analyzing the cost structure and profitability of a college hostel mess serving 650+ students. Ran 11 optimization scenarios across wastage reduction, efficiency, and pricing levers. Identified ₹2.11L/month in avoidable food wastage costs and demonstrated a path to improve profit margin from 17.6% to 27.1% (₹37.7L annualized improvement).
