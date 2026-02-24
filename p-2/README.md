# Campus Resource Utilization Analysis

> **Consulting-grade analytics** of campus facility utilization at an Indian engineering college.

---

## 📋 Project Objective

Analyze the utilization efficiency of 7 campus facilities (Gymnasium, Library, Computer Labs, Study Rooms, Auditorium, Sports Complex) and identify optimization opportunities through data-driven scenario modeling.

## 🔬 Methodology

| Phase | Description |
|-------|-------------|
| **Data Collection** | Built hourly usage dataset for 7 facilities across 18 time slots (6 AM – 11 PM) |
| **Utilization Analysis** | Computed peak, average, idle, and overcrowding metrics per resource |
| **Scenario Simulation** | Modeled 7 optimization scenarios across redistribution, scheduling, and right-sizing levers |
| **Insight Generation** | Identified root causes of inefficiency and quantified improvement potential |

## 🛠 Tools Used

- **Python** — Data modeling, scenario simulation, automation
- **openpyxl** — Excel workbook generation with formatted sheets and charts
- **Chart.js** — Interactive browser-based data visualizations
- **Operations Analytics** — Utilization analysis, capacity planning, scenario modeling

## 📊 Key Findings

| Metric | Current | Optimized | Improvement |
|--------|---------|-----------|-------------|
| Avg Utilization | 33.7% | 39.1% | +5.4 pp |
| Cost / Person-Hour | ₹11.22 | ₹9.79 | −₹1.43 |
| Idle Hours / Day | 42 | 34 | −8 hours |
| Most Under-utilised | Auditorium (17.3%) | — | Massive wastage |

## 📁 Deliverables

| File | Description |
|------|-------------|
| `resource_utilization_model.py` | Python model engine — generates all outputs |
| `Campus_Resource_Utilization_Analysis.xlsx` | 6-sheet Excel workbook (Dataset, Metrics, Scenarios, Dashboard, Insights, Summary) |
| `resource_dashboard.html` | Interactive browser dashboard with charts and insights |

## ▶️ How to Run

```bash
# Ensure Python 3.8+ is installed
python3 resource_utilization_model.py
# → Generates Campus_Resource_Utilization_Analysis.xlsx

# Open the dashboard
open resource_dashboard.html
```

## 🎯 Resume Summary

> **Campus Resource Utilization Analysis** — Built a consulting-grade analytics model analyzing hourly utilization patterns across 7 campus facilities at an Indian engineering college. Identified 42 daily idle hours and 17.3% Auditorium utilization as key inefficiencies. Modeled 7 optimization scenarios demonstrating a path to improve utilization from 33.7% to 39.1% and reduce cost per person-hour by ₹1.43. Delivered an executive dashboard with actionable recommendations.
