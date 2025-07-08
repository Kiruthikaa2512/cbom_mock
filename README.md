# Supplier Cost & Performance Report Generator

This Python-based reporting tool analyzes supplier-level spend, quantity, and lead time from a CBOM-style dataset. It automatically generates a formatted Excel report along with visual charts to support vendor review discussions—perfect for sourcing, CapEx, or GSM teams.

---

##  What It Does

- Cleans raw supplier data
- Calculates `Total Cost = Unit Cost × Quantity`
- Summarizes key metrics per supplier:
  - Total Quantity
  - Total Spend
  - Average Lead Time
- Creates a multi-tab Excel report
- Saves visualizations (pie, bar, and line charts)

---

##  Why It’s Useful

This tool helps teams:

| Use Case | Benefit |
|----------|---------|
| Track vendor spend | See who contributes most to CapEx costs |
| Compare supplier performance | Spot delivery lead time risks or inefficiencies |
| Automate reporting | Reduce manual analysis and speed up reviews |
| Enable better decisions | Share data in a clean, readable format |

---

## 🧾 Sample Output

**Excel File:** `Supplier_Report.xlsx`
- Tab 1: `Summary` → Spend, Quantity & Lead Time per supplier  
- Tab 2: `Cleaned_Data` → Processed records ready for audit  
- Top 3 spenders are conditionally highlighted

**Charts:**
- `pie_chart.png`: Share of spend per supplier
- `bar_chart.png`: Total cost by supplier
- `line_chart.png`: Cost vs. lead time trend

---

##  How to Use

1. Place your data CSV in the project folder (e.g., `cbom_sample.csv`)
2. Install dependencies:
   ```bash
   pip install pandas matplotlib seaborn xlsxwriter