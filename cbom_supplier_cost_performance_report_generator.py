import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
import os
from datetime import datetime

# === Step 1: Set up the output folder ===
timestamp = datetime.now().strftime('%Y-%m-%d_%H-%M-%S')
output_dir = f"outputs/report_{timestamp}"
os.makedirs(output_dir, exist_ok=True)

# === Step 2: Load the supplier data ===
df = pd.read_csv("cbom_sample.csv")

# === Step 3: Clean column names ===
df.columns = df.columns.str.strip().str.replace(" ", "_")

# === Step 4: Handle missing values ===
df_clean = df.dropna(subset=["Unit_Cost", "Quantity"]).copy()

# === Step 5: Add Total Cost column ===
df_clean["Total_Cost"] = df_clean["Unit_Cost"] * df_clean["Quantity"]

# === Step 6: Create supplier summary ===
supplier_summary = df_clean.groupby("Supplier").agg({
    "Quantity": "sum",
    "Total_Cost": "sum",
    "Lead_Time": "mean"
}).sort_values(by="Total_Cost", ascending=False).reset_index()

# === Step 7: Save to Excel with multiple tabs ===
excel_path = os.path.join(output_dir, "Supplier_Report.xlsx")
with pd.ExcelWriter(excel_path, engine="xlsxwriter") as writer:
    supplier_summary.to_excel(writer, sheet_name="Summary", index=False)
    df_clean.to_excel(writer, sheet_name="Cleaned_Data", index=False)

    # Optional: Add formatting (e.g., highlight top 3)
    workbook = writer.book
    summary_sheet = writer.sheets["Summary"]
    format_top = workbook.add_format({'bg_color': '#C6EFCE', 'font_color': '#006100'})
    summary_sheet.conditional_format('C2:C100', {
        'type': '3_top',
        'format': format_top,
        'criteria': '>=',
        'value': 0
    })

# === Step 8: Make charts ===
# Pie Chart
plt.figure(figsize=(6, 6))
plt.pie(supplier_summary["Total_Cost"], labels=supplier_summary["Supplier"], autopct="%1.1f%%", startangle=90)
plt.title("Supplier Spend Share")
plt.axis("equal")
plt.savefig(os.path.join(output_dir, "pie_chart.png"))
plt.close()

# Bar Chart
plt.figure(figsize=(8, 5))
sns.barplot(data=supplier_summary, x="Supplier", y="Total_Cost", palette="viridis")
plt.title("Total Cost by Supplier")
plt.ylabel("Total Cost ($)")
plt.xticks(rotation=45)
plt.tight_layout()
plt.savefig(os.path.join(output_dir, "bar_chart.png"))
plt.close()

# Line Chart
plt.figure(figsize=(10, 6))
sns.lineplot(data=df_clean, x="Lead_Time", y="Total_Cost", hue="Supplier", marker="o")
plt.title("Cost vs Lead Time by Supplier")
plt.xlabel("Lead Time (days)")
plt.ylabel("Total Cost ($)")
plt.tight_layout()
plt.savefig(os.path.join(output_dir, "line_chart.png"))
plt.close()

print(f"\n Report created at: {output_dir}")
print("- Supplier_Report.xlsx")
print("- pie_chart.png")
print("- bar_chart.png")
print("- line_chart.png")