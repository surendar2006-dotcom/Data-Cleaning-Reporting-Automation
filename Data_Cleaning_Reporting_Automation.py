"""
YUVA TASK: DATA CLEANING & REPORTING AUTOMATION
Run:
    pip install pandas matplotlib openpyxl
    python Data_Cleaning_Reporting_Automation.py

This single-file project creates a sample dataset, cleans it, and generates
CSV, Excel, chart, and text-summary outputs in the output_reports folder.
"""
from pathlib import Path
import pandas as pd
import matplotlib.pyplot as plt

OUTPUT_DIR = Path(__file__).resolve().parent / "output_reports"
OUTPUT_DIR.mkdir(exist_ok=True)

records = [
    ["ORD001", "2026-10-01", " Alice ", "Chennai", "Laptop", 1, 55000, "Completed"],
    ["ORD002", "02/10/2026", "bob", " chennai ", "Mouse", 2, 500, "completed"],
    ["ORD003", "2026-10-03", "CHARLIE", "Coimbatore", "Keyboard", 1, 1200, "Pending"],
    ["ORD004", "not available", "Diana", "Madurai", "Monitor", 1, 9000, "Shipped"],
    ["ORD005", "2026-10-05", "Alice", "Chennai", "Laptop", 1, 55000, "Completed"],
    ["ORD005", "2026-10-05", "Alice", "Chennai", "Laptop", 1, 55000, "Completed"],
    ["ORD006", "2026-10-06", None, "Trichy", "Headphones", 2, 1500, "Pending"],
    ["ORD007", "2026-10-07", "Ethan", None, "Mouse", 3, 500, "completed"],
    ["ORD008", "2026-10-08", "Fiona", "Salem", "Keyboard", None, 1200, "Shipped"],
    ["ORD009", "2026-10-09", "George", "Madurai", "Monitor", 1, None, "Pending"],
    ["ORD010", "2026-10-10", "Hana", "Coimbatore", "Laptop", 1, 55000, "COMPLETED"],
    ["ORD011", "2026-10-11", "Ivan", "Chennai", "Mouse", 2, 500, "Cancelled"],
    ["ORD012", "2026-10-12", "Julia", "Trichy", "Headphones", 1, 1500, " shipped "],
    ["ORD013", "2026-10-13", "Kevin", "Salem", "Monitor", 1, 9000, "Pending"],
    ["ORD014", "2026-10-14", "Liam", "Madurai", "Keyboard", 2, 1200, "Completed"],
]
columns = ["Order_ID", "Order_Date", "Customer", "City", "Product",
           "Quantity", "Unit_Price", "Status"]

def main():
    df = pd.DataFrame(records, columns=columns)
    rows_before = len(df)
    df.to_csv(OUTPUT_DIR / "sales_raw.csv", index=False)

    # Remove duplicate rows and standardize text.
    df = df.drop_duplicates().copy()
    for col in ["Order_ID", "Customer", "City", "Product", "Status"]:
        df[col] = df[col].astype("string").str.strip()
    for col in ["Customer", "City", "Product", "Status"]:
        df[col] = df[col].str.title()

    # Convert dates/numbers; invalid values become missing.
    df["Order_Date"] = pd.to_datetime(df["Order_Date"], errors="coerce", dayfirst=True)
    df["Quantity"] = pd.to_numeric(df["Quantity"], errors="coerce")
    df["Unit_Price"] = pd.to_numeric(df["Unit_Price"], errors="coerce")

    # Fill missing values using documented rules.
    for col in ["Customer", "City", "Status"]:
        df[col] = df[col].fillna("Unknown")
    df["Quantity"] = df["Quantity"].fillna(df["Quantity"].median())
    df["Unit_Price"] = df["Unit_Price"].fillna(df["Unit_Price"].median())
    df["Order_Date"] = df["Order_Date"].fillna(pd.Timestamp("2026-10-01"))

    df["Quantity"] = df["Quantity"].round().astype(int)
    df["Revenue"] = (df["Quantity"] * df["Unit_Price"]).round(2)
    df["Order_Date"] = df["Order_Date"].dt.strftime("%Y-%m-%d")
    df.to_csv(OUTPUT_DIR / "sales_cleaned.csv", index=False)

    status_summary = df.groupby("Status").agg(
        Orders=("Order_ID", "count"), Revenue=("Revenue", "sum")
    ).reset_index()
    product_summary = df.groupby("Product")["Revenue"].sum().sort_values(ascending=False)
    total_revenue = df["Revenue"].sum()
    avg_order = total_revenue / len(df) if len(df) else 0

    with pd.ExcelWriter(OUTPUT_DIR / "automated_sales_report.xlsx",
                        engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="Cleaned Data", index=False)
        status_summary.to_excel(writer, sheet_name="Status Summary", index=False)
        product_summary.rename("Revenue").to_frame().to_excel(
            writer, sheet_name="Product Revenue"
        )

    plt.figure(figsize=(8, 5))
    product_summary.sort_values().plot(kind="barh")
    plt.title("Revenue by Product")
    plt.xlabel("Revenue")
    plt.ylabel("Product")
    plt.tight_layout()
    plt.savefig(OUTPUT_DIR / "revenue_by_product.png", dpi=160)
    plt.close()

    summary = f"""DATA CLEANING & REPORTING AUTOMATION

Input rows: {rows_before}
Cleaned rows: {len(df)}
Duplicate rows removed: {rows_before - len(df)}
Total revenue: Rs. {total_revenue:,.2f}
Average order value: Rs. {avg_order:,.2f}

Cleaning performed:
- Removed duplicate records.
- Trimmed extra spaces and standardized text capitalization.
- Converted date and numeric columns to valid formats.
- Replaced missing text values with Unknown.
- Replaced missing quantity and unit price with their median values.
- Replaced invalid/missing dates with 2026-10-01 for this sample.
- Calculated Revenue = Quantity x Unit_Price.

Generated files:
sales_raw.csv, sales_cleaned.csv, automated_sales_report.xlsx,
revenue_by_product.png, summary_report.txt
"""
    (OUTPUT_DIR / "summary_report.txt").write_text(summary, encoding="utf-8")
    print("SUCCESS: Cleaning and reporting completed.")
    print(f"Open output folder: {OUTPUT_DIR}")
    print(summary)

if __name__ == "__main__":
    main()
