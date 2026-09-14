"""Script to generate test datasets for Automated Excel Data Analyst.

Generates:
1. sales.xlsx (Date, Region, Product, Revenue, Profit, Quantity)
2. employee_performance.csv (Employee, Efficiency in Percentage, Attended Calls, Active With Cust, Active Time)
3. messy_data.csv (Duplicates, nulls, '25%', '₹500', whitespace)
4. numeric_heavy.csv (Continuous sensor/financial telemetry)
5. categorical_heavy.csv (Customer survey and feedback)
6. multi_sheet.xlsx (Dashboard, Summary, Raw Data, Report sheets)
"""

import os
from pathlib import Path
import pandas as pd
import numpy as np
import openpyxl


TEST_DATA_DIR = Path(__file__).resolve().parent / "test_data"
os.makedirs(TEST_DATA_DIR, exist_ok=True)


def generate_datasets():
    # 1. Sales XLSX
    dates = pd.date_range(start="2024-01-01", periods=30, freq="D")
    regions = ["North", "South", "East", "West"]
    products = ["Laptop Pro", "Wireless Mouse", "4K Monitor", "USB-C Hub", "Mechanical Keyboard"]

    sales_rows = []
    np.random.seed(42)
    for d in dates:
        for _ in range(3):
            reg = np.random.choice(regions)
            prod = np.random.choice(products)
            qty = int(np.random.randint(1, 20))
            price = {"Laptop Pro": 1200, "Wireless Mouse": 25, "4K Monitor": 400, "USB-C Hub": 45, "Mechanical Keyboard": 90}[prod]
            rev = float(qty * price)
            profit = round(rev * float(np.random.uniform(0.15, 0.35)), 2)
            sales_rows.append({
                "Date": d.strftime("%Y-%m-%d"),
                "Region": reg,
                "Product": prod,
                "Revenue": rev,
                "Profit": profit,
                "Quantity": qty
            })
    df_sales = pd.DataFrame(sales_rows)
    sales_path = TEST_DATA_DIR / "sales.xlsx"
    df_sales.to_excel(sales_path, index=False)
    print(f"Generated {sales_path} ({len(df_sales)} rows)")

    # 2. Employee Performance CSV
    employees = [f"EMP_{i:03d}" for i in range(1, 26)]
    emp_rows = []
    for emp in employees:
        efficiency = round(float(np.random.uniform(65.0, 98.0)), 1)
        attended_calls = int(np.random.randint(40, 150))
        active_with_cust = round(float(np.random.uniform(180.0, 420.0)), 1)
        active_time = round(float(np.random.uniform(300.0, 480.0)), 1)  # Duration in minutes
        emp_rows.append({
            "Employee": emp,
            "Efficiency in Percentage": f"{efficiency}%",
            "Attended Calls": attended_calls,
            "Active With Cust": active_with_cust,
            "Active Time": active_time
        })
    df_emp = pd.DataFrame(emp_rows)
    emp_path = TEST_DATA_DIR / "employee_performance.csv"
    df_emp.to_csv(emp_path, index=False)
    print(f"Generated {emp_path} ({len(df_emp)} rows)")

    # 3. Messy CSV
    messy_rows = [
        {"Transaction_ID": "TX-101", " Customer Name ": " John Doe ", "Amount": " ₹1,500.00 ", "Discount": "10%", "Status": "Completed"},
        {"Transaction_ID": "TX-102", " Customer Name ": "Jane Smith", "Amount": "₹2,450.50", "Discount": "15%", "Status": "Completed"},
        {"Transaction_ID": "TX-103", " Customer Name ": " Bob Wilson ", "Amount": "null", "Discount": "-", "Status": "Pending"},
        {"Transaction_ID": "TX-104", " Customer Name ": "Alice Brown", "Amount": "₹890.00", "Discount": "5%", "Status": "Failed"},
        {"Transaction_ID": "TX-105", " Customer Name ": "Charlie Davis", "Amount": "₹3,200.00", "Discount": "25%", "Status": "Completed"},
        # Duplicate row
        {"Transaction_ID": "TX-101", " Customer Name ": " John Doe ", "Amount": " ₹1,500.00 ", "Discount": "10%", "Status": "Completed"},
        {"Transaction_ID": "TX-106", " Customer Name ": "NA", "Amount": "₹750.00", "Discount": "N/A", "Status": "Completed"},
        {"Transaction_ID": "TX-107", " Customer Name ": "Eve White", "Amount": "--", "Discount": "0%", "Status": "Unknown"}
    ]
    df_messy = pd.DataFrame(messy_rows)
    messy_path = TEST_DATA_DIR / "messy_data.csv"
    df_messy.to_csv(messy_path, index=False)
    print(f"Generated {messy_path} ({len(df_messy)} rows)")

    # 4. Numeric-Heavy CSV
    num_rows = []
    for i in range(1, 101):
        num_rows.append({
            "Sensor_ID": f"SENS_{i:04d}",
            "Temperature_C": round(float(np.random.normal(72, 5)), 2),
            "Pressure_PSI": round(float(np.random.normal(1013, 20)), 1),
            "Vibration_Hz": round(float(np.random.exponential(15)), 2),
            "Voltage_V": round(float(np.random.normal(230, 4)), 1),
            "Current_A": round(float(np.random.uniform(5, 25)), 2)
        })
    df_num = pd.DataFrame(num_rows)
    num_path = TEST_DATA_DIR / "numeric_heavy.csv"
    df_num.to_csv(num_path, index=False)
    print(f"Generated {num_path} ({len(df_num)} rows)")

    # 5. Categorical-Heavy CSV
    cat_rows = []
    departments = ["Engineering", "Sales", "Support", "Marketing", "Product"]
    satisfaction = ["Very Satisfied", "Satisfied", "Neutral", "Dissatisfied"]
    feedback_channel = ["Email", "Slack", "Portal", "Phone"]
    priority = ["High", "Medium", "Low"]

    for i in range(1, 61):
        cat_rows.append({
            "Ticket_ID": f"TKT-{i:04d}",
            "Department": np.random.choice(departments),
            "Satisfaction_Level": np.random.choice(satisfaction),
            "Feedback_Channel": np.random.choice(feedback_channel),
            "Priority": np.random.choice(priority),
            "Rating_Score": int(np.random.randint(1, 6))
        })
    df_cat = pd.DataFrame(cat_rows)
    cat_path = TEST_DATA_DIR / "categorical_heavy.csv"
    df_cat.to_csv(cat_path, index=False)
    print(f"Generated {cat_path} ({len(df_cat)} rows)")

    # 6. Multi-Sheet XLSX
    multi_path = TEST_DATA_DIR / "multi_sheet.xlsx"
    wb = openpyxl.Workbook()

    # Sheet 1: Dashboard (poor tabular, single notes/KPIs)
    ws_dash = wb.active
    ws_dash.title = "Dashboard"
    ws_dash["A1"] = "Executive Summary Dashboard"
    ws_dash["A2"] = "Target Achieved: 94%"
    ws_dash["A3"] = "Notes: Review quarterly numbers with CFO"

    # Sheet 2: Summary (Aggregated totals)
    ws_summ = wb.create_sheet(title="Summary")
    ws_summ.append(["Department", "Headcount"])
    ws_summ.append(["Engineering", 45])
    ws_summ.append(["Sales", 30])

    # Sheet 3: Raw Data (Rich tabular data - should be chosen!)
    ws_raw = wb.create_sheet(title="Raw Data")
    ws_raw.append(["Transaction_ID", "Date", "Region", "Revenue", "Units_Sold", "Customer_Rating"])
    for i in range(1, 51):
        ws_raw.append([
            f"TXN-{i:04d}",
            f"2024-03-{(i%28)+1:02d}",
            np.random.choice(regions),
            float(np.random.randint(100, 2000)),
            int(np.random.randint(1, 15)),
            round(float(np.random.uniform(3.0, 5.0)), 1)
        ])

    # Sheet 4: Report
    ws_rep = wb.create_sheet(title="Report")
    ws_rep.append(["Config", "Value"])
    ws_rep.append(["Version", "1.2"])

    wb.save(multi_path)
    wb.close()
    print(f"Generated {multi_path} with 4 worksheets.")


if __name__ == "__main__":
    generate_datasets()
