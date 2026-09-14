"""Automated Test Suite for Automated Excel Data Analyst.

Tests:
1. Sales XLSX (Date, Region, Product, Revenue, Profit, Quantity)
2. Employee Performance CSV (Employee, Efficiency, Active Time protection)
3. Messy CSV (Deduplication, Currency symbols, Imputation)
4. Numeric-Heavy CSV (Continuous distributions, correlation)
5. Categorical-Heavy CSV (Segment aggregations)
6. Multi-Sheet XLSX (Smart sheet selection of 'Raw Data')
7. File Validator (Rejection of invalid types and 0-byte files)
8. FastAPI Endpoints (/health, /analyze, /download)
"""

import os
import sys
import tempfile
import unittest
from pathlib import Path
import pandas as pd
import openpyxl

# Add project root to sys.path
PROJECT_ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(PROJECT_ROOT))

from src.file_validator import validate_file
from src.sheet_selector import select_best_sheet
from src.pipeline import run_pipeline
from src.output_validator import validate_output
from src.smart_column_detector import detect_column_roles


class TestExcelAnalystPipeline(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.test_data_dir = PROJECT_ROOT / "tests" / "test_data"
        cls.temp_dir = tempfile.TemporaryDirectory()

    @classmethod
    def tearDownClass(cls):
        cls.temp_dir.cleanup()

    def test_01_sales_xlsx(self):
        """Test Case 1: Sales XLSX with Date, Region, Product, Revenue, Profit, Quantity."""
        input_path = str(self.test_data_dir / "sales.xlsx")
        output_path = os.path.join(self.temp_dir.name, "sales_report.xlsx")

        res = run_pipeline(input_path, output_path)

        self.assertEqual(res["status"], "success", f"Pipeline failed: {res.get('error')}")
        self.assertEqual(res["input_type"], ".xlsx")

        # Verify column roles: Date recognized
        roles = res["column_roles"]
        self.assertIn("Date", roles)
        self.assertEqual(roles["Date"]["role"], "datetime")

        # Verify Revenue detected as currency / measure
        self.assertIn(roles["Revenue"]["role"], ["currency", "measure", "numeric"])

        # Verify KPIs
        kpis = res["selected_kpis"]
        self.assertGreaterEqual(len(kpis), 2)
        self.assertEqual(kpis[0]["label"], "Total Records")

        # Verify Charts
        charts = res["selected_charts"]
        self.assertGreater(len(charts), 0)

        # Verify Output Workbook
        val = res["output_validation"]
        self.assertTrue(val["valid"])
        self.assertGreater(val["chart_count"], 0)
        self.assertIn("Dashboard", val["sheet_names"])
        self.assertIn("Cleaned Data", val["sheet_names"])
        self.assertIn("Cleaning Report", val["sheet_names"])
        self.assertIn("Original Data", val["sheet_names"])
        self.assertIn("Chart Data", val["sheet_names"])

    def test_02_employee_performance_csv(self):
        """Test Case 2: Employee CSV. Active Time MUST remain numeric, NOT datetime!"""
        input_path = str(self.test_data_dir / "employee_performance.csv")
        output_path = os.path.join(self.temp_dir.name, "employee_report.xlsx")

        res = run_pipeline(input_path, output_path)

        self.assertEqual(res["status"], "success", f"Pipeline failed: {res.get('error')}")
        self.assertEqual(res["input_type"], ".csv")
        self.assertEqual(res["selected_sheet"], "CSV Data")

        roles = res["column_roles"]
        # CRITICAL TEST: Active Time must NEVER be datetime!
        self.assertIn("Active Time", roles)
        self.assertNotEqual(
            roles["Active Time"]["role"], "datetime",
            "CRITICAL BUG: 'Active Time' numeric metric was misclassified as datetime!"
        )
        self.assertTrue(
            roles["Active Time"]["is_numeric"],
            "'Active Time' must be recognized as numeric."
        )

        # Efficiency percentage check
        self.assertIn(roles["Efficiency in Percentage"]["role"], ["percentage", "numeric"])

        # Preserved ID check
        self.assertEqual(roles["Employee"]["role"], "identifier")

        # Output validation
        val = res["output_validation"]
        self.assertTrue(val["valid"])
        self.assertGreater(val["chart_count"], 0)

    def test_03_messy_data_csv(self):
        """Test Case 3: Messy CSV with duplicates, missing values, ₹ currency, and percentages."""
        input_path = str(self.test_data_dir / "messy_data.csv")
        output_path = os.path.join(self.temp_dir.name, "messy_report.xlsx")

        res = run_pipeline(input_path, output_path)

        self.assertEqual(res["status"], "success", f"Pipeline failed: {res.get('error')}")

        clean_stats = res["cleaning_report"]
        # Verify duplicate row was removed
        self.assertGreaterEqual(clean_stats["duplicates_removed"], 1)

        # Verify missing values were imputed
        self.assertEqual(clean_stats["missing_values_after"], 0)

        # Verify currency conversion
        self.assertTrue(len(clean_stats["currency_conversions"]) > 0 or "Amount" in str(clean_stats))

        # Output validation
        val = res["output_validation"]
        self.assertTrue(val["valid"])

    def test_04_numeric_heavy_csv(self):
        """Test Case 4: Sensor numeric-heavy dataset."""
        input_path = str(self.test_data_dir / "numeric_heavy.csv")
        output_path = os.path.join(self.temp_dir.name, "numeric_report.xlsx")

        res = run_pipeline(input_path, output_path)
        self.assertEqual(res["status"], "success")

        # Check that EDA contains correlations and distributions
        eda = res["eda_summary"]
        self.assertIn("correlations", eda)
        self.assertIn("distributions", eda)

        val = res["output_validation"]
        self.assertTrue(val["valid"])

    def test_05_categorical_heavy_csv(self):
        """Test Case 5: Survey categorical-heavy dataset."""
        input_path = str(self.test_data_dir / "categorical_heavy.csv")
        output_path = os.path.join(self.temp_dir.name, "cat_report.xlsx")

        res = run_pipeline(input_path, output_path)
        self.assertEqual(res["status"], "success")

        val = res["output_validation"]
        self.assertTrue(val["valid"])
        self.assertGreater(val["chart_count"], 0)

    def test_06_multi_sheet_xlsx(self):
        """Test Case 6: Multi-sheet XLSX. System must intelligently pick 'Raw Data'."""
        input_path = str(self.test_data_dir / "multi_sheet.xlsx")
        output_path = os.path.join(self.temp_dir.name, "multi_sheet_report.xlsx")

        sheet_res = select_best_sheet(input_path)
        self.assertEqual(
            sheet_res["selected_sheet"], "Raw Data",
            f"Expected 'Raw Data' sheet to be selected, got '{sheet_res['selected_sheet']}'"
        )

        res = run_pipeline(input_path, output_path)
        self.assertEqual(res["status"], "success")
        self.assertEqual(res["selected_sheet"], "Raw Data")

        val = res["output_validation"]
        self.assertTrue(val["valid"])

    def test_07_file_validator_rejection(self):
        """Test Case 7: File validator rejects non-excel/csv and 0-byte files."""
        # 0-byte file
        zero_file = os.path.join(self.temp_dir.name, "empty.csv")
        with open(zero_file, "w") as f:
            pass
        val_zero = validate_file(zero_file)
        self.assertFalse(val_zero["valid"])
        self.assertIn("empty", val_zero["error"].lower())

        # Unsupported extension
        txt_file = os.path.join(self.temp_dir.name, "data.txt")
        with open(txt_file, "w") as f:
            f.write("hello,world\n1,2")
        val_txt = validate_file(txt_file)
        self.assertFalse(val_txt["valid"])
        self.assertIn("unsupported", val_txt["error"].lower())

    def test_08_fastapi_endpoints(self):
        """Test Case 8: Test FastAPI API routes using TestClient."""
        from fastapi.testclient import TestClient
        from api import app

        client = TestClient(app)

        # Test GET /
        resp_root = client.get("/", headers={"Accept": "application/json"})
        self.assertEqual(resp_root.status_code, 200)
        self.assertEqual(resp_root.json()["status"], "running")

        # Test GET /health
        resp_health = client.get("/health")
        self.assertEqual(resp_health.status_code, 200)
        self.assertEqual(resp_health.json()["status"], "healthy")

        # Test POST /analyze with sales.xlsx
        sales_path = self.test_data_dir / "sales.xlsx"
        with open(sales_path, "rb") as f:
            resp_analyze = client.post(
                "/analyze",
                files={"file": ("sales.xlsx", f, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")}
            )

        self.assertEqual(resp_analyze.status_code, 200, resp_analyze.text)
        data = resp_analyze.json()
        self.assertEqual(data["status"], "success")
        self.assertIn("download_url", data)

        # Test GET /download/{filename}
        download_url = data["download_url"]
        resp_dl = client.get(download_url)
        self.assertEqual(resp_dl.status_code, 200)
        self.assertGreater(len(resp_dl.content), 1000)


if __name__ == "__main__":
    unittest.main()
