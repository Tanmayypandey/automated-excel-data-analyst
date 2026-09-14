"""Output Validator Module.

Validates the generated Excel workbook:
- Confirms file existence and non-zero size
- Verifies workbook loads cleanly in OpenPyXL
- Validates presence of all 5 required worksheets
- Validates row counts in data sheets
- Validates presence of embedded OpenPyXL charts on Dashboard
"""

import os
from typing import Dict, Any, List
import openpyxl


REQUIRED_SHEETS = [
    "Dashboard",
    "Cleaned Data",
    "Cleaning Report",
    "Original Data",
    "Chart Data"
]


def validate_output(output_path: str) -> Dict[str, Any]:
    """Validate that the generated .xlsx workbook meets all structural and visual criteria.

    Args:
        output_path: Path to the generated .xlsx workbook.

    Returns:
        Structured validation report.
    """
    if not os.path.exists(output_path):
        return {
            "valid": False,
            "error": f"Output file does not exist: {output_path}",
            "file_size": 0,
            "sheet_names": [],
            "chart_count": 0,
            "original_rows": 0,
            "cleaned_rows": 0
        }

    size = os.path.getsize(output_path)
    if size == 0:
        return {
            "valid": False,
            "error": "Generated output file is empty (0 bytes).",
            "file_size": 0,
            "sheet_names": [],
            "chart_count": 0,
            "original_rows": 0,
            "cleaned_rows": 0
        }

    try:
        wb = openpyxl.load_workbook(output_path, data_only=False)
        sheet_names = wb.sheetnames

        # Check required sheets
        missing_sheets = [s for s in REQUIRED_SHEETS if s not in sheet_names]
        if missing_sheets:
            wb.close()
            return {
                "valid": False,
                "error": f"Missing required sheets: {', '.join(missing_sheets)}",
                "file_size": size,
                "sheet_names": sheet_names,
                "chart_count": 0,
                "original_rows": 0,
                "cleaned_rows": 0
            }

        # Check data sheet rows
        ws_orig = wb["Original Data"]
        ws_clean = wb["Cleaned Data"]
        orig_rows = ws_orig.max_row or 0
        clean_rows = ws_clean.max_row or 0

        if orig_rows <= 1:
            wb.close()
            return {
                "valid": False,
                "error": "Original Data worksheet has no data rows.",
                "file_size": size,
                "sheet_names": sheet_names,
                "chart_count": 0,
                "original_rows": orig_rows,
                "cleaned_rows": clean_rows
            }

        if clean_rows <= 1:
            wb.close()
            return {
                "valid": False,
                "error": "Cleaned Data worksheet has no data rows.",
                "file_size": size,
                "sheet_names": sheet_names,
                "chart_count": 0,
                "original_rows": orig_rows,
                "cleaned_rows": clean_rows
            }

        # Check dashboard charts
        ws_dash = wb["Dashboard"]
        chart_count = len(getattr(ws_dash, "_charts", []))

        wb.close()

        return {
            "valid": True,
            "error": None,
            "file_size": size,
            "sheet_names": sheet_names,
            "chart_count": chart_count,
            "original_rows": orig_rows - 1,  # exclude header
            "cleaned_rows": clean_rows - 1   # exclude header
        }

    except Exception as e:
        return {
            "valid": False,
            "error": f"Failed to load or validate workbook: {str(e)}",
            "file_size": size,
            "sheet_names": [],
            "chart_count": 0,
            "original_rows": 0,
            "cleaned_rows": 0
        }
