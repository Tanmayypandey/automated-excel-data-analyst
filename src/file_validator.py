"""File Validator Module.

Validates input files (.xlsx and .csv) against size limits, extension rules,
row/column bounds, and structural readability.
"""

import os
from typing import Dict, Any
import openpyxl


MAX_FILE_SIZE_BYTES = 100 * 1024 * 1024  # 100 MB
MAX_CSV_ROWS = 500000
MAX_CSV_COLUMNS = 200
SUPPORTED_EXTENSIONS = {".xlsx", ".csv"}


def validate_file(file_path: str) -> Dict[str, Any]:
    """Validate uploaded file for type, size, readability, and dimensions.

    Args:
        file_path: Absolute or relative path to the file.

    Returns:
        Structured dictionary reporting validation outcome, file type, rows,
        columns, and any error message encountered.
    """
    if not os.path.exists(file_path):
        return {
            "valid": False,
            "file_type": "unknown",
            "rows": 0,
            "columns": 0,
            "error": f"File not found: {file_path}"
        }

    # Size check: reject zero-byte or over 100MB files
    file_size = os.path.getsize(file_path)
    if file_size == 0:
        return {
            "valid": False,
            "file_type": "unknown",
            "rows": 0,
            "columns": 0,
            "error": "The uploaded file is empty (0 bytes)."
        }

    if file_size > MAX_FILE_SIZE_BYTES:
        return {
            "valid": False,
            "file_type": "unknown",
            "rows": 0,
            "columns": 0,
            "error": f"File size ({file_size / (1024*1024):.2f} MB) exceeds maximum allowed size of 100 MB."
        }

    _, ext = os.path.splitext(file_path)
    ext_lower = ext.lower()

    if ext_lower not in SUPPORTED_EXTENSIONS:
        return {
            "valid": False,
            "file_type": ext_lower.lstrip("."),
            "rows": 0,
            "columns": 0,
            "error": f"Unsupported file type '{ext}'. Only .xlsx and .csv files are supported."
        }

    if ext_lower == ".csv":
        return _validate_csv(file_path)
    else:
        return _validate_xlsx(file_path)


def _validate_csv(file_path: str) -> Dict[str, Any]:
    """Validate CSV file structure, row and column limits."""
    encodings = ["utf-8", "utf-8-sig", "latin1", "cp1252"]
    success = False
    row_count = 0
    col_count = 0
    error_msg = None

    for enc in encodings:
        try:
            with open(file_path, "r", encoding=enc, errors="strict") as f:
                first_line = f.readline()
                if not first_line:
                    return {
                        "valid": False,
                        "file_type": "csv",
                        "rows": 0,
                        "columns": 0,
                        "error": "CSV file contains no data or header."
                    }
                # Count columns from header (handle comma, semicolon, tab)
                delimiters = [",", ";", "\t"]
                col_counts = {d: len(first_line.split(d)) for d in delimiters}
                best_delim = max(col_counts, key=col_counts.get)
                col_count = col_counts[best_delim]

                # Count rows
                f.seek(0)
                # Count remaining lines
                row_count = sum(1 for line in f if line.strip()) - 1  # Exclude header
                if row_count < 0:
                    row_count = 0

            success = True
            break
        except (UnicodeDecodeError, UnicodeError):
            continue
        except Exception as e:
            error_msg = f"Error reading CSV file: {str(e)}"
            break

    if not success and not error_msg:
        error_msg = "Could not decode CSV file using supported encodings (utf-8, utf-8-sig, latin1, cp1252)."

    if error_msg:
        return {
            "valid": False,
            "file_type": "csv",
            "rows": 0,
            "columns": 0,
            "error": error_msg
        }

    if col_count > MAX_CSV_COLUMNS:
        return {
            "valid": False,
            "file_type": "csv",
            "rows": row_count,
            "columns": col_count,
            "error": f"CSV has {col_count} columns, exceeding the maximum limit of {MAX_CSV_COLUMNS}."
        }

    if row_count > MAX_CSV_ROWS:
        return {
            "valid": False,
            "file_type": "csv",
            "rows": row_count,
            "columns": col_count,
            "error": f"CSV has {row_count} rows, exceeding the maximum limit of {MAX_CSV_ROWS}."
        }

    return {
        "valid": True,
        "file_type": "csv",
        "rows": row_count,
        "columns": col_count,
        "error": None
    }


def _validate_xlsx(file_path: str) -> Dict[str, Any]:
    """Validate XLSX workbook readability and sheet existence."""
    try:
        wb = openpyxl.load_workbook(file_path, read_only=True, data_only=True)
        sheet_names = wb.sheetnames

        if not sheet_names:
            wb.close()
            return {
                "valid": False,
                "file_type": "xlsx",
                "rows": 0,
                "columns": 0,
                "error": "Excel file contains no worksheets."
            }

        # Check that at least one sheet has data
        total_rows = 0
        total_cols = 0
        has_any_data = False

        for name in sheet_names:
            ws = wb[name]
            if ws.max_row and ws.max_row > 1 and ws.max_column and ws.max_column >= 1:
                has_any_data = True
                if ws.max_row > total_rows:
                    total_rows = ws.max_row - 1  # Excluding header row
                    total_cols = ws.max_column

        wb.close()

        if not has_any_data:
            return {
                "valid": False,
                "file_type": "xlsx",
                "rows": 0,
                "columns": 0,
                "error": "Excel workbook has no readable data rows."
            }

        return {
            "valid": True,
            "file_type": "xlsx",
            "rows": total_rows,
            "columns": total_cols,
            "error": None
        }

    except Exception as e:
        return {
            "valid": False,
            "file_type": "xlsx",
            "rows": 0,
            "columns": 0,
            "error": f"Unable to read Excel workbook: {str(e)}"
        }
