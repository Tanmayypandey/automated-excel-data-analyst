"""Data Loader Module.

Loads CSV and XLSX datasets into Pandas DataFrames with multi-encoding support
and intelligent sheet routing.
"""

import os
from typing import Tuple, Optional, Union
import pandas as pd
from src.sheet_selector import select_best_sheet


CSV_ENCODINGS = ["utf-8", "utf-8-sig", "latin1", "cp1252"]


def load_csv(file_path: str) -> pd.DataFrame:
    """Load CSV dataset with sequential encoding fallback and automatic delimiter detection.

    Args:
        file_path: Path to the .csv file.

    Returns:
        Pandas DataFrame containing loaded CSV data.

    Raises:
        ValueError: If file cannot be decoded using any supported encoding.
    """
    if not os.path.exists(file_path):
        raise FileNotFoundError(f"CSV file not found: {file_path}")

    last_err = None
    for enc in CSV_ENCODINGS:
        try:
            # Let pandas infer separator (comma, semicolon, tab) or default to comma
            df = pd.read_csv(file_path, encoding=enc, sep=None, engine="python", low_memory=False)
            return df
        except Exception as e:
            last_err = e
            # Fallback with standard comma delimiter
            try:
                df = pd.read_csv(file_path, encoding=enc, sep=",", low_memory=False)
                return df
            except Exception:
                continue

    raise ValueError(f"Failed to load CSV file '{file_path}'. Last error: {str(last_err)}")


def load_excel(file_path: str, sheet_name: Union[str, int] = 0) -> pd.DataFrame:
    """Load Excel sheet into a DataFrame.

    Args:
        file_path: Path to the .xlsx file.
        sheet_name: Worksheet name or 0-based index.

    Returns:
        Pandas DataFrame containing the sheet data.
    """
    if not os.path.exists(file_path):
        raise FileNotFoundError(f"Excel file not found: {file_path}")

    df = pd.read_excel(file_path, sheet_name=sheet_name)
    return df


def load_data(file_path: str, sheet_name: Optional[Union[str, int]] = None) -> Tuple[pd.DataFrame, str]:
    """Unified data loader for CSV and XLSX files.

    For CSV:
        Directly loads file and returns (df, "CSV Data").

    For XLSX:
        If sheet_name is provided, loads that sheet.
        If sheet_name is None, runs smart sheet selection and loads the best sheet.

    Args:
        file_path: Path to input dataset (.csv or .xlsx).
        sheet_name: Optional worksheet name for Excel files.

    Returns:
        Tuple of (DataFrame, selected_sheet_name).
    """
    _, ext = os.path.splitext(file_path)
    ext_lower = ext.lower()

    if ext_lower == ".csv":
        df = load_csv(file_path)
        return df, "CSV Data"
    elif ext_lower == ".xlsx":
        if sheet_name is None:
            selection_res = select_best_sheet(file_path)
            selected_sheet = selection_res["selected_sheet"]
        else:
            selected_sheet = sheet_name

        df = load_excel(file_path, sheet_name=selected_sheet)
        return df, str(selected_sheet)
    else:
        raise ValueError(f"Unsupported file format '{ext}'. Must be .xlsx or .csv")
