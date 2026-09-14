"""Data Cleaner Module.

Cleans and standardizes raw datasets:
- Standardizes column names and trims strings
- Normalizes missing indicators ("", "NA", "N/A", "null", "None", "-", "--")
- Preserves identifiers as clean strings
- Parses percentages ("25%" -> 0.25)
- Parses currency values (₹, $, €, £)
- Safely coerces numeric columns (>= 80% numeric-like values)
- Coerces true dates while strictly protecting numeric durations (active_time, etc.)
- Imputes missing data (numeric -> median, categorical -> mode, ID -> 'Unknown')
- Removes exact duplicate rows
- Generates a comprehensive audit trail of transformations
"""

import re
from typing import Dict, Any, Tuple, Optional, List
import pandas as pd
import numpy as np

from src.smart_column_detector import (
    DURATION_TIME_KEYWORDS,
    DATETIME_KEYWORDS,
    ID_KEYWORDS,
    CURRENCY_SYMBOLS,
    detect_column_roles
)


MISSING_VALUE_STRINGS = {
    "", "na", "n/a", "null", "none", "-", "--", "nan", "undefined", "?", "nil"
}


def clean_dataset(
    df: pd.DataFrame,
    detected_roles: Optional[Dict[str, Dict[str, Any]]] = None
) -> Tuple[pd.DataFrame, Dict[str, Any]]:
    """Clean and transform raw DataFrame into a structured, typed analysis-ready dataset.

    Args:
        df: Input raw pandas DataFrame.
        detected_roles: Optional pre-detected column roles. If None, detected dynamically.

    Returns:
        Tuple of (cleaned_df, cleaning_stats_dict).
    """
    df_clean = df.copy(deep=True)
    original_rows = len(df_clean)
    original_cols = len(df_clean.columns)

    # 1. Standardize column names
    col_name_mapping = {}
    for col in df_clean.columns:
        clean_col = str(col).strip()
        # Clean extra spaces and newlines
        clean_col = re.sub(r"\s+", " ", clean_col)
        col_name_mapping[col] = clean_col

    df_clean.rename(columns=col_name_mapping, inplace=True)

    # Detect roles if not provided
    if detected_roles is None:
        detected_roles = detect_column_roles(df_clean)

    # Audit tracking metrics
    stats: Dict[str, Any] = {
        "original_rows": original_rows,
        "cleaned_rows": 0,
        "original_columns": original_cols,
        "cleaned_columns": len(df_clean.columns),
        "duplicates_removed": 0,
        "missing_values_before": int(df_clean.isna().sum().sum()),
        "missing_values_after": 0,
        "columns_converted_to_numeric": [],
        "columns_converted_to_datetime": [],
        "percentage_conversions": [],
        "currency_conversions": [],
        "imputed_columns": {},
        "preserved_id_columns": [],
        "major_actions": []
    }

    # 2. Trim whitespace on all string cells and normalize missing value strings
    for col in df_clean.columns:
        if df_clean[col].dtype == "object":
            def _normalize_str(val):
                if pd.isna(val):
                    return np.nan
                s = str(val).strip()
                if s.lower() in MISSING_VALUE_STRINGS:
                    return np.nan
                return s

            df_clean[col] = df_clean[col].apply(_normalize_str)

    # Recalculate missing values before after string normalization
    stats["missing_values_before"] = int(df_clean.isna().sum().sum())

    # 3. Identify ID columns to strictly preserve
    id_columns = set()
    for col, meta in detected_roles.items():
        if col in df_clean.columns and meta.get("role") == "identifier":
            id_columns.add(col)
            stats["preserved_id_columns"].append(col)

    # Also double-check column names for ID keywords to guarantee preservation
    for col in df_clean.columns:
        col_lower = str(col).strip().lower()
        if any(re.search(pat, col_lower) for pat in ID_KEYWORDS):
            if col not in id_columns:
                id_columns.add(col)
                stats["preserved_id_columns"].append(col)

    if stats["preserved_id_columns"]:
        stats["major_actions"].append(
            f"Preserved identifier columns as clean text: {', '.join(stats['preserved_id_columns'])}"
        )

    # 4. Type conversions (Percentage, Currency, Numeric, Datetime)
    for col in df_clean.columns:
        if col in id_columns:
            # Preserve as string object
            df_clean[col] = df_clean[col].astype("object")
            continue

        col_lower = str(col).strip().lower()
        non_null_series = df_clean[col].dropna()
        if len(non_null_series) == 0:
            continue

        # A. Percentage conversion
        # Check if values end with '%' or column name indicates percentage
        sample_str = [str(x).strip() for x in non_null_series.head(50)]
        has_pct_symbol = any("%" in s for s in sample_str)

        if has_pct_symbol:
            def _parse_pct(val):
                if pd.isna(val):
                    return np.nan
                s = str(val).strip().replace("%", "").strip()
                try:
                    num = float(s)
                    return num / 100.0
                except ValueError:
                    return np.nan

            df_clean[col] = df_clean[col].apply(_parse_pct)
            stats["percentage_conversions"].append(col)
            stats["major_actions"].append(f"Converted percentage column '{col}' from text to decimal ratio.")
            continue

        # If already numeric between 0 and 100 with name like 'percentage' or 'efficiency in percentage'
        if any(w in col_lower for w in ["percentage", "pct", "%"]) and pd.api.types.is_numeric_dtype(df_clean[col]):
            if non_null_series.max() > 1.0 and non_null_series.max() <= 100.0:
                # Store normalized decimal for consistent analytics if specified as percentage
                # But keep note in stats
                df_clean[col] = df_clean[col] / 100.0
                stats["percentage_conversions"].append(col)
                stats["major_actions"].append(f"Normalized '{col}' values to decimal (0.0 to 1.0).")
                continue

        # B. Currency conversion (only when real currency symbols exist)
        has_curr_symbol = False
        detected_sym = None
        for sym in CURRENCY_SYMBOLS:
            if any(sym in s for s in sample_str):
                has_curr_symbol = True
                detected_sym = sym
                break

        if has_curr_symbol:
            def _parse_curr(val):
                if pd.isna(val):
                    return np.nan
                s = str(val).strip()
                for sym in CURRENCY_SYMBOLS:
                    s = s.replace(sym, "")
                s = s.replace(",", "").strip()
                try:
                    return float(s)
                except ValueError:
                    return np.nan

            df_clean[col] = df_clean[col].apply(_parse_curr)
            stats["currency_conversions"].append(f"{col} ({detected_sym})")
            stats["major_actions"].append(f"Parsed currency column '{col}' with symbol '{detected_sym}'.")
            continue

        # C. Date conversion (ONLY when column meaning strongly suggests true date)
        # CRITICAL CHECK: Never convert duration columns
        is_duration = any(re.search(pat, col_lower) for pat in DURATION_TIME_KEYWORDS)
        is_candidate_date = any(re.search(pat, col_lower) for pat in DATETIME_KEYWORDS)

        if not is_duration and is_candidate_date:
            try:
                import warnings
                with warnings.catch_warnings():
                    warnings.simplefilter("ignore", UserWarning)
                    parsed_dates = pd.to_datetime(df_clean[col], errors="coerce", format="mixed")
                if parsed_dates.notna().sum() / len(non_null_series) >= 0.75:
                    df_clean[col] = parsed_dates
                    stats["columns_converted_to_datetime"].append(col)
                    stats["major_actions"].append(f"Coerced '{col}' to datetime timestamps.")
                    continue
            except Exception:
                pass

        # D. Numeric conversion (when >=80% of values are numeric-like)
        if not pd.api.types.is_numeric_dtype(df_clean[col]) and not pd.api.types.is_datetime64_any_dtype(df_clean[col]):
            numeric_parsed = pd.to_numeric(df_clean[col].astype(str).str.replace(",", ""), errors="coerce")
            valid_ratio = numeric_parsed.notna().sum() / len(non_null_series)
            if valid_ratio >= 0.8:
                df_clean[col] = numeric_parsed
                stats["columns_converted_to_numeric"].append(col)
                stats["major_actions"].append(f"Coerced '{col}' to numeric ({valid_ratio*100:.1f}% numeric-like).")
                continue

    # 5. Missing value imputation
    for col in df_clean.columns:
        null_count = int(df_clean[col].isna().sum())
        if null_count > 0:
            if col in id_columns:
                df_clean[col] = df_clean[col].fillna("Unknown")
                stats["imputed_columns"][col] = "Filled with 'Unknown' (ID column)"
            elif pd.api.types.is_numeric_dtype(df_clean[col]):
                med = df_clean[col].median()
                if pd.isna(med):
                    med = 0
                df_clean[col] = df_clean[col].fillna(med)
                stats["imputed_columns"][col] = f"Imputed {null_count} missing values with median ({med:.2f})"
            elif pd.api.types.is_datetime64_any_dtype(df_clean[col]):
                # Fill missing dates with forward fill or most frequent
                mode_val = df_clean[col].mode()
                fill_date = mode_val.iloc[0] if len(mode_val) > 0 else pd.Timestamp.now()
                df_clean[col] = df_clean[col].fillna(fill_date)
                stats["imputed_columns"][col] = f"Imputed {null_count} missing dates with mode ({fill_date})"
            else:
                # Categorical
                mode_val = df_clean[col].mode()
                fill_cat = mode_val.iloc[0] if len(mode_val) > 0 else "Unknown"
                df_clean[col] = df_clean[col].fillna(fill_cat)
                stats["imputed_columns"][col] = f"Imputed {null_count} missing values with mode ('{fill_cat}')"

    # 6. Remove exact duplicate rows
    dupes_count = int(df_clean.duplicated().sum())
    if dupes_count > 0:
        df_clean.drop_duplicates(inplace=True)
        stats["duplicates_removed"] = dupes_count
        stats["major_actions"].append(f"Removed {dupes_count} exact duplicate rows.")

    df_clean.reset_index(drop=True, inplace=True)
    stats["cleaned_rows"] = len(df_clean)
    stats["missing_values_after"] = int(df_clean.isna().sum().sum())

    return df_clean, stats
