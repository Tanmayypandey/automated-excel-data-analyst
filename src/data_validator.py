"""Data Validator Module.

Inspects DataFrame structure, data types, missing values, duplicates,
constant columns, and potential identifiers.
"""

import re
from typing import Dict, Any, List
import pandas as pd


ID_NAME_PATTERNS = [
    r"^id$", r".*_id$", r"^.*_code$", r"^code$", r"^uuid$",
    r"^key$", r"^num$", r"^number$", r"^employee.*", r"^emp_.*",
    r"^customer.*", r"^order.*id.*", r"^transaction.*id.*",
    r"^invoice.*", r"^ticket.*", r"^account.*"
]


def validate_dataset(df: pd.DataFrame) -> Dict[str, Any]:
    """Perform structural health check and profile candidate columns of a DataFrame.

    Args:
        df: Input pandas DataFrame.

    Returns:
        Dictionary containing dataset dimensions, missing counts, duplicate counts,
        data types, constant columns, high cardinality columns, and potential IDs.
    """
    total_rows = len(df)
    total_columns = len(df.columns)

    if total_rows == 0:
        return {
            "total_rows": 0,
            "total_columns": total_columns,
            "missing_values_total": 0,
            "missing_values_per_column": {},
            "duplicate_rows": 0,
            "data_types": {},
            "fully_empty_columns": list(df.columns),
            "constant_columns": [],
            "high_cardinality_columns": [],
            "potential_id_columns": []
        }

    # Missing values
    null_counts = df.isna().sum().to_dict()
    total_nulls = sum(null_counts.values())

    # Duplicate rows
    duplicate_rows = int(df.duplicated().sum())

    # Data types as string
    dtypes = {str(col): str(dtype) for col, dtype in df.dtypes.items()}

    # Fully empty columns
    fully_empty = [str(col) for col in df.columns if df[col].isna().all()]

    # Constant columns (nunique == 1 ignoring or including nulls)
    constant_cols = []
    for col in df.columns:
        if df[col].nunique(dropna=False) <= 1:
            constant_cols.append(str(col))

    # Potential ID columns and High cardinality
    potential_ids: List[str] = []
    high_cardinality: List[str] = []

    for col in df.columns:
        col_str = str(col)
        col_lower = col_str.strip().lower()
        nuniques = df[col].nunique(dropna=True)

        # Check name pattern matches
        is_id_by_name = any(re.search(pat, col_lower) for pat in ID_NAME_PATTERNS)
        # Check if column is almost 100% unique strings or integers with no mathematical variance
        is_id_by_cardinality = (total_rows > 5 and nuniques == total_rows and df[col].dtype == "object")

        if is_id_by_name or is_id_by_cardinality:
            potential_ids.append(col_str)

        # High cardinality for categorical/object columns (excluding detected IDs)
        if df[col].dtype == "object" and not is_id_by_name:
            if (nuniques > 50 or (total_rows > 10 and nuniques / total_rows > 0.5)) and nuniques > 5:
                high_cardinality.append(col_str)

    return {
        "total_rows": total_rows,
        "total_columns": total_columns,
        "missing_values_total": int(total_nulls),
        "missing_values_per_column": {str(k): int(v) for k, v in null_counts.items()},
        "duplicate_rows": duplicate_rows,
        "data_types": dtypes,
        "fully_empty_columns": fully_empty,
        "constant_columns": constant_cols,
        "high_cardinality_columns": high_cardinality,
        "potential_id_columns": potential_ids
    }
