"""KPI Engine Module.

Calculates aggregated business metrics across columns based on detected roles.
"""

from typing import Dict, Any, List, Optional
import pandas as pd
import numpy as np


def compute_column_metrics(
    df: pd.DataFrame,
    col: str,
    role_meta: Dict[str, Any]
) -> List[Dict[str, Any]]:
    """Compute candidate KPI aggregates for a single column.

    Args:
        df: Cleaned pandas DataFrame.
        col: Column name.
        role_meta: Column metadata including role, is_duration, etc.

    Returns:
        List of candidate KPI dictionaries.
    """
    candidates: List[Dict[str, Any]] = []
    role = role_meta.get("role", "numeric")
    curr_sym = role_meta.get("currency_symbol")
    is_duration = role_meta.get("is_duration", False)

    series = df[col].dropna()
    if len(series) == 0 or series.nunique() <= 1:
        return candidates

    # Verify not all zeros
    if (series == 0).all():
        return candidates

    col_title = str(col).replace("_", " ").title()

    if role == "currency":
        total_val = float(series.sum())
        mean_val = float(series.mean())
        candidates.append({
            "label": f"Total {col_title}",
            "value": total_val,
            "role": "currency",
            "column": col,
            "aggregation": "sum",
            "currency_symbol": curr_sym or "₹",
            "priority": 90
        })
        candidates.append({
            "label": f"Avg {col_title}",
            "value": mean_val,
            "role": "currency",
            "column": col,
            "aggregation": "mean",
            "currency_symbol": curr_sym or "₹",
            "priority": 75
        })

    elif role == "percentage":
        mean_val = float(series.mean())
        candidates.append({
            "label": f"Avg {col_title}",
            "value": mean_val,
            "role": "percentage",
            "column": col,
            "aggregation": "mean",
            "currency_symbol": None,
            "priority": 85
        })

    elif role == "rating":
        mean_val = float(series.mean())
        candidates.append({
            "label": f"Avg {col_title}",
            "value": mean_val,
            "role": "rating",
            "column": col,
            "aggregation": "mean",
            "currency_symbol": None,
            "priority": 80
        })

    elif role == "count":
        total_val = float(series.sum())
        candidates.append({
            "label": f"Total {col_title}",
            "value": total_val,
            "role": "count",
            "column": col,
            "aggregation": "sum",
            "currency_symbol": None,
            "priority": 78
        })

    elif role == "measure" or is_duration:
        mean_val = float(series.mean())
        candidates.append({
            "label": f"Avg {col_title}",
            "value": mean_val,
            "role": "numeric",
            "column": col,
            "aggregation": "mean",
            "currency_symbol": None,
            "priority": 70
        })

    elif role == "numeric":
        mean_val = float(series.mean())
        candidates.append({
            "label": f"Avg {col_title}",
            "value": mean_val,
            "role": "numeric",
            "column": col,
            "aggregation": "mean",
            "currency_symbol": None,
            "priority": 60
        })

    return candidates
