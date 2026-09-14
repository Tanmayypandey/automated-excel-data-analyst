"""Smart KPI Selector Module.

Intelligently selects up to 4 high-value business KPIs dynamically from any dataset.
Always includes 'Total Records' followed by the top business measures (revenue, profit,
efficiency, volume, ratings).
"""

from typing import Dict, Any, List, Optional
import pandas as pd
from src.kpi_engine import compute_column_metrics
from src.kpi_formatter import format_kpi_value


def select_kpis(
    df: pd.DataFrame,
    detected_roles: Optional[Dict[str, Dict[str, Any]]] = None
) -> List[Dict[str, Any]]:
    """Select up to 4 most meaningful business KPIs from the cleaned dataset.

    Args:
        df: Cleaned pandas DataFrame.
        detected_roles: Semantic roles mapping for columns.

    Returns:
        List of up to 4 KPI objects, each containing label, value, formatted_value,
        role, aggregation, and currency symbol.
    """
    total_records = len(df)
    kpis: List[Dict[str, Any]] = []

    # KPI #1: Total Records is always present
    kpis.append({
        "label": "Total Records",
        "value": total_records,
        "formatted_value": format_kpi_value(total_records, role="count"),
        "role": "count",
        "column": None,
        "aggregation": "count",
        "currency_symbol": None
    })

    if detected_roles is None:
        from src.smart_column_detector import detect_column_roles
        detected_roles = detect_column_roles(df)

    # Collect candidate metrics from all non-ID, non-constant columns
    candidates: List[Dict[str, Any]] = []
    for col, meta in detected_roles.items():
        if col not in df.columns:
            continue
        if meta.get("role") in ["identifier", "datetime", "categorical"]:
            continue
        if df[col].nunique(dropna=True) <= 1:
            continue

        col_candidates = compute_column_metrics(df, col, meta)
        candidates.extend(col_candidates)

    # Sort candidates by priority descending
    candidates.sort(key=lambda x: x["priority"], reverse=True)

    # Select up to 3 more diverse KPIs (so max total is 4)
    chosen_cols = set()
    for cand in candidates:
        if len(kpis) >= 4:
            break
        # Prefer variety of columns so we don't pick both Sum and Avg of the same column
        col = cand["column"]
        if col not in chosen_cols:
            chosen_cols.add(col)
            cand["formatted_value"] = format_kpi_value(
                cand["value"],
                role=cand["role"],
                currency_symbol=cand["currency_symbol"]
            )
            kpis.append(cand)

    # If still fewer than 4 KPIs and we have candidates from existing cols (e.g. average of same col)
    if len(kpis) < 4:
        for cand in candidates:
            if len(kpis) >= 4:
                break
            if cand not in kpis:
                cand["formatted_value"] = format_kpi_value(
                    cand["value"],
                    role=cand["role"],
                    currency_symbol=cand["currency_symbol"]
                )
                kpis.append(cand)

    return kpis
