"""EDA Engine Module.

Generates automated exploratory data analysis profiles:
- Dataset shape and null counts
- Descriptive numeric statistics (mean, std, min, quantiles, max, skewness)
- Categorical frequency distributions
- Correlation matrix (excluding ID and constant columns)
- Metric histograms / distribution quantiles
"""

from typing import Dict, Any, List, Optional
import pandas as pd
import numpy as np


def run_eda(
    df: pd.DataFrame,
    detected_roles: Optional[Dict[str, Dict[str, Any]]] = None
) -> Dict[str, Any]:
    """Execute exploratory data analysis across cleaned DataFrame.

    Args:
        df: Cleaned pandas DataFrame.
        detected_roles: Optional column roles mapping to exclude IDs and constants.

    Returns:
        Structured EDA dictionary with statistical profiles and correlations.
    """
    total_rows = len(df)
    total_cols = len(df.columns)

    if detected_roles is None:
        from src.smart_column_detector import detect_column_roles
        detected_roles = detect_column_roles(df)

    id_cols = {col for col, meta in detected_roles.items() if meta.get("role") == "identifier"}

    # 1. Missing value summary
    missing_summary = {}
    for col in df.columns:
        nulls = int(df[col].isna().sum())
        missing_summary[str(col)] = {
            "null_count": nulls,
            "null_percentage": round((nulls / total_rows * 100) if total_rows > 0 else 0, 2)
        }

    # 2. Duplicate count
    duplicate_count = int(df.duplicated().sum())

    # 3. Numeric descriptive statistics
    numeric_stats: Dict[str, Dict[str, Any]] = {}
    numeric_cols = [
        col for col in df.columns
        if pd.api.types.is_numeric_dtype(df[col]) and col not in id_cols
    ]

    for col in numeric_cols:
        series = df[col].dropna()
        if len(series) == 0:
            continue
        try:
            mean_val = float(series.mean())
            std_val = float(series.std()) if len(series) > 1 else 0.0
            min_val = float(series.min())
            q25 = float(series.quantile(0.25))
            med_val = float(series.median())
            q75 = float(series.quantile(0.75))
            max_val = float(series.max())
            skew_val = float(series.skew()) if len(series) > 2 else 0.0

            numeric_stats[str(col)] = {
                "count": int(len(series)),
                "mean": round(mean_val, 3),
                "std": round(std_val, 3),
                "min": round(min_val, 3),
                "q25": round(q25, 3),
                "median": round(med_val, 3),
                "q75": round(q75, 3),
                "max": round(max_val, 3),
                "skewness": round(skew_val, 3)
            }
        except Exception:
            pass

    # 4. Categorical summaries
    categorical_summaries: Dict[str, Dict[str, Any]] = {}
    cat_cols = [
        col for col in df.columns
        if not pd.api.types.is_numeric_dtype(df[col]) and not pd.api.types.is_datetime64_any_dtype(df[col])
    ]

    for col in cat_cols:
        series = df[col].dropna().astype(str)
        if len(series) == 0:
            continue
        vc = series.value_counts().head(10)
        top_cats = []
        for val, count in vc.items():
            top_cats.append({
                "value": str(val),
                "count": int(count),
                "percentage": round((count / total_rows * 100) if total_rows > 0 else 0, 2)
            })
        categorical_summaries[str(col)] = {
            "unique_count": int(series.nunique()),
            "top_categories": top_cats
        }

    # 5. Correlation matrix (exclude IDs, datetime, and constant columns)
    correlations: Dict[str, Dict[str, float]] = {}
    valid_corr_cols = [
        c for c in numeric_cols
        if df[c].nunique(dropna=True) > 1
    ]

    if len(valid_corr_cols) >= 2:
        try:
            corr_df = df[valid_corr_cols].corr()
            for c1 in valid_corr_cols:
                correlations[str(c1)] = {}
                for c2 in valid_corr_cols:
                    val = corr_df.loc[c1, c2]
                    correlations[str(c1)][str(c2)] = round(float(val), 3) if not np.isnan(val) else 0.0
        except Exception:
            pass

    # 6. Important distributions (histogram binning for key numeric metrics)
    distributions: Dict[str, Any] = {}
    for col in numeric_cols[:5]:  # Top 5 numeric columns
        series = df[col].dropna()
        if len(series) >= 5 and series.nunique() > 2:
            try:
                counts, bin_edges = np.histogram(series, bins="auto")
                if len(counts) > 10:
                    counts, bin_edges = np.histogram(series, bins=10)
                distributions[str(col)] = {
                    "counts": [int(c) for c in counts],
                    "bin_edges": [round(float(b), 2) for b in bin_edges]
                }
            except Exception:
                pass

    return {
        "shape": {"rows": total_rows, "columns": total_cols},
        "missing_summary": missing_summary,
        "duplicate_count": duplicate_count,
        "numeric_statistics": numeric_stats,
        "categorical_summaries": categorical_summaries,
        "correlations": correlations,
        "distributions": distributions
    }
