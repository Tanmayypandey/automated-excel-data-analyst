"""Smart Chart Selector Module.

Intelligently selects up to 4 diverse, meaningful charts based on dataset structure:
- LINE: when a true datetime column exists, paired with primary metric
- BAR: categorical column (2-20 values) paired with top metric
- TOP-10 BAR: high-cardinality dimension (Employee, Product) paired with volume/metric
- PIE: low-cardinality categorical column (2-6 values)
- SCATTER: two correlated continuous numeric metrics
- HISTOGRAM: continuous numeric distribution with sensible binning
"""

import re
from typing import Dict, Any, List, Optional
import pandas as pd


def select_charts(
    df: pd.DataFrame,
    detected_roles: Optional[Dict[str, Dict[str, Any]]] = None
) -> List[Dict[str, Any]]:
    """Select up to 4 high-value, diverse charts tailored to dataset characteristics.

    Args:
        df: Cleaned pandas DataFrame.
        detected_roles: Semantic roles mapping for columns.

    Returns:
        List of up to 4 chart specifications.
    """
    if detected_roles is None:
        from src.smart_column_detector import detect_column_roles
        detected_roles = detect_column_roles(df)

    # Classify columns into functional groups
    date_cols: List[str] = []
    cat_low_cols: List[str] = []   # 2 to 6 unique values
    cat_mid_cols: List[str] = []   # 7 to 20 unique values
    cat_high_cols: List[str] = []  # > 20 unique values (e.g. Employee, Product)
    measure_cols: List[str] = []   # currency, count, measure, numeric
    currency_cols: List[str] = []
    percentage_cols: List[str] = []

    for col, meta in detected_roles.items():
        if col not in df.columns:
            continue
        role = meta.get("role")
        n_uniq = df[col].nunique(dropna=True)

        if role == "datetime":
            date_cols.append(col)
        elif role == "identifier":
            # If named employee or product, allow high-cardinality ranking
            col_l = str(col).lower()
            if any(k in col_l for k in ["employee", "emp", "agent", "product", "item", "rep"]):
                cat_high_cols.append(col)
        elif role == "categorical":
            if 2 <= n_uniq <= 6:
                cat_low_cols.append(col)
            elif 7 <= n_uniq <= 20:
                cat_mid_cols.append(col)
            elif n_uniq > 20:
                cat_high_cols.append(col)
        elif role == "currency":
            currency_cols.append(col)
            measure_cols.append(col)
        elif role == "percentage":
            percentage_cols.append(col)
            measure_cols.append(col)
        elif role in ["count", "measure", "numeric", "rating"]:
            if n_uniq > 1:
                measure_cols.append(col)

    # Rank measures by business priority: currency > count > measure > percentage > numeric
    ranked_measures: List[str] = []
    for c in currency_cols:
        if c not in ranked_measures:
            ranked_measures.append(c)
    for c, m in detected_roles.items():
        if m.get("role") == "count" and c in measure_cols and c not in ranked_measures:
            ranked_measures.append(c)
    for c in measure_cols:
        if c not in ranked_measures:
            ranked_measures.append(c)

    charts: List[Dict[str, Any]] = []
    used_types = set()
    used_col_pairs = set()

    # Rule 1: LINE CHART (if true datetime column exists and measure exists)
    if date_cols and ranked_measures:
        d_col = date_cols[0]
        m_col = ranked_measures[0]
        agg_func = "mean" if detected_roles.get(m_col, {}).get("role") in ["percentage", "rating"] else "sum"
        charts.append({
            "chart_id": f"chart_{len(charts)+1}",
            "chart_type": "line",
            "title": f"{m_col.replace('_', ' ').title()} Over Time",
            "x_col": d_col,
            "y_col": m_col,
            "aggregation": agg_func,
            "description": f"Time series trend of {m_col} across {d_col}."
        })
        used_types.add("line")
        used_col_pairs.add((d_col, m_col))

    # Rule 2: BAR CHART (Categorical column 2-20 values + Top measure)
    candidate_bar_cats = cat_mid_cols + cat_low_cols
    if candidate_bar_cats and ranked_measures:
        # Pick category not already used if possible
        bar_cat = candidate_bar_cats[0]
        # Pick measure
        bar_m = ranked_measures[0] if (bar_cat, ranked_measures[0]) not in used_col_pairs else (
            ranked_measures[1] if len(ranked_measures) > 1 else ranked_measures[0]
        )
        agg_func = "mean" if detected_roles.get(bar_m, {}).get("role") in ["percentage", "rating"] else "sum"
        agg_word = "Average" if agg_func == "mean" else "Total"

        charts.append({
            "chart_id": f"chart_{len(charts)+1}",
            "chart_type": "bar",
            "title": f"{agg_word} {bar_m.replace('_', ' ').title()} by {bar_cat.replace('_', ' ').title()}",
            "x_col": bar_cat,
            "y_col": bar_m,
            "aggregation": agg_func,
            "description": f"Comparison of {bar_m} across {bar_cat} categories."
        })
        used_types.add("bar")
        used_col_pairs.add((bar_cat, bar_m))

    # Rule 3: PIE CHART (Low cardinality 2-6 values)
    if cat_low_cols and ranked_measures and len(charts) < 4:
        pie_cat = None
        for c in cat_low_cols:
            if not any(chart.get("x_col") == c for chart in charts):
                pie_cat = c
                break
        if not pie_cat:
            pie_cat = cat_low_cols[0]

        pie_m = ranked_measures[0]
        # Pie charts only make sense for positive sums
        if df[pie_m].min() >= 0:
            charts.append({
                "chart_id": f"chart_{len(charts)+1}",
                "chart_type": "pie",
                "title": f"{pie_m.replace('_', ' ').title()} Distribution by {pie_cat.replace('_', ' ').title()}",
                "x_col": pie_cat,
                "y_col": pie_m,
                "aggregation": "sum",
                "description": f"Proportional share of {pie_m} across {pie_cat} segments."
            })
            used_types.add("pie")
            used_col_pairs.add((pie_cat, pie_m))

    # Rule 4: TOP 10 BAR CHART (for high-cardinality dimensions like Employee or Product)
    if cat_high_cols and ranked_measures and len(charts) < 4:
        high_cat = cat_high_cols[0]
        top_m = ranked_measures[0]
        agg_func = "mean" if detected_roles.get(top_m, {}).get("role") in ["percentage", "rating"] else "sum"

        charts.append({
            "chart_id": f"chart_{len(charts)+1}",
            "chart_type": "bar",
            "title": f"Top 10 {high_cat.replace('_', ' ').title()}s by {top_m.replace('_', ' ').title()}",
            "x_col": high_cat,
            "y_col": top_m,
            "aggregation": f"top10_{agg_func}",
            "description": f"Leaderboard ranking top 10 {high_cat} performers by {top_m}."
        })
        used_types.add("bar")
        used_col_pairs.add((high_cat, top_m))

    # Rule 5: SCATTER PLOT (two distinct continuous numeric measures)
    if len(ranked_measures) >= 2 and len(charts) < 4:
        scat_x = ranked_measures[0]
        scat_y = ranked_measures[1]
        # Avoid percentage vs percentage if possible
        charts.append({
            "chart_id": f"chart_{len(charts)+1}",
            "chart_type": "scatter",
            "title": f"{scat_y.replace('_', ' ').title()} vs {scat_x.replace('_', ' ').title()}",
            "x_col": scat_x,
            "y_col": scat_y,
            "aggregation": "raw",
            "description": f"Correlation relationship between {scat_x} and {scat_y}."
        })
        used_types.add("scatter")
        used_col_pairs.add((scat_x, scat_y))

    # Rule 6: HISTOGRAM (continuous metric distribution)
    if ranked_measures and len(charts) < 4:
        hist_m = None
        for m in ranked_measures:
            if not any(c.get("chart_type") == "histogram" and c.get("x_col") == m for c in charts):
                hist_m = m
                break
        if hist_m:
            charts.append({
                "chart_id": f"chart_{len(charts)+1}",
                "chart_type": "histogram",
                "title": f"{hist_m.replace('_', ' ').title()} Distribution",
                "x_col": hist_m,
                "y_col": None,
                "aggregation": "histogram",
                "description": f"Frequency distribution across binned intervals of {hist_m}."
            })
            used_types.add("histogram")

    # If still fewer than 4 charts and we have more categories/measures
    if len(charts) < 4 and len(cat_mid_cols) > 1 and len(ranked_measures) > 1:
        cat = cat_mid_cols[1]
        m = ranked_measures[1]
        if (cat, m) not in used_col_pairs:
            charts.append({
                "chart_id": f"chart_{len(charts)+1}",
                "chart_type": "bar",
                "title": f"{m.replace('_', ' ').title()} by {cat.replace('_', ' ').title()}",
                "x_col": cat,
                "y_col": m,
                "aggregation": "sum",
                "description": f"Breakdown of {m} by {cat}."
            })

    # Return at most 4 charts
    return charts[:4]
