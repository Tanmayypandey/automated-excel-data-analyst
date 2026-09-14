"""Smart Column Detector Module.

Analyzes column names, data types, sample values, and statistical distributions
to accurately infer business semantic roles (identifier, datetime, measure,
percentage, currency, rating, count, categorical).
"""

import re
from typing import Dict, Any, List, Optional
import pandas as pd
import numpy as np


# Duration / Metric keywords that must NEVER be misclassified as datetime
DURATION_TIME_KEYWORDS = [
    r"active.*time", r"response.*time", r"hold.*time", r"call.*time",
    r"wait.*time", r"elapsed.*time", r"time.*spent", r"duration",
    r"handle.*time", r"tat", r"turnaround.*time", r"cycle.*time",
    r"lead.*time", r"active_with_cust", r"talk_time", r"wrap_up_time"
]

DATETIME_KEYWORDS = [
    r"^date$", r".*_date$", r"^.*date.*", r"^dob$", r"^timestamp$",
    r"^created_?at$", r"^updated_?at$", r"^order_?date$", r"^sale_?date$",
    r"^purchase_?date$", r"^ship_?date$", r"^due_?date$", r"^month$",
    r"^year$", r"^day$", r"^week$", r"^period$"
]

PERCENTAGE_KEYWORDS = [
    r"percent", r"pct", r"%", r"rate$", r"efficiency",
    r"ratio", r"margin_pct", r"share_pct", r"proportion"
]

CURRENCY_KEYWORDS = [
    r"revenue", r"profit", r"sales", r"price", r"cost",
    r"salary", r"amount", r"spend", r"budget", r"fee",
    r"earning", r"income", r"mrr", r"arr", r"commission"
]

RATING_KEYWORDS = [
    r"rating", r"score", r"csat", r"nps", r"stars",
    r"feedback", r"grade", r"evaluation"
]

COUNT_KEYWORDS = [
    r"count", r"quantity", r"qty", r"units", r"orders",
    r"calls", r"attended", r"visits", r"clicks", r"tickets",
    r"items", r"sessions", r"views", r"inventory"
]

ID_KEYWORDS = [
    r"^id$", r".*_id$", r"^code$", r".*_code$", r"^uuid$",
    r"^key$", r"^sku$", r"^barcode$", r"^account_?no$",
    r"^invoice_?no$", r"^employee$", r"^customer_?id$", r"^order_?id$"
]

CURRENCY_SYMBOLS = {"₹": "INR", "$": "USD", "€": "EUR", "£": "GBP"}


def detect_column_roles(df: pd.DataFrame) -> Dict[str, Dict[str, Any]]:
    """Detect the semantic role for each column in the DataFrame.

    Args:
        df: Input pandas DataFrame.

    Returns:
        Mapping of column_name -> {
            "role": str,
            "data_type": str,
            "currency_symbol": Optional[str],
            "is_duration": bool,
            "unique_count": int,
            "is_numeric": bool
        }
    """
    total_rows = len(df)
    roles: Dict[str, Dict[str, Any]] = {}

    for col in df.columns:
        col_str = str(col)
        col_clean = col_str.strip().lower().replace(" ", "_")
        series = df[col].dropna()
        n_unique = int(series.nunique())
        dtype_str = str(df[col].dtype)

        # Check sample string representations for currency or percentage symbols
        sample_strings = [str(v).strip() for v in series.head(50)]
        has_percent_symbol = any("%" in s for s in sample_strings)
        detected_currency = None
        for sym in CURRENCY_SYMBOLS:
            if any(sym in s for s in sample_strings):
                detected_currency = sym
                break

        # Check for duration indicators
        is_duration = any(re.search(pat, col_clean) for pat in DURATION_TIME_KEYWORDS)

        # Check if column is numeric or mostly numeric
        is_numeric = pd.api.types.is_numeric_dtype(df[col])
        if not is_numeric and len(sample_strings) > 0:
            # Check if strings look like numbers or percentages or currencies
            numeric_like = 0
            for s in sample_strings:
                clean_s = re.sub(r"[₹\$€£,%]", "", s).strip()
                try:
                    float(clean_s)
                    numeric_like += 1
                except ValueError:
                    pass
            if numeric_like / len(sample_strings) >= 0.8:
                is_numeric = True

        # 1. Check Identifier
        is_id_name = any(re.search(pat, col_clean) for pat in ID_KEYWORDS)
        if is_id_name or (not is_numeric and n_unique == total_rows and total_rows > 5):
            roles[col_str] = {
                "role": "identifier",
                "data_type": dtype_str,
                "currency_symbol": None,
                "is_duration": False,
                "unique_count": n_unique,
                "is_numeric": is_numeric
            }
            continue

        # 2. Check Datetime
        # CRITICAL RULE: If is_duration is True, NEVER classify as datetime
        is_date_col = False
        if not is_duration:
            if pd.api.types.is_datetime64_any_dtype(df[col]):
                is_date_col = True
            else:
                has_date_name = any(re.search(pat, col_clean) for pat in DATETIME_KEYWORDS)
                # If name suggests date and values are not pure floating numbers
                if has_date_name:
                    is_date_col = True
                elif not is_numeric and len(sample_strings) > 0:
                    # Test date parsing on samples
                    try:
                        import warnings
                        with warnings.catch_warnings():
                            warnings.simplefilter("ignore", UserWarning)
                            parsed = pd.to_datetime(series.head(10), errors="coerce", format="mixed")
                        if parsed.notna().sum() >= 8:
                            # Verify they are not just integers interpreted as timestamps
                            if not all(s.isdigit() and len(s) < 5 for s in sample_strings[:10]):
                                is_date_col = True
                    except Exception:
                        pass

        if is_date_col:
            roles[col_str] = {
                "role": "datetime",
                "data_type": dtype_str,
                "currency_symbol": None,
                "is_duration": False,
                "unique_count": n_unique,
                "is_numeric": False
            }
            continue

        # 3. Check Percentage
        is_pct = has_percent_symbol or any(re.search(pat, col_clean) for pat in PERCENTAGE_KEYWORDS)
        if is_pct:
            roles[col_str] = {
                "role": "percentage",
                "data_type": dtype_str,
                "currency_symbol": None,
                "is_duration": False,
                "unique_count": n_unique,
                "is_numeric": True
            }
            continue

        # 4. Check Currency
        is_currency = (detected_currency is not None) or any(re.search(pat, col_clean) for pat in CURRENCY_KEYWORDS)
        if is_currency and is_numeric:
            roles[col_str] = {
                "role": "currency",
                "data_type": dtype_str,
                "currency_symbol": detected_currency or "₹",
                "is_duration": False,
                "unique_count": n_unique,
                "is_numeric": True
            }
            continue

        # 5. Check Rating / Score
        is_rating = any(re.search(pat, col_clean) for pat in RATING_KEYWORDS)
        if is_rating and is_numeric:
            roles[col_str] = {
                "role": "rating",
                "data_type": dtype_str,
                "currency_symbol": None,
                "is_duration": False,
                "unique_count": n_unique,
                "is_numeric": True
            }
            continue

        # 6. Check Count
        is_count = any(re.search(pat, col_clean) for pat in COUNT_KEYWORDS)
        if is_count and is_numeric:
            roles[col_str] = {
                "role": "count",
                "data_type": dtype_str,
                "currency_symbol": None,
                "is_duration": False,
                "unique_count": n_unique,
                "is_numeric": True
            }
            continue

        # 7. Check Duration / Measure
        if is_duration and is_numeric:
            roles[col_str] = {
                "role": "measure",
                "data_type": dtype_str,
                "currency_symbol": None,
                "is_duration": True,
                "unique_count": n_unique,
                "is_numeric": True
            }
            continue

        # 8. Numeric general
        if is_numeric:
            roles[col_str] = {
                "role": "numeric",
                "data_type": dtype_str,
                "currency_symbol": None,
                "is_duration": False,
                "unique_count": n_unique,
                "is_numeric": True
            }
            continue

        # 9. Categorical
        roles[col_str] = {
            "role": "categorical",
            "data_type": dtype_str,
            "currency_symbol": None,
            "is_duration": False,
            "unique_count": n_unique,
            "is_numeric": False
        }

    return roles
