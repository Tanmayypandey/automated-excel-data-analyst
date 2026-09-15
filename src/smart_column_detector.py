"""Smart Column Detector Module.

Fast semantic role detection for tabular datasets.

Detects:
- identifier
- datetime
- numeric
- percentage
- currency
- rating
- count
- measure/duration
- categorical

Designed to avoid expensive full-column operations on large datasets.
"""

import re
from typing import Dict, Any

import pandas as pd


# ---------------------------------------------------------
# KEYWORDS
# ---------------------------------------------------------

DURATION_TIME_KEYWORDS = [
    r"active.*time",
    r"response.*time",
    r"hold.*time",
    r"call.*time",
    r"wait.*time",
    r"elapsed.*time",
    r"time.*spent",
    r"duration",
    r"handle.*time",
    r"tat",
    r"turnaround.*time",
    r"cycle.*time",
    r"lead.*time",
    r"active_with_cust",
    r"talk_time",
    r"wrap_up_time",
]

DATETIME_KEYWORDS = [
    r"^date$",
    r".*_date$",
    r"^dob$",
    r"^timestamp$",
    r"^created_?at$",
    r"^updated_?at$",
    r"^order_?date$",
    r"^sale_?date$",
    r"^purchase_?date$",
    r"^ship_?date$",
    r"^due_?date$",
    r"datetime",
]

PERCENTAGE_KEYWORDS = [
    r"percent",
    r"percentage",
    r"pct",
    r"%",
    r"rate$",
    r"efficiency",
    r"ratio",
    r"margin_pct",
    r"share_pct",
    r"proportion",
]

CURRENCY_KEYWORDS = [
    r"revenue",
    r"profit",
    r"sales",
    r"price",
    r"cost",
    r"salary",
    r"amount",
    r"spend",
    r"budget",
    r"fee",
    r"earning",
    r"income",
    r"mrr",
    r"arr",
    r"commission",
]

RATING_KEYWORDS = [
    r"rating",
    r"score",
    r"csat",
    r"nps",
    r"stars",
    r"feedback",
    r"grade",
    r"evaluation",
]

COUNT_KEYWORDS = [
    r"count",
    r"quantity",
    r"qty",
    r"units",
    r"orders",
    r"calls",
    r"attended",
    r"visits",
    r"clicks",
    r"tickets",
    r"items",
    r"sessions",
    r"views",
    r"inventory",
]

ID_KEYWORDS = [
    r"^id$",
    r".*_id$",
    r"^code$",
    r".*_code$",
    r"^uuid$",
    r"^key$",
    r"^sku$",
    r"^barcode$",
    r"^account_?no$",
    r"^invoice_?no$",
    r"^employee_?id$",
    r"^customer_?id$",
    r"^order_?id$",
    r"^booking_?id$",
    r"^transaction_?id$",
]

CURRENCY_SYMBOLS = {
    "₹": "INR",
    "$": "USD",
    "€": "EUR",
    "£": "GBP",
}


# ---------------------------------------------------------
# HELPERS
# ---------------------------------------------------------

def _matches_any(patterns, text: str) -> bool:
    return any(re.search(pattern, text) for pattern in patterns)


def _sample_series(series: pd.Series, sample_size: int = 500) -> pd.Series:
    """Return a bounded non-null sample without scanning/copying excessively."""
    non_null = series.dropna()

    if len(non_null) <= sample_size:
        return non_null

    return non_null.iloc[:sample_size]


def _looks_numeric(sample: pd.Series) -> bool:
    if sample.empty:
        return False

    if pd.api.types.is_numeric_dtype(sample):
        return True

    as_str = (
        sample.astype(str)
        .str.strip()
        .str.replace(",", "", regex=False)
        .str.replace("%", "", regex=False)
        .str.replace("₹", "", regex=False)
        .str.replace("$", "", regex=False)
        .str.replace("€", "", regex=False)
        .str.replace("£", "", regex=False)
    )

    parsed = pd.to_numeric(as_str, errors="coerce")

    return parsed.notna().mean() >= 0.80


def _detect_currency_symbol(sample: pd.Series):
    if sample.empty:
        return None

    sample_strings = sample.astype(str).head(50)

    for symbol in CURRENCY_SYMBOLS:
        if sample_strings.str.contains(
            re.escape(symbol),
            regex=True,
            na=False
        ).any():
            return symbol

    return None


def _looks_datetime(sample: pd.Series) -> bool:
    """Detect actual date-like values from a bounded sample."""
    if sample.empty:
        return False

    if pd.api.types.is_datetime64_any_dtype(sample):
        return True

    # Prevent numeric columns such as 1, 2, 3 or durations
    if pd.api.types.is_numeric_dtype(sample):
        return False

    sample = sample.head(30)

    sample_strings = sample.astype(str).str.strip()

    # Prevent short pure integer values being interpreted as timestamps
    pure_small_numbers = sample_strings.str.match(r"^\d{1,4}$")

    if pure_small_numbers.mean() > 0.80:
        return False

    try:
        parsed = pd.to_datetime(
            sample_strings,
            errors="coerce",
            format="mixed"
        )

        return parsed.notna().mean() >= 0.75

    except Exception:
        return False


# ---------------------------------------------------------
# MAIN DETECTOR
# ---------------------------------------------------------

def detect_column_roles(
    df: pd.DataFrame,
    sample_size: int = 500
) -> Dict[str, Dict[str, Any]]:
    """Detect semantic roles efficiently.

    Large datasets are inspected using samples rather than repeatedly scanning
    complete columns.

    Args:
        df:
            Input DataFrame.

        sample_size:
            Maximum number of values used for semantic inspection.

    Returns:
        Mapping:
            column_name -> metadata dictionary
    """

    roles: Dict[str, Dict[str, Any]] = {}

    total_rows = len(df)

    for col in df.columns:
        col_name = str(col)

        col_clean = (
            col_name
            .strip()
            .lower()
            .replace(" ", "_")
        )

        series = df[col]

        dtype_str = str(series.dtype)

        sample = _sample_series(
            series,
            sample_size=sample_size
        )

        sample_count = len(sample)

        # Unique count only on sample for performance.
        # Exact count is unnecessary for semantic classification.
        sample_unique = (
            int(sample.nunique(dropna=True))
            if sample_count
            else 0
        )

        is_numeric_dtype = pd.api.types.is_numeric_dtype(series)

        is_numeric = (
            is_numeric_dtype
            or _looks_numeric(sample)
        )

        is_duration = _matches_any(
            DURATION_TIME_KEYWORDS,
            col_clean
        )

        # -------------------------------------------------
        # ID DETECTION
        # -------------------------------------------------

        is_id_name = _matches_any(
            ID_KEYWORDS,
            col_clean
        )

        # High uniqueness can support ID detection,
        # but only when column name is also somewhat ID-like.
        high_uniqueness = (
            sample_count >= 10
            and sample_unique / sample_count >= 0.98
        )

        weak_id_hint = any(
            token in col_clean
            for token in [
                "identifier",
                "serial",
                "reference",
                "ref_no",
                "number",
            ]
        )

        if is_id_name or (
            weak_id_hint and high_uniqueness
        ):
            roles[col_name] = {
                "role": "identifier",
                "data_type": dtype_str,
                "currency_symbol": None,
                "is_duration": False,
                "unique_count": sample_unique,
                "is_numeric": is_numeric,
            }

            continue

        # -------------------------------------------------
        # DATETIME DETECTION
        # -------------------------------------------------

        is_date_name = _matches_any(
            DATETIME_KEYWORDS,
            col_clean
        )

        is_datetime = False

        if not is_duration:

            if pd.api.types.is_datetime64_any_dtype(series):
                is_datetime = True

            elif is_date_name and _looks_datetime(sample):
                is_datetime = True

            elif (
                not is_numeric
                and _looks_datetime(sample)
            ):
                is_datetime = True

        if is_datetime:
            roles[col_name] = {
                "role": "datetime",
                "data_type": dtype_str,
                "currency_symbol": None,
                "is_duration": False,
                "unique_count": sample_unique,
                "is_numeric": False,
            }

            continue

        # -------------------------------------------------
        # PERCENTAGE
        # -------------------------------------------------

        sample_strings = (
            sample.astype(str).head(50)
            if sample_count
            else pd.Series(dtype="object")
        )

        has_percent_symbol = (
            sample_strings
            .str.contains("%", regex=False, na=False)
            .any()
            if sample_count
            else False
        )

        is_percentage_name = _matches_any(
            PERCENTAGE_KEYWORDS,
            col_clean
        )

        if has_percent_symbol or (
            is_percentage_name and is_numeric
        ):
            roles[col_name] = {
                "role": "percentage",
                "data_type": dtype_str,
                "currency_symbol": None,
                "is_duration": False,
                "unique_count": sample_unique,
                "is_numeric": True,
            }

            continue

        # -------------------------------------------------
        # CURRENCY
        # -------------------------------------------------

        detected_currency = _detect_currency_symbol(sample)

        currency_name = _matches_any(
            CURRENCY_KEYWORDS,
            col_clean
        )

        if (
            detected_currency is not None
            or (currency_name and is_numeric)
        ):
            roles[col_name] = {
                "role": "currency",
                "data_type": dtype_str,
                "currency_symbol": detected_currency,
                "currency_code": (
                    CURRENCY_SYMBOLS.get(detected_currency)
                    if detected_currency
                    else None
                ),
                "is_duration": False,
                "unique_count": sample_unique,
                "is_numeric": True,
            }

            continue

        # -------------------------------------------------
        # RATING
        # -------------------------------------------------

        if (
            _matches_any(RATING_KEYWORDS, col_clean)
            and is_numeric
        ):
            roles[col_name] = {
                "role": "rating",
                "data_type": dtype_str,
                "currency_symbol": None,
                "is_duration": False,
                "unique_count": sample_unique,
                "is_numeric": True,
            }

            continue

        # -------------------------------------------------
        # COUNT
        # -------------------------------------------------

        if (
            _matches_any(COUNT_KEYWORDS, col_clean)
            and is_numeric
        ):
            roles[col_name] = {
                "role": "count",
                "data_type": dtype_str,
                "currency_symbol": None,
                "is_duration": False,
                "unique_count": sample_unique,
                "is_numeric": True,
            }

            continue

        # -------------------------------------------------
        # DURATION / MEASURE
        # -------------------------------------------------

        if is_duration and is_numeric:
            roles[col_name] = {
                "role": "measure",
                "data_type": dtype_str,
                "currency_symbol": None,
                "is_duration": True,
                "unique_count": sample_unique,
                "is_numeric": True,
            }

            continue

        # -------------------------------------------------
        # GENERAL NUMERIC
        # -------------------------------------------------

        if is_numeric:
            roles[col_name] = {
                "role": "numeric",
                "data_type": dtype_str,
                "currency_symbol": None,
                "is_duration": False,
                "unique_count": sample_unique,
                "is_numeric": True,
            }

            continue

        # -------------------------------------------------
        # CATEGORICAL / TEXT
        # -------------------------------------------------

        roles[col_name] = {
            "role": "categorical",
            "data_type": dtype_str,
            "currency_symbol": None,
            "is_duration": False,
            "unique_count": sample_unique,
            "is_numeric": False,
        }

    return roles