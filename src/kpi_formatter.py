"""KPI Formatter Module.

Standardizes display formatting for metrics across both Web UI and Excel workbooks:
- percentage/rate/ratio -> 0.00%
- currency -> ₹#,##0.00 or $#,##0.00
- rating/score -> 0.00
- count/orders/customers/quantity/units/records -> #,##0
- default numeric -> #,##0.00
"""

from typing import Union, Optional


def format_kpi_value(
    value: Union[int, float],
    role: str = "numeric",
    currency_symbol: Optional[str] = "₹"
) -> str:
    """Format numeric KPI value into clean string representation.

    Args:
        value: Numeric value to format.
        role: Detected semantic role (percentage, currency, rating, count, numeric).
        currency_symbol: Detected or default currency symbol (e.g., '₹', '$', '€').

    Returns:
        Formatted string (e.g., '72.00%', '₹1,250.00', '4.50', '1,450').
    """
    if value is None:
        return "N/A"

    try:
        val = float(value)
    except (ValueError, TypeError):
        return str(value)

    role_lower = str(role).lower()
    symbol = currency_symbol or "₹"

    if role_lower in ["percentage", "rate", "ratio", "efficiency"]:
        # If value is already in decimal (e.g. 0.72)
        return f"{val * 100:.2f}%"

    elif role_lower in ["currency", "revenue", "profit", "sales", "cost"]:
        return f"{symbol}{val:,.2f}"

    elif role_lower in ["rating", "score", "csat", "nps"]:
        return f"{val:.2f}"

    elif role_lower in ["count", "records", "orders", "quantity", "units", "calls"]:
        return f"{int(round(val)):,}"

    else:
        # Default numeric
        if val.is_integer():
            return f"{int(val):,}"
        return f"{val:,.2f}"


def get_excel_number_format(
    role: str = "numeric",
    currency_symbol: Optional[str] = "₹"
) -> str:
    """Get the standard OpenPyXL number_format code for a given role.

    Args:
        role: Column semantic role.
        currency_symbol: Currency symbol.

    Returns:
        Excel format string (e.g. '0.00%', '"₹"#,##0.00', '#,##0', '#,##0.00').
    """
    role_lower = str(role).lower()
    symbol = currency_symbol or "₹"

    if role_lower in ["percentage", "rate", "ratio", "efficiency"]:
        return "0.00%"
    elif role_lower in ["currency", "revenue", "profit", "sales", "cost"]:
        return f'"{symbol}"#,##0.00'
    elif role_lower in ["rating", "score", "csat", "nps"]:
        return "0.00"
    elif role_lower in ["count", "records", "orders", "quantity", "units", "calls"]:
        return "#,##0"
    else:
        return "#,##0.00"
