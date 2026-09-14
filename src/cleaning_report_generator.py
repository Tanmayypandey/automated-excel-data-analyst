"""Cleaning Report Generator Module.

Formats the data transformation audit log and populates the 'Cleaning Report'
worksheet in the Excel workbook with styled summary tables and action logs.
"""

from typing import Dict, Any
from openpyxl.worksheet.worksheet import Worksheet
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter


def populate_cleaning_report(ws: Worksheet, cleaning_stats: Dict[str, Any]) -> None:
    """Populate and style the 'Cleaning Report' Excel worksheet.

    Args:
        ws: OpenPyXL Worksheet instance.
        cleaning_stats: Dictionary containing cleaning metrics and actions.
    """
    ws.title = "Cleaning Report"
    ws.views.sheetView[0].showGridLines = True

    # Styling definitions
    title_font = Font(name="Segoe UI", size=16, bold=True, color="1E293B")
    section_font = Font(name="Segoe UI", size=12, bold=True, color="0F172A")
    header_font = Font(name="Segoe UI", size=10, bold=True, color="FFFFFF")
    cell_font = Font(name="Segoe UI", size=10, color="334155")
    bold_cell_font = Font(name="Segoe UI", size=10, bold=True, color="1E293B")

    header_fill = PatternFill(start_color="3B82F6", end_color="3B82F6", fill_type="solid")
    alt_row_fill = PatternFill(start_color="F8FAFC", end_color="F8FAFC", fill_type="solid")
    accent_fill = PatternFill(start_color="EFF6FF", end_color="EFF6FF", fill_type="solid")

    thin_border_side = Side(style="thin", color="E2E8F0")
    cell_border = Border(
        left=thin_border_side, right=thin_border_side,
        top=thin_border_side, bottom=thin_border_side
    )

    # Title Banner
    ws["B2"] = "Data Cleaning & Transformation Audit Report"
    ws["B2"].font = title_font
    ws["B3"] = "Automated data hygiene, validation, type coercion, and imputation log"
    ws["B3"].font = Font(name="Segoe UI", size=10, italic=True, color="64748B")

    # Section 1: Summary Metrics Table
    ws["B5"] = "1. Summary Statistics"
    ws["B5"].font = section_font

    headers_1 = ["Metric", "Value Before", "Value After", "Net Change"]
    for col_idx, h_text in enumerate(headers_1, start=2):
        cell = ws.cell(row=6, column=col_idx, value=h_text)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = cell_border

    orig_rows = cleaning_stats.get("original_rows", 0)
    clean_rows = cleaning_stats.get("cleaned_rows", 0)
    orig_cols = cleaning_stats.get("original_columns", 0)
    clean_cols = cleaning_stats.get("cleaned_columns", 0)
    dupes = cleaning_stats.get("duplicates_removed", 0)
    nulls_before = cleaning_stats.get("missing_values_before", 0)
    nulls_after = cleaning_stats.get("missing_values_after", 0)

    rows_data = [
        ("Total Rows", orig_rows, clean_rows, clean_rows - orig_rows),
        ("Total Columns", orig_cols, clean_cols, clean_cols - orig_cols),
        ("Duplicate Rows", dupes, 0, -dupes),
        ("Missing / Null Values", nulls_before, nulls_after, nulls_after - nulls_before)
    ]

    for r_offset, (metric, v_before, v_after, v_diff) in enumerate(rows_data, start=7):
        fill_to_use = alt_row_fill if r_offset % 2 == 1 else PatternFill(fill_type=None)
        
        c1 = ws.cell(row=r_offset, column=2, value=metric)
        c1.font = bold_cell_font
        c1.fill = fill_to_use
        c1.border = cell_border

        c2 = ws.cell(row=r_offset, column=3, value=v_before)
        c2.font = cell_font
        c2.alignment = Alignment(horizontal="right")
        c2.fill = fill_to_use
        c2.border = cell_border

        c3 = ws.cell(row=r_offset, column=4, value=v_after)
        c3.font = cell_font
        c3.alignment = Alignment(horizontal="right")
        c3.fill = fill_to_use
        c3.border = cell_border

        c4 = ws.cell(row=r_offset, column=5, value=v_diff)
        c4.font = cell_font
        c4.alignment = Alignment(horizontal="right")
        c4.fill = fill_to_use
        c4.border = cell_border

    # Section 2: Conversions & Imputations
    curr_row = 12
    ws.cell(row=curr_row, column=2, value="2. Type Conversions & Feature Engineering").font = section_font
    curr_row += 1

    headers_2 = ["Category", "Columns / Details"]
    for col_idx, h_text in enumerate(headers_2, start=2):
        cell = ws.cell(row=curr_row, column=col_idx, value=h_text)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = cell_border
    curr_row += 1

    num_conv = cleaning_stats.get("columns_converted_to_numeric", [])
    date_conv = cleaning_stats.get("columns_converted_to_datetime", [])
    pct_conv = cleaning_stats.get("percentage_conversions", [])
    curr_conv = cleaning_stats.get("currency_conversions", [])
    id_cols = cleaning_stats.get("preserved_id_columns", [])

    details_list = [
        ("Numeric Conversions", ", ".join(num_conv) if num_conv else "None"),
        ("Datetime Conversions", ", ".join(date_conv) if date_conv else "None"),
        ("Percentage Conversions", ", ".join(pct_conv) if pct_conv else "None"),
        ("Currency Conversions", ", ".join(curr_conv) if curr_conv else "None"),
        ("Preserved ID Columns", ", ".join(id_cols) if id_cols else "None")
    ]

    for cat_name, details_text in details_list:
        c1 = ws.cell(row=curr_row, column=2, value=cat_name)
        c1.font = bold_cell_font
        c1.border = cell_border

        c2 = ws.cell(row=curr_row, column=3, value=details_text)
        c2.font = cell_font
        c2.border = cell_border
        curr_row += 1

    # Section 3: Missing Value Imputation Details
    curr_row += 1
    ws.cell(row=curr_row, column=2, value="3. Missing Value Imputation Strategy").font = section_font
    curr_row += 1

    headers_3 = ["Column", "Imputation Method Applied"]
    for col_idx, h_text in enumerate(headers_3, start=2):
        cell = ws.cell(row=curr_row, column=col_idx, value=h_text)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = cell_border
    curr_row += 1

    imputed = cleaning_stats.get("imputed_columns", {})
    if imputed:
        for col_name, method in imputed.items():
            c1 = ws.cell(row=curr_row, column=2, value=col_name)
            c1.font = bold_cell_font
            c1.border = cell_border

            c2 = ws.cell(row=curr_row, column=3, value=method)
            c2.font = cell_font
            c2.border = cell_border
            curr_row += 1
    else:
        c1 = ws.cell(row=curr_row, column=2, value="All Columns")
        c1.font = bold_cell_font
        c1.border = cell_border
        c2 = ws.cell(row=curr_row, column=3, value="No missing values required imputation")
        c2.font = cell_font
        c2.border = cell_border
        curr_row += 1

    # Section 4: Log of Major Actions
    curr_row += 1
    ws.cell(row=curr_row, column=2, value="4. Major Cleaning Actions Log").font = section_font
    curr_row += 1

    actions = cleaning_stats.get("major_actions", [])
    if actions:
        for act in actions:
            c = ws.cell(row=curr_row, column=2, value=f"• {act}")
            c.font = cell_font
            curr_row += 1
    else:
        c = ws.cell(row=curr_row, column=2, value="• Dataset was already pristine. No destructive modifications applied.")
        c.font = cell_font
        curr_row += 1

    # Adjust column widths
    ws.column_dimensions["A"].width = 4
    ws.column_dimensions["B"].width = 30
    ws.column_dimensions["C"].width = 35
    ws.column_dimensions["D"].width = 20
    ws.column_dimensions["E"].width = 20
