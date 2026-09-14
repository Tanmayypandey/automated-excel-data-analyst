"""Excel Exporter Module.

Constructs the complete 5-sheet professional .xlsx workbook:
1. Dashboard (KPIs, native charts, dataset summary)
2. Cleaned Data (standardized, typed, imputed tabular records)
3. Cleaning Report (transformation audit trail)
4. Original Data (raw untouched dataset)
5. Chart Data (source aggregations for native Excel charts)
"""

import os
from typing import Dict, Any, List, Tuple
import pandas as pd
import numpy as np
import openpyxl
from openpyxl.worksheet.worksheet import Worksheet
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

from src.cleaning_report_generator import populate_cleaning_report
from src.dashboard_generator import build_dashboard


def export_excel_report(
    raw_df: pd.DataFrame,
    cleaned_df: pd.DataFrame,
    cleaning_stats: Dict[str, Any],
    kpis: List[Dict[str, Any]],
    chart_specs: List[Dict[str, Any]],
    output_file_path: str
) -> str:
    """Generate and save the master 5-sheet Excel workbook.

    Args:
        raw_df: Raw original dataset before cleaning.
        cleaned_df: Cleaned and standardized dataset.
        cleaning_stats: Dictionary containing cleaning metrics and audit log.
        kpis: Selected KPIs list.
        chart_specs: Selected charts specification list.
        output_file_path: Destination path for final .xlsx file.

    Returns:
        Absolute path to the saved Excel workbook.
    """
    os.makedirs(os.path.dirname(os.path.abspath(output_file_path)), exist_ok=True)

    wb = openpyxl.Workbook()
    # Default sheet
    ws_default = wb.active

    # Create sheets in desired order
    ws_dashboard = ws_default
    ws_dashboard.title = "Dashboard"

    ws_cleaned = wb.create_sheet(title="Cleaned Data")
    ws_report = wb.create_sheet(title="Cleaning Report")
    ws_original = wb.create_sheet(title="Original Data")
    ws_chart_data = wb.create_sheet(title="Chart Data")

    # 1. Populate Original Data Table
    _write_dataframe_table(ws_original, raw_df, "Original Data (Raw Upload)")

    # 2. Populate Cleaned Data Table
    _write_dataframe_table(ws_cleaned, cleaned_df, "Cleaned & Standardized Data")

    # 3. Populate Cleaning Report
    populate_cleaning_report(ws_report, cleaning_stats)

    # 4. Prepare and Populate Chart Data
    chart_ranges = _populate_chart_data(ws_chart_data, cleaned_df, chart_specs)

    # 5. Build Dashboard Sheet
    dataset_summary = {
        "total_rows": len(cleaned_df),
        "total_columns": len(cleaned_df.columns),
        "duplicate_rows": cleaning_stats.get("duplicates_removed", 0),
        "missing_values_cleaned": (
            cleaning_stats.get("missing_values_before", 0) - cleaning_stats.get("missing_values_after", 0)
        )
    }

    build_dashboard(
        ws_dash=ws_dashboard,
        ws_chart_data=ws_chart_data,
        kpis=kpis,
        chart_specs=chart_specs,
        chart_data_ranges=chart_ranges,
        dataset_summary=dataset_summary
    )

    # Save workbook
    wb.save(output_file_path)
    wb.close()

    return os.path.abspath(output_file_path)


def _write_dataframe_table(ws: Worksheet, df: pd.DataFrame, title: str) -> None:
    """Format and write pandas DataFrame into worksheet with clean headers."""
    ws.views.sheetView[0].showGridLines = True

    header_fill = PatternFill(start_color="1E293B", end_color="1E293B", fill_type="solid")
    header_font = Font(name="Segoe UI", size=10, bold=True, color="FFFFFF")
    cell_font = Font(name="Segoe UI", size=9, color="334155")
    alt_fill = PatternFill(start_color="F8FAFC", end_color="F8FAFC", fill_type="solid")

    thin_border = Border(
        left=Side(style="thin", color="E2E8F0"),
        right=Side(style="thin", color="E2E8F0"),
        top=Side(style="thin", color="E2E8F0"),
        bottom=Side(style="thin", color="E2E8F0")
    )

    # Write column headers
    cols = list(df.columns)
    for c_idx, col_name in enumerate(cols, start=1):
        cell = ws.cell(row=1, column=c_idx, value=str(col_name))
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = thin_border

    # Write rows (limit to 100,000 for excel performance)
    max_export_rows = min(len(df), 100000)
    for r_idx, row in enumerate(df.iloc[:max_export_rows].itertuples(index=False), start=2):
        use_alt = (r_idx % 2 == 1)
        for c_idx, val in enumerate(row, start=1):
            cell = ws.cell(row=r_idx, column=c_idx)
            # Format datetime if applicable
            if isinstance(val, (pd.Timestamp, np.datetime64)):
                cell.value = str(val)[:19]
            elif pd.isna(val):
                cell.value = ""
            else:
                cell.value = val

            cell.font = cell_font
            cell.border = thin_border
            if use_alt:
                cell.fill = alt_fill

    # Adjust column widths
    for c_idx, col_name in enumerate(cols, start=1):
        max_len = max(len(str(col_name)), 10)
        c_letter = get_column_letter(c_idx)
        ws.column_dimensions[c_letter].width = min(max_len + 5, 40)


def _populate_chart_data(
    ws: Worksheet,
    df: pd.DataFrame,
    chart_specs: List[Dict[str, Any]]
) -> List[Dict[str, Any]]:
    """Generate aggregated tables in Chart Data worksheet and return cell ranges."""
    ws.views.sheetView[0].showGridLines = True
    header_fill = PatternFill(start_color="3B82F6", end_color="3B82F6", fill_type="solid")
    header_font = Font(name="Segoe UI", size=9, bold=True, color="FFFFFF")
    cell_font = Font(name="Segoe UI", size=9, color="334155")

    ranges: List[Dict[str, Any]] = []
    # Layout up to 4 tables side by side (Cols 1-2, 4-5, 7-8, 10-11)
    col_offsets = [(1, 2), (4, 5), (7, 8), (10, 11)]

    for idx, spec in enumerate(chart_specs[:4]):
        cat_c, val_c = col_offsets[idx]
        c_type = spec.get("chart_type", "bar")
        x_col = spec.get("x_col")
        y_col = spec.get("y_col")
        agg = spec.get("aggregation", "sum")

        table_df = pd.DataFrame()

        try:
            if c_type == "line" and x_col and y_col:
                # Group by date
                temp = df[[x_col, y_col]].dropna().copy()
                temp[x_col] = pd.to_datetime(temp[x_col], errors="coerce").dt.strftime("%Y-%m-%d")
                if agg == "mean":
                    table_df = temp.groupby(x_col, as_index=False)[y_col].mean().sort_values(by=x_col)
                else:
                    table_df = temp.groupby(x_col, as_index=False)[y_col].sum().sort_values(by=x_col)
                table_df = table_df.head(30)  # max 30 time intervals for readability

            elif c_type == "bar" and x_col and y_col:
                temp = df[[x_col, y_col]].dropna().copy()
                if "top10" in agg:
                    if "mean" in agg:
                        table_df = temp.groupby(x_col, as_index=False)[y_col].mean()
                    else:
                        table_df = temp.groupby(x_col, as_index=False)[y_col].sum()
                    table_df = table_df.sort_values(by=y_col, ascending=False).head(10)
                else:
                    if agg == "mean":
                        table_df = temp.groupby(x_col, as_index=False)[y_col].mean()
                    else:
                        table_df = temp.groupby(x_col, as_index=False)[y_col].sum()
                    table_df = table_df.sort_values(by=y_col, ascending=False).head(20)

            elif c_type == "pie" and x_col and y_col:
                temp = df[[x_col, y_col]].dropna().copy()
                table_df = temp.groupby(x_col, as_index=False)[y_col].sum().sort_values(by=y_col, ascending=False).head(6)

            elif c_type == "scatter" and x_col and y_col:
                temp = df[[x_col, y_col]].dropna().copy()
                # Sample up to 100 rows for scatter readability
                if len(temp) > 100:
                    temp = temp.sample(n=100, random_state=42)
                table_df = temp.sort_values(by=x_col)

            elif c_type == "histogram" and x_col:
                series = df[x_col].dropna()
                counts, bin_edges = np.histogram(series, bins=8)
                bin_labels = [f"{bin_edges[i]:.1f}-{bin_edges[i+1]:.1f}" for i in range(len(counts))]
                table_df = pd.DataFrame({"Interval": bin_labels, "Frequency": counts})

        except Exception:
            # Fallback table if aggregation failed
            table_df = pd.DataFrame({"Category": ["Sample"], "Value": [1]})

        if len(table_df) == 0:
            table_df = pd.DataFrame({"Category": ["No Data"], "Value": [0]})

        # Write table into worksheet
        headers = list(table_df.columns)
        c1_cell = ws.cell(row=1, column=cat_c, value=str(headers[0]))
        c1_cell.font = header_font
        c1_cell.fill = header_fill

        c2_cell = ws.cell(row=1, column=val_c, value=str(headers[1]))
        c2_cell.font = header_font
        c2_cell.fill = header_fill

        for r_offset, (k, v) in enumerate(zip(table_df.iloc[:, 0], table_df.iloc[:, 1]), start=2):
            cell_k = ws.cell(row=r_offset, column=cat_c, value=str(k))
            cell_k.font = cell_font

            cell_v = ws.cell(row=r_offset, column=val_c, value=float(v) if isinstance(v, (int, float, np.number)) else str(v))
            cell_v.font = cell_font

        end_r = 1 + len(table_df)
        ranges.append({
            "cat_col": cat_c,
            "val_col": val_c,
            "start_row": 1,
            "end_row": end_r
        })

        # Set column width
        ws.column_dimensions[get_column_letter(cat_c)].width = 16
        ws.column_dimensions[get_column_letter(val_c)].width = 14

    return ranges
