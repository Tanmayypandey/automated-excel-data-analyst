"""Dashboard Generator Module.

Builds a professional, clean Excel Dashboard worksheet using OpenPyXL:
- Top Title Banner
- 4 Styled KPI metric cards
- Up to 4 native Excel charts (Bar, Line, Pie, Scatter, Histogram)
- Dataset Summary section
"""

from typing import Dict, Any, List, Optional
import openpyxl
from openpyxl.worksheet.worksheet import Worksheet
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.chart import (
    BarChart, LineChart, PieChart, ScatterChart,
    Reference, Series
)
from src.kpi_formatter import get_excel_number_format


def build_dashboard(
    ws_dash: Worksheet,
    ws_chart_data: Worksheet,
    kpis: List[Dict[str, Any]],
    chart_specs: List[Dict[str, Any]],
    chart_data_ranges: List[Dict[str, Any]],
    dataset_summary: Dict[str, Any]
) -> None:
    """Construct professional Excel Dashboard sheet with KPI cards and charts.

    Args:
        ws_dash: Target 'Dashboard' worksheet.
        ws_chart_data: Source 'Chart Data' worksheet containing aggregated tables.
        kpis: List of selected KPIs (max 4).
        chart_specs: Selected chart specifications.
        chart_data_ranges: Cell coordinates where chart data tables are located.
        dataset_summary: Summary metrics (rows, columns, duplicates, missing).
    """
    ws_dash.title = "Dashboard"
    ws_dash.views.sheetView[0].showGridLines = False

    # Define color scheme (Clean Indigo/Slate Executive Theme)
    brand_blue = "2563EB"
    brand_dark = "0F172A"
    card_bg = "F8FAFC"
    border_color = "CBD5E1"
    subtext_color = "64748B"

    thin_border_side = Side(style="thin", color=border_color)
    card_border = Border(
        left=thin_border_side, right=thin_border_side,
        top=thin_border_side, bottom=thin_border_side
    )

    # 1. Dashboard Title Banner (Rows 1-2)
    ws_dash["B2"] = "AUTOMATED DATA ANALYSIS DASHBOARD"
    ws_dash["B2"].font = Font(name="Segoe UI", size=16, bold=True, color=brand_dark)
    ws_dash["B3"] = "Executive overview generated automatically from raw dataset"
    ws_dash["B3"].font = Font(name="Segoe UI", size=9, italic=True, color=subtext_color)

    # 2. 4 KPI Cards (Rows 5 to 7)
    # Card layout across columns:
    # Card 1: B5:D7
    # Card 2: F5:H7
    # Card 3: J5:L7
    # Card 4: N5:P7
    card_col_offsets = [
        (2, 4),    # B to D
        (6, 8),    # F to H
        (10, 12),  # J to L
        (14, 16)   # N to P
    ]

    card_fill = PatternFill(start_color=card_bg, end_color=card_bg, fill_type="solid")
    accent_fill = PatternFill(start_color="EFF6FF", end_color="EFF6FF", fill_type="solid")

    for idx, kpi in enumerate(kpis[:4]):
        start_c, end_c = card_col_offsets[idx]
        label = kpi.get("label", "Metric")
        value = kpi.get("value", 0)
        role = kpi.get("role", "numeric")
        sym = kpi.get("currency_symbol", "₹")

        # Merge cells for card label (Row 5)
        ws_dash.merge_cells(start_row=5, start_column=start_c, end_row=5, end_column=end_c)
        lbl_cell = ws_dash.cell(row=5, column=start_c, value=label.upper())
        lbl_cell.font = Font(name="Segoe UI", size=8, bold=True, color=subtext_color)
        lbl_cell.alignment = Alignment(horizontal="center", vertical="center")

        # Merge cells for card value (Row 6 to 7)
        ws_dash.merge_cells(start_row=6, start_column=start_c, end_row=7, end_column=end_c)
        val_cell = ws_dash.cell(row=6, column=start_c, value=value)
        val_cell.font = Font(name="Segoe UI", size=16, bold=True, color=brand_blue)
        val_cell.alignment = Alignment(horizontal="center", vertical="center")
        val_cell.number_format = get_excel_number_format(role, sym)

        # Apply borders and background to card region
        for r in range(5, 8):
            for c in range(start_c, end_c + 1):
                cell = ws_dash.cell(row=r, column=c)
                cell.fill = card_fill
                cell.border = card_border

    # 3. Add OpenPyXL Native Charts
    # Grid positions: Chart 1 at B10, Chart 2 at J10, Chart 3 at B26, Chart 4 at J26
    chart_positions = ["B10", "J10", "B26", "J26"]

    for idx, (spec, c_range) in enumerate(zip(chart_specs[:4], chart_data_ranges[:4])):
        pos = chart_positions[idx]
        chart_obj = _create_native_chart(spec, c_range, ws_chart_data)
        if chart_obj:
            ws_dash.add_chart(chart_obj, pos)

    # 4. Dataset Summary Section (Rows 43 to 45)
    summary_row = 43
    ws_dash.cell(row=summary_row, column=2, value="DATASET SUMMARY").font = Font(
        name="Segoe UI", size=10, bold=True, color=brand_dark
    )
    
    summary_metrics = [
        ("Total Rows", dataset_summary.get("total_rows", 0)),
        ("Total Columns", dataset_summary.get("total_columns", 0)),
        ("Duplicate Rows Removed", dataset_summary.get("duplicate_rows", 0)),
        ("Missing Values Cleaned", dataset_summary.get("missing_values_cleaned", 0))
    ]

    for offset, (m_lbl, m_val) in enumerate(summary_metrics):
        col_pos = 2 + (offset * 3)
        ws_dash.merge_cells(start_row=summary_row + 1, start_column=col_pos, end_row=summary_row + 1, end_column=col_pos + 1)
        lbl_c = ws_dash.cell(row=summary_row + 1, column=col_pos, value=f"{m_lbl}: {m_val:,}")
        lbl_c.font = Font(name="Segoe UI", size=9, bold=True, color=subtext_color)
        lbl_c.fill = accent_fill
        lbl_c.border = card_border

    # Adjust layout column widths
    ws_dash.column_dimensions["A"].width = 3
    for c_letter in ["B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P"]:
        ws_dash.column_dimensions[c_letter].width = 11


def _create_native_chart(
    spec: Dict[str, Any],
    c_range: Dict[str, Any],
    ws_source: Worksheet
) -> Optional[Any]:
    """Instantiate and configure OpenPyXL native chart object."""
    c_type = spec.get("chart_type", "bar")
    title = spec.get("title", "Analysis Chart")

    cat_col = c_range["cat_col"]
    val_col = c_range["val_col"]
    start_row = c_range["start_row"]
    end_row = c_range["end_row"]

    if end_row <= start_row:
        return None

    if c_type == "line":
        chart = LineChart()
        chart.title = title
        chart.style = 13
        chart.width = 14.5
        chart.height = 7.5
        chart.legend = None

        data = Reference(ws_source, min_col=val_col, min_row=start_row, max_row=end_row)
        cats = Reference(ws_source, min_col=cat_col, min_row=start_row + 1, max_row=end_row)
        chart.add_data(data, titles_from_data=True)
        chart.set_categories(cats)
        return chart

    elif c_type == "pie":
        chart = PieChart()
        chart.title = title
        chart.style = 10
        chart.width = 14.5
        chart.height = 7.5

        data = Reference(ws_source, min_col=val_col, min_row=start_row, max_row=end_row)
        cats = Reference(ws_source, min_col=cat_col, min_row=start_row + 1, max_row=end_row)
        chart.add_data(data, titles_from_data=True)
        chart.set_categories(cats)
        return chart

    elif c_type == "scatter":
        chart = ScatterChart()
        chart.title = title
        chart.style = 11
        chart.width = 14.5
        chart.height = 7.5
        chart.legend = None

        xvalues = Reference(ws_source, min_col=cat_col, min_row=start_row + 1, max_row=end_row)
        yvalues = Reference(ws_source, min_col=val_col, min_row=start_row, max_row=end_row)
        series = Series(yvalues, xvalues, title_from_data=True)
        series.marker.symbol = "circle"
        series.marker.size = 5
        series.graphicalProperties.line.noFill = True
        chart.series.append(series)
        return chart

    elif c_type in ["bar", "histogram"]:
        chart = BarChart()
        chart.type = "col"
        chart.style = 10
        chart.title = title
        chart.width = 14.5
        chart.height = 7.5
        chart.legend = None

        data = Reference(ws_source, min_col=val_col, min_row=start_row, max_row=end_row)
        cats = Reference(ws_source, min_col=cat_col, min_row=start_row + 1, max_row=end_row)
        chart.add_data(data, titles_from_data=True)
        chart.set_categories(cats)
        return chart

    return None
