"""Pipeline Orchestrator Module.

Coordinates end-to-end execution:
1. File validation
2. Sheet selection (XLSX) or Direct loading (CSV)
3. Dataset profiling and structural validation
4. Semantic column role detection
5. Data cleaning, typing, and imputation
6. Exploratory data analysis (EDA)
7. Smart KPI selection (max 4)
8. Smart Chart selection (max 4)
9. Master Excel workbook generation (Dashboard, Data, Report, Charts)
10. Output verification
"""

import os
from typing import Dict, Any
import pandas as pd

from src.file_validator import validate_file
from src.sheet_selector import select_best_sheet
from src.data_loader import load_data
from src.data_validator import validate_dataset
from src.smart_column_detector import detect_column_roles
from src.data_cleaner import clean_dataset
from src.eda_engine import run_eda
from src.smart_kpi_selector import select_kpis
from src.smart_chart_selector import select_charts
from src.excel_exporter import export_excel_report
from src.output_validator import validate_output


def run_pipeline(input_file_path: str, output_file_path: str) -> Dict[str, Any]:
    """Execute the complete Automated Excel Data Analyst workflow end-to-end.

    Args:
        input_file_path: Path to the uploaded raw file (.xlsx or .csv).
        output_file_path: Path where the analyzed .xlsx report will be saved.

    Returns:
        Structured dictionary containing all pipeline outputs, metrics, and validation states.
    """
    input_file_base = os.path.basename(input_file_path)
    _, ext = os.path.splitext(input_file_path)
    input_type = ext.lower()

    # Step 1: File Validation
    file_val = validate_file(input_file_path)
    if not file_val["valid"]:
        return {
            "status": "error",
            "input_file": input_file_base,
            "input_type": input_type,
            "selected_sheet": None,
            "error": file_val.get("error", "File validation failed."),
            "file_validation": file_val,
            "validation_report": {},
            "column_roles": {},
            "selected_kpis": [],
            "selected_charts": [],
            "cleaning_report": {},
            "dashboard_result": {},
            "output_validation": {"valid": False}
        }

    # Step 2: Sheet Selection & Loading
    selected_sheet = "CSV Data"
    try:
        if input_type == ".xlsx":
            sheet_meta = select_best_sheet(input_file_path)
            selected_sheet = sheet_meta["selected_sheet"]
            df_raw, _ = load_data(input_file_path, sheet_name=selected_sheet)
        else:
            df_raw, selected_sheet = load_data(input_file_path)
    except Exception as e:
        return {
            "status": "error",
            "input_file": input_file_base,
            "input_type": input_type,
            "selected_sheet": selected_sheet,
            "error": f"Failed to load dataset: {str(e)}",
            "file_validation": file_val,
            "validation_report": {},
            "column_roles": {},
            "selected_kpis": [],
            "selected_charts": [],
            "cleaning_report": {},
            "dashboard_result": {},
            "output_validation": {"valid": False}
        }

    # Step 3: Dataset Structural Validation
    try:
        validation_report = validate_dataset(df_raw)
    except Exception as e:
        validation_report = {"error": str(e)}

    # Step 4: Semantic Column Role Detection
    try:
        column_roles = detect_column_roles(df_raw)
    except Exception as e:
        column_roles = {}

    # Step 5: Data Cleaning & Imputation
    try:
        df_cleaned, cleaning_stats = clean_dataset(df_raw, detected_roles=column_roles)
    except Exception as e:
        return {
            "status": "error",
            "input_file": input_file_base,
            "input_type": input_type,
            "selected_sheet": selected_sheet,
            "error": f"Data cleaning failed: {str(e)}",
            "file_validation": file_val,
            "validation_report": validation_report,
            "column_roles": column_roles,
            "selected_kpis": [],
            "selected_charts": [],
            "cleaning_report": {},
            "dashboard_result": {},
            "output_validation": {"valid": False}
        }

    # Update column roles on cleaned data
    column_roles = detect_column_roles(df_cleaned)

    # Step 6: Automated EDA
    try:
        eda_summary = run_eda(df_cleaned, detected_roles=column_roles)
    except Exception as e:
        eda_summary = {"error": str(e)}

    # Step 7: Smart KPI Selection (max 4)
    try:
        selected_kpis = select_kpis(df_cleaned, detected_roles=column_roles)
    except Exception as e:
        selected_kpis = [{
            "label": "Total Records",
            "value": len(df_cleaned),
            "formatted_value": f"{len(df_cleaned):,}",
            "role": "count",
            "aggregation": "count"
        }]

    # Step 8: Smart Chart Selection (max 4)
    try:
        selected_charts = select_charts(df_cleaned, detected_roles=column_roles)
    except Exception as e:
        selected_charts = []

    # Step 9: Export Final Excel Workbook
    try:
        export_excel_report(
            raw_df=df_raw,
            cleaned_df=df_cleaned,
            cleaning_stats=cleaning_stats,
            kpis=selected_kpis,
            chart_specs=selected_charts,
            output_file_path=output_file_path
        )
        dashboard_result = {
            "charts_count": len(selected_charts),
            "kpis_count": len(selected_kpis),
            "output_path": os.path.basename(output_file_path)
        }
    except Exception as e:
        return {
            "status": "error",
            "input_file": input_file_base,
            "input_type": input_type,
            "selected_sheet": selected_sheet,
            "error": f"Excel generation failed: {str(e)}",
            "file_validation": file_val,
            "validation_report": validation_report,
            "column_roles": column_roles,
            "selected_kpis": selected_kpis,
            "selected_charts": selected_charts,
            "cleaning_report": cleaning_stats,
            "dashboard_result": {},
            "output_validation": {"valid": False}
        }

    # Step 10: Validate Final Output Workbook
    output_val = validate_output(output_file_path)

    return {
        "status": "success" if output_val.get("valid") else "error",
        "input_file": input_file_base,
        "input_type": input_type,
        "selected_sheet": selected_sheet,
        "error": output_val.get("error"),
        "file_validation": file_val,
        "validation_report": validation_report,
        "column_roles": column_roles,
        "eda_summary": eda_summary,
        "selected_kpis": selected_kpis,
        "selected_charts": selected_charts,
        "cleaning_report": cleaning_stats,
        "dashboard_result": dashboard_result,
        "output_validation": output_val
    }
