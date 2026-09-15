"""Pipeline Orchestrator Module.

Coordinates end-to-end execution:
1. File validation
2. Sheet selection / data loading
3. Dataset validation
4. Semantic column detection
5. Data cleaning
6. Role re-detection
7. EDA
8. KPI selection
9. Chart selection
10. Excel generation
11. Output validation

Includes detailed timing logs for performance debugging.
"""

import os
import time
from typing import Dict, Any

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


def _elapsed(start_time: float) -> float:
    """Return elapsed seconds."""
    return time.perf_counter() - start_time


def _print_step(step: str, seconds: float) -> None:
    """Print formatted timing information."""
    print(
        f"{step:<45} {seconds:>8.2f} sec",
        flush=True
    )


def run_pipeline(
    input_file_path: str,
    output_file_path: str
) -> Dict[str, Any]:

    """Execute complete Automated Excel Data Analyst pipeline."""

    pipeline_start = time.perf_counter()

    print("\n" + "=" * 70, flush=True)
    print("AUTOMATED EXCEL DATA ANALYST - PIPELINE STARTED", flush=True)
    print("=" * 70, flush=True)

    input_file_base = os.path.basename(
        input_file_path
    )

    _, ext = os.path.splitext(
        input_file_path
    )

    input_type = ext.lower()

    print(
        f"Input file : {input_file_base}",
        flush=True
    )

    try:
        file_size_mb = (
            os.path.getsize(input_file_path)
            / (1024 * 1024)
        )

        print(
            f"File size  : {file_size_mb:.2f} MB",
            flush=True
        )

    except Exception:
        pass

    print(
        f"File type  : {input_type}",
        flush=True
    )

    print("-" * 70, flush=True)


    # =====================================================
    # STEP 1 - FILE VALIDATION
    # =====================================================

    step_start = time.perf_counter()

    print(
        "[1] Starting file validation...",
        flush=True
    )

    try:
        file_val = validate_file(
            input_file_path
        )

    except Exception as e:

        _print_step(
            "[1] File validation FAILED",
            _elapsed(step_start)
        )

        return {
            "status": "error",
            "input_file": input_file_base,
            "input_type": input_type,
            "selected_sheet": None,
            "error": f"File validation failed: {str(e)}",
            "file_validation": {},
            "validation_report": {},
            "column_roles": {},
            "selected_kpis": [],
            "selected_charts": [],
            "cleaning_report": {},
            "dashboard_result": {},
            "output_validation": {
                "valid": False
            }
        }


    _print_step(
        "[1] File validation completed",
        _elapsed(step_start)
    )


    if not file_val.get("valid"):

        return {
            "status": "error",
            "input_file": input_file_base,
            "input_type": input_type,
            "selected_sheet": None,
            "error": file_val.get(
                "error",
                "File validation failed."
            ),
            "file_validation": file_val,
            "validation_report": {},
            "column_roles": {},
            "selected_kpis": [],
            "selected_charts": [],
            "cleaning_report": {},
            "dashboard_result": {},
            "output_validation": {
                "valid": False
            }
        }


    # =====================================================
    # STEP 2 - SHEET SELECTION
    # =====================================================

    selected_sheet = "CSV Data"

    if input_type == ".xlsx":

        step_start = time.perf_counter()

        print(
            "[2] Starting Excel sheet selection...",
            flush=True
        )

        try:

            sheet_meta = select_best_sheet(
                input_file_path
            )

            selected_sheet = (
                sheet_meta["selected_sheet"]
            )

        except Exception as e:

            _print_step(
                "[2] Sheet selection FAILED",
                _elapsed(step_start)
            )

            return {
                "status": "error",
                "input_file": input_file_base,
                "input_type": input_type,
                "selected_sheet": None,
                "error": (
                    f"Failed to select Excel sheet: {str(e)}"
                ),
                "file_validation": file_val,
                "validation_report": {},
                "column_roles": {},
                "selected_kpis": [],
                "selected_charts": [],
                "cleaning_report": {},
                "dashboard_result": {},
                "output_validation": {
                    "valid": False
                }
            }


        _print_step(
            "[2] Excel sheet selection completed",
            _elapsed(step_start)
        )

        print(
            f"    Selected sheet: {selected_sheet}",
            flush=True
        )

    else:

        print(
            "[2] CSV detected - sheet selection skipped",
            flush=True
        )


    # =====================================================
    # STEP 3 - DATA LOADING
    # =====================================================

    step_start = time.perf_counter()

    print(
        "[3] Starting dataset loading...",
        flush=True
    )

    try:

        if input_type == ".xlsx":

            df_raw, _ = load_data(
                input_file_path,
                sheet_name=selected_sheet
            )

        else:

            df_raw, selected_sheet = load_data(
                input_file_path
            )

    except Exception as e:

        _print_step(
            "[3] Dataset loading FAILED",
            _elapsed(step_start)
        )

        return {
            "status": "error",
            "input_file": input_file_base,
            "input_type": input_type,
            "selected_sheet": selected_sheet,
            "error": (
                f"Failed to load dataset: {str(e)}"
            ),
            "file_validation": file_val,
            "validation_report": {},
            "column_roles": {},
            "selected_kpis": [],
            "selected_charts": [],
            "cleaning_report": {},
            "dashboard_result": {},
            "output_validation": {
                "valid": False
            }
        }


    _print_step(
        "[3] Dataset loading completed",
        _elapsed(step_start)
    )

    print(
        f"    Dataset shape: "
        f"{len(df_raw):,} rows x "
        f"{len(df_raw.columns):,} columns",
        flush=True
    )

    try:

        memory_mb = (
            df_raw.memory_usage(
                deep=True
            ).sum()
            / (1024 * 1024)
        )

        print(
            f"    DataFrame memory: "
            f"{memory_mb:.2f} MB",
            flush=True
        )

    except Exception:
        pass


    # =====================================================
    # STEP 4 - DATASET VALIDATION
    # =====================================================

    step_start = time.perf_counter()

    print(
        "[4] Starting dataset validation...",
        flush=True
    )

    try:

        validation_report = validate_dataset(
            df_raw
        )

    except Exception as e:

        validation_report = {
            "error": str(e)
        }

        print(
            f"    WARNING: {e}",
            flush=True
        )


    _print_step(
        "[4] Dataset validation completed",
        _elapsed(step_start)
    )


    # =====================================================
    # STEP 5 - COLUMN ROLE DETECTION
    # =====================================================

    step_start = time.perf_counter()

    print(
        "[5] Starting semantic column detection...",
        flush=True
    )

    try:

        column_roles = detect_column_roles(
            df_raw
        )

    except Exception as e:

        column_roles = {}

        print(
            f"    WARNING: Column detection failed: {e}",
            flush=True
        )


    _print_step(
        "[5] Column detection completed",
        _elapsed(step_start)
    )


    # =====================================================
    # STEP 6 - DATA CLEANING
    # =====================================================

    step_start = time.perf_counter()

    print(
        "[6] Starting data cleaning...",
        flush=True
    )

    try:

        df_cleaned, cleaning_stats = (
            clean_dataset(
                df_raw,
                detected_roles=column_roles
            )
        )

    except Exception as e:

        _print_step(
            "[6] Data cleaning FAILED",
            _elapsed(step_start)
        )

        return {
            "status": "error",
            "input_file": input_file_base,
            "input_type": input_type,
            "selected_sheet": selected_sheet,
            "error": (
                f"Data cleaning failed: {str(e)}"
            ),
            "file_validation": file_val,
            "validation_report": validation_report,
            "column_roles": column_roles,
            "selected_kpis": [],
            "selected_charts": [],
            "cleaning_report": {},
            "dashboard_result": {},
            "output_validation": {
                "valid": False
            }
        }


    _print_step(
        "[6] Data cleaning completed",
        _elapsed(step_start)
    )

    print(
        f"    Cleaned shape: "
        f"{len(df_cleaned):,} rows x "
        f"{len(df_cleaned.columns):,} columns",
        flush=True
    )


    # =====================================================
    # STEP 7 - ROLE RE-DETECTION
    # =====================================================

    step_start = time.perf_counter()

    print(
        "[7] Re-detecting roles after cleaning...",
        flush=True
    )

    try:

        column_roles = detect_column_roles(
            df_cleaned
        )

    except Exception as e:

        print(
            f"    WARNING: Role re-detection failed: {e}",
            flush=True
        )


    _print_step(
        "[7] Role re-detection completed",
        _elapsed(step_start)
    )


    # =====================================================
    # STEP 8 - EDA
    # =====================================================

    step_start = time.perf_counter()

    print(
        "[8] Starting EDA...",
        flush=True
    )

    try:

        eda_summary = run_eda(
            df_cleaned,
            detected_roles=column_roles
        )

    except Exception as e:

        eda_summary = {
            "error": str(e)
        }

        print(
            f"    WARNING: EDA failed: {e}",
            flush=True
        )


    _print_step(
        "[8] EDA completed",
        _elapsed(step_start)
    )


    # =====================================================
    # STEP 9 - KPI SELECTION
    # =====================================================

    step_start = time.perf_counter()

    print(
        "[9] Starting KPI selection...",
        flush=True
    )

    try:

        selected_kpis = select_kpis(
            df_cleaned,
            detected_roles=column_roles
        )

    except Exception as e:

        print(
            f"    WARNING: KPI selection failed: {e}",
            flush=True
        )

        selected_kpis = [{
            "label": "Total Records",
            "value": len(df_cleaned),
            "formatted_value": (
                f"{len(df_cleaned):,}"
            ),
            "role": "count",
            "aggregation": "count"
        }]


    _print_step(
        "[9] KPI selection completed",
        _elapsed(step_start)
    )

    print(
        f"    KPIs selected: "
        f"{len(selected_kpis)}",
        flush=True
    )


    # =====================================================
    # STEP 10 - CHART SELECTION
    # =====================================================

    step_start = time.perf_counter()

    print(
        "[10] Starting chart selection...",
        flush=True
    )

    try:

        selected_charts = select_charts(
            df_cleaned,
            detected_roles=column_roles
        )

    except Exception as e:

        selected_charts = []

        print(
            f"    WARNING: Chart selection failed: {e}",
            flush=True
        )


    _print_step(
        "[10] Chart selection completed",
        _elapsed(step_start)
    )

    print(
        f"    Charts selected: "
        f"{len(selected_charts)}",
        flush=True
    )


    # =====================================================
    # STEP 11 - EXCEL GENERATION
    # =====================================================

    step_start = time.perf_counter()

    print(
        "[11] Starting Excel workbook generation...",
        flush=True
    )

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
            "charts_count":
                len(selected_charts),

            "kpis_count":
                len(selected_kpis),

            "output_path":
                os.path.basename(
                    output_file_path
                )
        }


    except Exception as e:

        _print_step(
            "[11] Excel generation FAILED",
            _elapsed(step_start)
        )

        print(
            f"    ERROR: {e}",
            flush=True
        )

        return {
            "status": "error",
            "input_file": input_file_base,
            "input_type": input_type,
            "selected_sheet": selected_sheet,
            "error": (
                f"Excel generation failed: {str(e)}"
            ),
            "file_validation": file_val,
            "validation_report": validation_report,
            "column_roles": column_roles,
            "selected_kpis": selected_kpis,
            "selected_charts": selected_charts,
            "cleaning_report": cleaning_stats,
            "dashboard_result": {},
            "output_validation": {
                "valid": False
            }
        }


    _print_step(
        "[11] Excel generation completed",
        _elapsed(step_start)
    )


    # =====================================================
    # STEP 12 - OUTPUT VALIDATION
    # =====================================================

    step_start = time.perf_counter()

    print(
        "[12] Starting output validation...",
        flush=True
    )

    try:

        output_val = validate_output(
            output_file_path
        )

    except Exception as e:

        output_val = {
            "valid": False,
            "error": str(e)
        }


    _print_step(
        "[12] Output validation completed",
        _elapsed(step_start)
    )


    # =====================================================
    # TOTAL TIME
    # =====================================================

    total_time = _elapsed(
        pipeline_start
    )


    print("-" * 70, flush=True)

    print(
        f"TOTAL PIPELINE TIME: "
        f"{total_time:.2f} seconds "
        f"({total_time / 60:.2f} minutes)",
        flush=True
    )


    if os.path.exists(
        output_file_path
    ):

        output_size_mb = (
            os.path.getsize(
                output_file_path
            )
            / (1024 * 1024)
        )

        print(
            f"Output workbook size: "
            f"{output_size_mb:.2f} MB",
            flush=True
        )


    print(
        "=" * 70,
        flush=True
    )

    print(
        "PIPELINE FINISHED",
        flush=True
    )

    print(
        "=" * 70 + "\n",
        flush=True
    )


    # =====================================================
    # RESPONSE
    # =====================================================

    return {
        "status": (
            "success"
            if output_val.get("valid")
            else "error"
        ),

        "input_file":
            input_file_base,

        "input_type":
            input_type,

        "selected_sheet":
            selected_sheet,

        "error":
            output_val.get("error"),

        "file_validation":
            file_val,

        "validation_report":
            validation_report,

        "column_roles":
            column_roles,

        "eda_summary":
            eda_summary,

        "selected_kpis":
            selected_kpis,

        "selected_charts":
            selected_charts,

        "cleaning_report":
            cleaning_stats,

        "dashboard_result":
            dashboard_result,

        "output_validation":
            output_val,

        "processing_time_seconds":
            round(total_time, 2)
    }