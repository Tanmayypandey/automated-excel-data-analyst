"""Sheet Selector Module.

Intelligently scores and selects the best tabular data worksheet from multi-sheet
Excel workbooks, ignoring empty, summary, notes, or dashboard-only sheets.
"""

import os
import re
from typing import Dict, Any, List
import openpyxl


PENALIZED_SHEET_NAMES = [
    r"dashboard", r"summary", r"report", r"chart", r"pivot",
    r"notes?", r"cover", r"intro", r"readme", r"config",
    r"metadata", r"overview", r"kpis?"
]

FAVORED_SHEET_NAMES = [
    r"raw", r"data", r"records?", r"transactions?", r"details?",
    r"sales", r"orders?", r"master", r"table", r"main"
]


def select_best_sheet(file_path: str) -> Dict[str, Any]:
    """Inspect all sheets in an Excel workbook and select the best tabular data sheet.

    Args:
        file_path: Path to the .xlsx file.

    Returns:
        Dictionary containing the chosen sheet name, individual sheet scores,
        and evaluation details.
    """
    if not os.path.exists(file_path):
        raise FileNotFoundError(f"File not found: {file_path}")

    wb = openpyxl.load_workbook(file_path, read_only=True, data_only=True)
    sheet_names = wb.sheetnames

    if not sheet_names:
        wb.close()
        raise ValueError("The Excel workbook contains no sheets.")

    if len(sheet_names) == 1:
        wb.close()
        return {
            "selected_sheet": sheet_names[0],
            "sheet_scores": {sheet_names[0]: 100.0},
            "reason": "Single worksheet available."
        }

    scores: Dict[str, float] = {}
    details: Dict[str, Dict[str, Any]] = {}

    for name in sheet_names:
        ws = wb[name]
        score, meta = _score_sheet(ws, name)
        scores[name] = score
        details[name] = meta

    wb.close()

    # Pick the sheet with the highest score
    best_sheet = max(scores, key=scores.get)

    return {
        "selected_sheet": best_sheet,
        "sheet_scores": scores,
        "details": details,
        "reason": f"Sheet '{best_sheet}' scored highest ({scores[best_sheet]:.1f}) based on tabular structure, row/column count, and content density."
    }


def _score_sheet(ws, sheet_name: str) -> (float, Dict[str, Any]):
    """Calculate a tabular quality score for an individual worksheet."""
    # Sample up to first 100 rows
    sample_rows: List[List[Any]] = []
    max_r = 0
    max_c = 0

    try:
        for r_idx, row in enumerate(ws.iter_rows(values_only=True), start=1):
            sample_rows.append(list(row))
            if len(row) > max_c:
                max_c = len(row)
            if r_idx >= 150:
                break
        max_r = ws.max_row or len(sample_rows)
    except Exception:
        pass

    if not sample_rows:
        return -1000.0, {"rows": 0, "columns": 0, "non_empty_ratio": 0.0}

    # Count actual columns with data
    num_sample_rows = len(sample_rows)
    total_cells = 0
    non_empty_cells = 0
    col_has_data = set()

    for r in sample_rows:
        for c_idx, val in enumerate(r):
            total_cells += 1
            if val is not None and str(val).strip() != "":
                non_empty_cells += 1
                col_has_data.add(c_idx)

    num_active_cols = len(col_has_data)
    non_empty_ratio = (non_empty_cells / total_cells) if total_cells > 0 else 0.0

    # Start scoring
    score = 0.0

    # 1. Row count scoring
    if max_r <= 2:
        score -= 150.0  # Barely any data
    elif max_r < 10:
        score += max_r * 2.0
    elif max_r < 100:
        score += 30.0 + (max_r * 0.5)
    else:
        score += 80.0 + min(max_r * 0.05, 50.0)

    # 2. Column count scoring
    if num_active_cols <= 1:
        score -= 100.0
    elif num_active_cols == 2:
        score -= 20.0
    elif 3 <= num_active_cols <= 30:
        score += 40.0 + (num_active_cols * 1.5)
    else:
        score += 50.0

    # 3. Density scoring (tabular data should be fairly dense)
    if non_empty_ratio < 0.15:
        score -= 60.0
    elif non_empty_ratio > 0.5:
        score += non_empty_ratio * 40.0

    # 4. Header detection (first non-empty row has multiple strings)
    header_found = False
    for r in sample_rows[:5]:
        non_empty = [v for v in r if v is not None and str(v).strip() != ""]
        if len(non_empty) >= 3 and all(isinstance(v, str) for v in non_empty[:3]):
            header_found = True
            score += 25.0
            break

    # 5. Penalize summary / dashboard / report names
    clean_name = sheet_name.strip().lower()
    for pattern in PENALIZED_SHEET_NAMES:
        if re.search(pattern, clean_name):
            score -= 75.0
            break

    # 6. Favor explicit raw/data/records names
    for pattern in FAVORED_SHEET_NAMES:
        if re.search(pattern, clean_name):
            score += 35.0
            break

    meta = {
        "rows": max_r,
        "columns": num_active_cols,
        "non_empty_ratio": round(non_empty_ratio, 3),
        "header_detected": header_found,
        "calculated_score": round(score, 1)
    }

    return score, meta
