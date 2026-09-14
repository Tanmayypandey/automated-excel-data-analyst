# Automated Excel Data Analyst

An autonomous, full-stack, production-grade data analytics application that transforms raw `.xlsx` and `.csv` files into interactive web insights and executive, multi-sheet downloadable Excel workbooks featuring native embedded charts.

---

## 1. Overview & Business Problem

Organizations and business analysts often receive raw spreadsheets that are unstandardized, missing key records, poorly formatted, or spread across multiple worksheets with ad-hoc summary tables. Converting these raw extracts into executive-ready reports traditionally requires hours of manual formula writing, data munging, pivot table generation, and chart formatting.

**Automated Excel Data Analyst** solves this by completely automating the data engineering and reporting lifecycle locally:
1. Validating file boundaries and worksheet integrity.
2. Intelligently scoring and selecting the best tabular data sheet from complex multi-sheet workbooks.
3. Cleansing data: standardizing headers, stripping punctuation, converting currencies and percentages, handling missing values, and deduplicating rows.
4. Preserving identifiers and protecting duration metrics (e.g. `active_time`, `hold_time`) from improper datetime coercion.
5. Performing statistical Exploratory Data Analysis (EDA).
6. Dynamically identifying and calculating high-impact business KPIs.
7. Designing up to 4 native OpenPyXL charts tailored to dataset distributions (line, bar, top-10 leaderboard, pie, scatter, histogram).
8. Building an executive 5-sheet workbook:
   - **Dashboard** (KPI cards, native charts, dataset summary)
   - **Cleaned Data** (typed and imputed tabular records)
   - **Cleaning Report** (before/after audit trail)
   - **Original Data** (untouched raw upload)
   - **Chart Data** (clean source tables powering native charts)
9. Presenting results on a responsive SaaS web dashboard and providing instant one-click report downloads.

---

## 2. Key Features

- **Dual File Format Support**: Seamless processing for Microsoft Excel (`.xlsx`) and Delimited Data (`.csv`).
- **Heuristic Sheet Selector**: Inspects multi-sheet Excel files and automatically picks the highest-quality tabular data sheet, skipping notes, cover pages, and empty summaries.
- **Robust Multi-Encoding CSV Ingestion**: Handles `utf-8`, `utf-8-sig`, `latin1`, and `cp1252` encoding variations.
- **Intelligent Semantic Role Inference**: Classifies columns into identifiers, datetimes, percentages, currencies, counts, ratings, and measures.
- **Duration Protection Guardrail**: Explicitly prevents numeric duration/latency columns (e.g., `active_time`, `response_time`, `call_time`) from being misclassified as datetimes.
- **Automated Data Cleaning & Imputation**:
  - Missing strings (`"NA"`, `"null"`, `"-"`, `"--"`) normalized to nulls.
  - Percentage strings (`"25%"`) parsed to ratios (`0.25`).
  - Currency symbols (`₹`, `$`, `€`, `£`) stripped and recorded for formatting.
  - Numeric missing values imputed with median; categoricals with mode; identifiers with `"Unknown"`.
  - Exact duplicate rows removed.
- **Executive OpenPyXL Dashboard**: Native Excel charts generated directly inside the workbook—no static image screenshots—allowing users to edit, refresh, and style charts directly within Microsoft Excel.
- **Audit Trail Report**: Detailed worksheet summarizing all transformation metrics, net row/column changes, and imputation steps.
- **Modern SaaS Web Interface**: Clean light theme with drag-and-drop file upload, animated progress stepper, KPI stat cards, and instant download buttons.
- **100% Local & Free**: Zero external API dependencies, zero cloud database requirements, and zero data leakage.

---

## 3. Technology Stack

- **Backend**: Python 3.10+, FastAPI, Uvicorn, Python-Multipart
- **Data Engineering**: Pandas, NumPy, SciPy
- **Excel Generation**: OpenPyXL, XlsxWriter
- **Frontend**: HTML5, CSS3 (Modern SaaS Light Theme), Vanilla JavaScript (ES6+)
- **Testing**: Unittest, Pytest, FastAPI TestClient

---

## 4. Architecture & Pipeline Flow

```text
User Upload (.xlsx / .csv)
       │
       ▼
[File Validator] ──── Validates file size, extension, row/col bounds
       │
       ▼
[Sheet Selector] ─── (XLSX only) Evaluates density, dimensions, and names
       │              to choose the optimal tabular worksheet
       ▼
[Data Loader] ────── Multi-encoding ingestion (utf-8, latin1, etc.)
       │
       ▼
[Data Validator] ─── Profiles nulls, dtypes, duplicates, and cardinality
       │
       ▼
[Column Detector] ── Inferred semantic roles (measures, currencies, IDs)
       │
       ▼
[Data Cleaner] ───── Normalizes types, imputes missing values, deduplicates
       │
       ▼
[EDA Engine] ─────── Statistical distributions, percentiles, correlations
       │
       ▼
[KPI Selector] ───── Dynamic selection of max 4 high-value executive metrics
       │
       ▼
[Chart Selector] ─── Tailors up to 4 diverse charts (line, bar, pie, scatter)
       │
       ▼
[Excel Exporter] ─── Assembles 5-sheet workbook with OpenPyXL charts
       │
       ▼
[Output Validator] ─ Verifies workbook structure, sheets, and charts
       │
       ▼
Web UI & Download ── Interactive dashboard display + Downloadable .xlsx
```

---

## 5. Project Structure

```text
automated_excel_data_analyst/
│
├── data/
│   ├── input/                     # Uploaded files (.xlsx, .csv)
│   └── output/                    # Generated analyzed workbooks (.xlsx)
│
├── src/
│   ├── __init__.py
│   ├── file_validator.py          # File type, size, readability checks
│   ├── sheet_selector.py          # Intelligent XLSX sheet scoring & selection
│   ├── data_loader.py             # Robust CSV & Excel loading (multi-encoding)
│   ├── data_validator.py          # Dataset schema, nulls, duplicates, cardinality
│   ├── data_cleaner.py            # Type inference, currency/percent conversions, imputation
│   ├── smart_column_detector.py   # Semantic roles (measures, durations, IDs, dates)
│   ├── eda_engine.py              # Statistical profiling, correlations, distributions
│   ├── kpi_engine.py              # Aggregation engine for KPI calculations
│   ├── smart_kpi_selector.py      # Selects max 4 business-relevant KPIs dynamically
│   ├── kpi_formatter.py           # Formats currencies, percentages, counts, scores
│   ├── smart_chart_selector.py    # Recommends max 4 diverse charts (bar, line, pie, etc.)
│   ├── cleaning_report_generator.py # Formats audit report of transformations
│   ├── excel_exporter.py          # Master openpyxl multi-sheet workbook builder
│   ├── dashboard_generator.py     # Native openpyxl KPI cards & chart placement
│   ├── output_validator.py        # Validates generated .xlsx structure & chart integrity
│   └── pipeline.py                # End-to-end execution pipeline
│
├── static/
│   ├── style.css                  # Professional SaaS light theme styling
│   └── script.js                  # Vanilla JS upload, stepper animations, results
│
├── templates/
│   └── index.html                 # Modern responsive HTML5 UI
│
├── tests/
│   ├── __init__.py
│   ├── generate_test_data.py      # Synthetic test datasets generator
│   ├── test_pipeline.py           # Automated test suite covering all 6 test scenarios
│   └── test_data/                 # Generated test datasets
│
├── api.py                         # FastAPI application & REST endpoints
├── app.py                         # Standalone runner entrypoint (python app.py)
├── requirements.txt               # Dependencies
├── README.md                      # Documentation
└── .gitignore                     # Git ignore rules for data, cache, and venvs
```

---

## 6. Installation & Setup

### Prerequisites
- Python 3.10 or higher
- pip package manager

### 1. Clone or Navigate to Project
```powershell
cd "c:\Users\hp\Desktop\automated Excel\automated_excel_data_analyst"
```

### 2. Create and Activate Virtual Environment
```powershell
python -m venv venv
venv\Scripts\activate
```

### 3. Install Dependencies
```powershell
pip install -r requirements.txt
```

---

## 7. Running the Application

### Option A: Via Application Runner
```powershell
python app.py
```

### Option B: Via Uvicorn Directly
```powershell
uvicorn api:app --reload
```

Open your browser and navigate to:
```text
http://127.0.0.1:8000/
```

---

## 8. API Reference

| Method | Endpoint | Description |
|---|---|---|
| `GET` | `/` | Web UI Dashboard (HTML) or API status (JSON) |
| `GET` | `/health` | Server health check `{"status": "healthy"}` |
| `POST` | `/analyze` | Multipart upload of `.xlsx` or `.csv`, executes analysis pipeline |
| `GET` | `/download/{filename}` | Secure download of generated `.xlsx` workbook |

### Example Response (`POST /analyze`):
```json
{
  "status": "success",
  "message": "Data analysis completed successfully.",
  "original_filename": "sales.xlsx",
  "input_type": ".xlsx",
  "selected_sheet": "Sales Data",
  "charts_generated": 4,
  "output_filename": "sales_analyzed_792c3a59f1.xlsx",
  "download_url": "/download/sales_analyzed_792c3a59f1.xlsx",
  "kpis": [
    {
      "label": "Total Records",
      "value": 90,
      "formatted_value": "90",
      "role": "count",
      "aggregation": "count"
    },
    {
      "label": "Total Revenue",
      "value": 245200.0,
      "formatted_value": "₹245,200.00",
      "role": "currency",
      "aggregation": "sum"
    }
  ],
  "cleaning_report": {
    "original_rows": 90,
    "cleaned_rows": 90,
    "duplicates_removed": 0,
    "missing_values_before": 0,
    "missing_values_after": 0
  }
}
```

---

## 9. Automated Testing

Run the end-to-end test suite covering all 6 test scenarios:
```powershell
python -m unittest tests/test_pipeline.py
```

Test scenarios covered:
1. **Sales XLSX**: Verifies date detection, revenue KPIs, and line/bar chart construction.
2. **Employee Performance CSV**: Proves duration columns like `Active Time` are protected as numeric metrics and not converted to datetimes.
3. **Messy CSV**: Verifies duplicate removal, imputation, and `₹` currency normalization.
4. **Numeric Heavy CSV**: Tests continuous distribution calculations and correlation matrices.
5. **Categorical Heavy CSV**: Tests categorical segmentation and proportions.
6. **Multi-Sheet XLSX**: Confirms automatic scoring selects `Raw Data` over summary/notes sheets.
7. **File Validator**: Tests rejection of invalid extensions and empty files.
8. **FastAPI Endpoints**: Tests `/health`, `/analyze`, and `/download` with `TestClient`.

---

## 10. Limitations & Future Roadmap

- **Row Limitations**: Standard OpenPyXL performance is optimized for datasets up to 100,000 rows.
- **Future Enhancements**:
  - Support for multi-table extraction from single worksheets.
  - Natural-language query interface over the generated dataset.
  - Customizable theme colors and corporate logo embedding for Excel templates.
