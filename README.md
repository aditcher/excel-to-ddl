# Excel → DDL Automator

> Upload any messy spreadsheet. Get a production-ready `CREATE TABLE` statement, a full DML script, a data quality audit, and a data dictionary — in seconds.

[![GitHub release](https://img.shields.io/github/v/release/aditcher/excel-to-ddl)](https://github.com/aditcher/excel-to-ddl/releases)
[![Python](https://img.shields.io/badge/python-3.9%2B-blue)](https://www.python.org/)
[![License: MIT](https://img.shields.io/badge/License-MIT-green.svg)](LICENSE)

---

## Why this exists

Every data analyst knows the pain: a business stakeholder hands you a spreadsheet with no schema documentation, inconsistent formatting, mixed data types, and blank cells scattered throughout. Before you can load it into a warehouse, you need a table definition — and then you need to actually get the data in.

This tool automates the entire intake workflow — schema generation, data cleaning, DML scripting, and documentation — the part that usually takes 20–30 minutes of manual work.

---

## What it does

### 1. Generates production-ready DDL across 5 SQL dialects

Paste straight into your query editor and run:

| Dialect | Example type names |
|---|---|
| **Snowflake** | `TIMESTAMP_NTZ`, `FLOAT`, `VARIANT` |
| **SQL Server (T-SQL)** | `DATETIME2`, `NVARCHAR(255)`, `DECIMAL(18,4)` |
| **MySQL** | `DATETIME`, `TINYINT(1)`, `JSON` |
| **PostgreSQL** | `NUMERIC(18,4)`, `BOOLEAN`, `JSONB` |
| **Oracle** | `NUMBER(10)`, `VARCHAR2(255)`, `CLOB` |

---

### 2. Generates a complete DML script *(new in v1.1.0)*

The **DML Output** tab produces a ready-to-use SQL script with 5 sections — dialect-aware for every supported platform:

| Section | What it generates |
|---|---|
| **1. Batch INSERTs** | One `INSERT` per row (SQL Server uses efficient 1,000-row `VALUES` blocks). Row limit is user-configurable: 100 / 500 / 1,000 / All rows |
| **2. Snowflake COPY INTO** | Full stage setup (`CREATE STAGE`), `PUT` command for SnowSQL, `COPY INTO` with error handling, and `RESULT_SCAN` validation *(Snowflake only)* |
| **3. CSV Staging Export** | Header row and column map — matches the DDL schema exactly. Use **Save Staged CSV** to export the cleaned file ready for bulk load |
| **4. UPDATE template** | Pre-built `UPDATE … SET … WHERE` with `<placeholder>` values for every non-PK column |
| **5. MERGE / UPSERT** | Dialect-aware: standard `MERGE INTO` for Snowflake / SQL Server / Oracle, `INSERT … ON DUPLICATE KEY UPDATE` for MySQL, `INSERT … ON CONFLICT DO UPDATE` for PostgreSQL |

---

### 3. Cleans your data automatically before generating

- Removes exact duplicate rows
- Strips leading/trailing whitespace from all cells
- Converts `"NA"`, `"N/A"`, `"null"`, `"none"` strings to real NULL
- Detects zeros used as null placeholders in categorical columns (color, size, style, type) and converts them to NULL
- Standardizes 10+ date formats to `YYYY-MM-DD` — handles `03/14/1985`, `14.11.1978`, `30/05/1995` and more
- Normalizes company/name columns — strips Inc., LLC, Ltd., Corp. suffixes and title-cases values
- Collapses empty strings to NULL

---

### 4. Audits data quality at a glance

- Row count, column count, total null cells
- Null % per column shown as a color-coded bar chart (blue → amber → red by severity)
- Duplicate primary key detection
- Foreign key candidate hints (columns named `_id`, `_key`, `_code`, `_ref`)
- Mixed-type column detection → flags as `VARIANT` / `JSONB`

---

### 5. Exports deliverables a BA or analyst would actually hand to a team

| Export | Description |
|---|---|
| **Save .sql** | `CREATE TABLE` statement, ready to execute |
| **Save DML .sql** | Full DML script — all 5 sections for the selected dialect |
| **Save Staged CSV** | Cleaned dataset with DDL-aligned column headers, ready for bulk load |
| **Save Dict .csv** | Lightweight data dictionary for any tool |
| **Save Dict .xlsx** | Formatted Excel data dictionary with bold headers, color-coded null %, PK/FK highlights, and a Cleaning Log sheet |
| **Row Preview tab** | First 50 cleaned rows before committing to the schema |
| **Column Name Map tab** | Before/after log of every column rename (`"Phone #"` → `phone_num`) |

---

## Quick Start

### Run from source (Mac / Windows / Linux)

```bash
# 1. Clone
git clone https://github.com/aditcher/excel-to-ddl.git
cd excel-to-ddl

# 2. Install dependencies (no database required)
pip install -r requirements.txt

# 3. Run
python3 excel_ddl_automator.py
```

### Download a built release

Head to [Releases](https://github.com/aditcher/excel-to-ddl/releases) and grab:

- **Windows:** `ExcelToDDL.exe` — double-click, no install needed
- **macOS:** `ExcelToDDL_mac.dmg` — drag to Applications

---

## Usage

1. Click **Browse** and select any `.xlsx`, `.xls`, or `.csv` file
   *(or hit **Load Demo Data** to try it instantly)*
2. If the workbook has multiple sheets, pick one from the dropdown
3. Set your **table name**, **SQL dialect**, and **nullability** preference
4. Set your **DML row limit** (100 / 500 / 1,000 / All rows)
5. Click **⚡ Generate DDL + DML**
6. Review the **DDL Output**, **DML Output**, **Data Dictionary**, and **Column Name Map** tabs
7. Use the toolbar to **Copy DDL**, **Copy DML**, **Save .sql**, **Save DML .sql**, **Save Staged CSV**, or **Save Dict .xlsx**

---

## Supported Input Formats

| Format | Extension |
|--------|-----------|
| Excel (Open XML) | `.xlsx` |
| Legacy Excel | `.xls` (via openpyxl) |
| Comma-separated | `.csv` |

---

## Type Inference Logic

| Data pattern | Inferred type |
|---|---|
| All whole numbers | `INTEGER` / `INT` / `NUMBER(10)` |
| Any decimal | `FLOAT` / `DECIMAL(18,4)` / `NUMERIC(18,4)` |
| `true/false/1/0/yes/no` | `BOOLEAN` / `BIT` / `TINYINT(1)` |
| `YYYY-MM-DD HH:MM` | `TIMESTAMP_NTZ` / `DATETIME2` / `TIMESTAMP` |
| `YYYY-MM-DD` | `DATE` |
| Short strings ≤50 chars | `VARCHAR(50)` / `NVARCHAR(50)` / `VARCHAR2(50)` |
| Strings ≤255 chars | `VARCHAR(255)` / `NVARCHAR(255)` |
| Long text | `TEXT` / `NVARCHAR(MAX)` / `CLOB` |
| Mixed types (3+ kinds) | `VARIANT` / `JSONB` / `JSON` / `CLOB` |

---

## Build from Source

### macOS
```bash
bash build_mac.sh
# Output: dist/ExcelToDDL.app  +  dist/ExcelToDDL_mac.dmg
```

### Windows EXE / macOS DMG (via GitHub Actions)
Push a version tag to trigger an automated build:
```bash
git tag v1.1.0
git push origin v1.1.0
```
GitHub Actions builds both the `.exe` and `.dmg` and attaches them to the release automatically.

---

## Project Structure

```
excel-to-ddl/
├── excel_ddl_automator.py     # Application entry point + all logic
├── requirements.txt           # openpyxl only — no database drivers needed
├── build_mac.sh               # Local macOS PyInstaller build script
├── .github/
│   └── workflows/
│       └── release.yml        # CI/CD: Windows EXE + macOS DMG on git tag
└── README.md
```

---

## Tech Stack

- **Python 3.9+** — standard library (`tkinter`, `csv`, `re`, `datetime`, `threading`)
- **openpyxl** — Excel file reading and formatted dictionary export (only non-stdlib dependency)
- **PyInstaller** — packaging to native executables
- **GitHub Actions** — cross-platform CI/CD release pipeline

---

## Changelog

### v1.1.0
- Added full DML generation engine (`build_dml()`) with 5 sections per dialect
- New **DML Output** tab with syntax highlighting
- Snowflake COPY INTO template with stage setup and load validation
- Dialect-aware MERGE / UPSERT (standard MERGE, ON DUPLICATE KEY, ON CONFLICT)
- UPDATE statement template with placeholder values
- User-configurable DML row limit (100 / 500 / 1,000 / All rows)
- New toolbar buttons: Copy DML, Save DML .sql, Save Staged CSV
- `format_sql_value()` helper for correct NULL / numeric / string quoting

### v1.0.0
- Initial release: DDL generation across 5 dialects
- Automated data cleaning engine
- Data quality audit with null % bar chart
- Formatted Excel data dictionary export
- Row Preview and Column Name Map tabs
- GitHub Actions CI/CD pipeline (Windows EXE + macOS DMG)

---

## Author

**Aaron Ditcher** — Senior Data & Analytics Engineer  
[github.com/aditcher](https://github.com/aditcher) · [Portfolio Dashboards](https://aditcher.github.io/Dashboard)

---

## License

MIT — free to use, fork, and adapt.
