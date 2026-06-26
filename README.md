# Excel → DDL Automator

> Upload any messy spreadsheet. Get a production-ready `CREATE TABLE` statement, a data quality audit, and a data dictionary — in seconds.

[![GitHub release](https://img.shields.io/github/v/release/aditcher/excel-to-ddl)](https://github.com/aditcher/excel-to-ddl/releases)
[![Python](https://img.shields.io/badge/python-3.9%2B-blue)](https://www.python.org/)
[![License: MIT](https://img.shields.io/badge/License-MIT-green.svg)](LICENSE)

---

## Why this exists

Every data analyst knows the pain: a business stakeholder hands you a spreadsheet with no schema documentation, inconsistent formatting, mixed data types, and blank cells scattered throughout. Before you can load it into a warehouse, you need a table definition.

This tool automates that intake workflow — the part that usually takes 20–30 minutes of manual inspection and typing.

---

## Features

| Feature | Detail |
|---|---|
| **Type inference** | INTEGER, FLOAT, VARCHAR(50/255), TEXT, BOOLEAN, DATE, TIMESTAMP, VARIANT/JSON for mixed-type columns |
| **Null % audit** | Visual bar chart + counts per column — instant data quality snapshot |
| **Duplicate PK detection** | Warns if your candidate primary key has duplicates before you commit to a schema |
| **FK candidate hints** | Scans column names for `_id`, `_key`, `_code`, `_ref` patterns and flags them in DDL comments |
| **Multi-sheet support** | Dropdown selector for workbooks with multiple tabs |
| **Column name sanitization** | Converts `"Phone #"` → `phone_num`, `"Date of Birth"` → `date_of_birth` — with a before/after log |
| **Data dictionary export** | One-click `.csv` export: column names, types, null %, unique counts, sample values |
| **Sample INSERTs** | Optional `INSERT INTO` statements for the first 5 rows |
| **5 SQL dialects** | Snowflake · SQL Server (T-SQL) · MySQL · PostgreSQL · Oracle |
| **Copy or save** | Copy DDL to clipboard or save as `.sql` |

---

## Screenshots

> *Add screenshots here after first build*

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
python main.py
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
4. Optionally check **Include sample INSERTs**
5. Click **⚡ Generate DDL + Data Dictionary**
6. Review the **DDL Output**, **Data Dictionary**, and **Column Name Map** tabs
7. **Copy**, **Save .sql**, or **Save Dict .csv**

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

### Windows EXE (via GitHub Actions)
Push a version tag to trigger an automated build:
```bash
git tag v1.0.0
git push origin v1.0.0
```
GitHub Actions builds both the `.exe` and `.dmg` and attaches them to the release automatically.

---

## Project Structure

```
excel-to-ddl/
├── main.py                    # Application entry point + all logic
├── requirements.txt           # openpyxl only — no database drivers needed
├── build_mac.sh               # Local macOS PyInstaller build script
├── .github/
│   └── workflows/
│       └── release.yml        # CI/CD: Windows EXE + macOS DMG on git tag
└── README.md
```

---

## Tech Stack

- **Python 3.9+** — standard library (`tkinter`, `csv`, `re`, `datetime`)
- **openpyxl** — Excel file reading (only non-stdlib dependency)
- **PyInstaller** — packaging to native executables
- **GitHub Actions** — cross-platform CI/CD release pipeline

---

## Author

**Aaron Ditcher** — Senior Data & Analytics Engineer  
[github.com/aditcher](https://github.com/aditcher) · [Portfolio Dashboards](https://aditcher.github.io/Dashboard)

---

## License

MIT — free to use, fork, and adapt.
