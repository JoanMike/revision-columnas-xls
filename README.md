# XLS Column Verifier

![License](https://img.shields.io/badge/license-PolyForm%20NC%201.0.0-blue)
![Python](https://img.shields.io/badge/python-3.12%2B-blue?logo=python&logoColor=white)
![Platform](https://img.shields.io/badge/platform-Windows-0078D6?logo=windows&logoColor=white)

> Check that every XLS/XLSX report downloaded from SAP in a folder has the same number of columns with data — differences are highlighted so you can spot them fast.

## Overview

This Python script scans a folder of Excel reports exported from SAP and verifies that all files share the same column count. When differences are found, it clearly reports the file name and column count, highlighted in color so they are easy to locate. Corrupt or unreadable files are reported with their error without stopping the analysis of the rest.

<div align="center">
	<img width="743" height="716" alt="XLS Column Verifier output" src="https://github.com/user-attachments/assets/7b95ff85-74c7-4b4e-adce-54f0b1a3073a" />
</div>

## Features

- **Real-time progress bar** to follow the analysis.
- **Color-coded results table**: files with differences are shown first and in red; unreadable files display the error reason.
- **Highlighted panel** with the exact names of the files that differ.
- **Parallel processing** (multi-threaded) for speed with many files.
- **Robust against corrupt files**: a damaged file is reported with its error and does not stop the rest of the analysis.
- **Header comparison** to detect renamed, reordered, or duplicated columns.
- **Recursive scanning** of subfolders (files are shown with their relative path to distinguish repeated names).
- **Export** results to CSV or JSON.
- **Quiet/verbose modes** for automation or detailed inspection.
- Supports native Excel files (.xls, .xlsx) and SAP text exports (UTF-16 LE and BE).

## Tech Stack

- **Python 3.12+** with `pandas`, `openpyxl`, `xlrd`, `rich` and `tkinter` (bundled with Python).
- Tests with `pytest`; dependency audit with `pip-audit`.

## Requirements

- Python 3.12 or higher.
- Libraries: `pandas`, `openpyxl` (for .xlsx), `xlrd` (for legacy .xls), `rich`.

## Installation

```bash
pip install -r requirements.txt
```

Or manually:

```bash
pip install pandas openpyxl xlrd rich
```

## Usage

### Interactive mode (graphical dialog)

```bash
python check_columns.py
```

A dialog opens to select the folder containing the XLS files.

### Command-line mode (CLI)

```bash
python check_columns.py --carpeta "C:\path\to\my\folder"
python check_columns.py -c "C:\path\to\my\folder"
```

### Available options

| Option | Short | Description |
|--------|-------|-------------|
| `--carpeta` | `-c` | Path to the folder with XLS files |
| `--recursivo` | `-r` | Also scan subfolders |
| `--comparar-headers` | | Compare column names in addition to the count |
| `--exportar` | `-e` | Export results to a file (.csv or .json) |
| `--verbose` | `-v` | Show extra metadata (rows, size, format, time) |
| `--quiet` | `-q` | Silent mode: one summary line + files with differences |
| `--workers` | `-w` | Parallel processing threads (integer >= 1, default: 4) |

### Examples

```bash
# Basic analysis with graphical dialog
python check_columns.py

# Analysis with progress bar and extra details
python check_columns.py -c "C:\SAP\Reports" -v

# Recursive scan comparing headers
python check_columns.py -c "C:\SAP\Reports" -r --comparar-headers

# Export results to JSON
python check_columns.py -c "C:\SAP\Reports" -e results.json

# Silent mode for batch scripts
python check_columns.py -c "C:\SAP\Reports" -q
```

### Exit codes

| Code | Meaning |
|------|---------|
| 0 | Success: all files have the same number of columns |
| 1 | Error: no folder selected, no files, or read failure |
| 2 | Differences found between files |

## Notes

- "Columns with data" means columns with at least one non-empty value in any row.
- If a file cannot be read as Excel, the script tries to read it as CSV (common in SAP exports).
- If a file cannot be read, the error reason is shown in the table and the file is skipped without stopping the rest of the analysis.
- The script reads the first sheet of each Excel file, or the full CSV file.
- Files with differences are shown **first** in the table and in a separate red panel for easy spotting.
- On legacy Windows consoles (cmd with cp1252/cp850), emojis and symbols degrade to `?` instead of failing; on Windows Terminal or any modern terminal they render correctly.

## Development

To contribute or modify the script, also install the development dependencies:

```bash
pip install -r requirements-dev.txt
```

Run the tests (they generate Excel fixtures on the fly, no external data needed):

```bash
pytest tests/ -v
```

Audit dependencies for vulnerabilities:

```bash
pip-audit -r requirements.txt
```

## License

Distributed under the **PolyForm Noncommercial License 1.0.0** — free for
noncommercial use only. See [LICENSE](LICENSE) for the full license text.

Copyright (c) 2026 Jose Miguel Maldonado Garcia

## Author

**Jose Miguel Maldonado Garcia** — [@JoanMike](https://github.com/JoanMike)
