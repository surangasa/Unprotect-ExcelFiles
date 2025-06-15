# Unprotect-ExcelFiles

A cross-platform Python tool to remove workbook, worksheet, VBA, structure,
and file-open passwords from Microsoft Excel and OpenDocument spreadsheets.

## Features

- Supports `.xlsx`, `.xlsm`, `.xltx`, `.xltm`, `.xls`, `.xlsb`, and `.ods`
- Clears workbook protection, sheet protection, VBA project passwords, and
  file-open encryption
- Progress-bar feedback with detailed error messages
- Runs on Windows, macOS, and Linux

## Installation

```bash
pip install -r requirements.txt
```

## Usage

```bash
python UnprotectExcel_v1.py /path/to/locked.xlsx
```

The script creates `/path/to/locked_unlocked.xlsx` in the same directory.

## Testing

After installing the dependencies, run the unit tests:

```bash
pytest -q
```

## Building executables

A GitHub Actions workflow builds a standalone Windows EXE and macOS app
using PyInstaller. Built binaries are pushed to the `gh-pages` branch so they
can be downloaded from GitHub Pages. See `.github/workflows/build.yml` for
details.
