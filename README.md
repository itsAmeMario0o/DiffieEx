# DiffieEx - Excel File Comparison Script

This script compares two Excel files with multiple sheets and identifies overlapping data between corresponding sheets. The overlapping data is written into a new Excel file.

## Features

- Compares all sheets with matching names in both Excel files.
- Finds overlapping rows by performing an **inner join** on all columns.
- Writes the overlapping data to a new Excel file, preserving the original sheet names.

## Prerequisites

The script requires Python 3.6 or later, and the following libraries:

- `pandas` (for data manipulation)
- `openpyxl` (for reading and writing `.xlsx` files)

You can install the required dependencies by running:

```bash
pip install -r requirements.txt