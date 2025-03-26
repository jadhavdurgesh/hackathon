# PDF Table Extractor

## Overview
This script extracts tables from a PDF file and saves them as an Excel file. It first attempts to detect tables using a strict method that looks for a well-defined structure and falls back to a flexible method that considers consistent column counts.

## Features
- Extracts tables from PDF files using `PyMuPDF` (fitz).
- Detects tables using both strict and flexible heuristics.
- Cleans extracted text to remove illegal characters.
- Saves extracted tables as separate sheets in an Excel file using `pandas`.

## Requirements
Ensure you have the required dependencies installed:

```sh
pip install pymupdf pandas openpyxl
```

## How It Works
1. The script reads a PDF file and processes each page to detect tables.
2. It uses two methods to detect tables:
   - **Strict Detection**: Looks for a table with at least 5 columns and aligned rows.
   - **Flexible Detection**: Groups rows with similar column counts if strict detection fails.
3. The detected tables are stored in a structured format and cleaned.
4. Extracted tables are saved as an Excel file, with each table in a separate sheet.

## Usage
Run the script with the PDF file as an argument:

```sh
python pdf_table_extractor.py <pdf_file>
```

Example:

```sh
python pdf_table_extractor.py sample.pdf
```

## Output
- The script generates an Excel file with the extracted tables.
- If the input file is `sample.pdf`, the output will be `sample_tables.xlsx`.

## File Structure
```
.
├── pdf_table_extractor.py  # Main script
├── README.md               # Documentation
└── requirements.txt        # List of dependencies
```

## Notes
- If no tables are detected, the script will notify the user and exit.
- Tables with missing values are padded with empty cells to maintain structure.

## License
This project is licensed under the MIT License.

