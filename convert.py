#!/usr/bin/env python3
"""
Batch File Converter — PDF -> xlsx  |  xls -> xlsx

Usage:
    python convert.py <input_dir> [output_dir]

    input_dir  : folder containing PDF and/or XLS files to convert
    output_dir : folder where Excel files are saved (default: ./output)

Each PDF is saved as a separate .xlsx file with one sheet per PDF page.
Each XLS is saved as a separate .xlsx file preserving all sheets.
"""

import sys
import argparse
from pathlib import Path

import pdfplumber
import pandas as pd
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from tqdm import tqdm


def extract_tables_from_pdf(pdf_path: Path) -> dict[int, list[pd.DataFrame]]:
    """Return {page_number: [DataFrame, ...]} for every table found in the PDF."""
    pages: dict[int, list[pd.DataFrame]] = {}
    with pdfplumber.open(pdf_path) as pdf:
        for page_num, page in enumerate(pdf.pages, start=1):
            tables = page.extract_tables()
            if not tables:
                continue
            dfs = []
            for table in tables:
                if not table:
                    continue
                df = pd.DataFrame(table[1:], columns=table[0])
                dfs.append(df)
            if dfs:
                pages[page_num] = dfs
    return pages


def write_to_excel(pages: dict[int, list[pd.DataFrame]], output_path: Path) -> None:
    """Write extracted tables to an Excel file, one sheet per page."""
    wb = Workbook()
    wb.remove(wb.active)  # remove default empty sheet

    for page_num, dfs in pages.items():
        ws = wb.create_sheet(title=f"Page {page_num}")
        current_row = 1
        for df in dfs:
            for r_idx, row in enumerate(dataframe_to_rows(df, index=False, header=True)):
                for c_idx, value in enumerate(row, start=1):
                    ws.cell(row=current_row, column=c_idx, value=value)
                current_row += 1
            current_row += 1  # blank row between tables

    wb.save(output_path)


def convert_xls_to_xlsx(xls_path: Path, output_dir: Path) -> str:
    """Convert a single .xls file to .xlsx, preserving all sheets."""
    try:
        xls = pd.ExcelFile(xls_path, engine="xlrd")
        output_path = output_dir / (xls_path.stem + ".xlsx")
        with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
            for sheet_name in xls.sheet_names:
                df = xls.parse(sheet_name)
                df.to_excel(writer, sheet_name=sheet_name, index=False)
        return f"ok ({len(xls.sheet_names)} sheet(s)) -> {output_path.name}"
    except Exception as exc:
        return f"ERROR: {exc}"


def convert_pdf_to_excel(pdf_path: Path, output_dir: Path) -> str:
    """Convert a single PDF file to Excel. Returns a status string."""
    try:
        pages = extract_tables_from_pdf(pdf_path)
        if not pages:
            return "skipped (no tables found)"

        output_path = output_dir / (pdf_path.stem + ".xlsx")
        write_to_excel(pages, output_path)
        total_tables = sum(len(dfs) for dfs in pages.values())
        return f"ok ({len(pages)} page(s), {total_tables} table(s)) -> {output_path.name}"
    except Exception as exc:
        return f"ERROR: {exc}"


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Convert PDF and XLS files to Excel (.xlsx)."
    )
    parser.add_argument("input_dir", help="Directory containing PDF and/or XLS files")
    parser.add_argument(
        "output_dir",
        nargs="?",
        default="output",
        help="Directory for output Excel files (default: ./output)",
    )
    args = parser.parse_args()

    input_dir = Path(args.input_dir).expanduser().resolve()
    output_dir = Path(args.output_dir).expanduser().resolve()

    if not input_dir.is_dir():
        print(f"Error: '{input_dir}' is not a directory.")
        sys.exit(1)

    pdf_files = sorted(input_dir.glob("*.pdf"))
    xls_files = sorted(input_dir.glob("*.xls"))
    all_files = pdf_files + xls_files

    if not all_files:
        print(f"No PDF or XLS files found in '{input_dir}'.")
        sys.exit(0)

    output_dir.mkdir(parents=True, exist_ok=True)
    print(f"Found {len(pdf_files)} PDF and {len(xls_files)} XLS file(s). Output -> {output_dir}\n")

    results = {}
    for path in tqdm(all_files, desc="Converting", unit="file"):
        if path.suffix.lower() == ".pdf":
            results[path.name] = convert_pdf_to_excel(path, output_dir)
        else:
            results[path.name] = convert_xls_to_xlsx(path, output_dir)

    print("\nSummary:")
    for name, status in results.items():
        print(f"  {name}: {status}")


if __name__ == "__main__":
    main()
