#!/usr/bin/env python3
"""Convert a CSV file to Excel (.xlsx).

Default behavior:
- Input:  团建联络员打分结果_自动评估.csv
- Output: 团建联络员打分结果_自动评估.xlsx
- Sheet:  Sheet1

Usage:
    python csv_to_excel.py
    python csv_to_excel.py --input 团建联络员打分结果_自动评估.csv
    python csv_to_excel.py --input input.csv --output output.xlsx
    python csv_to_excel.py --sheet "打分结果"

Requirements:
    pip install openpyxl
"""

from __future__ import annotations

import argparse
import csv
import sys
from pathlib import Path


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Convert .csv to .xlsx")
    parser.add_argument(
        "--input",
        default="团建联络员打分结果_自动评估.csv",
        help="Input .csv file path (default: 团建联络员打分结果_自动评估.csv)",
    )
    parser.add_argument(
        "--output",
        default=None,
        help="Output .xlsx file path (default: same name as input, .xlsx)",
    )
    parser.add_argument(
        "--sheet",
        default="Sheet1",
        help="Worksheet name in output file (default: Sheet1)",
    )
    parser.add_argument(
        "--overwrite",
        action="store_true",
        help="Overwrite output file if it already exists.",
    )
    return parser.parse_args()


def build_output_path(input_path: Path, output_arg: str | None) -> Path:
    if output_arg:
        return Path(output_arg).expanduser().resolve()
    return input_path.with_suffix(".xlsx")


def require_openpyxl():
    try:
        from openpyxl import Workbook
    except ImportError as exc:
        raise RuntimeError(
            "Missing dependency 'openpyxl'. Install it with: pip install openpyxl"
        ) from exc
    return Workbook


def iter_csv_rows(csv_path: Path):
    # utf-8-sig allows reading files saved with UTF-8 BOM from Excel.
    with csv_path.open("r", encoding="utf-8-sig", newline="") as f:
        reader = csv.reader(f)
        for row in reader:
            yield row


def write_xlsx(csv_path: Path, output_path: Path, sheet_name: str) -> None:
    Workbook = require_openpyxl()

    wb = Workbook()
    ws = wb.active
    ws.title = sheet_name[:31] if sheet_name else "Sheet1"

    for row in iter_csv_rows(csv_path):
        ws.append(row)

    wb.save(output_path)


def main() -> int:
    args = parse_args()

    input_path = Path(args.input).expanduser().resolve()
    if not input_path.exists() or not input_path.is_file():
        print(f"Error: input file not found: {input_path}", file=sys.stderr)
        return 1

    if input_path.suffix.lower() != ".csv":
        print("Error: input file must be a .csv file.", file=sys.stderr)
        return 1

    output_path = build_output_path(input_path, args.output)

    if output_path.exists() and not args.overwrite:
        print(
            f"Error: output already exists: {output_path}\\n"
            "Use --overwrite to replace it.",
            file=sys.stderr,
        )
        return 1

    output_path.parent.mkdir(parents=True, exist_ok=True)

    try:
        write_xlsx(input_path, output_path, args.sheet)
    except RuntimeError as exc:
        print(f"Error: {exc}", file=sys.stderr)
        return 1
    except Exception as exc:  # noqa: BLE001
        print(f"Error: conversion failed ({exc})", file=sys.stderr)
        return 2

    print(f"OK: {input_path} -> {output_path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
