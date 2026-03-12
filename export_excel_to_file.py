#!/usr/bin/env python3
"""
Export each worksheet in an Excel workbook to separate CSV files.

Key capabilities:
• Per‑sheet export with robust delimiter/quoting/NA options.
• Skip empty or non‑tabular sheets automatically.
• Validate that header names do not contain the chosen delimiter (can be disabled).
• Safe filename construction and timestamp suffix when an output would be overwritten.
• Optional logging with configurable level.

CLI examples:
export_excel_to_file.py SalesData.xlsx --out-dir ./out --prefix run01 \
    --delimiter tab --quoting minimal --na-rep N/A --log-level INFO

Output filename pattern:
{prefix}_{sanitized_sheet_name}_{sanitized_base}.csv

Exit codes:
0 success; 1 on fatal error.
"""
from __future__ import annotations

import argparse
import csv
import logging
import re
import sys
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Iterable, Optional, Sequence

import pandas as pd

# ----------------------------- Utilities -----------------------------

def setup_logging(level: str) -> None:
    numeric_level = getattr(logging, level.upper(), None)
    if not isinstance(numeric_level, int):
        raise ValueError(f"Invalid log level: {level}")
    logging.basicConfig(
        level=numeric_level,
        format="%(asctime)s %(levelname)s %(message)s",
        handlers=[logging.StreamHandler(sys.stdout)],
    )


def sanitize_for_filename(name: str, replacement: str = "_", max_len: int = 150) -> str:
    """Make a string safe as a filename component across platforms."""
    safe = re.sub(r'[\<>:"/\\\n?*\x00-\x1F]', replacement, name)
    safe = re.sub(r"\s+", replacement, safe)
    safe = re.sub(rf"{re.escape(replacement)}+", replacement, safe)
    safe = safe.strip(" ._")
    if not safe:
        safe = "sheet"
    if len(safe) > max_len:
        safe = safe[:max_len].rstrip(" ._")
    return safe


def dataframe_is_effectively_empty(df: pd.DataFrame) -> bool:
    """Consider empty if all rows/cols are NaN after dropping all-empty rows/cols."""
    if df.empty:
        return True
    df2 = df.dropna(how="all").dropna(axis=1, how="all")
    return df2.empty


def validate_headers_for_delimiter(df: pd.DataFrame, delimiter: str) -> None:
    for col in df.columns:
        if delimiter in str(col):
            raise ValueError(
                f"Delimiter '{delimiter}' found in column name '{col}'. "
                "This may corrupt CSV structure."
            )


def _parse_basic_escapes(token: str) -> str:
    mapping = {
        "\\t": "\t",
        "\\n": "\n",
        "\\r": "\r",
        "\\\\": "\\",
    }
    return mapping.get(token, token)


def parse_delimiter(value: Optional[str]) -> str:
    """Normalize textual tokens/escapes to a one-character delimiter."""
    if not value:
        return ","
    token = value.strip().lower()
    keywords = {
        "tab": "\t",
        "\\t": "\t",
        "pipe": "|",
        "comma": ",",
        "semicolon": ";",
        "space": " ",
    }
    if token in keywords:
        return keywords[token]
    token = _parse_basic_escapes(value)
    if len(token) == 1:
        return token
    raise ValueError(
        f"Unrecognized delimiter: {value!r}. Try one of: ',', '\\t', '|', ';', "
        f"'tab', 'pipe', 'comma', 'semicolon', 'space'."
    )


def parse_quoting(value: Optional[str]) -> int:
    if not value:
        return csv.QUOTE_MINIMAL
    token = value.strip().lower()
    mapping = {
        "minimal": csv.QUOTE_MINIMAL,
        "all": csv.QUOTE_ALL,
        "nonnumeric": csv.QUOTE_NONNUMERIC,
        "none": csv.QUOTE_NONE,
    }
    if token not in mapping:
        raise ValueError("Choose quoting from: minimal, all, nonnumeric, none.")
    return mapping[token]

# ----------------------------- Core export -----------------------------

@dataclass
class ExportOptions:
    prefix: str = ""
    sheet_names: Optional[Sequence[str]] = None
    header_row: Optional[int] = 0  # None means no header
    encoding: str = "utf-8"
    delimiter: str = ","
    quoting: int = csv.QUOTE_MINIMAL
    quotechar: str = '"'
    escapechar: Optional[str] = None
    doublequote: bool = True
    na_rep: str = ""
    allow_delimiter_in_header: bool = False
    log_level: Optional[str] = None


def export_excel_sheets_to_csv(
    excel_path: Path,
    out_dir: Path,
    opts: ExportOptions,
) -> list[Path]:
    """Export each (or selected) sheet to CSVs. Returns list of written file paths."""
    logger = logging.getLogger(__name__)
    if opts.log_level:
        setup_logging(opts.log_level)
    if not excel_path.exists():
        raise FileNotFoundError(f"Excel file not found: {excel_path}")
    out_dir.mkdir(parents=True, exist_ok=True)

    base_name = excel_path.stem
    safe_base = sanitize_for_filename(base_name)

    with pd.ExcelFile(excel_path, engine="openpyxl") as xls:
        available_sheets = xls.sheet_names
        if opts.sheet_names:
            missing = [s for s in opts.sheet_names if s not in available_sheets]
            if missing:
                raise ValueError(
                    f"Requested sheet(s) not in workbook: {missing}\n"
                    f"Available: {available_sheets}"
                )
            target_sheets = list(opts.sheet_names)
        else:
            target_sheets = available_sheets

        # Auto-assign escapechar if quoting is NONE and none provided
        effective_escapechar = opts.escapechar
        auto_assigned_escape = False
        if opts.quoting == csv.QUOTE_NONE and not effective_escapechar:
            effective_escapechar = "\\"
            auto_assigned_escape = True

        written: list[Path] = []
        for sheet in target_sheets:
            logger.info("Processing sheet: %s", sheet)
            header = opts.header_row if opts.header_row is not None else None
            df = pd.read_excel(xls, sheet_name=sheet, header=header)

            if dataframe_is_effectively_empty(df):
                logger.warning("Skipping empty/non-tabular sheet: %s", sheet)
                continue

            if opts.header_row is not None and not opts.allow_delimiter_in_header:
                validate_headers_for_delimiter(df, opts.delimiter)

            safe_sheet = sanitize_for_filename(str(sheet))
            out_name = f"{opts.prefix}_{safe_sheet}_{safe_base}.csv" if opts.prefix else f"{safe_sheet}_{safe_base}.csv"
            out_path = out_dir / out_name

            # Collision avoidance: append timestamp if exists
            if out_path.exists():
                ts = datetime.now().strftime("%Y%m%d_%H%M%S")
                out_name = f"{opts.prefix + '_' if opts.prefix else ''}{safe_sheet}_{safe_base}_{ts}.csv"
                out_path = out_dir / out_name
                logger.warning("File exists; appending timestamp: %s", out_path.name)

            df.to_csv(
                out_path,
                index=False,
                encoding=opts.encoding,
                sep=opts.delimiter,
                quoting=opts.quoting,
                quotechar=opts.quotechar,
                escapechar=effective_escapechar,
                doublequote=opts.doublequote,
                na_rep=opts.na_rep,
            )
            written.append(out_path)

        if auto_assigned_escape:
            print(
                "Note: quoting=none requested without --escapechar; auto-set escapechar='\\' to ensure valid output."
            )

        return written

# ----------------------------- CLI -----------------------------

def build_arg_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(
        description=(
            "Export each worksheet in an Excel file to separate CSV files. "
            "Output filenames follow: {prefix}_{sheet}_{base}.csv"
        )
    )
    p.add_argument("excel", type=Path, help="Path to the Excel file (.xlsx or .xls)")
    p.add_argument(
        "-o", "--out-dir", type=Path, default=Path("./csv_out"),
        help="Directory to write CSV files (default: ./csv_out)",
    )
    p.add_argument("--prefix", type=str, default="", help="Filename prefix (default: '')")
    p.add_argument("--sheets", nargs="+", help="Optional list of sheet names to export (default: all)")
    p.add_argument(
        "--header-row", type=int, default=0,
        help="0-based header row index, or -1 for no header (default: 0)",
    )
    p.add_argument("--encoding", type=str, default="utf-8", help="CSV encoding (default: utf-8)")
    p.add_argument(
        "--delimiter", "--sep", dest="delimiter", type=str, default=",",
        help=(
            "Output delimiter (default ','). Examples: ',', '\\t', '|', ';', "
            "'tab', 'pipe', 'comma', 'semicolon', 'space'."
        ),
    )
    p.add_argument(
        "--quoting", type=str, choices=["minimal", "all", "nonnumeric", "none"], default="minimal",
        help="CSV field quoting strategy (default: minimal)",
    )
    p.add_argument("--quotechar", type=str, default='"', help='Quote character (default: ")')
    p.add_argument(
        "--escapechar", type=str, default=None,
        help="Escape character (optional). If quoting=none and not provided, defaults to '\\'",
    )
    p.add_argument(
        "--no-doublequote", action="store_true", help="Disable doubling of quote characters inside fields",
    )
    p.add_argument(
        "--na-rep", type=str, default="",
        help="String to use for NA/NaN cells (default: empty). Supports \\t, \\n, \\r, \\.",
    )
    p.add_argument(
        "--allow-delimiter-in-header", action="store_true",
        help="Do not error if delimiter appears in header names",
    )
    p.add_argument(
        "--log-level", type=str, default=None,
        help="Optional logging level (e.g., DEBUG, INFO, WARNING)",
    )
    return p


def main() -> None:
    args = build_arg_parser().parse_args()

    header_row = None if args.header_row == -1 else args.header_row

    try:
        delimiter = parse_delimiter(args.delimiter)
        quoting = parse_quoting(args.quoting)
        na_rep = _parse_basic_escapes(args.na_rep)
        escapechar = _parse_basic_escapes(args.escapechar) if args.escapechar else None

        opts = ExportOptions(
            prefix=args.prefix,
            sheet_names=args.sheets,
            header_row=header_row,
            encoding=args.encoding,
            delimiter=delimiter,
            quoting=quoting,
            quotechar=args.quotechar,
            escapechar=escapechar,
            doublequote=not args.no_doublequote,
            na_rep=na_rep,
            allow_delimiter_in_header=args.allow_delimiter_in_header,
            log_level=args.log_level,
        )

        written = export_excel_sheets_to_csv(
            excel_path=args.excel,
            out_dir=args.out_dir,
            opts=opts,
        )
        if not written:
            print("No CSVs written (all target sheets empty or skipped).")
        else:
            print("Written files:")
            for p in written:
                print(f"- {p}")
    except Exception:
        logging.exception("Fatal error during execution")
        sys.exit(1)


if __name__ == "__main__":
    main()
