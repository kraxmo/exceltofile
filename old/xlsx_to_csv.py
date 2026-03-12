#!/usr/bin/env python3
"""
Convert each worksheet in an .xlsx file to a CSV, using only Python's standard library.

Features:
- Parses OpenXML (.xlsx/.xlsm) via zipfile + xml.etree.ElementTree
- Handles shared strings, inline strings, booleans, numbers, (formula values if present)
- Interprets dates/times for known Excel number formats
- Skips hidden sheets by default (use --include-hidden to process them)
- Streams worksheets with iterparse to keep memory usage modest
- Configurable CSV quoting via --quoting and optional --escapechar

Limitations:
- Does NOT support legacy .xls (BIFF) files
- Formulas are not calculated; if cached <v> exists, it is used; otherwise cell is blank
- Date detection relies on style numFmtId heuristics and may not match Excel locale formatting

Author: M365 Copilot
"""

from __future__ import annotations

import argparse
import csv
import datetime as dt
import os
import re
import sys
import unicodedata
import zipfile
from typing import Dict, Iterable, List, Optional, Set, Tuple
import xml.etree.ElementTree as ET


SPREADSHEETML_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
CONTENT_TYPES = "http://schemas.openxmlformats.org/package/2006/content-types"
NS = {"ss": SPREADSHEETML_NS, "rel": REL_NS}

# Known/builtin Excel number format IDs that represent dates/times.
BUILTIN_DATE_FORMAT_IDS: Set[int] = {
    14, 15, 16, 17, 18, 19, 20, 21, 22,
    27, 30, 36, 45, 46, 47, 50, 57
}

# Regex to detect custom date/time formats in styles
DATE_FMT_TOKENS = re.compile(r"[ymdhYsHMs]", re.IGNORECASE)  # crude heuristic


def q(tag: str) -> str:
    """Fully qualified tag for spreadsheetml."""
    return f"{{{SPREADSHEETML_NS}}}{tag}"


def sanitize_filename(name: str) -> str:
    """Make a filename safe across OSes."""
    name = unicodedata.normalize("NFKC", name)
    name = re.sub(r"[\\/:*?\"<>|\r\n\t]+", "_", name).strip()
    return name or "Sheet"


def ensure_unique_path(base_path: str) -> str:
    """If path exists, append (2), (3), ... until unique."""
    if not os.path.exists(base_path):
        return base_path
    root, ext = os.path.splitext(base_path)
    i = 2
    while True:
        candidate = f"{root} ({i}){ext}"
        if not os.path.exists(candidate):
            return candidate
        i += 1


def load_shared_strings(zf: zipfile.ZipFile) -> List[str]:
    """Read shared strings table; return list indexed by sst index."""
    sst_path = "xl/sharedStrings.xml"
    shared: List[str] = []
    if sst_path not in zf.namelist():
        return shared

    with zf.open(sst_path) as f:
        for event, elem in ET.iterparse(f, events=("end",)):
            if elem.tag == q("si"):
                shared.append(_si_text(elem))
                elem.clear()
    return shared


def _si_text(si_elem: ET.Element) -> str:
    """Extract text from a shared string item <si> (plain or rich text)."""
    parts: List[str] = []
    for t in si_elem.findall(q("t")):
        parts.append(_text_with_space(t))
    for r in si_elem.findall(q("r")):
        t = r.find(q("t"))
        if t is not None:
            parts.append(_text_with_space(t))
    return "".join(parts)


def _text_with_space(t_elem: ET.Element) -> str:
    text = t_elem.text or ""
    if t_elem.get("{http://www.w3.org/XML/1998/namespace}space") == "preserve":
        return text
    return text.strip()


def parse_workbook(zf: zipfile.ZipFile) -> Tuple[List[Tuple[str, str, str]], Dict[str, str], bool]:
    """
    Parse workbook to get sheets [(name, rId, state)], relationships (rId->target),
    and date1904 flag. Map ALL relationships, not only 'worksheet' type.
    """
    with zf.open("xl/workbook.xml") as f:
        wb = ET.parse(f).getroot()

    # date1904
    wbpr = wb.find(q("workbookPr"))
    date1904 = False
    if wbpr is not None and wbpr.get("date1904") in ("1", "true", "True"):
        date1904 = True

    # Collect <sheet> nodes
    sheets_info: List[Tuple[str, str, str]] = []
    sheets_parent = wb.find(q("sheets"))
    if sheets_parent is not None:
        for s in sheets_parent.findall(q("sheet")):
            name = s.get("name", "Sheet")
            rid = s.get(f"{{{REL_NS}}}id")  # r:id
            state = s.get("state", "visible")
            if not rid:
                continue
            sheets_info.append((name, rid, state))

    # Map ALL relationships in workbook.xml.rels (not just worksheet)
    rels_path = "xl/_rels/workbook.xml.rels"
    rmap: Dict[str, str] = {}
    if rels_path in zf.namelist():
        with zf.open(rels_path) as f:
            rels_root = ET.parse(f).getroot()
            for rel in rels_root.findall(f"{{{REL_NS}}}Relationship"):
                rid = rel.get("Id")
                target = rel.get("Target", "")
                if rid and target:
                    # Keep every mapping; we'll validate the target later.
                    rmap[rid] = target
    return sheets_info, rmap, date1904

def parse_styles(zf: zipfile.ZipFile) -> Tuple[Set[int], Dict[int, int]]:
    """Parse styles to find which xf indices correspond to date/time formats."""
    styles_path = "xl/styles.xml"
    date_xf_ids: Set[int] = set()
    xf_to_numfmt: Dict[int, int] = {}

    if styles_path not in zf.namelist():
        return date_xf_ids, xf_to_numfmt

    with zf.open(styles_path) as f:
        root = ET.parse(f).getroot()

    custom_date_numfmt_ids: Set[int] = set()
    numFmts = root.find(q("numFmts"))
    if numFmts is not None:
        for numFmt in numFmts.findall(q("numFmt")):
            try:
                numFmtId = int(numFmt.get("numFmtId", ""))
            except ValueError:
                continue
            fmt_code = numFmt.get("formatCode", "")
            if DATE_FMT_TOKENS.search(fmt_code):
                custom_date_numfmt_ids.add(numFmtId)

    cellXfs = root.find(q("cellXfs"))
    if cellXfs is not None:
        for idx, xf in enumerate(cellXfs.findall(q("xf"))):
            try:
                numFmtId = int(xf.get("numFmtId", "0"))
            except ValueError:
                numFmtId = 0
            xf_to_numfmt[idx] = numFmtId
            if numFmtId in BUILTIN_DATE_FORMAT_IDS or numFmtId in custom_date_numfmt_ids:
                date_xf_ids.add(idx)

    return date_xf_ids, xf_to_numfmt


def excel_serial_to_datetime(value: float, date1904: bool) -> dt.datetime:
    """Convert Excel serial date/time to Python datetime."""
    base = dt.datetime(1904, 1, 1) if date1904 else dt.datetime(1899, 12, 30)
    return base + dt.timedelta(days=float(value))


def cell_coordinate_to_col_index(r: str) -> int:
    """Convert a cell reference like 'C5' to zero-based column index (C -> 2)."""
    m = re.match(r"([A-Z]+)", r.upper())
    if not m:
        return 0
    letters = m.group(1)
    col = 0
    for ch in letters:
        col = col * 26 + (ord(ch) - ord("A") + 1)
    return col - 1


def read_cell_value(
    c_elem: ET.Element,
    shared_strings: List[str],
    date_xf_ids: Set[int],
    date1904: bool
) -> str:
    """Extract a human-readable cell value from <c>."""
    t = c_elem.get("t")
    s_idx = c_elem.get("s")
    xf_idx = int(s_idx) if s_idx is not None and s_idx.isdigit() else None

    v_elem = c_elem.find(q("v"))
    f_elem = c_elem.find(q("f"))
    is_elem = c_elem.find(q("is"))

    if t == "inlineStr" and is_elem is not None:
        return _si_text(is_elem)

    if t == "s" and v_elem is not None and (v_elem.text or "").strip().isdigit():
        sst_idx = int(v_elem.text.strip())
        return shared_strings[sst_idx] if 0 <= sst_idx < len(shared_strings) else ""

    if t == "b" and v_elem is not None:
        return "TRUE" if v_elem.text and v_elem.text.strip() == "1" else "FALSE"

    if t in ("str", "e"):
        return (v_elem.text or "").strip() if v_elem is not None else ""

    if v_elem is not None and v_elem.text is not None:
        raw = v_elem.text.strip()
        if raw == "":
            return ""
        if xf_idx is not None and xf_idx in date_xf_ids:
            try:
                num = float(raw)
                dt_val = excel_serial_to_datetime(num, date1904)
                if abs(num - int(num)) > 1e-10:
                    return dt_val.replace(microsecond=0).isoformat(sep=" ")
                else:
                    return dt_val.date().isoformat()
            except Exception:
                return raw
        return raw

    if f_elem is not None and v_elem is None:
        return ""

    return ""


def iter_sheet_rows(
    zf: zipfile.ZipFile,
    sheet_path: str,
    shared_strings: List[str],
    date_xf_ids: Set[int],
    date1904: bool
) -> Iterable[List[str]]:
    """Stream rows from a worksheet XML, yielding list[str] for each row."""
    with zf.open(f"xl/{sheet_path}") as f:
        context = ET.iterparse(f, events=("end",))
        for event, elem in context:
            if elem.tag == q("row"):
                cells: Dict[int, str] = {}
                for c in elem.findall(q("c")):
                    r = c.get("r", "")
                    col_idx = cell_coordinate_to_col_index(r)
                    cells[col_idx] = read_cell_value(c, shared_strings, date_xf_ids, date1904)
                row: List[str] = []
                if cells:
                    max_col = max(cells.keys())
                    row = [cells.get(i, "") for i in range(max_col + 1)]
                else:
                    row = []
                yield row
                elem.clear()


def parse_quoting(name: str) -> int:
    """Map user-friendly quoting names to csv module constants."""
    name = (name or "").strip().lower()
    if name == "minimal":
        return csv.QUOTE_MINIMAL
    if name == "all":
        return csv.QUOTE_ALL
    if name == "none":
        return csv.QUOTE_NONE
    if name == "nonnumeric":
        return csv.QUOTE_NONNUMERIC
    raise ValueError("Invalid quoting value. Choose from: minimal, all, none, nonnumeric.")


def write_csv(
    rows: Iterable[List[str]],
    out_path: str,
    delimiter: str = ",",
    quotechar: str = '"',
    encoding: str = "utf-8",
    quoting: int = csv.QUOTE_MINIMAL,
    escapechar: Optional[str] = None,
    doublequote: bool = True,
) -> None:
    os.makedirs(os.path.dirname(out_path), exist_ok=True)
    # Validate escapechar if provided
    if escapechar is not None and len(escapechar) != 1:
        raise ValueError("--escapechar must be a single character")
    try:
        with open(out_path, "w", newline="", encoding=encoding) as f:
            writer = csv.writer(
                f,
                delimiter=delimiter,
                quotechar=quotechar,
                quoting=quoting,
                lineterminator="\n",
                escapechar=escapechar,
                doublequote=doublequote,
            )
            for row in rows:
                writer.writerow(row)
    except csv.Error as e:
        # Common pitfall: QUOTE_NONE used without an escapechar when data contains delimiters/quotes.
        if quoting == csv.QUOTE_NONE and escapechar is None:
            raise ValueError(
                "csv.QUOTE_NONE requires --escapechar when data contains delimiter, "
                "quotechar, or newlines."
            ) from e
        raise



def convert_xlsx_to_csvs(
    xlsx_path: str,
    outdir: Optional[str],
    include_hidden: bool,
    delimiter: str,
    quotechar: str,
    encoding: str,
    quoting_name: str = "minimal",
    escapechar: Optional[str] = None,
    debug: bool = False,
) -> List[str]:
    if not os.path.isfile(xlsx_path):
        raise FileNotFoundError(f"Input file not found: {xlsx_path}")

    quoting = parse_quoting(quoting_name)

    base_name = os.path.splitext(os.path.basename(xlsx_path))[0]
    outdir = outdir or os.path.dirname(os.path.abspath(xlsx_path)) or "."
    written: List[str] = []

    try:
        with zipfile.ZipFile(xlsx_path) as zf:
            sheets_info, rmap, date1904 = parse_workbook(zf)
            shared_strings = load_shared_strings(zf)
            date_xf_ids, _ = parse_styles(zf)

        if debug:
            print(f"[DEBUG] Found {len(sheets_info)} sheet entries in workbook.xml")
            print(f"[DEBUG] Found {len(rmap)} relationship entries in workbook.xml.rels")
            if rmap:
                print("[DEBUG] Relationship Ids available:", ", ".join(sorted(rmap.keys())))

        for sheet_name, rid, state in sheets_info:
            if debug:
                print(f"[DEBUG] Sheet found: name='{sheet_name}', rId='{rid}', state='{state}'")

            if not include_hidden and state in ("hidden", "veryHidden"):
                if debug:
                    print(f"[DEBUG] Skipping hidden sheet: {sheet_name}")
                continue

            target = rmap.get(rid)
            if not target:
                if debug:
                    print(f"[DEBUG] rId={rid} not found in workbook.xml.rels; attempting heuristic fallback")

                # Try to match by sheetId ordering (if present) or scan worksheets
                # First, find the zero-based ordinal of this sheet in 'sheets_info'
                # (This is not guaranteed to align to filenames, but works for many generators.)
                try:
                    ordinal = [r for r in sheets_info].index((sheet_name, rid, state))
                except ValueError:
                    ordinal = None

                candidate_paths = []
                if ordinal is not None:
                    # Common names are sheet1.xml, sheet2.xml, ...
                    candidate_paths.append(f"xl/worksheets/sheet{ordinal+1}.xml")
                # Also add a scan of actual worksheet parts present in the zip
                with zipfile.ZipFile(xlsx_path) as _zf_scan:
                    for name in _zf_scan.namelist():
                        if name.startswith("xl/worksheets/") and name.endswith(".xml"):
                            candidate_paths.append(name)

                # Select the first candidate that exists
                chosen = None
                with zipfile.ZipFile(xlsx_path) as _zf_scan2:
                    for path in candidate_paths:
                        if path in _zf_scan2.namelist():
                            chosen = path
                            break

                if chosen:
                    if debug:
                        print(f"[DEBUG] Heuristic matched sheet '{sheet_name}' to: '{chosen}'")
                    zip_path = chosen
                else:
                    if debug:
                        print(f"[DEBUG] Could not heuristically resolve sheet '{sheet_name}'; skipping")
                    continue
            else:
                zip_path = resolve_sheet_zip_path(target)
            if debug:
                print(f"[DEBUG] rId={rid} target='{target}' -> resolved='{zip_path}'")

            # Verify the part exists in the package
            with zipfile.ZipFile(xlsx_path) as _zf_check:
                if zip_path not in _zf_check.namelist():
                    if debug:
                        print(f"[DEBUG] Resolved sheet path not found in ZIP: '{zip_path}'; trying fallback")
                    # Fallback: If target already contained 'xl/', try without it, or vice versa
                    fallback_paths = set()
                    if zip_path.startswith("xl/"):
                        fallback_paths.add(zip_path[len("xl/"):])  # remove "xl/"
                    else:
                        fallback_paths.add("xl/" + zip_path)       # add "xl/"
                    # Any fallback hit?
                    hit = None
                    for fp in fallback_paths:
                        if fp in _zf_check.namelist():
                            hit = fp
                            break
                    if hit:
                        if debug:
                            print(f"[DEBUG] Fallback resolved to: '{hit}'")
                        zip_path = hit
                    else:
                        if debug:
                            print(f"[DEBUG] No valid path found for sheet '{sheet_name}'; skipping")
                        continue

            safe_sheet = sanitize_filename(sheet_name)
            out_filename = f"{base_name} - {safe_sheet}.csv"
            out_path = ensure_unique_path(os.path.join(outdir, out_filename))

            # Now open using the resolved path (which already includes 'xl/' if needed)
            with zipfile.ZipFile(xlsx_path) as zf2:
                rows = iter_sheet_rows(zf2, zip_path.replace("xl/", "", 1) if zip_path.startswith("xl/") else zip_path, shared_strings, date_xf_ids, date1904)

            write_csv(
                rows,
                out_path,
                delimiter=delimiter,
                quotechar=quotechar,
                encoding=encoding,
                quoting=quoting,
                escapechar=escapechar,
                doublequote=True,
            )
            written.append(out_path)

    except zipfile.BadZipFile:
        raise ValueError("The file is not a valid .xlsx/.xlsm (OpenXML) package.")
    except KeyError as ex:
        raise ValueError(f"Missing expected workbook part in the .xlsx: {ex}")

    return written


def resolve_sheet_zip_path(target: str) -> str:
    """
    Normalize the target from workbook.xml.rels to an actual path in the ZIP.
    Accepts things like: 'worksheets/sheet1.xml', '/worksheets/sheet1.xml', 'xl/worksheets/sheet1.xml', '../worksheets/sheet1.xml'
    """
    t = target.strip()
    # Remove a leading slash
    t = t.lstrip("/")
    # Collapse simple '../' prefixes that sometimes appear
    while t.startswith("../"):
        t = t[3:]
    # Ensure we have exactly one 'xl/' prefix
    if not t.startswith("xl/"):
        t = f"xl/{t}"
    return t


def main(argv: Optional[List[str]] = None) -> int:
    parser = argparse.ArgumentParser(
        description="Convert each worksheet in an .xlsx file to a CSV (standard library only)."
    )
    parser.add_argument("xlsx", help="Path to the input .xlsx or .xlsm file")
    parser.add_argument("--outdir", default=None, help="Output directory (default: alongside the input file)")
    parser.add_argument("--delimiter", default=",", help="CSV delimiter (default: ,)")
    parser.add_argument("--quotechar", default='"', help='CSV quote character (default: ")')
    parser.add_argument("--encoding", default="utf-8", help="Output file encoding (default: utf-8)")
    parser.add_argument(
        "--include-hidden",
        action="store_true",
        help="Include hidden/veryHidden sheets (default: false)"
    )
    parser.add_argument(
        "--quoting",
        default="minimal",
        choices=["minimal", "all", "none", "nonnumeric"],
        help="CSV quoting style: minimal (default), all, none, nonnumeric"
    )
    parser.add_argument(
        "--escapechar",
        default=None,
        help="Single character used to escape when quoting=none (e.g., \\)."
    )

    parser.add_argument(
        "--debug",
        action="store_true",
        help="Print diagnostic information about detected sheets and resolved paths"
    )

    args = parser.parse_args(argv)

    try:
        outputs = convert_xlsx_to_csvs(
            xlsx_path=args.xlsx,
            outdir=args.outdir,
            include_hidden=args.include_hidden,
            delimiter=args.delimiter,
            quotechar=args.quotechar,
            encoding=args.encoding,
            quoting_name=args.quoting,
            escapechar=args.escapechar,
            debug=args.debug,
        )
        if outputs:
            print("Wrote:")
            for p in outputs:
                print("  ", p)
        else:
            print("No worksheets exported (perhaps all were hidden?).")
        return 0
    except Exception as e:
        print(f"ERROR: {e}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
