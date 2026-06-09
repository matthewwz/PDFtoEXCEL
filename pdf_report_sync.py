"""
Write the **Report No.** value into Intertek-style PDF AcroForm fields (top-right).

The newer Word→PDF forms use field names like ``Report No du rapport``,
``Report No du rapport_3``, etc. All matching fields are set to the same value
so repeated headers stay consistent.

Typical flow: row in Excel (``File`` + ``Report No.`` columns, export layout)
→ pick PDF → save updated copy.

CLI::

    python pdf_report_sync.py form.pdf mybook.xlsx -o form-updated.pdf
    python pdf_report_sync.py form.pdf --report "2026-WZ-1084" -o form-updated.pdf
    python pdf_report_sync.py form.pdf mybook.xlsx --in-place
"""

from __future__ import annotations

import argparse
import os
import tempfile

import pandas as pd
from pypdf import PdfReader, PdfWriter

# AcroForm partial name from Adobe / Word export (see ``PdfReader.get_fields()``).
REPORT_FIELD_NAME_PREFIX = "Report No du rapport"


def report_form_field_keys(reader: PdfReader) -> list[str]:
    fields = reader.get_fields()
    if not fields:
        return []
    return sorted(
        k for k in fields if str(k).startswith(REPORT_FIELD_NAME_PREFIX)
    )


def update_pdf_report_fields(
    pdf_in: str,
    pdf_out: str,
    report_no: str,
) -> int:
    """Set every ``Report No du rapport*`` widget to *report_no*. Returns number of fields updated.

    Raises *ValueError* if the PDF has no matching form fields.
    """
    report_no = str(report_no).strip()
    if not report_no:
        raise ValueError("Report number is empty.")

    reader = PdfReader(pdf_in, strict=False)
    keys = report_form_field_keys(reader)
    if not keys:
        raise ValueError(
            f"No AcroForm fields starting with {REPORT_FIELD_NAME_PREFIX!r} were found. "
            "This PDF may use a different template."
        )

    updates = {k: report_no for k in keys}
    writer = PdfWriter()
    writer.append(reader)
    writer.set_need_appearances_writer()
    for page in writer.pages:
        writer.update_page_form_field_values(page, updates)

    out_dir = os.path.dirname(os.path.abspath(pdf_out)) or "."
    os.makedirs(out_dir, exist_ok=True)
    with open(pdf_out, "wb") as fh:  # noqa: SIM115
        writer.write(fh)
    return len(keys)


def read_report_no_for_pdf_row(
    excel_path: str,
    pdf_path: str,
    sheet_name: str | None = None,
    file_column: str = "File",
    report_column: str = "Report No.",
) -> str:
    """Return *report_column* from the first row whose *file_column* matches this PDF (basename or path)."""
    pdf_base = os.path.basename(pdf_path)
    df = pd.read_excel(
        excel_path,
        sheet_name=sheet_name if sheet_name else 0,
        engine="openpyxl",
    )
    if file_column not in df.columns:
        raise ValueError(
            f"Column {file_column!r} not found in {excel_path!r}. "
            f"Columns: {list(df.columns)}"
        )
    if report_column not in df.columns:
        raise ValueError(
            f"Column {report_column!r} not found in {excel_path!r}. "
            f"Columns: {list(df.columns)}"
        )

    def matches(cell) -> bool:
        if pd.isna(cell):
            return False
        s = str(cell).strip()
        if not s:
            return False
        return os.path.basename(s) == pdf_base or s == pdf_path or s == pdf_base

    hit = df[df[file_column].apply(matches)]
    if hit.empty:
        raise ValueError(
            f"No Excel row where {file_column!r} matches {pdf_base!r} (full path is also accepted)."
        )
    raw = hit.iloc[0][report_column]
    if pd.isna(raw):
        raise ValueError(f"Report cell is empty for file {pdf_base!r}.")
    s = str(raw).strip()
    if not s:
        raise ValueError(f"Report cell is blank for file {pdf_base!r}.")
    return s


def sync_pdf_report_from_excel(
    excel_path: str,
    pdf_in: str,
    pdf_out: str,
    sheet_name: str | None = None,
) -> tuple[str, int]:
    """Read report number from Excel, write into *pdf_out*; returns ``(report_no, n_fields_updated)``."""
    report = read_report_no_for_pdf_row(
        excel_path, pdf_in, sheet_name=sheet_name
    )
    n = update_pdf_report_fields(pdf_in, pdf_out, report)
    return report, n


def _write_inplace(pdf_path: str, report_no: str) -> int:
    """Write via a temp file in the same directory, then replace *pdf_path*."""
    d = os.path.dirname(os.path.abspath(pdf_path)) or "."
    fd, tmp = tempfile.mkstemp(suffix=".pdf", dir=d)
    os.close(fd)
    try:
        n = update_pdf_report_fields(pdf_path, tmp, report_no)
        os.replace(tmp, pdf_path)
        return n
    except BaseException:
        try:
            os.unlink(tmp)
        except OSError:
            pass
        raise


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(
        description="Copy Report No. from an Excel row (matched by File name) into PDF form fields.",
    )
    parser.add_argument("pdf", help="Input .pdf")
    parser.add_argument(
        "excel",
        nargs="?",
        default=None,
        help=".xlsx workbook (needs File + Report No. columns); omit if you pass --report",
    )
    parser.add_argument(
        "-o",
        "--output",
        metavar="PATH",
        help="Output .pdf (default: add _report-filled before .pdf)",
    )
    parser.add_argument(
        "--in-place",
        action="store_true",
        help="Overwrite the input PDF (use a backup first)",
    )
    parser.add_argument(
        "--sheet",
        metavar="NAME",
        default=None,
        help="Worksheet name (default: first sheet)",
    )
    parser.add_argument(
        "--report",
        metavar="TEXT",
        default=None,
        help="If set, use this report number instead of reading Excel",
    )
    args = parser.parse_args(argv)

    pdf_in = args.pdf
    sheet = args.sheet.strip() if args.sheet else None

    if args.report is not None:
        report = str(args.report).strip()
        if not report:
            parser.error("--report must not be empty")
    else:
        excel = args.excel
        if not excel:
            parser.error(
                "Pass the Excel .xlsx after the PDF (e.g. pdf_report_sync.py form.pdf book.xlsx), "
                "or use --report instead of Excel"
            )
        report = read_report_no_for_pdf_row(excel, pdf_in, sheet_name=sheet)

    if args.in_place:
        n = _write_inplace(pdf_in, report)
        print(f"Updated {n} field(s) in-place → {pdf_in!r} (report {report!r})")
        return 0

    out = args.output
    if not out:
        root, ext = os.path.splitext(pdf_in)
        out = f"{root}_report-filled{ext or '.pdf'}"

    n = update_pdf_report_fields(pdf_in, out, report)
    print(f"Updated {n} field(s) → {out!r} (report {report!r})")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
