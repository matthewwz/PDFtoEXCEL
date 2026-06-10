# PDF to Excel

Desktop app that reads **the first page** of specific bilingual (EN/FR) PDF reports for header fields, scans **all pages** for batch size and serial/label tables, and appends **one row per PDF** to an Excel workbook.

This is **not** a general-purpose PDF table extractor; it expects Intertek-style bilingual layouts. The older **Submitter / Réquérant** header and the newer **Company | Entreprise** pipe-style form (including Word→PDF **AcroForm** fills where values are not in the text layer) are both supported. See `fileTest*.pdf` (legacy) and newer `SI#…` samples.

**Columns written** depend on the workbook:

- **Existing master list** (any workbook whose first sheet does **not** have `File` in cell A1): finds the **first row where C and D are both empty and neither has a background fill** (so highlighted or banded cells are skipped); then writes **Report No.** in **A** when **A** is empty (otherwise leaves **A** unchanged), **Submitter** in **C**, **Model** in **D**, **E** = numeric part of the label after stripping a leading **C-** (e.g. `2307970`), **F** = **E** + batch size − 1 (same as **E** when batch size is 1; e.g. batch 3 → **F** = `2307972`), **Batch Size** in **J**, PDF **file name** in **K**, and **removes fill** on A/C/D for those writes. If every scanned row is in use or has fill on C or D, a **new row** is added after the last row. Column **B** is set only when you pass a date (CLI `--column-b-date "…"` or the optional field in the GUI); otherwise **B** is left as-is on that row (typically blank). Use **`--sheet`** / the **Sheet name** field to choose the worksheet; if omitted, the workbook’s **active** sheet (last selected tab in Excel) is used.
- **Create New Excel File** in the GUI (or the first time *Process* writes to a path that does not exist yet): creates the **same master-list layout** as selecting an existing master workbook — row **1** headers **A** Report No., **B** Date, **C** Submitter, **D** Model, **E**/**F** serial range, **J** Batch Size, **K** File (PDF base name). **Process** then fills rows the same way as for any other master list (including **Report No.** in **A** when **A** is empty).
- **Legacy export workbook** (the chosen sheet’s cell **A1** is exactly **`File`**): wide table — File, Date, Model, Description, Submitter, Report No., Label, Batch Size. Use this only if you still use that older export format. You can target a specific sheet with the same **Sheet name** option; other sheets are left in place when appending.

## Setup

Use Python 3.10+ (3.12 tested).

```bash
cd /path/to/PDFtoExcel
python3 -m venv .venv
source .venv/bin/activate   # Windows: .venv\Scripts\activate
pip install -r requirements.txt
```

## Run (from source)

```bash
python gui.py
```

The GUI uses a **PDF queue**: add files, reorder with **Move up** / **Move down**, then **Process queue** to write rows in that order. Successful files are removed from the queue; **failed files stay** in the same order so you can fix the issue and run **Process queue** again. Use **Remove from queue** (with rows selected) or **Clear queue** to drop files without processing.

**Write Report No. into the PDF:** use **Write Report No. from Excel into PDF form…** (or the CLI below). The workbook needs columns **`File`** (PDF base name, e.g. column **K** on the master list) and **`Report No.`** (e.g. column **A**). The tool finds the row whose **File** matches the PDF’s file name and copies **Report No.** into every AcroForm field whose name starts with **`Report No du rapport`**. You always save a **new** PDF; the original is not modified unless you use the CLI **`--in-place`**.

## Saving while Excel has the workbook open

Microsoft Excel and macOS **Preview** often **lock** the `.xlsx` file, which blocks other programs from saving changes. This app **retries** the save several times; if it still cannot write, **close that workbook** (or quit Excel), run the tool again, then reopen the file. Excel does **not** live-reload rows written by another program—you need to reopen the workbook to see updates.

The app does **not** automatically close Excel for you (doing so could lose unsaved edits in any open workbook).

## Windows without Python (standalone `.exe`)

End users can run a single executable with no Python or `pip` on their PC. The **GitHub Actions** workflow builds it with **PyInstaller**, which embeds Python and libraries.

**Build on GitHub**

1. Push this repo to GitHub.
2. **Actions** → **“Build Windows executable”** → **Run workflow**.
3. Download the artifact **PDFtoExcel-windows-exe** (`PDFtoExcel.exe`).

First launch may be slow while Windows scans the file; some antivirus tools flag unsigned PyInstaller binaries (code signing helps).

## CLI / batch (optional)

From the project directory; writes `output.xlsx` using `fileTest*.pdf` in the same folder:

```bash
python extractor.py
python extractor.py --column-b-date "June 5th , 2026"
python extractor.py --sheet "Data"
python extractor.py --sheet "Data" --column-b-date "June 5th , 2026"
```

**Write Report No. into a PDF** (`pdf_report_sync.py`): needs **`File`** + **`Report No.`** in the workbook (unless you pass **`--report`** to set the value explicitly). Writes all AcroForm fields named **`Report No du rapport*`** and saves a copy (or **`--in-place`**).

```bash
python pdf_report_sync.py "SI#2026-WZ-1084 smart flex Lighting.pdf" mybook.xlsx -o "SI#2026-WZ-1084-updated.pdf"
python pdf_report_sync.py form.pdf mybook.xlsx --sheet "Sheet1"
python pdf_report_sync.py form.pdf --report "2026-WZ-1084" -o out.pdf
```

## Dependencies

- [pdfplumber](https://github.com/jsvine/pdfplumber) — PDF text
- [pypdf](https://github.com/py-pdf/pypdf) — PDF form field updates
- [pandas](https://pandas.pydata.org/) + [openpyxl](https://openpyxl.readthedocs.io/) — `.xlsx`
