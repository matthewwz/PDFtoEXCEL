import os
import tkinter as tk
from tkinter import filedialog, messagebox

from extractor import create_master_list_workbook, process_pdf, update_excel
from pdf_report_sync import sync_pdf_report_from_excel

excel_file_path = None
pdf_queue: list[str] = []
pdf_listbox: tk.Listbox | None = None


def refresh_pdf_listbox() -> None:
    if pdf_listbox is None:
        return
    pdf_listbox.delete(0, tk.END)
    for i, p in enumerate(pdf_queue):
        pdf_listbox.insert(tk.END, f"{i + 1}. {os.path.basename(p)}")


def select_excel_file():
    global excel_file_path
    if excel_file_path is not None:
        messagebox.showwarning(
            "Warning", "An Excel file has already been selected or created!"
        )
        return

    temp_path = filedialog.askopenfilename(
        title="Select Excel file",
        filetypes=[("Excel files", "*.xlsx")],
        defaultextension=".xlsx",
    )
    if temp_path:
        excel_file_path = temp_path
        messagebox.showinfo("Success", f"Excel file selected: {excel_file_path}")
    else:
        messagebox.showinfo("Cancelled", "No file was selected.")


def write_report_no_into_pdf_form():
    """Excel row (File + Report No.) → PDF AcroForm *Report No du rapport* fields → save as new PDF."""
    xlsx = excel_file_path
    if not xlsx:
        xlsx = filedialog.askopenfilename(
            title="Excel workbook (needs File + Report No. columns)",
            filetypes=[("Excel files", "*.xlsx")],
            defaultextension=".xlsx",
        )
    if not xlsx:
        return
    pdf_in = filedialog.askopenfilename(
        title="PDF to update (top-right report # form fields)",
        filetypes=[("PDF files", "*.pdf")],
        defaultextension=".pdf",
    )
    if not pdf_in:
        return
    base = os.path.splitext(os.path.basename(pdf_in))[0]
    pdf_out = filedialog.asksaveasfilename(
        title="Save updated PDF as…",
        filetypes=[("PDF files", "*.pdf")],
        defaultextension=".pdf",
        initialfile=f"{base}_report-filled.pdf",
    )
    if not pdf_out:
        return
    raw_sheet = sheet_name_var.get().strip()
    sheet = raw_sheet if raw_sheet else None
    try:
        report, n = sync_pdf_report_from_excel(
            xlsx, pdf_in, pdf_out, sheet_name=sheet
        )
    except Exception as e:
        messagebox.showerror("Report → PDF failed", str(e))
        return
    messagebox.showinfo(
        "Report → PDF done",
        f"Wrote report {report!r} into {n} form field(s).\n\nSaved to:\n{pdf_out}",
    )


def create_excel_file():
    global excel_file_path
    if excel_file_path is not None:
        messagebox.showwarning(
            "Warning", "An Excel file has already been selected or created!"
        )
        return

    temp_path = filedialog.asksaveasfilename(
        title="Create New Excel File",
        filetypes=[("Excel files", "*.xlsx")],
        defaultextension=".xlsx",
    )
    if temp_path:
        excel_file_path = temp_path
        create_master_list_workbook(excel_file_path)
        messagebox.showinfo(
            "Success",
            "New master-list workbook created (columns A–F, J–K row 1 headers).\n"
            f"{excel_file_path}",
        )
    else:
        messagebox.showinfo("Cancelled", "File creation was cancelled.")


def add_pdfs_to_queue():
    paths = filedialog.askopenfilenames(
        title="Add PDF files to the queue", filetypes=[("PDF files", "*.pdf")]
    )
    if not paths:
        return
    added = 0
    for p in paths:
        if not p.lower().endswith(".pdf"):
            continue
        if p not in pdf_queue:
            pdf_queue.append(p)
            added += 1
    if added == 0:
        messagebox.showinfo(
            "Queue unchanged",
            "No new PDF files were added (empty selection, non-PDF paths, or duplicates).",
        )
    refresh_pdf_listbox()


def remove_selected_from_queue():
    if pdf_listbox is None:
        return
    sel = list(pdf_listbox.curselection())
    if not sel:
        messagebox.showinfo("Remove from queue", "Select one or more rows in the list, then click Remove from queue.")
        return
    for i in sorted(sel, reverse=True):
        if 0 <= i < len(pdf_queue):
            pdf_queue.pop(i)
    refresh_pdf_listbox()


def move_pdf_up():
    if pdf_listbox is None:
        return
    sel = pdf_listbox.curselection()
    if len(sel) != 1:
        messagebox.showinfo("Reorder", "Select exactly one row to move up.")
        return
    i = sel[0]
    if i <= 0:
        return
    pdf_queue[i - 1], pdf_queue[i] = pdf_queue[i], pdf_queue[i - 1]
    refresh_pdf_listbox()
    pdf_listbox.selection_set(i - 1)


def move_pdf_down():
    if pdf_listbox is None:
        return
    sel = pdf_listbox.curselection()
    if len(sel) != 1:
        messagebox.showinfo("Reorder", "Select exactly one row to move down.")
        return
    i = sel[0]
    if i >= len(pdf_queue) - 1:
        return
    pdf_queue[i], pdf_queue[i + 1] = pdf_queue[i + 1], pdf_queue[i]
    refresh_pdf_listbox()
    pdf_listbox.selection_set(i + 1)


def clear_pdf_queue():
    if not pdf_queue:
        return
    if not messagebox.askyesno("Clear queue", "Remove all PDFs from the queue?"):
        return
    pdf_queue.clear()
    refresh_pdf_listbox()


def process_queue():
    global excel_file_path
    if excel_file_path is None:
        messagebox.showwarning(
            "Warning", "Please select or create an Excel file first."
        )
        return
    if not pdf_queue:
        messagebox.showwarning(
            "Nothing to process", "Add one or more PDF files to the queue first."
        )
        return

    batch = list(pdf_queue)
    ok: list[str] = []
    failed: list[tuple[str, str]] = []

    raw_b = column_b_date_var.get().strip()
    b_date = raw_b if raw_b else None
    raw_sheet = sheet_name_var.get().strip()
    sheet = raw_sheet if raw_sheet else None

    for file_path in batch:
        try:
            new_data = process_pdf(file_path)
            update_excel(
                new_data,
                excel_file_path,
                column_b_date=b_date,
                sheet_name=sheet,
            )
            ok.append(file_path)
        except Exception as e:
            failed.append((file_path, str(e)))

    lines: list[str] = []
    if ok:
        lines.append(f"Successfully processed {len(ok)} file(s) (in queue order):")
        lines.extend(f"  • {os.path.basename(p)}" for p in ok)
    if failed:
        lines.append("")
        lines.append(
            f"Failed ({len(failed)}) — left in the queue so you can fix and retry:"
        )
        for p, err in failed[:12]:
            lines.append(f"  • {os.path.basename(p)}")
            lines.append(f"    {err}")
        if len(failed) > 12:
            lines.append(f"  … and {len(failed) - 12} more failure(s)")

    # Drop successes; keep only failures in the same order for another run
    pdf_queue.clear()
    pdf_queue.extend(p for p, _ in failed)
    refresh_pdf_listbox()

    if not lines:
        messagebox.showinfo("Processing complete", "No files were processed.")
        return

    body = "\n".join(lines)
    title = "Processing complete"
    if failed and not ok:
        messagebox.showerror(title, body)
    elif failed:
        messagebox.showwarning(title, body)
    else:
        messagebox.showinfo(title, body)


root = tk.Tk()
root.title("PDF to Excel")
# Tall default + minimum so list, Process button, and help text fit on first open (esp. macOS).
root.geometry("720x780")
root.minsize(640, 660)

column_b_date_var = tk.StringVar()
sheet_name_var = tk.StringVar()

main = tk.Frame(root)
main.pack(fill=tk.BOTH, expand=True, padx=12, pady=10)

tk.Button(main, text="Select Excel File", command=select_excel_file).pack(pady=(0, 4))
tk.Button(main, text="Create New Excel File", command=create_excel_file).pack(pady=(0, 4))
tk.Button(
    main,
    text="Write Report No. from Excel into PDF form…",
    command=write_report_no_into_pdf_form,
).pack(pady=(0, 8))

bf = tk.Frame(main)
bf.pack(pady=(0, 6), fill=tk.X)
tk.Label(bf, text="Date for column B (master sheet only, optional):").pack(anchor=tk.W)
tk.Entry(bf, textvariable=column_b_date_var, width=62).pack(anchor=tk.W, fill=tk.X)

sf = tk.Frame(main)
sf.pack(pady=(0, 6), fill=tk.X)
tk.Label(sf, text="Sheet name (optional; blank = active sheet in Excel):").pack(anchor=tk.W)
tk.Entry(sf, textvariable=sheet_name_var, width=62).pack(anchor=tk.W, fill=tk.X)

queue_lab = tk.Label(main, text="PDF queue (top = processed first):")
queue_lab.pack(anchor=tk.W, pady=(4, 4))

btn_row = tk.Frame(main)
btn_row.pack(fill=tk.X, pady=(0, 6))
tk.Button(btn_row, text="Add PDFs…", command=add_pdfs_to_queue).pack(side=tk.LEFT, padx=(0, 8))
tk.Button(btn_row, text="Move up", command=move_pdf_up).pack(side=tk.LEFT, padx=8)
tk.Button(btn_row, text="Move down", command=move_pdf_down).pack(side=tk.LEFT, padx=8)
tk.Button(btn_row, text="Remove", command=remove_selected_from_queue).pack(side=tk.LEFT, padx=8)
tk.Button(btn_row, text="Clear", command=clear_pdf_queue).pack(side=tk.LEFT, padx=8)

list_frame = tk.Frame(main)
list_frame.pack(pady=(0, 8), fill=tk.BOTH, expand=True)
scroll = tk.Scrollbar(list_frame)
scroll.pack(side=tk.RIGHT, fill=tk.Y)
pdf_listbox = tk.Listbox(
    list_frame,
    height=14,
    selectmode=tk.EXTENDED,
    yscrollcommand=scroll.set,
    font=("TkDefaultFont", 11),
)
pdf_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
scroll.config(command=pdf_listbox.yview)

tk.Button(
    main,
    text="Process",
    command=process_queue,
).pack(pady=(0, 8), fill=tk.X)

tk.Label(
    main,
    text=(
        "1. Select an Excel file, or create a new master-list workbook (same A–K header layout as an existing master sheet)\n"
        "2. Add PDFs, reorder with Move up/down, use Remove from queue to drop rows\n"
        "3. Process queue — successes are removed; failures stay in the queue to retry\n"
        "4. Optional: “Write Report No. from Excel into PDF form…” — needs File + Report No. columns; "
        "matches the PDF by file name; saves a new PDF (AcroForm templates only)"
    ),
    justify=tk.LEFT,
    wraplength=680,
).pack(anchor=tk.W, fill=tk.X)

root.mainloop()
