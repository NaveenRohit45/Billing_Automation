# billing_logic.py
import os
import fitz  # PyMuPDF
import pandas as pd
import tempfile
import shutil
from docx2pdf import convert
from PyPDF2 import PdfReader


# -------------------------------------------------
# DOCX PAGE COUNT
# -------------------------------------------------
def get_docx_page_count(docx_path):
    try:
        temp_dir = tempfile.mkdtemp()
        temp_pdf = os.path.join(temp_dir, "temp.pdf")

        convert(docx_path, temp_pdf)
        pages = len(PdfReader(temp_pdf).pages)

        shutil.rmtree(temp_dir)
        return pages
    except Exception:
        return "NA"


# -------------------------------------------------
# FIND ORIGINAL FOLDER (ROBUST)
# -------------------------------------------------
def find_original_folder(job_path):
    """
    Finds folder ending with 'original'
    Supports: Original, 1.Original, 01.Original, 1. Original
    """
    for name in os.listdir(job_path):
        full = os.path.join(job_path, name)
        if os.path.isdir(full) and name.lower().replace(" ", "").endswith("original"):
            return full
    return None


# -------------------------------------------------
# PROCESS FILES INSIDE ONE EDITS FOLDER
# -------------------------------------------------
def process_folder(folder_path, job_name, edit_folder):
    summary_rows = []
    detail_rows = []

    for file in os.listdir(folder_path):
        full = os.path.join(folder_path, file)
        if not os.path.isfile(full):
            continue

        # ---------- PDF ----------
        if file.lower().endswith(".pdf"):
            doc = fitz.open(full)
            total_comments = 0
            pages_with_comments = set()

            for page_index in range(len(doc)):
                annots = doc[page_index].annots()
                if annots:
                    for annot in annots:
                        total_comments += 1
                        pages_with_comments.add(page_index + 1)

                        detail_rows.append({
                            "Job Name": job_name,
                            "Edit Folder": edit_folder,
                            "File Name": file,
                            "File Type": "PDF",
                            "Page Number": page_index + 1,
                            "Comment Type": annot.type[1] if annot.type else "Unknown",
                            "Author": annot.info.get("title", ""),
                            "Comment Text": annot.info.get("content", "").strip()
                        })

            summary_rows.append({
                "Job Name": job_name,
                "Edit Folder": edit_folder,
                "File Name": file,
                "File Type": "PDF",
                "Total Pages": len(doc),
                "Pages with Comments": len(pages_with_comments),
                "Total Comments": total_comments
            })

        # ---------- DOCX ----------
        elif file.lower().endswith(".docx"):
            pages = get_docx_page_count(full)
            summary_rows.append({
                "Job Name": job_name,
                "Edit Folder": edit_folder,
                "File Name": file,
                "File Type": "DOCX",
                "Total Pages": pages,
                "Pages with Comments": "NA",
                "Total Comments": "NA"
            })

    return summary_rows, detail_rows


# -------------------------------------------------
# SINGLE FOLDER MODE
# -------------------------------------------------
def run_single_mode(folder_path):
    return process_folder(
        folder_path,
        job_name="SingleJob",
        edit_folder="NA"
    )


# -------------------------------------------------
# MULTI JOB (BATCH) MODE
# -------------------------------------------------
def run_batch_mode(root_folder):
    all_summary = []
    all_details = []

    for job in os.listdir(root_folder):
        job_path = os.path.join(root_folder, job)
        if not os.path.isdir(job_path):
            continue

        original = find_original_folder(job_path)
        if not original:
            print(f"⚠️ Skipping {job} (Original folder not found)")
            continue

        for sub in os.listdir(original):
            if not sub.startswith("Edits_"):
                continue

            edits_path = os.path.join(original, sub)
            if not os.path.isdir(edits_path):
                continue

            print(f"✅ Processing: {job} → {sub}")

            summary, details = process_folder(
                edits_path,
                job_name=job,
                edit_folder=sub
            )

            all_summary.extend(summary)
            all_details.extend(details)

    return all_summary, all_details


# -------------------------------------------------
# MASTER EXCEL OUTPUT
# -------------------------------------------------
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill


def generate_master_excel(summary_rows, detail_rows, output_folder):
    output = os.path.join(output_folder, "Master_Billing_Report.xlsx")

    wb = Workbook()
    ws = wb.active
    ws.title = "Summary"

    headers = [
        "Job Name",
        "Edit Folder",
        "File Name",
        "File Type",
        "Total Pages",
        "Pages with Comments",
        "Total Comments"
    ]
    ws.append(headers)

    header_fill = PatternFill("solid", fgColor="DDDDDD")
    for col in range(1, len(headers) + 1):
        ws.cell(row=1, column=col).fill = header_fill
        ws.cell(row=1, column=col).font = Font(bold=True)

    # ---------------- GROUP BY JOB ----------------
    jobs = {}
    for row in summary_rows:
        jobs.setdefault(row["Job Name"], []).append(row)

    yellow_fill = PatternFill("solid", fgColor="FFFF00")
    bold_font = Font(bold=True)

    for job_name, rows in jobs.items():
        # ===== JOB HEADER =====
        ws.append([f"JOB NAME ---- {job_name}"])
        r = ws.max_row
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=7)
        ws.cell(row=r, column=1).fill = yellow_fill
        ws.cell(row=r, column=1).font = bold_font

        total_docx_pages = 0
        total_pdf_pages_with_comments = 0
        total_pdf_comments = 0

        # ===== FILE ROWS =====
        for rdata in rows:
            ws.append([
                rdata["Job Name"],
                rdata["Edit Folder"],
                rdata["File Name"],
                rdata["File Type"],
                rdata["Total Pages"],
                rdata["Pages with Comments"],
                rdata["Total Comments"]
            ])

            if rdata["File Type"] == "DOCX" and rdata["Total Pages"] != "NA":
                total_docx_pages += int(rdata["Total Pages"])

            if rdata["File Type"] == "PDF":
                if rdata["Pages with Comments"] != "NA":
                    total_pdf_pages_with_comments += int(rdata["Pages with Comments"])
                if rdata["Total Comments"] != "NA":
                    total_pdf_comments += int(rdata["Total Comments"])

        # ===== TOTAL ROW =====
        ws.append([
            "Total",
            "",
            "",
            "DOCX pages",
            total_docx_pages,
            total_pdf_pages_with_comments,
            total_pdf_comments
        ])

        total_row = ws.max_row
        for c in range(1, 8):
            ws.cell(row=total_row, column=c).font = bold_font

        ws.append([])  # blank line between jobs

    # ---------------- PDF COMMENTS SHEET ----------------
    if detail_rows:
        ws2 = wb.create_sheet("PDF_Comments")
        ws2.append(list(detail_rows[0].keys()))
        for row in detail_rows:
            ws2.append(list(row.values()))

    wb.save(output)
    return output


