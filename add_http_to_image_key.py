"""
Image Key -> Full URL Adder
----------------------------
Ei script ta Excel file er "Image 1", "Image 2", ... eরকম header thaka
sob column e ghure ghure, jei cell e image key ache tar age tumi deya
base URL ta bosiye dibe. Jodi kono cell already http/https diye shuru
hoy, oita skip kore jabe (double add hobe na).

Kivabe run korbe (cmd/terminal e):
    python add_image_url.py

Tারপর ও tomake 2 ta jinis চাইবে:
    1) Excel file er full path (e.g. C:\\Users\\Limon\\Desktop\\catalogue.xlsx)
    2) Base URL (e.g. https://sl-dev-s3.s3.amazonaws.com/product/)

Output: original file thaka jei folder e ache, oi shei folder e
"image_http_url_add.xlsx" name e notun file banabe (original file
touch kore na, tai risk nai). Notun file e 2 ta tab thakbe:
    1) Data tab(s) - jekhane URL bosano hoyeche (original sheet name
       shei rakha hoyeche)
    2) "Report" tab - choto summary report (kotogulo cell update
       hoyeche, kotogulo age theke URL chilo, kon sheet e ki hoyeche)
"""

import os
import sys
from datetime import datetime

try:
    from openpyxl import load_workbook
    from openpyxl.styles import Font, PatternFill
except ImportError:
    print("openpyxl install kora nai. Age ei command ta run koro:")
    print("    pip install openpyxl")
    sys.exit(1)


def process_file(input_path: str, base_url: str, output_dir: str = None) -> str:
    if not os.path.isfile(input_path):
        raise FileNotFoundError(f"File pawa jayni: {input_path}")

    if not base_url.endswith("/"):
        base_url += "/"

    wb = load_workbook(input_path)  # formatting preserve kore

    total_updated = 0
    total_already_url = 0
    total_blank = 0
    sheet_reports = []  # (sheet_name, image_columns_count, updated, already_url, blank)

    for sheet in wb.worksheets:
        max_row = sheet.max_row
        max_col = sheet.max_column
        if max_row < 2:
            continue  # data row nai

        # Header row (row 1) theke বের korbo kon kon column "Image" diye শুরু
        image_columns = []
        for col_idx in range(1, max_col + 1):
            header_val = sheet.cell(row=1, column=col_idx).value
            if header_val and "image" in str(header_val).strip().lower():
                image_columns.append(col_idx)

        if not image_columns:
            continue

        sheet_updated = 0
        sheet_already_url = 0
        sheet_blank = 0

        # Row 2 theke last row porjonto, প্রতিটা image column er cell update korbo
        for row_idx in range(2, max_row + 1):
            for col_idx in image_columns:
                cell = sheet.cell(row=row_idx, column=col_idx)
                value = cell.value
                if value is None:
                    sheet_blank += 1
                    continue
                value_str = str(value).strip()
                if not value_str:
                    sheet_blank += 1
                    continue
                if value_str.lower().startswith("http://") or value_str.lower().startswith("https://"):
                    sheet_already_url += 1
                    continue  # already full URL, skip
                cell.value = base_url + value_str
                sheet_updated += 1

        total_updated += sheet_updated
        total_already_url += sheet_already_url
        total_blank += sheet_blank
        sheet_reports.append((sheet.title, len(image_columns), sheet_updated, sheet_already_url, sheet_blank))

    # ---- Report tab banano ----
    report = wb.create_sheet("Report")

    bold = Font(bold=True)
    header_fill = PatternFill(start_color="12B5E5", end_color="12B5E5", fill_type="solid")
    header_font = Font(bold=True, color="FFFFFF")

    report["A1"] = "Image URL Add - Summary Report"
    report["A1"].font = Font(bold=True, size=14)

    report["A3"] = "Source file:"
    report["B3"] = os.path.basename(input_path)
    report["A4"] = "Base URL used:"
    report["B4"] = base_url
    report["A5"] = "Run time:"
    report["B5"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    report["A7"] = "TOTAL cells updated:"
    report["B7"] = total_updated
    report["A8"] = "TOTAL already had URL (skipped):"
    report["B8"] = total_already_url
    report["A9"] = "TOTAL blank cells (skipped):"
    report["B9"] = total_blank
    for r in range(7, 10):
        report.cell(row=r, column=1).font = bold

    # Per-sheet breakdown table
    start_row = 11
    headers = ["Sheet Name", "Image Columns Found", "Cells Updated", "Already URL", "Blank"]
    for col_idx, h in enumerate(headers, start=1):
        c = report.cell(row=start_row, column=col_idx, value=h)
        c.font = header_font
        c.fill = header_fill

    for i, (sheet_name, img_col_count, updated, already, blank) in enumerate(sheet_reports, start=1):
        row = start_row + i
        report.cell(row=row, column=1, value=sheet_name)
        report.cell(row=row, column=2, value=img_col_count)
        report.cell(row=row, column=3, value=updated)
        report.cell(row=row, column=4, value=already)
        report.cell(row=row, column=5, value=blank)

    # Column width thik kora
    report.column_dimensions["A"].width = 32
    report.column_dimensions["B"].width = 28
    report.column_dimensions["C"].width = 16
    report.column_dimensions["D"].width = 14
    report.column_dimensions["E"].width = 10

    if output_dir is None:
        output_dir = os.path.dirname(os.path.abspath(input_path))
    else:
        os.makedirs(output_dir, exist_ok=True)
    output_path = os.path.join(output_dir, "image_http_url_add.xlsx")
    wb.save(output_path)

    print(f"\nDone! Total {len(sheet_reports)} data sheet(s) e Image column pawa geche.")
    print(f"Total {total_updated} ta cell update kora hoyeche.")
    print(f"Notun file save hoyeche ekhane:\n    {output_path}\n")
    return output_path


if __name__ == "__main__":
    print("=== Image Key -> Full URL Adder ===\n")
    excel_path = input("Excel file er full path dao: ").strip().strip('"')
    url = input("Base URL dao (e.g. https://sl-dev-s3.s3.amazonaws.com/product/): ").strip()

    try:
        process_file(excel_path, url)
    except Exception as e:
        print(f"\nError hoyeche: {e}")
        sys.exit(1)
