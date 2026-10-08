"""
Merge Mapped Image URLs into HTML Columns
------------------------------------------
Ei script 2 ta Excel file lagbe:

  1) MAIN FILE - jekhane onek column ache (Product ID, Name, Description,
     Highlights, etc). Ei file er 4 ta column e HTML content ache, ar
     shei HTML er ভিতরে <img src="..."> tag diye purono external image
     (Daraz cdn etc.) attach kora:
         - Highlights (English)
         - Highlights (Bengali)
         - Description (Bengali)
         - Description (English)

     NOTE: Column heading exact-match na hoile o (jemon
     "Highlights(English)" vs "Highlights (English)" - space thakuk
     ba na thakuk, capital/small jekono hok) script ekhon automatic
     match kore nibe. Manual ar kichu change korte hobe na.

  2) MAPPING FILE - jeta "add_image_url.py" script diye age generate
     kora hoyeche (e.g. image_http_url_add.xlsx). Ei file e "Product ID"
     ar "Image 1", "Image 2"... column gulo te already-complete URL
     ache.

Ei script ki kore:
  - Product ID diye match kore, MAIN FILE er oi 4 ta column er HTML er
    ভিতরে jekono <img src="..."> pabe, segulo (order onujayi: 1st img
    tag -> Image 1, 2nd img tag -> Image 2, ...) MAPPING FILE er URL
    diye replace kore dibe.
  - Jodi kono product er HTML e img tag beshi thake mapped image
    thaka, tahole ba-ki tag gulo (jegulor jonno notun image nai)
    OGE-ROKOM ROOF thakbe (change hobe na) - data loss hobe na.
  - Jodi kono Product ID mapping file e na pawa jay, oi row er 4 ta
    column touch kora hobe na.
  - Report tab e shob kichur summary thakbe.

Kivabe run korbe:
    python merge_images_into_html.py

Tারপর script tomake 2 ta path chaibe:
    1) Main file er path
    2) Mapping file er path (image_http_url_add.xlsx type file)

Output: main file jei folder e ache oi shei folder e
"<mainfile>_image_replaced.xlsx" name e notun file banabe.
"""

import os
import re
import sys
from datetime import datetime

try:
    from openpyxl import load_workbook
    from openpyxl.styles import Font, PatternFill
except ImportError:
    print("openpyxl install kora nai. Age ei command ta run koro:")
    print("    pip install openpyxl")
    sys.exit(1)


# Ei list e "canonical" naam gula ache. Actual file er header ei naam
# gular shathe EXACT match na holeo cholbe - normalize_header() function
# space/parenthesis/case shob ignore kore match kore dey.
TARGET_COLUMNS = [
    "Highlights (English)",
    "Highlights (Bengali)",
    "Description (Bengali)",
    "Description (English)",
]

PRODUCT_ID_NAMES = ["Product ID", "ProductID", "Product Id"]

IMG_SRC_PATTERN = re.compile(r'(<img[^>]*\ssrc=)(["\'])(.*?)\2', re.IGNORECASE)


def normalize_header(value):
    """Header string ke ekta 'canonical' form e niye ashe, jate
    space thakuk ba na thakuk, capital/small jekono hok, extra
    punctuation thakuk - shob e match kore. Only a-z0-9 rakha hoy.

    Example: "Highlights(English)" -> "highlightsenglish"
             "Highlights (English)" -> "highlightsenglish"
             "  highlights  english " -> "highlightsenglish"
    """
    if value is None:
        return ""
    s = str(value).strip().lower()
    s = re.sub(r"[^a-z0-9]", "", s)
    return s


def build_header_index(headers):
    """headers list theke {normalized_name: column_number} dict banay.
    Column number 1-indexed."""
    idx = {}
    for i, h in enumerate(headers):
        if h is None:
            continue
        key = normalize_header(h)
        if key and key not in idx:  # prothom match ta rakhbo
            idx[key] = i + 1
    return idx


def find_column(norm_idx, candidate_names):
    """candidate_names list er moddhe first ja match kore, oita
    column number return korbe. Na paile None."""
    for name in candidate_names:
        c = norm_idx.get(normalize_header(name))
        if c:
            return c
    return None


def normalize_pid(value):
    """Product ID ke ekta consistent string e convert kore, jate
    1064210 ar 1064210.0 ar '1064210' shob match kore."""
    if value is None:
        return None
    s = str(value).strip()
    if s.endswith(".0"):
        s = s[:-2]
    return s


def build_mapping(mapping_path: str):
    wb = load_workbook(mapping_path, data_only=True)

    # "Report" naam er sheet bad diye, jei sheet e "Product ID" +
    # "Image" column ache oita khuje ber korbo
    target_sheet = None
    norm_idx = None
    for sheet in wb.worksheets:
        if sheet.title.strip().lower() == "report":
            continue
        headers = [c.value for c in sheet[1]]
        idx = build_header_index(headers)
        if find_column(idx, PRODUCT_ID_NAMES):
            target_sheet = sheet
            norm_idx = idx
            break

    if target_sheet is None:
        raise ValueError("Mapping file e 'Product ID' column pawa jayni.")

    pid_col = find_column(norm_idx, PRODUCT_ID_NAMES)

    headers = [c.value for c in target_sheet[1]]
    img_cols = []
    for i, h in enumerate(headers):
        if h and "image" in normalize_header(h):
            img_cols.append(i + 1)
    img_cols.sort()

    pid_to_urls = {}
    for row in range(2, target_sheet.max_row + 1):
        pid = normalize_pid(target_sheet.cell(row=row, column=pid_col).value)
        if pid is None:
            continue
        urls = []
        for c in img_cols:
            v = target_sheet.cell(row=row, column=c).value
            if v:
                urls.append(str(v).strip())
        if urls:
            pid_to_urls[pid] = urls

    return pid_to_urls


def replace_images_in_html(html: str, urls: list):
    """HTML er প্রতিটা <img src="..."> ke order onujayi mapping URL
    diye replace kore. Return: (new_html, replaced_count, total_tags)."""
    state = {"pos": 0, "replaced": 0, "total": 0}

    def _sub(m):
        state["total"] += 1
        i = state["pos"]
        if i < len(urls):
            state["pos"] += 1
            state["replaced"] += 1
            return f"{m.group(1)}{m.group(2)}{urls[i]}{m.group(2)}"
        return m.group(0)

    new_html = IMG_SRC_PATTERN.sub(_sub, html)
    return new_html, state["replaced"], state["total"]


def process(main_path: str, mapping_path: str, output_dir: str = None) -> str:
    if not os.path.isfile(main_path):
        raise FileNotFoundError(f"Main file pawa jayni: {main_path}")
    if not os.path.isfile(mapping_path):
        raise FileNotFoundError(f"Mapping file pawa jayni: {mapping_path}")

    pid_to_urls = build_mapping(mapping_path)

    wb = load_workbook(main_path)
    ws = wb.active  # main data sheet dhore nicchi active sheet

    headers = [c.value for c in ws[1]]
    norm_idx = build_header_index(headers)

    pid_col = find_column(norm_idx, PRODUCT_ID_NAMES)
    if pid_col is None:
        raise ValueError("Main file e 'Product ID' column pawa jayni.")

    found_cols = {}
    for col_name in TARGET_COLUMNS:
        c = norm_idx.get(normalize_header(col_name))
        if c:
            found_cols[col_name] = c
        else:
            print(f"WARNING: Main file e '{col_name}' er moto column pawa jayni, skip kora hocche.")

    total_rows = 0
    rows_no_mapping = 0
    tags_replaced_total = 0
    tags_found_total = 0
    rows_partial = 0  # img tag beshi, mapped image kom

    # row_status[row_number] = "no_mapping" | "partial" | (matched hole key thakbe na)
    row_status = {}

    for row in range(2, ws.max_row + 1):
        pid_val = ws.cell(row=row, column=pid_col).value
        if pid_val is None:
            continue
        total_rows += 1
        pid = normalize_pid(pid_val)
        urls = pid_to_urls.get(pid)

        if not urls:
            rows_no_mapping += 1
            row_status[row] = "No Mapping Found"
            continue

        row_had_shortfall = False
        for col_name, col_idx in found_cols.items():
            cell = ws.cell(row=row, column=col_idx)
            html = cell.value
            if not html:
                continue
            new_html, replaced, total_tags = replace_images_in_html(str(html), urls)
            if total_tags > 0:
                tags_found_total += total_tags
                tags_replaced_total += replaced
                if replaced < total_tags:
                    row_had_shortfall = True
                if new_html != html:
                    cell.value = new_html

        if row_had_shortfall:
            rows_partial += 1
            row_status[row] = "Insufficient Mapped Images"

    # ---- Unmatched / partial row gulo Sheet1 theke সরিয়ে "Needs Review" sheet e নেওয়া ----
    max_col = ws.max_column
    headers_full = [ws.cell(row=1, column=c).value for c in range(1, max_col + 1)]

    moved_rows = []  # (issue_label, [cell values...])
    flagged_rows_sorted = sorted(row_status.keys(), reverse=True)  # bottom theke upore delete korte hobe
    for row in flagged_rows_sorted:
        values = [ws.cell(row=row, column=c).value for c in range(1, max_col + 1)]
        moved_rows.append((row_status[row], values))
        ws.delete_rows(row, 1)

    moved_rows.reverse()  # abar original top-to-bottom order e ana

    review_ws = wb.create_sheet("Needs Review")
    header_fill = PatternFill(start_color="ED0A5E", end_color="ED0A5E", fill_type="solid")
    header_font = Font(bold=True, color="FFFFFF")
    for c, h in enumerate(headers_full, start=1):
        cell = review_ws.cell(row=1, column=c, value=h)
        cell.font = header_font
        cell.fill = header_fill
    issue_col = max_col + 1
    issue_header = review_ws.cell(row=1, column=issue_col, value="Issue")
    issue_header.font = header_font
    issue_header.fill = header_fill

    for r_offset, (issue_label, values) in enumerate(moved_rows, start=2):
        for c, v in enumerate(values, start=1):
            review_ws.cell(row=r_offset, column=c, value=v)
        review_ws.cell(row=r_offset, column=issue_col, value=issue_label)

    # column width original sheet theke copy kora, plus Issue column
    for c in range(1, max_col + 1):
        letter = ws.cell(row=1, column=c).column_letter
        src_dim = ws.column_dimensions.get(letter)
        if src_dim and src_dim.width:
            review_ws.column_dimensions[letter].width = src_dim.width
    review_ws.column_dimensions[review_ws.cell(row=1, column=issue_col).column_letter].width = 26
    review_ws.freeze_panes = "A2"

    # ---- Report tab ----
    report = wb.create_sheet("Merge Report")
    report["A1"] = "Image Merge into HTML Columns - Summary Report"
    report["A1"].font = Font(bold=True, size=14)

    report["A3"] = "Main file:"
    report["B3"] = os.path.basename(main_path)
    report["A4"] = "Mapping file:"
    report["B4"] = os.path.basename(mapping_path)
    report["A5"] = "Run time:"
    report["B5"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    rows_info = [
        ("Total product rows processed:", total_rows),
        ("Rows fully matched & kept in main sheet:", total_rows - len(moved_rows)),
        ("Rows moved to 'Needs Review' sheet (total):", len(moved_rows)),
        ("  -> Rows with NO mapping found (Product ID not in mapping file):", rows_no_mapping),
        ("  -> Rows with Insufficient Mapped Images (some img tags left old):", rows_partial),
        ("Total <img> tags found (across 4 columns):", tags_found_total),
        ("Total <img> tags replaced with new URL:", tags_replaced_total),
        ("Columns processed:", ", ".join(found_cols.keys()) if found_cols else "(kono column pawa jayni!)"),
    ]
    start_row = 7
    bold = Font(bold=True)
    for i, (label, value) in enumerate(rows_info):
        r = start_row + i
        report.cell(row=r, column=1, value=label).font = bold
        report.cell(row=r, column=2, value=value)

    report.column_dimensions["A"].width = 65
    report.column_dimensions["B"].width = 40

    base, ext = os.path.splitext(os.path.basename(main_path))
    if output_dir is None:
        output_dir = os.path.dirname(os.path.abspath(main_path))
    else:
        os.makedirs(output_dir, exist_ok=True)
    output_path = os.path.join(output_dir, f"{base}_image_replaced.xlsx")
    wb.save(output_path)

    print(f"\nDone!")
    print(f"Total product rows: {total_rows}")
    print(f"Fully matched (main sheet e roilo): {total_rows - len(moved_rows)}")
    print(f"'Needs Review' sheet e move hoyeche: {len(moved_rows)} (No mapping: {rows_no_mapping}, Partial: {rows_partial})")
    print(f"Total img tags found: {tags_found_total}, replaced: {tags_replaced_total}")
    print(f"Notun file save hoyeche ekhane:\n    {output_path}\n")
    return output_path


if __name__ == "__main__":
    print("=== Merge Mapped Image URLs into HTML Columns ===\n")
    main_file = input("Main file er full path dao: ").strip().strip('"')
    mapping_file = input("Mapping file er full path dao (image_http_url_add.xlsx): ").strip().strip('"')

    try:
        process(main_file, mapping_file)
    except Exception as e:
        print(f"\nError hoyeche: {e}")
        sys.exit(1)