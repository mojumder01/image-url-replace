"""Step 7: build upload-ready files from the official template + the approval file.

1. Takes every row of 'Basic combine.xlsx' (Workings sheet, all columns).
2. Keeps ONLY products Claude returned, and puts the cleaned content into the
   HTML columns. Everything else is left out and listed in not_uploaded.xlsx:
     - Needs Review rows (no mapping / too few images)
     - rows Claude did not return (stops unless allow_missing_cleaned_rows)
3. Re-orders columns to match the official export template (one of the files
   downloaded in step 2: sheet 'Product Update', header in row 2, data from row 3).
4. Splits rows evenly into upload files (at least min_files, at most
   max_rows_per_file rows each). Template formatting and dropdowns are kept.
5. Writes Product-approval.xlsx: every uploaded Product ID with Approval Status = 1.
"""

import os
from copy import copy

from openpyxl import load_workbook
from openpyxl.utils import get_column_letter

from ..common import (PID_NAMES, StepError, ask_yes_no, find_col, info, interactive, list_xlsx,
                      norm_header, norm_pid, read_table, warn, write_table)
from .s6_merge import sent_product_ids

TEMPLATE_SHEET = "Product Update"


# ---------------------------------------------------------------------------
# Template helpers
# ---------------------------------------------------------------------------
def find_template(downloads_dir):
    """First export file with a 'Product Update' sheet and Product ID in row 2."""
    for path in list_xlsx(downloads_dir):
        try:
            wb = load_workbook(path, read_only=True)
            ws = _template_sheet(wb)
            rows = list(ws.iter_rows(min_row=1, max_row=2, values_only=True))
            wb.close()
        except Exception:
            continue
        if len(rows) == 2 and find_col(rows[1], *PID_NAMES) is not None:
            return path
    return None


def _template_sheet(wb):
    for name in wb.sheetnames:
        if norm_header(name) == norm_header(TEMPLATE_SHEET):
            return wb[name]
    return wb[wb.sheetnames[0]]


def _extend_validations(ws, orig_max_row, new_max_row):
    """Stretch dropdown ranges that ended at the template's last row."""
    for dv in list(ws.data_validations.dataValidation):
        ranges, changed = [], False
        for rng in list(dv.sqref.ranges):
            if rng.max_row == orig_max_row:
                ranges.append(f"{get_column_letter(rng.min_col)}{rng.min_row}:"
                              f"{get_column_letter(rng.max_col)}{new_max_row}")
                changed = True
            else:
                ranges.append(str(rng))
        if changed:
            dv.sqref = " ".join(ranges)


def fill_sheet(ws, rows, first_data_row):
    """Clear old values (keeping formats/dropdowns), write rows, extend styles/dropdowns."""
    orig_max_row = ws.max_row
    max_col = ws.max_column
    for r in range(first_data_row, orig_max_row + 1):
        for c in range(1, max_col + 1):
            ws.cell(row=r, column=c).value = None
    for i, values in enumerate(rows):
        r = first_data_row + i
        if r > orig_max_row:
            for c in range(1, max_col + 1):
                ws.cell(row=r, column=c)._style = copy(ws.cell(row=orig_max_row, column=c)._style)
        for c, v in enumerate(values, start=1):
            ws.cell(row=r, column=c).value = v
    new_max_row = first_data_row + len(rows) - 1
    if new_max_row > orig_max_row:
        _extend_validations(ws, orig_max_row, new_max_row)


def split_evenly(rows, max_per_file, min_files):
    n = max(min_files, -(-len(rows) // max_per_file))
    n = max(1, min(n, len(rows)))
    size, extra = divmod(len(rows), n)
    chunks, start = [], 0
    for i in range(n):
        end = start + size + (1 if i < extra else 0)
        chunks.append(rows[start:end])
        start = end
    return chunks


def build_approval(template_path, product_ids, out_path):
    wb = load_workbook(template_path)
    ws = wb[wb.sheetnames[0]]
    header = [ws.cell(1, c).value for c in range(1, ws.max_column + 1)]
    pid_col = find_col(header, *PID_NAMES)
    status_col = find_col(header, "Approval Status")
    if pid_col is None or status_col is None:
        raise StepError("Approval template needs 'Product ID' and 'Approval Status' columns.")
    rows = []
    for pid in product_ids:
        row = [None] * len(header)
        row[pid_col] = pid
        row[status_col] = 1
        rows.append(row)
    fill_sheet(ws, rows, first_data_row=2)
    wb.save(out_path)


# ---------------------------------------------------------------------------
def run(ctx):
    run_, cfg = ctx.run, ctx.cfg
    html_names = cfg["columns"]["html"]

    # Main data (all columns) and cleaned content
    headers, rows = read_table(run_.file("combine"), sheet="Workings")
    pid_col = find_col(headers, *PID_NAMES)
    if pid_col is None:
        raise StepError("No 'Product ID' column in Basic combine.xlsx.")
    html_cols = {name: find_col(headers, name) for name in html_names}
    html_cols = {k: v for k, v in html_cols.items() if v is not None}

    c_headers, c_rows = read_table(run_.file("cleaned"))
    c_pid = find_col(c_headers, *PID_NAMES)
    c_cols = {name: find_col(c_headers, name) for name in html_cols}
    cleaned = {}
    for r in c_rows:
        pid = norm_pid(r[c_pid])
        if pid:
            cleaned[pid] = {name: r[i] for name, i in c_cols.items() if i is not None}

    # Why a product is not uploaded
    review = {}
    r_headers, r_rows = read_table(run_.file("images"), sheet="Needs Review")
    r_pid, r_issue = find_col(r_headers, *PID_NAMES), find_col(r_headers, "Issue")
    if r_pid is not None:
        review = {norm_pid(r[r_pid]): (r[r_issue] if r_issue is not None else "Needs Review") for r in r_rows}
    sent = sent_product_ids(run_.folder("for_claude"))

    upload_rows, skipped = [], []
    for row in rows:
        pid = norm_pid(row[pid_col])
        if pid in review:
            skipped.append((pid, f"Needs Review: {review[pid]}", row))
        elif pid not in sent:
            skipped.append((pid, "Not sent to Claude", row))
        elif pid not in cleaned:
            skipped.append((pid, "Missing from Claude output", row))
        else:
            for name, i in html_cols.items():
                if name in cleaned[pid]:
                    row[i] = cleaned[pid][name]
            upload_rows.append(row)

    missing = [s for s in skipped if s[1] == "Missing from Claude output"]
    if missing:
        warn(f"{len(missing)} product(s) were sent to Claude but are not in its output. "
             f"Example: {[m[0] for m in missing[:5]]}")
        allowed = cfg["upload"]["allow_missing_cleaned_rows"] or ctx.options.get("allow_missing")
        if not allowed and not (interactive() and ask_yes_no("Leave them out and continue?")):
            raise StepError("Stopped. Add the missing rows to 6_claude_done, or run finish with "
                            "--allow-missing (or set allow_missing_cleaned_rows = true).")
    if not upload_rows:
        raise StepError("Nothing to upload - no cleaned product matched Basic combine.xlsx.")

    # Template
    downloads = os.path.join(run_.folder("export"), "downloads")
    template = find_template(downloads)
    if not template:
        raise StepError(f"No export template found in {downloads} "
                        f"(need a file with a '{TEMPLATE_SHEET}' sheet).")
    wb = load_workbook(template, read_only=True)
    t_header = list(list(_template_sheet(wb).iter_rows(min_row=2, max_row=2, values_only=True))[0])
    wb.close()
    t_index = {norm_header(h): i for i, h in enumerate(t_header) if h is not None}
    mapping = [(i, t_index[norm_header(h)]) for i, h in enumerate(headers)
               if h is not None and norm_header(h) in t_index]
    not_in_template = [h for h in headers if h is not None and norm_header(h) not in t_index]
    if not_in_template:
        info(f"Columns not in the template (left out): {not_in_template}")
    arranged = []
    for row in upload_rows:
        out = [None] * len(t_header)
        for src, dst in mapping:
            out[dst] = row[src]
        arranged.append(out)

    # Upload files
    up = cfg["upload"]
    out_dir = run_.folder("upload")
    for old in os.listdir(out_dir):
        if old.lower().startswith("update_part") and old.lower().endswith(".xlsx"):
            os.remove(os.path.join(out_dir, old))
    chunks = split_evenly(arranged, int(up["max_rows_per_file"]), int(up["min_files"]))
    for n, chunk in enumerate(chunks, start=1):
        out_wb = load_workbook(template)
        fill_sheet(_template_sheet(out_wb), chunk, first_data_row=3)
        path = os.path.join(out_dir, f"update_part{n}.xlsx")
        out_wb.save(path)
        info(f"Wrote {os.path.basename(path)} ({len(chunk)} rows)")

    # Approval file
    approval_tpl = cfg.path(up["approval_template"])
    if os.path.isfile(approval_tpl):
        build_approval(approval_tpl, [norm_pid(r[pid_col]) for r in upload_rows], run_.file("approval"))
        info(f"Wrote Product-approval.xlsx ({len(upload_rows)} products, Approval Status = 1)")
    else:
        warn(f"Approval template not found ({approval_tpl}) - approval file not made.")

    # Report what was left out
    seller_col = find_col(headers, "Seller Code")
    name_col = find_col(headers, "Name (English)")
    write_table(run_.file("not_uploaded"), ["Product ID", "Seller Code", "Name (English)", "Reason"],
                [[pid, row[seller_col] if seller_col is not None else None,
                  row[name_col] if name_col is not None else None, reason]
                 for pid, reason, row in skipped], text_cols=[0])

    info(f"\nUploaded: {len(upload_rows)} products in {len(chunks)} file(s). "
         f"Left out: {len(skipped)} (see not_uploaded.xlsx)")
    info(f"Upload folder: {out_dir}")
    return {"uploaded": len(upload_rows), "files": len(chunks), "left_out": len(skipped),
            "missing_from_claude": len(missing), "template": os.path.basename(template)}
