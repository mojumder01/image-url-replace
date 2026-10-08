"""Step 5 (automatic mode): clean the HTML with the Python rules instead of Claude.

Reads the 'Data' sheet of step 4 and writes 7_merged/Basic combine cleaned.xlsx
(the same file the manual Claude route produces), plus 'Cleaning Report' and
'Check Problems' sheets. The prompt's 6 self-checks run on every row.
"""

import os
from collections import Counter

from ..cleaner import check_row, clean_rows
from ..common import (PID_NAMES, StepError, add_report_sheet, add_table_sheet, find_col, info,
                      new_workbook, read_table, timestamp, warn)


def clean_file(images_path, keep_columns, html_columns, out_path):
    headers, rows = read_table(images_path, sheet="Data")
    if not rows:
        raise StepError("No rows to clean (the 'Data' sheet of step 4 is empty).")
    idx = {name: find_col(headers, name) for name in keep_columns}
    idx[keep_columns[0]] = find_col(headers, *PID_NAMES)
    name_col = find_col(headers, "Name (English)", "*Name (English)")
    he, hb, de, db = (find_col(headers, n) for n in html_columns)
    if None in (he, hb, de, db):
        raise StepError(f"HTML columns not found. Expected: {html_columns}")

    cleaned, stats, notes = clean_rows([(r[name_col] if name_col is not None else "", r[he], r[hb], r[de], r[db])
                                        for r in rows])
    out_rows, problems = [], []
    pid_col = idx[keep_columns[0]]
    for row, new, note in zip(rows, cleaned, notes):
        out = [row[idx[c]] if idx[c] is not None else None for c in keep_columns]
        for name, value in zip(html_columns, new):
            out[keep_columns.index(name)] = value
        out_rows.append(out)
        bad = check_row([row[he], row[hb], row[de], row[db]], list(new))
        if bad:
            problems.append([row[pid_col], ", ".join(bad)])

    counts = Counter(p for _, text in problems for p in text.split(", "))
    wb = new_workbook()
    add_table_sheet(wb, "Data", keep_columns, out_rows, text_cols=[0])
    add_report_sheet(wb, "Cleaning Report", "HTML Cleaning (Python rules) - Summary", [
        ("Source:", os.path.basename(images_path)),
        ("Run time:", timestamp()),
        ("Rows cleaned:", len(out_rows)),
        ("Highlights made from the description:", stats["highlights_from_description"]),
        ("Description made from the highlights:", stats["description_from_highlights"]),
        ("Generic highlights replaced by own description bullets:", stats["highlights_enriched"]),
        ("Rows failing a self-check:", len(problems)),
    ] + [(f"  - {name}:", n) for name, n in counts.items()])
    add_table_sheet(wb, "Check Problems", ["Product ID", "Problem"], problems, text_cols=[0])
    wb.save(out_path)
    return len(out_rows), stats, problems


def run(ctx):
    out = ctx.run.file("cleaned")
    count, stats, problems = clean_file(ctx.run.file("images"), ctx.cfg["columns"]["keep"],
                                        ctx.cfg["columns"]["html"], out)
    info(f"Cleaned {count} rows -> {out}")
    if problems:
        warn(f"{len(problems)} row(s) fail a self-check - see the 'Check Problems' sheet.")
    return {"rows": count, "check_problems": len(problems),
            "highlights_from_description": stats["highlights_from_description"],
            "description_from_highlights": stats["description_from_highlights"],
            "highlights_enriched": stats["highlights_enriched"]}
