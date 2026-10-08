"""Step 6: join Claude's cleaned files into one 'Basic combine cleaned.xlsx'.

Files come from the run's 6_claude_done folder (or, if that is empty, the
project's fallback folder, e.g. H&D_Processed_Input). Columns are matched by
name loosely ("Highlights (English)" = "Highlights(English)"), so Claude's small
header changes do not matter. Product IDs are compared with what was sent:
missing, extra and duplicate IDs are reported.
"""

import os

from ..common import (PID_NAMES, StepError, find_col, info, list_xlsx, norm_pid, read_table, warn,
                      write_table)


def sent_product_ids(for_claude_dir):
    ids = set()
    for path in list_xlsx(for_claude_dir):
        headers, rows = read_table(path)
        col = find_col(headers, *PID_NAMES)
        if col is not None:
            ids.update(p for p in (norm_pid(r[col]) for r in rows) if p)
    return ids


def merge_cleaned(files, columns, out_path):
    """columns = canonical output columns (config keep list)."""
    out_rows, problems = [], []
    for path in files:
        headers, rows = read_table(path)
        idx = [find_col(headers, *(PID_NAMES if i == 0 else (name,))) for i, name in enumerate(columns)]
        missing = [columns[i] for i, c in enumerate(idx) if c is None]
        if idx[0] is None:
            raise StepError(f"{os.path.basename(path)} has no 'Product ID' column.")
        if missing:
            problems.append(f"{os.path.basename(path)}: missing columns {missing}")
        for row in rows:
            out = [row[c] if c is not None else None for c in idx]
            out[0] = norm_pid(out[0])
            out_rows.append(out)
        info(f"  {os.path.basename(path)}: {len(rows)} rows")
    write_table(out_path, columns, out_rows, text_cols=[0])
    return out_rows, problems


def run(ctx):
    run_, cfg = ctx.run, ctx.cfg
    files = list_xlsx(run_.folder("claude_done"))
    if not files:
        fallback = cfg.path(cfg["claude"]["fallback_done_folder"])
        files = list_xlsx(fallback)
        if files:
            info(f"6_claude_done is empty - using files from {fallback}")
    if not files:
        raise StepError(f"No cleaned files found. Put Claude's output files in:\n    {run_.folder('claude_done')}")

    columns = ["Product ID"] + [c for c in cfg["columns"]["keep"] if find_col([c], *PID_NAMES) is None]
    info(f"Merging {len(files)} cleaned file(s):")
    rows, problems = merge_cleaned(files, columns, run_.file("cleaned"))
    for p in problems:
        warn(p)

    cleaned = [r[0] for r in rows if r[0]]
    duplicates = len(cleaned) - len(set(cleaned))
    sent = sent_product_ids(run_.folder("for_claude"))
    missing = sorted(sent - set(cleaned))
    extra = sorted(set(cleaned) - sent)
    info(f"Rows sent to Claude: {len(sent)} | returned: {len(set(cleaned))} | "
         f"missing: {len(missing)} | extra: {len(extra)} | duplicates: {duplicates}")
    if duplicates:
        warn(f"{duplicates} duplicate Product ID(s) in Claude's output - the last one is used.")
    if extra:
        warn(f"{len(extra)} Product ID(s) were not sent to Claude (ignored at upload). Example: {extra[:5]}")
    if missing:
        warn(f"{len(missing)} Product ID(s) sent to Claude are missing from its output. Example: {missing[:5]}")
    info(f"-> {run_.file('cleaned')}")
    return {"files": len(files), "rows": len(rows), "missing_from_claude": len(missing),
            "extra": len(extra), "duplicates": duplicates}
