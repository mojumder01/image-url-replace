"""Step 5: split the good rows into parts for Claude, and add the prompt + instructions."""

import os
import shutil

from ..common import PID_NAMES, StepError, find_col, info, read_table, warn, write_table


def split(images_path, rows_per_part, out_dir):
    headers, rows = read_table(images_path, sheet="Data")
    if not rows:
        raise StepError("No rows to send to Claude (the 'Data' sheet is empty).")
    for old in os.listdir(out_dir):
        if old.lower().startswith("part-") and old.lower().endswith(".xlsx"):
            os.remove(os.path.join(out_dir, old))
    pid_col = find_col(headers, *PID_NAMES)
    text_cols = [pid_col] if pid_col is not None else []
    parts = []
    for n, start in enumerate(range(0, len(rows), rows_per_part), start=1):
        path = os.path.join(out_dir, f"Part-{n}.xlsx")
        write_table(path, headers, rows[start:start + rows_per_part], sheet="Sheet1", text_cols=text_cols)
        parts.append((os.path.basename(path), min(rows_per_part, len(rows) - start)))
    return parts, len(rows)


INSTRUCTIONS = """HOW TO CLEAN THESE FILES WITH CLAUDE
=====================================

1. Open Claude. For each Part file:
     - upload the Part file AND {prompt}
     - ask Claude to process the file following the prompt
     - download the cleaned .xlsx it gives back
2. Save every cleaned file into this folder:
     {done}
   (any file names are fine; keep the same columns)
3. When all parts are done, double-click manual_claude\2_Finish.bat
   (or run:  python catprep.py finish)

Parts in this folder:
{parts}
Total rows: {total}
"""


def run(ctx):
    out_dir = ctx.run.folder("for_claude")
    parts, total = split(ctx.run.file("images"), int(ctx.cfg["claude"]["rows_per_part"]), out_dir)

    prompt = ctx.cfg.path(ctx.cfg["claude"]["prompt_file"])
    prompt_name = os.path.basename(prompt)
    if os.path.isfile(prompt):
        shutil.copy2(prompt, os.path.join(out_dir, prompt_name))
    else:
        warn(f"Prompt file not found: {prompt}")
    with open(os.path.join(out_dir, "README - what to do next.txt"), "w", encoding="utf-8") as f:
        f.write(INSTRUCTIONS.format(
            prompt=prompt_name, done=ctx.run.folder("claude_done"), total=total,
            parts="\n".join(f"  {name}  ({n} rows)" for name, n in parts)))

    info(f"{total} rows -> {len(parts)} part(s) in {out_dir}")
    return {"parts": len(parts), "rows": total}
