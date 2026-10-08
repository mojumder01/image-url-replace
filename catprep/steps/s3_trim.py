"""Step 3: keep only the columns Claude needs (config [columns] keep)."""

from ..common import StepError, find_col, info, read_table, warn, write_table


def trim(combine_path, keep, out_path):
    headers, rows = read_table(combine_path, sheet="Workings")
    picked = [(name, find_col(headers, name)) for name in keep]
    missing = [name for name, idx in picked if idx is None]
    picked = [(name, idx) for name, idx in picked if idx is not None]
    if not picked:
        raise StepError("None of the expected columns were found. Check the export file headers.")
    if missing:
        warn(f"Columns not found and skipped: {missing}")
    out_rows = [[row[idx] for _, idx in picked] for row in rows]
    write_table(out_path, [name for name, _ in picked], out_rows, text_cols=[0])
    return len(out_rows), missing


def run(ctx):
    out = ctx.run.file("trim")
    count, missing = trim(ctx.run.file("combine"), ctx.cfg["columns"]["keep"], out)
    info(f"{count} rows, {len(ctx.cfg['columns']['keep']) - len(missing)} columns -> {out}")
    return {"rows": count, "missing_columns": missing}
