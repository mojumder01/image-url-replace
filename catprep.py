"""Catalogue Prep v2

Commands:
  python catprep.py prepare            new run: steps urls -> export -> trim -> images -> split
  python catprep.py prepare --resume   continue the newest run from the first unfinished step
  python catprep.py finish             merge Claude's files + build upload files (newest run)
  python catprep.py step <name>        re-run one step on the newest run
  python catprep.py status             show the newest run's progress
  python catprep.py runs               list all runs
  python catprep.py check <file> [file2]   look inside an Excel file (read-only)

Add --run "<run folder>" to work on an older run instead of the newest.
Settings: config.toml
"""

import argparse
import os
import shutil
import sys
import traceback
from datetime import datetime

from catprep import config as config_mod
from catprep.common import (PID_NAMES, StepError, date_from_filename, find_col, find_mapping_files,
                            heading, info, norm_pid, warn)
from catprep.run_folder import Run
from catprep.steps import FINISH, PREPARE, STEPS

PROJECT_DIR = os.path.dirname(os.path.abspath(__file__))


class Context:
    def __init__(self, run, cfg, options):
        self.run, self.cfg, self.options = run, cfg, options


def run_steps(ctx, names):
    """Run steps in order; stop at the first failure. Returns True when all succeed."""
    for name in names:
        module, title = STEPS[name]
        heading(f"STEP {list(STEPS).index(name) + 1}/{len(STEPS)}: {title}  [{name}]")
        try:
            summary = module.run(ctx) or {}
        except StepError as e:
            ctx.run.step_failed(name, str(e))
            info(f"\nSTOPPED at step '{name}':\n{e}")
            return False
        except KeyboardInterrupt:
            ctx.run.step_failed(name, "interrupted")
            info("\nInterrupted. Continue later with:  python catprep.py prepare --resume")
            return False
        except Exception as e:
            traceback.print_exc()
            ctx.run.step_failed(name, f"{type(e).__name__}: {e}")
            info(f"\nUnexpected error in step '{name}' (details above).")
            return False
        ctx.run.step_done(name, **summary)
    return True


def open_folder(path):
    if os.name == "nt":
        try:
            os.startfile(path)  # Windows Explorer
        except OSError:
            pass


# ---------------------------------------------------------------------------
def cmd_prepare(args, cfg):
    if args.resume:
        run = Run.open(cfg["output_root"], args.run)
        todo = [s for s in PREPARE if not run.is_done(s)]
        info(f"Resuming run: {run.root}")
        if not todo:
            info("All prepare steps are already done. Next: python catprep.py finish")
            return True
    else:
        mapping = args.mapping
        if not mapping:
            folder = cfg.path(cfg["mapping"]["folder"])
            found = find_mapping_files([folder])
            if not found:
                raise StepError(f"No mapping file in {folder}\n"
                                "Its name must contain 'image', 'link' and 'mapping' (or 'bakaul').")
            mapping = found[0]
            if len(found) > 1:
                info(f"{len(found)} mapping files found - using the newest.")
        if not os.path.isfile(mapping):
            raise StepError(f"Mapping file not found: {mapping}")
        session_date = date_from_filename(mapping)
        if not session_date:
            warn("No date in the mapping file name - using today's date.")
            session_date = datetime.now()
        run = Run.create(cfg["output_root"], session_date)
        shutil.copy2(mapping, run.folder("input"))
        run.set(mapping_file=os.path.basename(mapping), mapping_source=os.path.abspath(mapping))
        run.log(f"New run for mapping file: {os.path.basename(mapping)}")
        info(f"Run folder: {run.root}")
        todo = PREPARE

    ctx = Context(run, cfg, {"manual_export": args.manual_export})
    if not run_steps(ctx, todo):
        info(f"\nFix the problem, then continue with:  python catprep.py prepare --resume")
        return False
    heading("PREPARE COMPLETE")
    info(f"Files for Claude:   {run.folder('for_claude')}")
    info(f"Put cleaned files:  {run.folder('claude_done')}")
    info("Then run 2_Finish.bat  (python catprep.py finish)")
    open_folder(run.folder("for_claude"))
    return True


def cmd_finish(args, cfg):
    run = Run.open(cfg["output_root"], args.run)
    info(f"Run folder: {run.root}")
    not_ready = [s for s in PREPARE if not run.is_done(s)]
    if not_ready:
        raise StepError(f"Prepare is not finished for this run (missing: {not_ready}). "
                        "Run: python catprep.py prepare --resume")
    ctx = Context(run, cfg, {"allow_missing": args.allow_missing})
    if not run_steps(ctx, FINISH):
        return False
    heading("FINISH COMPLETE - ready to upload")
    info(f"Upload folder: {run.folder('upload')}")
    open_folder(run.folder("upload"))
    return True


def cmd_step(args, cfg):
    run = Run.open(cfg["output_root"], args.run)
    info(f"Run folder: {run.root}")
    ctx = Context(run, cfg, {"manual_export": args.manual_export, "allow_missing": args.allow_missing})
    return run_steps(ctx, [args.name])


def cmd_status(args, cfg):
    run = Run.open(cfg["output_root"], args.run)
    st = run.status()
    heading(f"RUN: {run.name}")
    info(f"Folder:       {run.root}")
    info(f"Mapping file: {st.get('mapping_file', '-')}")
    info(f"Created:      {st.get('created_at', '-')}")
    info("")
    for n, name in enumerate(STEPS, start=1):
        s = st.get("steps", {}).get(name, {})
        state = s.get("state", "not started")
        extra = {k: v for k, v in s.items() if k not in ("state", "finished_at")}
        details = ", ".join(f"{k}={v}" for k, v in extra.items())
        info(f"  {n}. {name:<7} {state:<12} {s.get('finished_at', ''):<20} {details}")
        if name == PREPARE[-1]:
            info("     --- Claude cleaning (manual) ---")
    return True


def cmd_runs(args, cfg):
    runs = Run.all_runs(cfg["output_root"])
    if not runs:
        info(f"No runs in {cfg['output_root']}")
    for r in runs:
        steps = r.status().get("steps", {})
        done = sum(1 for s in steps.values() if s.get("state") == "done")
        info(f"  {r.name:<40} {done}/{len(STEPS)} steps done")
    return True


def cmd_check(args, cfg):
    from openpyxl import load_workbook

    def ids_of(path):
        wb = load_workbook(path, read_only=True, data_only=True)
        ids = set()
        for ws in wb.worksheets:
            rows = ws.iter_rows(values_only=True)
            for _ in range(2):  # header can be in row 1 or 2
                header = next(rows, None) or []
                col = find_col(header, *PID_NAMES)
                if col is not None:
                    ids.update(p for p in (norm_pid(r[col]) for r in rows if col < len(r)) if p)
                    break
        wb.close()
        return ids

    for path in [args.file] + ([args.file2] if args.file2 else []):
        wb = load_workbook(path, read_only=True, data_only=True)
        heading(os.path.basename(path))
        for ws in wb.worksheets:
            first = [list(r) for r in ws.iter_rows(min_row=1, max_row=2, values_only=True)]
            info(f"Sheet '{ws.title}': {ws.max_row} rows x {ws.max_column} columns")
            for i, row in enumerate(first, start=1):
                info(f"  row {i}: {[v for v in row if v is not None][:12]}")
        wb.close()
        info(f"Unique Product IDs: {len(ids_of(path))}")
    if args.file2:
        a, b = ids_of(args.file), ids_of(args.file2)
        info(f"\nIn both: {len(a & b)} | only in first: {len(a - b)} | only in second: {len(b - a)}")
    return True


# ---------------------------------------------------------------------------
def main(argv=None):
    p = argparse.ArgumentParser(description="Catalogue Prep v2",
                                formatter_class=argparse.RawDescriptionHelpFormatter, epilog=__doc__)
    sub = p.add_subparsers(dest="command", required=True)

    sp = sub.add_parser("prepare", help="steps 1-5: mapping file -> parts for Claude")
    sp.add_argument("--resume", action="store_true", help="continue the newest (or --run) run")
    sp.add_argument("--mapping", help="use this mapping file instead of the newest in Data/")
    sp.add_argument("--manual-export", action="store_true", help="put export files in yourself (no browser)")
    sp.add_argument("--run", help="run folder (only with --resume)")

    sp = sub.add_parser("finish", help="steps 6-7: Claude's files -> upload files")
    sp.add_argument("--run", help="run folder (default: newest)")
    sp.add_argument("--allow-missing", action="store_true",
                    help="continue even if Claude's output is missing some rows")

    sp = sub.add_parser("step", help="re-run one step")
    sp.add_argument("name", choices=list(STEPS))
    sp.add_argument("--run", help="run folder (default: newest)")
    sp.add_argument("--manual-export", action="store_true")
    sp.add_argument("--allow-missing", action="store_true")

    for name in ("status", "runs"):
        sp = sub.add_parser(name)
        sp.add_argument("--run", help="run folder (default: newest)")

    sp = sub.add_parser("check", help="look inside an Excel file (read-only)")
    sp.add_argument("file")
    sp.add_argument("file2", nargs="?", help="second file: compare Product IDs")

    args = p.parse_args(argv)
    commands = {"prepare": cmd_prepare, "finish": cmd_finish, "step": cmd_step,
                "status": cmd_status, "runs": cmd_runs, "check": cmd_check}
    try:
        cfg = config_mod.load(PROJECT_DIR)
        ok = commands[args.command](args, cfg)
    except StepError as e:
        info(f"\nERROR: {e}")
        ok = False
    except KeyboardInterrupt:
        info("\nInterrupted.")
        ok = False
    return 0 if ok else 1


if __name__ == "__main__":
    sys.exit(main())
