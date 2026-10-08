"""One run = one dated folder holding every input, output, status and log."""

import json
import os
from datetime import datetime

from .common import StepError, info, timestamp

# Sub-folders in the order the work happens.
FOLDERS = {
    "input": "0_input",
    "urls": "1_urls",
    "export": "2_export",
    "trim": "3_trim",
    "images": "4_images",
    "for_claude": "5_for_claude",
    "claude_done": "6_claude_done",
    "merged": "7_merged",
    "upload": "8_upload",
}

# Fixed output file names.
FILES = {
    "urls": ("urls", "image_http_url_add.xlsx"),
    "combine": ("export", "Basic combine.xlsx"),
    "export_report": ("export", "export_report.xlsx"),
    "trim": ("trim", "Basic combine H & D.xlsx"),
    "images": ("images", "Basic combine H & D_image_replaced.xlsx"),
    "cleaned": ("merged", "Basic combine cleaned.xlsx"),
    "approval": ("upload", "Product-approval.xlsx"),
    "not_uploaded": ("upload", "not_uploaded.xlsx"),
}

STATUS_NAME = "run.json"
LOG_NAME = "run.log"


class Run:
    def __init__(self, root):
        self.root = root

    # ----- creating / finding -------------------------------------------
    @classmethod
    def create(cls, output_root, session_date, now=None):
        """Folder name: '<Month Year>_<DD-MM-YYYY>_<HH-MM>' (run start time)."""
        now = now or datetime.now()
        name = f"{session_date:%B %Y}_{session_date:%d-%m-%Y}_{now:%H-%M}"
        root = os.path.join(output_root, name)
        n = 2
        while os.path.exists(root):
            root = os.path.join(output_root, f"{name}-{n}")
            n += 1
        os.makedirs(root)
        run = cls(root)
        run._save({"created_at": timestamp(), "session_date": f"{session_date:%d-%m-%Y}", "steps": {}})
        return run

    @classmethod
    def all_runs(cls, output_root):
        """All run folders, newest first."""
        runs = []
        if os.path.isdir(output_root):
            for name in os.listdir(output_root):
                root = os.path.join(output_root, name)
                if os.path.isfile(os.path.join(root, STATUS_NAME)):
                    runs.append(cls(root))
        runs.sort(key=lambda r: r.status().get("created_at", ""), reverse=True)
        return runs

    @classmethod
    def open(cls, output_root, path=None):
        """A given run folder, or the newest one."""
        if path:
            if not os.path.isfile(os.path.join(path, STATUS_NAME)):
                raise StepError(f"Not a run folder (no {STATUS_NAME}): {path}")
            return cls(path)
        runs = cls.all_runs(output_root)
        if not runs:
            raise StepError(f"No run folder found in {output_root}. Run 'prepare' first.")
        return runs[0]

    # ----- paths ---------------------------------------------------------
    @property
    def name(self):
        return os.path.basename(self.root)

    def folder(self, key):
        path = os.path.join(self.root, FOLDERS[key])
        os.makedirs(path, exist_ok=True)
        return path

    def file(self, key):
        folder_key, filename = FILES[key]
        return os.path.join(self.folder(folder_key), filename)

    def mapping_file(self):
        name = self.status().get("mapping_file")
        path = os.path.join(self.folder("input"), name) if name else None
        if not path or not os.path.isfile(path):
            raise StepError(f"Mapping file is missing from {self.folder('input')}")
        return path

    # ----- status + log --------------------------------------------------
    def status(self):
        try:
            with open(os.path.join(self.root, STATUS_NAME), encoding="utf-8") as f:
                return json.load(f)
        except (OSError, ValueError):
            return {"steps": {}}

    def _save(self, data):
        with open(os.path.join(self.root, STATUS_NAME), "w", encoding="utf-8") as f:
            json.dump(data, f, indent=2, ensure_ascii=False)

    def set(self, **values):
        data = self.status()
        data.update(values)
        self._save(data)

    def step_done(self, step, **summary):
        data = self.status()
        data.setdefault("steps", {})[step] = {"state": "done", "finished_at": timestamp(), **summary}
        self._save(data)
        details = ", ".join(f"{k}={v}" for k, v in summary.items())
        self.log(f"{step} done" + (f" ({details})" if details else ""))

    def step_failed(self, step, reason):
        data = self.status()
        data.setdefault("steps", {})[step] = {"state": "failed", "finished_at": timestamp(), "error": reason}
        self._save(data)
        self.log(f"{step} FAILED: {reason}")

    def is_done(self, step):
        return self.status().get("steps", {}).get(step, {}).get("state") == "done"

    def log(self, msg):
        line = f"[{timestamp()}] {msg}"
        with open(os.path.join(self.root, LOG_NAME), "a", encoding="utf-8") as f:
            f.write(line + "\n")
        info(line)
