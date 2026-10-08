"""Load config.toml and fill in defaults for anything missing."""

import copy
import os
import tomllib

from .common import StepError

DEFAULTS = {
    "output_root": r"D:\Image URL change",
    "base_url": "",
    "mapping": {"folder": "Data", "data_sheet": "Data Mapping", "seller_sheet": "Seller Info"},
    "export": {
        "batch_size": 4, "poll_seconds": 10, "timeout_seconds": 300,
        "headless": False, "slowmo_ms": 150,
        "credentials_file": "credentials.json", "auth_state_file": "auth_state.json",
    },
    "claude": {
        "rows_per_part": 3500,
        "prompt_file": r"templates\H&D_improved_prompt.md",
        "fallback_done_folder": "H&D_Processed_Input",
    },
    "upload": {
        "max_rows_per_file": 7500, "min_files": 2,
        "allow_missing_cleaned_rows": False, "skip_rows_failing_checks": True,
        "approval_template": r"templates\Product-approval-template.xlsx",
    },
    "otp": {
        "enabled": False, "imap_host": "imap.gmail.com", "sender": "", "subject_contains": "",
        "code_pattern": "", "wait_seconds": 120, "input_selector": "", "submit_selector": "",
    },
    "auto": {"watch_minutes": 10, "daily_time": "09:00", "max_attempts": 3, "headless_when_scheduled": True},
    "notify": {"enabled": False, "to": "", "smtp_host": "smtp.gmail.com", "smtp_port": 465},
    "site_upload": {
        "enabled": False, "upload_approval": True, "wait_seconds": 600,
        "update": {"menu": ["Products", "Multi Seller Bulk"], "file_input": "input[type=file]",
                   "submit_button": "button:has-text('Upload')", "success_text": ["success", "completed"],
                   "error_text": ["failed", "error"]},
        "approval": {"menu": [], "file_input": "input[type=file]", "submit_button": "button:has-text('Upload')",
                     "success_text": ["success", "completed"], "error_text": ["failed", "error"]},
    },
    "columns": {
        "keep": ["Product ID", "Seller Code", "Name (English)", "Highlights(English)",
                 "Highlights(Bengali)", "Description (English)", "Description (Bengali)"],
        "html": ["Highlights(English)", "Highlights(Bengali)",
                 "Description (English)", "Description (Bengali)"],
    },
}


def _merge(base, extra):
    for key, value in extra.items():
        if isinstance(value, dict) and isinstance(base.get(key), dict):
            _merge(base[key], value)
        else:
            base[key] = value
    return base


class Config(dict):
    """dict with a project_dir and a helper to resolve project-relative paths."""

    def __init__(self, data, project_dir):
        super().__init__(data)
        self.project_dir = project_dir

    def path(self, value):
        """Relative paths are taken from the project folder. Both / and \\ work."""
        value = value.replace("\\", os.sep).replace("/", os.sep)
        return value if os.path.isabs(value) else os.path.join(self.project_dir, value)


def load(project_dir, filename="config.toml"):
    data = copy.deepcopy(DEFAULTS)
    path = os.path.join(project_dir, filename)
    if os.path.isfile(path):
        try:
            with open(path, "rb") as f:
                _merge(data, tomllib.load(f))
        except tomllib.TOMLDecodeError as e:
            raise StepError(f"config.toml has a mistake: {e}")
    return Config(data, project_dir)
