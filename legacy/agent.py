# -*- coding: utf-8 -*-
"""
Catalogue Prep Agent
--------------------
Ei agent ta apnar pipeline (Step 1 - Step 5) run korbe, validate korbe,
progress mone rakhbe, ar Claude handoff / merge-back help korbe.
Apnar original 5 ta script ekdom unchanged thakbe - ei agent shudhu
shegulo theke function import kore call kore.

Kivabe run korbe:
    cd "Image URL work agent"
    python agent.py
"""

import argparse
import json
import os
import re
import shutil
import subprocess
import sys
import traceback
from copy import copy
from datetime import datetime

# --auto mode er jonno config - parse_args() te set hoy
AUTO_CONFIG = {}
AUTO_MODE = False

try:
    import pandas as pd
except ImportError:
    print("pandas install kora nai. Age ei command run koro:")
    print("    pip install pandas openpyxl")
    sys.exit(1)

try:
    from openpyxl import Workbook, load_workbook
except ImportError:
    print("openpyxl install kora nai. Age ei command run koro:")
    print("    pip install openpyxl")
    sys.exit(1)

from openpyxl.worksheet.cell_range import MultiCellRange

# ---------------------------------------------------------------------
# Original pipeline scripts import kora (unchanged thakbe).
# Import korlei interactive prompt chalu hoy na - oigulo
# "if __name__ == '__main__'" er bhitore.
# ---------------------------------------------------------------------
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
if BASE_DIR not in sys.path:
    sys.path.insert(0, BASE_DIR)

from add_http_to_image_key import process_file as step1_add_http
from trim_columns import trim_file as step3_trim
from merge_images_into_html import process as step4_merge
from split_excel import split_excel as step5_split  # noqa: F401  (option 4 uses run_step5 directly)

STATUS_FILE = os.path.join(BASE_DIR, "pipeline_status.json")
LOG_FILE = os.path.join(BASE_DIR, "agent_log.txt")

STEP1_DEFAULT_OUT = os.path.join(BASE_DIR, "image_http_url_add.xlsx")
STEP3_DEFAULT_OUT = os.path.join(BASE_DIR, "Basic combine H & D.xlsx")

# Step 2 automation script (Playwright export + combine -> Basic combine.xlsx)
STEP2_SCRIPT = os.path.join(BASE_DIR, "seller_bulk_export_automation.py")

STATUS_LABELS = {
    1: "Image URL add (Step 1)",
    2: "Basic combine.xlsx ready? (Step 2, automated)",
    3: "Trim columns (Step 3)",
    4: "Merge images into HTML (Step 4)",
    5: "Split for Claude (Step 5)",
    6: "Content update + split for bulk update (Step 6)",
}


# =====================================================================
# Small helpers
# =====================================================================
def ts():
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def clean_input(prompt_text, default=None):
    """input() + strip + surrounding quotes remove (drag & drop path).
    AUTO_MODE=True hole default value return kore, kono jiggesh na kore."""
    if AUTO_MODE and default is not None:
        print(f"{prompt_text}{default}  [auto]")
        return default
    try:
        val = input(prompt_text).strip().strip('"').strip("'")
    except EOFError:
        val = ""
    if not val and default is not None:
        return default
    return val


def ask_yes_no(prompt_text, default_yes=False):
    suffix = "[Y/n]: " if default_yes else "[y/N]: "
    if AUTO_MODE:
        # Auto mode te overwrite_existing config theke nebo
        auto_ans = AUTO_CONFIG.get("overwrite_existing", True)
        print(f"{prompt_text}{suffix}{'y' if auto_ans else 'n'}  [auto]")
        return auto_ans
    val = clean_input(prompt_text + suffix).lower()
    if not val:
        return default_yes
    return val in ("y", "yes")


def confirm_overwrite_if_exists(path, context=""):
    """Output file already thakle user er permission chara overwrite korbe na."""
    if os.path.exists(path):
        ctx = f" ({context})" if context else ""
        print(f"\nWARNING: ei file already ache{ctx}: {path}")
        if not ask_yes_no("Overwrite korte chan?"):
            print("Skip kora holo - apnar kono file change hoy nai.")
            return False
    return True


def log(msg):
    line = f"[{ts()}] {msg}"
    try:
        with open(LOG_FILE, "a", encoding="utf-8") as f:
            f.write(line + "\n")
    except OSError:
        pass
    print(line)


def normalize_pid(value):
    """Product ID ke consistent string e convert kore.
    Same convention: merge_images_into_html.py er normalize_pid().
    2261883 / 2261883.0 / '2261883' shob -> '2261883'.
    Blank/NaN -> None (kono 'nan' string likhbe na)."""
    try:
        if value is None or pd.isna(value):
            return None
    except (TypeError, ValueError):
        pass
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    s = str(value).strip()
    if s.endswith(".0"):
        s = s[:-2]
    return s


# =====================================================================
# Status tracking (pipeline_status.json + agent_log.txt)
# =====================================================================
# Ei session/run e je date use hobe output folder banate.
# Mapping file er naam theke date paua gele shei date use hobe -
# system er "aajker" date noy - jate track rakha jay kon kaj kobe
# process hoyeche (mapping file er date onujayi).
_SESSION_DATE = None  # datetime object, run_steps_1_to_5() e set hoy


def extract_date_from_filename(filename):
    """Filename e DD.MM.YYYY / DD-MM-YYYY / (DD.MM.YYYY) pattern khoje.
    Pele datetime object return kore, na pele None."""
    if not filename:
        return None
    patterns = [
        r"(\d{2})[.\-](\d{2})[.\-](\d{4})",  # 21.09.2026 ba 21-09-2026
        r"(\d{4})[.\-](\d{2})[.\-](\d{2})",  # 2026-09-21
    ]
    base = os.path.basename(filename)
    for pat in patterns:
        m = re.search(pat, base)
        if not m:
            continue
        g = m.groups()
        try:
            if len(g[0]) == 4:  # YYYY-MM-DD pattern
                y, mo, d = g
            else:  # DD-MM-YYYY pattern
                d, mo, y = g
            return datetime(int(y), int(mo), int(d))
        except ValueError:
            continue
    return None


def set_session_date_from_mapping(mapping_path):
    """run_steps_1_to_5() shuru hobar shoshoi ekbar call hoy.
    Mapping file er naam theke date ber kore _SESSION_DATE set kore.
    Date na pele system er aajker date fallback hisebe use hoy."""
    global _SESSION_DATE
    found = extract_date_from_filename(mapping_path)
    if found:
        _SESSION_DATE = found
        print(f"[info] Mapping file er date theke session date set: {found.strftime('%d-%m-%Y')}")
    else:
        _SESSION_DATE = datetime.now()
        print(f"[info] Mapping file e date pawa jayni - aajker date use hocche: {_SESSION_DATE.strftime('%d-%m-%Y')}")


def _auto_detect_session_date():
    """Kono prompt/jiggesh chhara, chupchap Data folder + main folder e
    mapping file khuje shei file er naam theke date ber kore dey.
    Ei function ta get_d_drive_output_folder() theke call hoy jokhon
    _SESSION_DATE ekhono set kora hoyni - jate PIPELINE-er baire theke
    (jemon shudhu Step 2/combine export alada kore chalale) o shothik
    date-wise folder e output jay, aajker system date e na."""
    try:
        data_folder = os.path.join(BASE_DIR, "Data")
        files = fuzzy_scan_xlsx([data_folder, BASE_DIR])
        candidates = match_mapping_files(files)
        if candidates:
            full, fname = candidates[0]
            found = extract_date_from_filename(fname)
            if found:
                return found
    except Exception:
        pass
    return None


def get_d_drive_output_folder():
    """Output folder banay: D:\\Image URL change\\Month Year\\DD-MM-YYYY
    _SESSION_DATE set kora thakle (mapping file er date) shetai use hoy.
    Na thakle (jemon kono submenu shorashori chalano hoyeche, full
    pipeline er baire theke) - chupchap mapping file khuje date ber
    korar chesta kore (_auto_detect_session_date). Kono mapping file
    na pawa gele shobsheshe aajker (system) date fallback hisebe use
    hoy. Ei karone 'September' er poriborte notun 'September 2026'
    duplicate folder ar banbe na - mapping file er month/year onujayi
    shothik purono folder e giye poribe."""
    global _SESSION_DATE
    base = r"D:\Image URL change"
    if _SESSION_DATE is not None:
        now = _SESSION_DATE
    else:
        auto_found = _auto_detect_session_date()
        now = auto_found if auto_found else datetime.now()
        _SESSION_DATE = now  # cache kore rakhi, baar baar file scan na hoy
    month_str = now.strftime("%B %Y")  # e.g., "September 2026"
    date_str = now.strftime("%d-%m-%Y")  # e.g., "21-09-2026"

    folder = os.path.join(base, month_str, date_str)
    os.makedirs(folder, exist_ok=True)
    return folder


def load_status():
    if os.path.isfile(STATUS_FILE):
        try:
            with open(STATUS_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        except (OSError, ValueError):
            pass
    return {}


def save_status(status):
    with open(STATUS_FILE, "w", encoding="utf-8") as f:
        json.dump(status, f, indent=2, ensure_ascii=False)


def update_status(step, **kwargs):
    status = load_status()
    entry = status.setdefault(str(step), {})
    entry["label"] = STATUS_LABELS.get(step, f"Step {step}")
    entry["updated_at"] = ts()
    for k, v in kwargs.items():
        entry[k] = v
    save_status(status)


def show_status():
    status = load_status()
    print("\n" + "-" * 60)
    print("CURRENT PIPELINE STATUS")
    print("-" * 60)
    if not status:
        print("Ekho no run hoyeche (status file nai / khali).")
    for step in range(1, 6):
        entry = status.get(str(step), {})
        state = entry.get("state", "not started")
        when = entry.get("updated_at", "")
        output = entry.get("output", "")
        line = f"Step {step}: {state}"
        if when:
            line += f"  ({when})"
        if output:
            line += f"\n         output: {output}"
        print(line)
    print("-" * 60)


# =====================================================================
# Read-only validation
# =====================================================================
STEP3_CANONICAL = [
    "Product ID", "Seller Code", "Name (English)",
    "Highlights(English)", "Highlights(Bengali)",
    "Description (English)", "Description (Bengali)",
]


def norm_header(value):
    if value is None:
        return ""
    s = str(value).strip().lower()
    return "".join(ch for ch in s if ch.isalnum())


def validate_file():
    path = clean_input("Excel file er full path dao: ")
    if not os.path.isfile(path):
        print(f"File pawa jayni: {path}")
        return
    print()
    try:
        wb = pd.read_excel(path, sheet_name=None, header=None, nrows=2)
    except Exception as e:
        print(f"File open kora jay ni ({e}).")
        return

    print(f"File OK: {path}")
    for sheet_name, df in wb.items():
        if df.empty:
            print(f"  Sheet '{sheet_name}': (khali)")
            continue
        headers = [norm_header(h) for h in df.iloc[0].tolist()]
        cols = ", ".join(str(h) for h in df.iloc[0].tolist() if str(h) != "nan")
        print(f"  Sheet '{sheet_name}': {len(headers)} column (approx, 1st row = header)")
        if cols:
            print(f"    headers: {cols}")

    # Known file types check
    pid_norm = norm_header("Product ID")
    img_headers = [h for h in wb[next(iter(wb))].iloc[0].tolist() if h and "image" in norm_header(h)]
    first_row = wb[next(iter(wb))].iloc[0].tolist()
    norm_first = [norm_header(h) for h in first_row]
    has_pid = pid_norm in norm_first
    if has_pid and img_headers:
        print("\n  --> Mone hocche MAPPING file (Product ID + Image columns ache).")
    elif has_pid:
        matched = [c for c in STEP3_CANONICAL if norm_header(c) in norm_first]
        print("\n  --> Mone hocche 'Basic combine' type file.")
        print(f"      Expected 7 column er moddhe pawa geche: {len(matched)}/7")
        missing = [c for c in STEP3_CANONICAL if norm_header(c) not in norm_first]
        if missing:
            print(f"      Missing: {missing}")
    else:
        print("\n  --> Product ID column pawa jay ni row 1 e.")

    # Cross-check 2 ta file er Product ID (optional)
    other = clean_input("\nAro ekta file er sathe Product ID cross-check korte chan? Path dao (Enter = na): ")
    if other:
        try:
            if not os.path.isfile(other):
                print(f"File pawa jay ni: {other}")
                return
            a = set()
            for sheet in pd.read_excel(path, sheet_name=None, dtype=str).values():
                headers = [norm_header(h) for h in list(sheet.columns)]
                if pid_norm in headers:
                    a |= set(
                        normalize_pid(v) for v in sheet.iloc[:, headers.index(pid_norm)].dropna()
                    )
            b = set()
            for sheet in pd.read_excel(other, sheet_name=None, dtype=str).values():
                headers = [norm_header(h) for h in list(sheet.columns)]
                if pid_norm in headers:
                    b |= set(
                        normalize_pid(v) for v in sheet.iloc[:, headers.index(pid_norm)].dropna()
                    )
            both = a & b
            print(f"\n  File A: {len(a)} unique Product ID | File B: {len(b)} unique")
            print(f"  Match: {len(both)}")
            print(f"  A-te ache kintu B-te nai: {len(a - b)}")
            print(f"  B-te ache kintu A-te nai: {len(b - a)}")
        except Exception as e:
            print(f"Cross-check korte somossa: {e}")


# =====================================================================
# File auto-detection (fuzzy matching)
# =====================================================================
import re as _re

# Mapping file: ei keyword gulo er kono ekta thakle match (case-insensitive)
MAPPING_KEYWORDS = ["image", "mapping", "maping", "link", "bakaul"]
# Basic combine: duita word-i thakte hobe (kono order e)
COMBINE_KEYWORDS = {"basic", "combine"}
# Intermediate output files: keyword -> kon step er output
INTERMEDIATE_KEYWORDS = {
    "http_url": "Step 1 output (image_http_url_add.xlsx)",
    "h & d": "Step 3 output (Basic combine H & D.xlsx)",
    "h&d": "Step 3 output (Basic combine H & D.xlsx)",
    "image_replaced": "Step 4 output (..._image_replaced.xlsx)",
    "split": "Step 5 split folder (..._split)",
}


def _name_tokens(name):
    """Filename ke choto-hater word token e vag kore (lowercase)."""
    return set(t for t in _re.split(r"[^a-z0-9&]+", str(name).lower()) if t)


def _squashed(name):
    """Space/punct chara ek line e squash kora, 'h&d' match er jonno."""
    s = _re.sub(r"\s+", "", str(name).lower())
    return s


def fuzzy_scan_xlsx(folders, exclude_self=False):
    """Folder gulo scan kore shob .xlsx file er (fullpath, filename) list dao.
    Temp file (~$...) ba agent er nijer output baad."""
    found = []
    for folder in folders:
        if not os.path.isdir(folder):
            continue
        for f in sorted(os.listdir(folder)):
            if not f.lower().endswith(".xlsx") or f.startswith("~$"):
                continue
            if exclude_self and os.path.abspath(os.path.join(folder, f)) == os.path.abspath(sys.argv[0]):
                continue
            found.append((os.path.join(folder, f), f))
    return found


def match_mapping_files(files):
    """Mapping file candidates: naam e TEEN TA keyword-i thakte hobe (AND logic):
      1. 'mapping' ba 'maping' ba 'bakaul'
      2. 'image'
      3. 'link' (substring - 'links' o match korbe)
    Multiple match hole most recently modified ta first.
    Example match: 'image_links_mapping (21.09.2026) Bakaul.xlsx'"""
    scored = []
    for full, fname in files:
        toks = _name_tokens(fname)
        sq = _squashed(fname)  # squash kora naam - 'links' e 'link' check er jonno

        has_mapping = any(k in toks for k in ("mapping", "maping", "bakaul"))
        has_image = "image" in toks
        has_link = "link" in sq  # 'link' ba 'links' duitai dhore

        if not (has_mapping and has_image and has_link):
            continue  # tini ta keyword-er ekta-o miss korle baad

        try:
            mt = os.path.getmtime(full)
        except OSError:
            mt = 0
        scored.append((-mt, full, fname))  # newest age (mtime descending)

    scored.sort()
    return [(full, fname) for _, full, fname in scored]


def match_combine_files(files):
    """Basic combine candidates: 'basic' AND 'combine' duitai thakte hobe.
    Ranked: extra token kom mane age (shudhu 'Basic combine.xlsx' type file
    upore), tarpor newest age."""
    scored = []
    for full, fname in files:
        toks = _name_tokens(fname)
        if not {"basic", "combine"} <= toks:
            continue
        extra = len(toks - {"basic", "combine"})
        try:
            mt = os.path.getmtime(full)
        except OSError:
            mt = 0
        scored.append((extra, -mt, full, fname))
    scored.sort()
    return [(full, fname) for _, _, full, fname in scored]


def match_intermediate(files, kind):
    """kind: 'http_url' | 'h&d' | 'image_replaced' | 'split'
    Keyword squash-name e substring hisebe check kora hoy."""
    out = []
    for full, fname in files:
        sq = _squashed(fname)
        if kind in sq:
            out.append((full, fname))
    return out


def pick_file(candidates, prompt_header, allow_new=True, default_newest=True):
    """Candidate list theke file pick korano.
    - 1 ta match: "Auto-detected: X - use korbo?" (Enter = yes, path dile sheta)
    - beshi match: numbered list (newest first), number dao / path dao
    - kono match nai: full path jiggesh (fallback)
    Return: full path ba None (cancel)."""
    def _mtime(full):
        try:
            return os.path.getmtime(full)
        except OSError:
            return 0

    # Callers already rank candidates (specificity first, then newest)
    candidates = list(candidates)

    if not candidates:
        if not allow_new:
            return None
        print(f"\n{prompt_header}")
        print("Kono matching file pawa jay ni - full path likhe din.")
        path = clean_input("Full path: ")
        return path or None

    if len(candidates) == 1:
        full, fname = candidates[0]
        print(f"\n{prompt_header}")
        ans = clean_input(f"Auto-detected: {fname} - use korbo? [Enter = Yes, ba notun path dao]: ")
        if not ans:
            return full
        return ans

    print(f"\n{prompt_header}")
    print("Ei file gula pawa geche - kon ta use korbo? (best match age):")
    for i, (full, fname) in enumerate(candidates, start=1):
        mtime = datetime.fromtimestamp(_mtime(full)).strftime("%Y-%m-%d %H:%M")
        print(f"  {i:2}. {fname}   (modified: {mtime})")
    ans = clean_input("Number dao (ba notun path type koro) [Enter = 1 no.]: ")
    if not ans:
        return candidates[0][0]  # Enter = best-ranked candidate
    if ans.isdigit() and 1 <= int(ans) <= len(candidates):
        return candidates[int(ans) - 1][0]
    return ans  # path hisebe treat korbe (validation pore caller kore)


def detect_mapping_file():
    """Mapping file auto-detect.
    AUTO_MODE: Data folder scan kore AND-logic match (image+link+mapping),
    newest file automatically select kore - kono jiggesh na kore.
    Interactive mode: pick_file() diye user select korte parbe."""
    data_folder = os.path.join(BASE_DIR, "Data")
    files = fuzzy_scan_xlsx([data_folder, BASE_DIR])
    candidates = match_mapping_files(files)

    if AUTO_MODE:
        if candidates:
            full, fname = candidates[0]
            print(f"[auto] Mapping file detected: {fname}")
            if len(candidates) > 1:
                print(f"       ({len(candidates)} ta match, newest ta select kora holo)")
            return full
        # config e hardcoded path ache kina check (fallback)
        cfg_path = AUTO_CONFIG.get("mapping_file", "").strip()
        if cfg_path and os.path.isfile(cfg_path):
            print(f"[auto] Mapping file (config fallback): {cfg_path}")
            return cfg_path
        print(f"[auto] ERROR: Data folder e kono mapping file pawa jay ni!")
        print(f"       '{data_folder}' e 'image', 'link', 'mapping' keyword wala .xlsx rakho.")
        return None

    return pick_file(candidates, "Mapping file detection:")


def detect_combine_file():
    d_drive = get_d_drive_output_folder()
    files = fuzzy_scan_xlsx([d_drive, BASE_DIR, os.path.join(BASE_DIR, "Data")])
    return pick_file(match_combine_files(files), "Basic combine file detection:")


def detect_cleaned_file():
    """'Basic combine cleaned.xlsx' (Claude/H&D cleaned merge-back output)
    auto-detect kore. AUTO_MODE hole newest matching file silently pick hoy."""
    d_drive = get_d_drive_output_folder()
    folders = [d_drive, BASE_DIR, os.path.join(BASE_DIR, "Data")]
    files = fuzzy_scan_xlsx(folders)
    candidates = match_intermediate(files, "cleaned")
    candidates = sorted(
        candidates,
        key=lambda x: -os.path.getmtime(x[0]) if os.path.isfile(x[0]) else 0,
    )
    if AUTO_MODE:
        if candidates:
            full, fname = candidates[0]
            print(f"[auto] Cleaned content file detected: {fname}")
            if len(candidates) > 1:
                print(f"       ({len(candidates)} ta match, newest ta select kora holo)")
            return full
        print("[auto] ERROR: 'cleaned' naam wala kono file pawa jay ni (Basic combine cleaned.xlsx)!")
        return None
    return pick_file(candidates, "Cleaned content file (Basic combine cleaned.xlsx) detection:")


def detect_template_file():
    """'downloads' subfolder (Step 2 AUTOMATED / menu 11 er raw per-batch
    seller export download) theke ekta .xlsx template khuje ber kore.
    Ei template-i final upload file-er 'shell' hisebe use hobe (sheets,
    dropdown validation, formatting hubohu rakhar jonno) - batch number
    matter kore na, shob template e shomo structure thake."""
    downloads_dir = os.path.join(get_d_drive_output_folder(), "downloads")
    if not os.path.isdir(downloads_dir):
        return None
    files = sorted(
        f for f in os.listdir(downloads_dir)
        if f.lower().endswith(".xlsx") and not f.startswith("~$")
    )
    if not files:
        return None
    return os.path.join(downloads_dir, files[0])


def detect_intermediate(kind, prefer_folder=None, exclude=None):
    """kind: 'http_url' | 'h&d' | 'image_replaced'.
    exclude: eei substring-gula jei filename e ache sheigulo baad."""
    d_drive = get_d_drive_output_folder()
    folders = [prefer_folder, d_drive, os.path.join(BASE_DIR, "Data"), BASE_DIR]
    folders = [f for f in folders if f]
    files = fuzzy_scan_xlsx(folders)
    label = INTERMEDIATE_KEYWORDS.get(kind, kind)
    cands = [
        (f, n) for (f, n) in match_intermediate(files, kind)
        if not exclude or not any(x in _squashed(n) for x in exclude)
    ]
    return pick_file(cands, f"{label} detection:")


def detect_split_folder():
    """Step 5 split folder (*_split) detection - handoff check er jonno."""
    d_drive = get_d_drive_output_folder()
    folders = []
    for d in (d_drive, os.path.join(BASE_DIR, "Data"), BASE_DIR):
        if os.path.isdir(d):
            for f in sorted(os.listdir(d)):
                full = os.path.join(d, f)
                if os.path.isdir(full) and f.lower().endswith("_split"):
                    folders.append((full, f))
    return pick_file(folders, "Split folder detection:")


def load_base_url():
    """Last-used Base URL porhe - status file er top-level entry,
    ba Step 1 er nijer entry (base_url field) theke."""
    st = load_status()
    top = st.get("base_url", {})
    if isinstance(top, dict) and top.get("value"):
        return top["value"]
    s1 = st.get("1", {})
    if isinstance(s1, dict) and s1.get("base_url"):
        return s1["base_url"]
    return ""


def ask_base_url():
    """Base URL jiggesh - ager stored URL default hisebe dekhay (Enter = same).
    AUTO_MODE=True hole config file er base_url use kore, na thakle stored URL use kore."""
    if AUTO_MODE:
        cfg_url = AUTO_CONFIG.get("base_url", "").strip()
        if cfg_url:
            print(f"Base URL: {cfg_url}  [auto - config file]")
            return cfg_url
        # config e nai, stored theke nebo
        prev = load_base_url()
        if prev:
            print(f"Base URL: {prev}  [auto - last used]")
            return prev
        print("ERROR: auto mode e base_url config e nai ebong stored-o nai!")
        print("       agent_config.json er 'base_url' field e URL bosao.")
        sys.exit(1)
    prev = load_base_url()
    if prev:
        ans = clean_input(f"Base URL dao [Enter = ager ta: {prev}]: ")
        return ans or prev
    return clean_input("Base URL dao (e.g. https://sl-dev-s3.s3.amazonaws.com/product/): ")



def save_base_url(url):
    if url:
        update_status("base_url", value=url, label="Last used Base URL")


# =====================================================================
# Output-post-parsing helpers (status numbers ber korar jonno)
# =====================================================================
def count_split_parts(folder):
    if not os.path.isdir(folder):
        return 0
    return len([f for f in os.listdir(folder) if f.endswith(".xlsx")])


def needs_review_count(path):
    try:
        xl = pd.ExcelFile(path)
        if "Needs Review" in xl.sheet_names:
            df = pd.read_excel(path, sheet_name="Needs Review")
            return len(df)
    except Exception:
        pass
    return 0


# =====================================================================
# Steps (original function + agent wrapper)
# =====================================================================
def run_step1(mapping_path=None, base_url=None):
    """Step 1: Image Links Mapping file e base URL bosiye
    image_http_url_add.xlsx banano."""
    print("\n--- STEP 1: Image URL add ---")
    if mapping_path is None:
        mapping_path = detect_mapping_file()
    if not mapping_path or not os.path.isfile(mapping_path):
        print(f"File pawa jay ni: {mapping_path} - Step 1 cancel.")
        return None
    if base_url is None:
        base_url = ask_base_url()
    if not base_url:
        print("Base URL dorkar - Step 1 cancel.")
        return None
    save_base_url(base_url)

    output_dir = get_d_drive_output_folder()
    expected_out = os.path.join(output_dir, "image_http_url_add.xlsx")
    if not confirm_overwrite_if_exists(expected_out, "Step 1 output"):
        return None

    try:
        out = step1_add_http(mapping_path, base_url, output_dir=output_dir)
    except Exception as e:
        print(f"Step 1 fail: {e}")
        return None

    update_status(1, state="done", input=mapping_path, output=out, base_url=base_url)
    log(f"Step 1 done -> {out}")
    return out


def _verify_basic_combine(path):
    """Basic combine.xlsx er column verify kore (7/7 match).
    Return: (matched_count, missing_list) ba None (file porte pari nai)."""
    try:
        first = pd.read_excel(path, header=None, nrows=1)
    except Exception as e:
        print(f"File open kora jay ni ({e}) - Step 2 verify fail.")
        return None
    norm_first = [norm_header(h) for h in first.iloc[0].tolist()]
    matched = [c for c in STEP3_CANONICAL if norm_header(c) in norm_first]
    missing = [c for c in STEP3_CANONICAL if norm_header(c) not in norm_first]
    return matched, missing


def run_step2():
    """Step 2: Basic combine.xlsx confirm kora.
    Notun: file ta auto-detect hoy (fuzzy match), ar jodi na thake,
    automated export script (menu 11) offer kora hoy."""
    print("\n--- STEP 2: Basic combine.xlsx (automated export available) ---")
    print("Ei step ta ekhon AUTOMATED - 'Multi Seller Bulk Update' page theke")
    print("export download + combine hoy menu 11 (automated export) diye.")

    detected = detect_combine_file()
    if detected and os.path.isfile(detected):
        res = _verify_basic_combine(detected)
        if res:
            matched, missing = res
            print(f"Verify: {len(matched)}/7 expected column pawa geche.")
            if missing:
                print(f"Missing column (Step 3 te skip hobe): {missing}")
            if not ask_yes_no("Ei file take Step 2 complete hisebe mark korte chan?"):
                print("Ei file baad - menu 11 diye notun export nite paren.")
                return None
            update_status(2, state="done", output=detected, matched_columns=f"{len(matched)}/7")
            log(f"Step 2 confirmed: {detected} ({len(matched)}/7 columns)")
            return detected
    else:
        print("\nBasic combine.xlsx pawa jay ni.")
        if ask_yes_no("Automated export script chalate chan (Playwright browser khulbe, OTP apni dibe)?"):
            return run_automated_step2()
        print("Automated skip holo - manual kaj shesh hole abar try koro.")
    return None


def run_automated_step2(mapping_path=None):
    """seller_bulk_export_automation.py ke subprocess hisebe chalay.
    Interactive (OTP + credentials apni dibe). Shesh hole Basic combine.xlsx
    verify kore Step 2 complete kore.
    mapping_path dile --mapping hisebe pass hobe (jate script abar
    jiggesh na kore); na dile script nijei auto-detect + memory use korbe."""
    if not os.path.isfile(STEP2_SCRIPT):
        print(f"Script pawa jay ni: {STEP2_SCRIPT}")
        return None
    try:
        import playwright  # noqa: F401
    except ImportError:
        print("Playwright install kora nai. Age ei command gulo run koro:")
        print("    pip install playwright")
        print("    python -m playwright install chromium")
        return None
    print("\n--- STEP 2 AUTOMATED: seller_bulk_export_automation.py ---")
    print("Browser khulbe (not headless). Login + OTP apni browser e dibe.")
    print("Shob batch export + combine shesh hole ei menu firbe.")
    if not ask_yes_no("Shuru korbo?"):
        return None

    d_drive_dir = get_d_drive_output_folder()
    outdir = os.path.join(d_drive_dir, "downloads")
    os.makedirs(outdir, exist_ok=True)
    out_path = os.path.join(d_drive_dir, "Basic combine.xlsx")

    cmd = [
        sys.executable, STEP2_SCRIPT, 
        "--outdir", outdir,
        "--combine-output", out_path
    ]
    if mapping_path and os.path.isfile(mapping_path):
        cmd += ["--mapping", mapping_path]
    print(f"\nCommand: {os.path.basename(sys.executable)} seller_bulk_export_automation.py --outdir '{outdir}' --combine-output '{out_path}'")
    print("(Script nijei credentials + OTP jiggesh korbe - ekhanei type korte parben.)\n")
    try:
        result = subprocess.run(cmd, cwd=BASE_DIR)
    except Exception as e:
        print(f"Script chalate somossa: {e}")
        return None
    if result.returncode != 0:
        print(f"\nAutomated export FAIL hoiche (exit code {result.returncode}).")
        return None
    if not os.path.isfile(out_path):
        detected = detect_combine_file()
        if not detected or not os.path.isfile(detected):
            print("Combine output pawa jay ni - manually path din.")
            p = clean_input("'Basic combine.xlsx' er full path: ")
            if not p or not os.path.isfile(p):
                return None
            out_path = p
        else:
            out_path = detected

    res = _verify_basic_combine(out_path)
    if res is None:
        return None
    matched, missing = res
    print(f"\nVerify: {len(matched)}/7 expected column pawa geche: {out_path}")
    if missing:
        print(f"Missing column (Step 3 te skip hobe): {missing}")
        if not ask_yes_no("Tao Step 2 complete korbe?"):
            return None
    update_status(2, state="done", output=out_path, matched_columns=f"{len(matched)}/7", mode="automated")
    log(f"Step 2 automated done: {out_path} ({len(matched)}/7 columns)")
    return out_path


def run_step2_manual():
    """Purano manual flow: file nijei banaya shesh kore path dibe."""
    print("\n--- STEP 2 (MANUAL MODE) - Basic combine.xlsx ---")
    print("File nijei banaya shesh kore Enter chapun - ami detect + verify korbo.")
    input("Kaj shesh hole Enter chapun continue korte...")
    path = detect_combine_file()
    if not path or not os.path.isfile(path):
        print(f"File pawa jay ni: {path} - Step 2 verify fail.")
        return None
    res = _verify_basic_combine(path)
    if res is None:
        return None
    matched, missing = res
    print(f"\nVerify: {len(matched)}/7 expected column pawa geche.")
    if not matched:
        print("Kono expected column nai - file ta check koro.")
        return None
    if missing:
        print(f"Missing column (skip hobe Step 3 te): {missing}")
        if not ask_yes_no("Tao Step 2 complete korbe?"):
            return None
    update_status(2, state="done", output=path, matched_columns=f"{len(matched)}/7")
    log(f"Step 2 confirmed: {path} ({len(matched)}/7 columns)")
    return path


def run_step3(basic_combine_path=None):
    """Step 3: Basic combine.xlsx -> Basic combine H & D.xlsx"""
    print("\n--- STEP 3: Trim columns ---")
    if basic_combine_path is None:
        basic_combine_path = detect_combine_file()
    if not os.path.isfile(basic_combine_path):
        print(f"File pawa jay ni: {basic_combine_path} - Step 3 cancel.")
        return None

    output_dir = get_d_drive_output_folder()
    expected_out = os.path.join(output_dir, "Basic combine H & D.xlsx")
    if not confirm_overwrite_if_exists(expected_out, "Step 3 output"):
        return None

    try:
        out = step3_trim(basic_combine_path, output_path=expected_out)
    except Exception as e:
        print(f"Step 3 fail: {e}")
        return None

    update_status(3, state="done", input=basic_combine_path, output=out)
    log(f"Step 3 done -> {out}")
    return out


def run_step4(trimmed_path=None, mapping_out=None):
    """Step 4: image URL gula HTML column e merge kora."""
    print("\n--- STEP 4: Merge images into HTML ---")
    if trimmed_path is None:
        trimmed_path = detect_intermediate("h&d", exclude=["image_replaced"])
    if not trimmed_path or not os.path.isfile(trimmed_path):
        print(f"File pawa jay ni: {trimmed_path} - Step 4 cancel.")
        return None
    if mapping_out is None:
        mapping_out = detect_intermediate("http_url")
    if not mapping_out or not os.path.isfile(mapping_out):
        print(f"File pawa jay ni: {mapping_out} - Step 4 cancel.")
        return None

    base, ext = os.path.splitext(os.path.basename(trimmed_path))
    output_dir = get_d_drive_output_folder()
    expected_out = os.path.join(output_dir, f"{base}_image_replaced.xlsx")
    if not confirm_overwrite_if_exists(expected_out, "Step 4 output"):
        return None

    try:
        out = step4_merge(trimmed_path, mapping_out, output_dir=output_dir)
    except Exception as e:
        print(f"Step 4 fail: {e}")
        return None

    nr = needs_review_count(out)
    update_status(4, state="done", input=f"{trimmed_path} + {mapping_out}", output=out,
                  needs_review_rows=nr)
    log(f"Step 4 done -> {out} (Needs Review rows: {nr})")
    return out


def run_step5(merged_path=None, header_rows=None, chunk_size=None):
    """Step 5: merged file ke choto choto part e split kora."""
    print("\n--- STEP 5: Split for Claude ---")
    if merged_path is None:
        merged_path = detect_intermediate("image_replaced")
    if not merged_path or not os.path.isfile(merged_path):
        print(f"File pawa jay ni: {merged_path} - Step 5 cancel.")
        return None

    if header_rows is None:
        if AUTO_MODE:
            header_rows = int(AUTO_CONFIG.get("header_rows", 1))
            print(f"Header rows: {header_rows}  [auto - config]")
        else:
            header_rows = int(clean_input("Koyta header row? (usually 1) [Enter = 1]: ", default="1"))
    if chunk_size is None:
        if AUTO_MODE:
            chunk_size = int(AUTO_CONFIG.get("chunk_size", 3500))
            print(f"Chunk size: {chunk_size}  [auto - config]")
        else:
            chunk_size = int(clean_input("Protita split file e koyta row thakbe? (e.g. 3500) [Enter = 3500]: ", default="3500"))


    base_no_ext = os.path.splitext(os.path.basename(merged_path))[0]
    output_dir = get_d_drive_output_folder()
    outdir = os.path.join(output_dir, f"{base_no_ext}_split")

    existing_parts = count_split_parts(outdir)
    if existing_parts > 0:
        print(f"\nWARNING: split folder e already {existing_parts} ta file ache: {outdir}")
        if not ask_yes_no("Overwrite korte chan?"):
            print("Skip kora holo - apnar kono file change hoy nai.")
            return None

    try:
        step5_split(merged_path, header_rows, chunk_size, output_dir=output_dir)
    except Exception as e:
        print(f"Step 5 fail: {e}")
        return None

    parts = count_split_parts(outdir)
    update_status(5, state="done", input=merged_path, output=outdir,
                  parts=parts, header_rows=header_rows, chunk_size=chunk_size)
    log(f"Step 5 done -> {outdir} ({parts} parts)")
    return outdir


# =====================================================================
# Claude handoff
# =====================================================================
HANDOFF_PROMPT_FILENAME = "H&D_improved_prompt.md"


def run_handoff_check():
    print("\n--- CLAUDE HANDOFF CHECK ---")
    folder = detect_split_folder()
    if not folder or not os.path.isdir(folder):
        print(f"Folder pawa jay ni: {folder}")
        return

    parts = sorted(
        f for f in os.listdir(folder) if f.endswith(".xlsx")
    )
    if not parts:
        print("Oi folder e kono .xlsx part file nai.")
        return

    header_rows = int(clean_input("Koyta header row chilo? (usually 1) [Enter = 1]: ", default="1"))

    print(f"\n{len(parts)} ta part file pawa geche:")
    manifest_lines = [
        "# Claude Handoff Manifest",
        "",
        f"Generated: {ts()}",
        f"Split folder: `{folder}`",
        "",
        "| # | File | Data rows |",
        "|---|------|-----------|",
    ]
    total_rows = 0
    from openpyxl import load_workbook as _load_wb
    for i, part in enumerate(parts, start=1):
        full = os.path.join(folder, part)
        rows = 0
        try:
            wb_p = _load_wb(full, read_only=True)
            rows = max(0, wb_p.active.max_row - header_rows)
            wb_p.close()
        except Exception:
            pass
        total_rows += rows
        print(f"  {i:2}. {part}  ({rows} data rows)")
        manifest_lines.append(f"| {i} | `{part}` | {rows} |")
    manifest_lines += [
        "",
        f"**Total data rows:** {total_rows}",
        "",
        "## Upload checklist",
        "",
        "1. [ ] Protita part file Claude-e upload koro, sathe `H&D_improved_prompt.md`",
        "2. [ ] Claude-er output thik kore nin (manual review)",
        "3. [ ] Shob cleaned/updated file ekta alada folder e joma koro",
        "4. [ ] Agent er 'Merge back cleaned files' option use koro",
        "",
    ]

    prompt_path = os.path.join(BASE_DIR, HANDOFF_PROMPT_FILENAME)
    if os.path.isfile(prompt_path):
        print(f"Prompt file pawa geche: {prompt_path}")
        manifest_lines.insert(4, f"Prompt file: `{HANDOFF_PROMPT_FILENAME}` (found)")
    else:
        print(f"WARNING: prompt file pawa jay ni: {prompt_path}")
        print("         Claude-e deyar age oita ready korte hobe.")
        manifest_lines.insert(4, f"Prompt file: `{HANDOFF_PROMPT_FILENAME}` (NOT FOUND - prepare it!)")

    manifest_path = os.path.join(folder, "handoff_manifest.md")
    with open(manifest_path, "w", encoding="utf-8") as f:
        f.write("\n".join(manifest_lines))
    print(f"\nManifest likha holo: {manifest_path}")


# =====================================================================
# Merge back cleaned files
# =====================================================================
def run_merge_back():
    print("\n--- MERGE BACK CLEANED FILES ---")
    
    # Auto-mode er jonno default folder path
    default_folder = os.path.join(BASE_DIR, "H&D_Processed_Input")
    
    if AUTO_MODE:
        if not os.path.isdir(default_folder):
            # Folder na thakle auto banaye dibe jate user pore file paste korte pare
            os.makedirs(default_folder, exist_ok=True)
        print("Claude theke pawa cleaned file gulo ei folder e rakhun:")
        print(f"-> {default_folder}")
        folder = clean_input("Cleaned files folder er full path dao: ", default=default_folder)
    else:
        print("Claude theke pawa cleaned file gulo ekta folder e rakho.")
        folder = clean_input("Cleaned files folder er full path dao [Enter = H&D_Processed_Input]: ", default=default_folder)

    if not os.path.isdir(folder):
        print(f"Folder pawa jay ni: {folder}")
        return

    files = sorted(f for f in os.listdir(folder) if f.endswith(".xlsx") and not f.startswith("~$"))
    if not files:
        print(f"Oi folder e kono .xlsx file nai: {folder}")
        return

    print(f"\n{len(files)} ta file pawa geche:")
    for i, f in enumerate(files, start=1):
        print(f"  {i:2}. {f}")

    header_rows = int(clean_input("Koyta header row? (usually 1) [Enter = 1]: ", default="1"))

    # ---- Columns same kina check ----
    header_sets = {}
    for f in files:
        full = os.path.join(folder, f)
        try:
            df = pd.read_excel(full, header=header_arg(header_rows), nrows=0)
            header_sets[f] = list(df.columns)
        except Exception as e:
            print(f"  {f} open kora jay ni ({e}) - merge back cancel.")
            return

    ref_file, ref_cols = next(iter(header_sets.items()))
    mismatch = [f for f, cols in header_sets.items() if cols != ref_cols]
    if mismatch:
        print("\nWARNING: ei file gulor column same na:")
        for f in mismatch:
            print(f"  {f}: {header_sets[f]}")
        print(f"Reference ({ref_file}): {ref_cols}")
        if not ask_yes_no("Tao merge back korbe? (mismatch column drop hobe)"):
            print("Merge back cancel.")
            return

    out_name = clean_input("Merged output file er naam dao [Enter = 'Basic combine cleaned.xlsx']: ", default="Basic combine cleaned.xlsx")
    if not out_name:
        print("Naam dorkar - merge back cancel.")
        return
    if not out_name.lower().endswith(".xlsx"):
        out_name += ".xlsx"
    out_path = os.path.join(get_d_drive_output_folder(), out_name)
    if not confirm_overwrite_if_exists(out_path, "merge back output"):
        return

    # ---- Read + concat (Product ID string hisebe, .0 chara) ----
    frames = []
    for f in files:
        full = os.path.join(folder, f)
        df = pd.read_excel(full, header=header_arg(header_rows))
        # Product ID column ke string-based clean kora (.0 remove etc.)
        pid_norm = norm_header("Product ID")
        norm_cols = [norm_header(c) for c in df.columns]
        if pid_norm in norm_cols:
            pid_col = df.columns[norm_cols.index(pid_norm)]
            df[pid_col] = df[pid_col].map(normalize_pid)
        frames.append(df)

    combined = pd.concat(frames, ignore_index=True)
    combined = combined.where(pd.notna(combined), None)

    try:
        with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
            combined.to_excel(writer, index=False, sheet_name="Data")
            # Product ID column ke text hisebe force kora (Excel e '2261883' dekhabe,
            # kono numeric conversion hobe na)
            ws = writer.sheets["Data"]
            pid_norm = norm_header("Product ID")
            norm_cols = [norm_header(c) for c in combined.columns]
            if pid_norm in norm_cols:
                col_idx = norm_cols.index(pid_norm) + 1  # 1-based, A=1
                from openpyxl.utils import get_column_letter
                letter = get_column_letter(col_idx)
                for row in range(2, ws.max_row + 1):
                    c = ws.cell(row=row, column=col_idx)
                    if c.value is not None:
                        c.value = str(c.value)
                        c.number_format = "@"
    except Exception as e:
        print(f"Merge back fail: {e}")
        return

    print(f"\nDone! {len(frames)} ta file theke {len(combined)} ta row merge holo.")
    print(f"Output: {out_path}")
    log(f"Merge back: {len(frames)} files -> {out_path} ({len(combined)} rows)")


def header_arg(header_rows):
    """split_excel.py er same convention: 1 header row -> header=0,
    2 -> header=[0,1] ityadi."""
    return 0 if header_rows == 1 else list(range(header_rows))


# =====================================================================
# Step 6: Content update + fill into official template + split
# =====================================================================
STEP6_CONTENT_COLUMNS = [
    "Highlights(English)", "Highlights(Bengali)",
    "Description (English)", "Description (Bengali)",
]

# Rows-per-file default - experience onujayi ~15k row porjonto 2 ta
# file e (mane ~7500/file) safely, fast update hoy. agent_config.json
# er "step6": {"max_rows_per_file": N, "min_files": N} diye override
# kora jay.
STEP6_MAX_ROWS_PER_FILE_DEFAULT = 7500
STEP6_MIN_FILES_DEFAULT = 2


def find_downloads_template():
    """get_d_drive_output_folder()/downloads folder theke pehla .xlsx
    file ke 'template shell' hisebe pick kore - eita ekta official
    downloaded file (header row + dropdown validation + DropdownData/
    Policy sheet shoho), splitted update file gula ei shell er upor
    base kore banano hobe jate dropdown/format thik thake (batch
    number matter kore na - shob template-i same structure)."""
    downloads_dir = os.path.join(get_d_drive_output_folder(), "downloads")
    if not os.path.isdir(downloads_dir):
        return None
    files = sorted(
        f for f in os.listdir(downloads_dir)
        if f.lower().endswith(".xlsx") and not f.startswith("~$")
    )
    if not files:
        return None
    return os.path.join(downloads_dir, files[0])


def _find_product_update_sheet(wb):
    """'Product Update' sheet khoje (naam sob shomoy hubohu na o hote
    pare), na pele first sheet fallback hisebe."""
    for name in wb.sheetnames:
        if norm_header(name) == norm_header("Product Update"):
            return name
    return wb.sheetnames[0]


def _extend_validations(ws, orig_max_row, new_max_row):
    """Dropdown (data validation) range gula orig_max_row theke
    new_max_row porjonto extend kore, jate template-er original row
    count-er baire notun row likhleo dropdown thik kaj kore. Protita
    validation-er nijer min_row thake, shudhu max_row update hoy
    (jodi shei validation-er max_row orig_max_row-er shathe mile)."""
    if new_max_row <= orig_max_row:
        return
    from openpyxl.utils import get_column_letter
    for dv in list(ws.data_validations.dataValidation):
        new_ranges = []
        changed = False
        for rng in list(dv.sqref.ranges):
            if rng.max_row == orig_max_row:
                col1 = get_column_letter(rng.min_col)
                col2 = get_column_letter(rng.max_col)
                new_ranges.append(f"{col1}{rng.min_row}:{col2}{new_max_row}")
                changed = True
            else:
                new_ranges.append(str(rng))
        if changed:
            dv.sqref = " ".join(new_ranges)


def _clear_and_fill_sheet(ws, rows_to_write, orig_max_row, min_data_row=3):
    """ws-er data area (min_data_row theke orig_max_row) er shob value
    None kore dey (format/dropdown touch kora hoy na - shudhu value
    clear), tarpor rows_to_write (protita row template column-order
    onujayi already sajano) min_data_row theke likhe dey. rows_to_write
    template-er original row count-er cheye beshi hole, baki notun
    row gulor jonno last template row theke cell style copy kore
    banano hoy (jate format e kono farak na thake), ar dropdown
    validation range o shei porjonto extend kora hoy."""
    from copy import copy as _copy

    for r in range(min_data_row, orig_max_row + 1):
        for c in range(1, ws.max_column + 1):
            ws.cell(row=r, column=c).value = None

    n = len(rows_to_write)
    for i, row_vals in enumerate(rows_to_write):
        r = min_data_row + i
        if r > orig_max_row:
            for c in range(1, ws.max_column + 1):
                src_cell = ws.cell(row=orig_max_row, column=c)
                dst_cell = ws.cell(row=r, column=c)
                dst_cell._style = _copy(src_cell._style)
        for c, val in enumerate(row_vals, start=1):
            ws.cell(row=r, column=c).value = val

    new_max_row = (min_data_row + n - 1) if n > 0 else (min_data_row - 1)
    if new_max_row > orig_max_row:
        _extend_validations(ws, orig_max_row, new_max_row)
    return new_max_row


def detect_approval_template():
    """QC approval template ('Product-approval-template.xlsx') khoje -
    ei file ta project er 'templates' folder e ekbar shob kore rakha
    thake (Seller ID / Product ID / Approval Status / Product Tags /
    Reject Reason column shoho), protibar notun kore lagena."""
    candidates = [
        os.path.join(BASE_DIR, "templates", "Product-approval-template.xlsx"),
        os.path.join(BASE_DIR, "Product-approval-template.xlsx"),
    ]
    for c in candidates:
        if os.path.isfile(c):
            return c
    try:
        for f in os.listdir(BASE_DIR):
            if "approval" in f.lower() and f.lower().endswith(".xlsx"):
                return os.path.join(BASE_DIR, f)
    except OSError:
        pass
    return None


def build_product_approval_file(product_ids, update_dir):
    """Upload-ready file e je Product ID gula gelo, shei shob ID niye
    QC-approval file banay (Approval Status = 1 shob-er jonno).
    Row-count niye kono limit/split lage na - ekta file e shob thake."""
    template_path = detect_approval_template()
    if not template_path:
        print("[warn] QC approval template pawa jayni (templates/Product-approval-template.xlsx) - approval file skip kora holo.")
        return None
    try:
        wb = load_workbook(template_path, data_only=False)
    except Exception as e:
        print(f"[warn] Approval template open kora jay ni: {e}")
        return None
    ws = wb[wb.sheetnames[0]]
    header = [ws.cell(1, c).value for c in range(1, ws.max_column + 1)]
    header_norm = [norm_header(h) for h in header]
    pid_norm = norm_header("Product ID")
    appr_norm = norm_header("Approval Status")
    if pid_norm not in header_norm or appr_norm not in header_norm:
        print("[warn] Approval template e 'Product ID'/'Approval Status' column pawa jayni - approval file skip kora holo.")
        return None
    pid_col = header_norm.index(pid_norm)
    appr_col = header_norm.index(appr_norm)
    n_cols = len(header)

    rows_to_write = []
    for pid in product_ids:
        r = [None] * n_cols
        r[pid_col] = pid
        r[appr_col] = 1
        rows_to_write.append(r)

    orig_max_row = ws.max_row
    _clear_and_fill_sheet(ws, rows_to_write, orig_max_row, min_data_row=2)

    out_path = os.path.join(update_dir, "Product-approval.xlsx")
    wb.save(out_path)
    print(f"Wrote {out_path}  ({len(rows_to_write)} rows, Approval Status = 1)")
    return out_path


def run_step6_update_and_split(combine_path=None, cleaned_path=None, template_path=None):
    """STEP 6: 'Basic combine.xlsx' (Workings sheet) er 4 ta content
    column 'Basic combine cleaned.xlsx' diye update kore, tarpor SHOB
    column shoho (full refresh) shei updated data ke ekta official
    downloaded template ('downloads' folder er jekono 1 ta batch file
    - header/dropdown/format shoho) er 'Product Update' sheet e
    bosiye dey, purono batch-er data completely clear kore. Onek row
    hole (max_rows_per_file er beshi) template-er multiple copy e
    split hoy, protita copy-i pura format/dropdown shoho thake (min 2
    file, jate ekekbar update dile shomoy kom lage, timeout/error na
    hoy)."""
    print("\n--- STEP 6: Update content + fill into official template + split ---")

    if combine_path is None:
        combine_path = detect_combine_file()
    if not combine_path or not os.path.isfile(combine_path):
        print("'Basic combine.xlsx' pawa jay ni - Step 6 cancel.")
        return None

    if cleaned_path is None:
        cleaned_path = detect_cleaned_file()
    if not cleaned_path or not os.path.isfile(cleaned_path):
        print("'Basic combine cleaned.xlsx' (cleaned content file) pawa jay ni - Step 6 cancel.")
        print("Age menu 9 ('Merge back cleaned files') diye Claude-processed part gula merge koro.")
        return None

    if template_path is None:
        template_path = find_downloads_template()
    if not template_path or not os.path.isfile(template_path):
        print("Official template file pawa jay ni ('downloads' folder e kono .xlsx nai) - Step 6 cancel.")
        return None

    print(f"Main file (shob column) : {combine_path}")
    print(f"Cleaned content file    : {cleaned_path}")
    print(f"Template shell (format) : {template_path}")

    # ---- Main file er 'Workings' sheet load kora ----
    try:
        wb = load_workbook(combine_path, data_only=True)
    except Exception as e:
        print(f"'{combine_path}' open kora jay ni: {e}")
        return None
    sheet_name = "Workings" if "Workings" in wb.sheetnames else wb.sheetnames[0]
    if sheet_name != "Workings":
        print(f"[warn] 'Workings' sheet pawa jayni - '{sheet_name}' sheet use kora hocche.")
    ws = wb[sheet_name]
    rows = list(ws.iter_rows(values_only=True))
    if not rows:
        print("Main file er sheet khali - Step 6 cancel.")
        return None
    header = list(rows[0])
    data_rows = [list(r) for r in rows[1:]]

    norm_header_list = [norm_header(h) for h in header]
    pid_norm = norm_header("Product ID")
    if pid_norm not in norm_header_list:
        print("Main file e 'Product ID' column pawa jay ni - Step 6 cancel.")
        return None
    pid_idx = norm_header_list.index(pid_norm)

    col_idx_map = {}
    for col in STEP6_CONTENT_COLUMNS:
        cn = norm_header(col)
        if cn in norm_header_list:
            col_idx_map[col] = norm_header_list.index(cn)
        else:
            print(f"[warn] Main file e '{col}' column pawa jayni - eita skip kora hobe.")
    if not col_idx_map:
        print("Kono content column match korlo na - Step 6 cancel.")
        return None

    # ---- Cleaned file load kora ----
    try:
        cdf = pd.read_excel(cleaned_path, dtype=str)
    except Exception as e:
        print(f"'{cleaned_path}' open kora jay ni: {e}")
        return None
    cdf_cols_norm = {norm_header(c): c for c in cdf.columns}
    if pid_norm not in cdf_cols_norm:
        print("Cleaned file e 'Product ID' column pawa jay ni - Step 6 cancel.")
        return None
    cdf_pid_col = cdf_cols_norm[pid_norm]
    cdf["_pid_norm"] = cdf[cdf_pid_col].apply(normalize_pid)

    cleaned_col_map = {}
    for col in col_idx_map:
        cn = norm_header(col)
        if cn in cdf_cols_norm:
            cleaned_col_map[col] = cdf_cols_norm[cn]
    if not cleaned_col_map:
        print("Cleaned file e kono matching content column pawa jayni - Step 6 cancel.")
        return None

    lookup = {}
    dup_pids = set()
    for _, row in cdf.iterrows():
        pid = row["_pid_norm"]
        if not pid:
            continue
        if pid in lookup:
            dup_pids.add(pid)
        lookup[pid] = {col: row[cdf_col] for col, cdf_col in cleaned_col_map.items()}
    if dup_pids:
        print(f"[warn] Cleaned file e {len(dup_pids)} ta duplicate Product ID pawa geche - "
              f"sheshbar-er value use kora hocche.")

    # ---- Content column gula replace kora ----
    updated_count = 0
    missing_pids = []
    for row in data_rows:
        pid = normalize_pid(row[pid_idx])
        if pid and pid in lookup:
            vals = lookup[pid]
            for col, idx in col_idx_map.items():
                v = vals.get(col)
                if v is None:
                    row[idx] = None
                elif isinstance(v, float) and pd.isna(v):
                    row[idx] = None
                else:
                    row[idx] = v
            updated_count += 1
        else:
            missing_pids.append(pid)

    print(f"\nContent update kora holo: {updated_count} / {len(data_rows)} row.")
    if missing_pids:
        shown = [p for p in missing_pids[:10] if p]
        print(f"[warn] {len(missing_pids)} ta row-er Product ID cleaned file e pawa jayni "
              f"(egula original/purono content shoho thakbe). Example PID: {shown}")

    # ---- Template-er column order onujayi shob column re-arrange ----
    try:
        tmpl_probe = load_workbook(template_path, data_only=True)
    except Exception as e:
        print(f"Template file open kora jay ni: {e}")
        return None
    tmpl_sheet_name = _find_product_update_sheet(tmpl_probe)
    tmpl_ws_probe = tmpl_probe[tmpl_sheet_name]
    tmpl_rows_probe = list(tmpl_ws_probe.iter_rows(min_row=1, max_row=2, values_only=True))
    if len(tmpl_rows_probe) < 2:
        print(f"Template '{tmpl_sheet_name}' sheet e 2 ta header row pawa jayni - Step 6 cancel.")
        return None
    tmpl_header = list(tmpl_rows_probe[1])  # row 2 = actual column header
    tmpl_header_norm = [norm_header(h) for h in tmpl_header]
    tmpl_orig_max_row = tmpl_ws_probe.max_row

    if pid_norm not in tmpl_header_norm:
        print("Template e 'Product ID' column pawa jayni - Step 6 cancel.")
        return None

    missing_in_template = [
        header[i] for i in range(len(header))
        if norm_header_list[i] and norm_header_list[i] not in tmpl_header_norm
    ]
    if missing_in_template:
        print(f"[info] Main file er ei column gula template e nai (skip hobe): {missing_in_template}")

    src_to_tmpl_col = {}
    for i, hn in enumerate(norm_header_list):
        if hn and hn in tmpl_header_norm:
            src_to_tmpl_col[i] = tmpl_header_norm.index(hn) + 1  # 1-based

    n_tmpl_cols = len(tmpl_header)
    rearranged_rows = []
    for row in data_rows:
        out_row = [None] * n_tmpl_cols
        for src_i, tmpl_c in src_to_tmpl_col.items():
            out_row[tmpl_c - 1] = row[src_i]
        rearranged_rows.append(out_row)

    total_rows = len(rearranged_rows)
    if total_rows == 0:
        print("Kono data row nai - Step 6 cancel.")
        return None

    # ---- Split count: min files, max rows/file ----
    cfg = AUTO_CONFIG.get("step6", {}) if isinstance(AUTO_CONFIG.get("step6"), dict) else {}
    max_rows_per_file = int(cfg.get("max_rows_per_file", STEP6_MAX_ROWS_PER_FILE_DEFAULT))
    min_files = int(cfg.get("min_files", STEP6_MIN_FILES_DEFAULT))

    num_files = max(min_files, -(-total_rows // max_rows_per_file))  # ceil division
    num_files = min(num_files, total_rows)

    base_size = total_rows // num_files
    remainder = total_rows % num_files
    chunks = []
    start = 0
    for i in range(num_files):
        size = base_size + (1 if i < remainder else 0)
        if size == 0:
            continue
        chunks.append(rearranged_rows[start:start + size])
        start += size

    update_dir = os.path.join(get_d_drive_output_folder(), "update")
    os.makedirs(update_dir, exist_ok=True)

    written = []
    for i, chunk in enumerate(chunks, start=1):
        try:
            out_wb = load_workbook(template_path, data_only=False)
        except Exception as e:
            print(f"[error] Template part {i}-er jonno abar khulte parlam na: {e}")
            continue
        out_sheet_name = _find_product_update_sheet(out_wb)
        out_ws = out_wb[out_sheet_name]
        orig_max_row = out_ws.max_row
        _clear_and_fill_sheet(out_ws, chunk, orig_max_row, min_data_row=3)

        out_name = f"update_part{i}.xlsx"
        out_path = os.path.join(update_dir, out_name)
        out_wb.save(out_path)
        written.append(out_path)
        print(f"Wrote {out_path}  ({len(chunk)} rows, template format+dropdown shoho)")

    print(f"\nDone. {total_rows} rows -> {len(written)} template file(s) split.")
    print(f"Update folder: {update_dir}")

    # ---- QC Approval file: shob Product ID, Approval Status = 1 ----
    all_product_ids = [row[pid_idx] for row in data_rows]
    approval_path = build_product_approval_file(all_product_ids, update_dir)

    update_status(
        6, state="done", input=combine_path, cleaned_input=cleaned_path,
        template=template_path, output=update_dir, files=len(written),
        updated_rows=updated_count, missing_rows=len(missing_pids),
        approval_file=approval_path,
    )
    log(f"Step 6 done -> {update_dir} ({len(written)} template files, {updated_count} content-updated, "
        f"{len(missing_pids)} missing in cleaned file, approval file: {bool(approval_path)})")
    return update_dir


# =====================================================================
# Menu
# =====================================================================
def sweep_stray_files():
    """Main folder / Data folder e bhul jaygay theke jawa purono
    intermediate file gula (seller_export_report_*.xlsx root e,
    image_http_url_add.xlsx Data/ e) delete na kore ekta
    '_archive_old_files' folder e sorie dey - jate main directory
    poriskar thake kintu kono data accidentally miss na hoy."""
    archive_dir = os.path.join(BASE_DIR, "_archive_old_files")
    stray_patterns = [
        (BASE_DIR, re.compile(r"^seller_export_report_.*\.xlsx$", re.I)),
        (os.path.join(BASE_DIR, "Data"), re.compile(r"^image_http_url_add\.xlsx$", re.I)),
    ]
    moved = []
    for folder, pat in stray_patterns:
        if not os.path.isdir(folder):
            continue
        for fname in os.listdir(folder):
            if pat.match(fname):
                src = os.path.join(folder, fname)
                if not os.path.isfile(src):
                    continue
                os.makedirs(archive_dir, exist_ok=True)
                dst = os.path.join(archive_dir, fname)
                # jodi already same naam e file thake archive e, timestamp jog kora
                if os.path.exists(dst):
                    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                    name, ext = os.path.splitext(fname)
                    dst = os.path.join(archive_dir, f"{name}_{stamp}{ext}")
                try:
                    os.replace(src, dst)
                    moved.append(fname)
                except OSError as e:
                    print(f"[cleanup] '{fname}' sorate parlam na: {e}")
    if moved:
        print(f"[cleanup] {len(moved)} ta purono stray file '_archive_old_files' folder e sorano holo: {', '.join(moved)}")


def run_steps_1_to_5():
    print("\n" + "=" * 60)
    print("RUNNING STEPS 1-5 SEQUENTIALLY")
    print("=" * 60)

    sweep_stray_files()

    mapping_path = detect_mapping_file()
    if not mapping_path or not os.path.isfile(mapping_path):
        print("\nMapping file nai - pipeline stop.")
        return
    set_session_date_from_mapping(mapping_path)
    base_url = ask_base_url()
    if not base_url:
        print("\nBase URL dorkar - pipeline stop.")
        return
    save_base_url(base_url)

    out1 = run_step1(mapping_path, base_url)
    if out1 is None:
        print("\nStep 1 fail/cancel - pipeline stop.")
        return

    print("\nStep 2 kivabe korbo?")
    print("  1. AUTOMATED - browser e export script chalbe (menu 11 er moto)")
    print("  2. MANUAL - ami file ready kore rekhechi, confirm korte chai")
    mode = clean_input("1/2 [Enter = 1]: ", default="1")
    if mode == "2":
        out2 = run_step2_manual()
    else:
        out2 = run_automated_step2(mapping_path)
    if out2 is None:
        print("\nStep 2 fail/cancel - pipeline stop.")
        return

    out3 = run_step3(out2)
    if out3 is None:
        print("\nStep 3 fail/cancel - pipeline stop.")
        return

    out4 = run_step4(out3, out1)
    if out4 is None:
        print("\nStep 4 fail/cancel - pipeline stop.")
        return

    out5 = run_step5(out4)
    if out5 is None:
        print("\nStep 5 fail/cancel - pipeline stop.")
        return

    print("\n" + "=" * 60)
    print("STEP 1-5 COMPLETE!")
    print("=" * 60)
    print(f"\nSplit file gula ekhane: {out5}")
    print("Ekhon Claude handoff check option (menu 8) use korte paren.")


def load_auto_config():
    """agent_config.json theke auto_mode config load kore.
    File na thakle ba parse error hole default value use kore."""
    config_path = os.path.join(BASE_DIR, "agent_config.json")
    if not os.path.isfile(config_path):
        print(f"[auto] agent_config.json pawa jay ni: {config_path}")
        print("[auto] Default value diye cholbo (base_url stored theke nebo).")
        return {}
    try:
        with open(config_path, "r", encoding="utf-8") as f:
            data = json.load(f)
        cfg = data.get("auto_mode", data)  # top-level ba nested 'auto_mode'
        print(f"[auto] Config loaded: {config_path}")
        return cfg
    except Exception as e:
        print(f"[auto] Config load kora jay ni ({e}) - default value use korbo.")
        return {}


def main():
    global AUTO_MODE, AUTO_CONFIG

    # ---- Argument parsing ----
    parser = argparse.ArgumentParser(
        description="Catalogue Prep Agent - pipeline Steps 1-5",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  python agent.py             # Normal interactive menu
  python agent.py --auto      # Auto mode: Steps 1-5 bina jiggesh kore (config file theke value nebe)
  python agent.py --auto --step 3   # Shudhu Step 3 auto chalabe
        """
    )
    parser.add_argument(
        "--auto", action="store_true",
        help="Auto mode: agent_config.json theke value nebe, kono jiggesh korbe na"
    )
    parser.add_argument(
        "--step", type=int, choices=[1, 2, 3, 4, 5, 6, 7, 9, 11], default=None,
        help="Specific step chalao (6 = Update from cleaned + split, 7 = Steps 1-5 sequentially, 9 = Merge back). Default: 7 (--auto er sathe)"
    )
    args = parser.parse_args()

    if args.auto:
        AUTO_MODE = True
        AUTO_CONFIG = load_auto_config()

        target_step = args.step if args.step is not None else 7

        print("=" * 60)
        print("CATALOGUE PREP AGENT  -  AUTO MODE")
        print("=" * 60)
        print(f"Working folder: {BASE_DIR}")
        show_status()
        print(f"\n[auto] Chalacchi: Step {target_step}")
        print("-" * 60)

        try:
            if target_step == 1:
                run_step1()
            elif target_step == 2:
                run_step2()
            elif target_step == 3:
                run_step3()
            elif target_step == 4:
                run_step4()
            elif target_step == 5:
                run_step5()
            elif target_step == 6:
                run_step6_update_and_split()
            elif target_step == 7:
                run_steps_1_to_5()
            elif target_step == 9:
                run_merge_back()
            elif target_step == 11:
                run_automated_step2()
        except KeyboardInterrupt:
            print("\n\n[auto] Interrupted - status save hoye geche. Bye!")
            sys.exit(0)
        except Exception:
            traceback.print_exc()
            print("\n[auto] Unexpected error (upore details).")
            sys.exit(1)

        print("\n[auto] Done!")
        return

    # ---- Normal interactive mode ----
    print("=" * 60)
    print("CATALOGUE PREP AGENT  -  pipeline Steps 1-5 + validation")
    print("=" * 60)
    print(f"Working folder: {BASE_DIR}")
    show_status()

    while True:
        print("\n--- MAIN MENU ---")
        print("1. Validate files (read-only)")
        print("2. Step 1  - Image URL add")
        print("3. Step 2  - confirm Basic combine.xlsx (auto-detect / manual)")
        print("4. Step 3  - Trim columns")
        print("5. Step 4  - Merge images into HTML")
        print("6. Step 5  - Split for Claude")
        print("7. Run Steps 1-5 sequentially")
        print("8. Claude handoff check (manifest + prompt file check)")
        print("9. Merge back cleaned files")
        print("10. Show status")
        print("11. Step 2 AUTOMATED - seller export + combine (browser)")
        print("12. Step 6 - Update content from cleaned file + split for bulk update")
        print("0. Exit")

        choice = clean_input("Choice: ")
        print()
        try:
            if choice == "1":
                validate_file()
            elif choice == "2":
                run_step1()
            elif choice == "3":
                run_step2()
            elif choice == "4":
                run_step3()
            elif choice == "5":
                run_step4()
            elif choice == "6":
                run_step5()
            elif choice == "7":
                run_steps_1_to_5()
            elif choice == "8":
                run_handoff_check()
            elif choice == "9":
                run_merge_back()
            elif choice == "10":
                show_status()
            elif choice == "11":
                run_automated_step2()
            elif choice == "12":
                run_step6_update_and_split()
            elif choice == "0":
                print("Bye!")
                break
            else:
                print("Bujhi nai - 0-12 er moddhe ekta number dao.")
        except KeyboardInterrupt:
            print("\n\nInterrupted - status save hoye geche. Bye!")
            break
        except EOFError:
            print("\nInput shesh - bye!")
            break
        except Exception:
            traceback.print_exc()
            print("\nUnexpected error (upore details). Status safe - abar try korte paro.")


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\nInterrupted - bye!")
        sys.exit(0)
