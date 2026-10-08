# seller_bulk_export_automation.py
# Real flow: seller-admin.cartup.com -> Products -> Multi Seller Bulk Update
# -> Export Excel File tab -> comma-separated Seller Code(s) batch -> Download
# -> See Export History -> filter by "Downloaded By" -> wait for "completed"
# -> click that row's Download link -> save file.
#
# Login/OTP handling is reused from the earlier working script (same site,
# same PrimeReact-based UI helpers).
#
# Shesh e shob downloaded file eksathe merge kore "Basic combine.xlsx" banay.

import argparse
import getpass
import json
import re
import sys
import time
import traceback
from datetime import datetime
from pathlib import Path

import pandas as pd
from playwright.sync_api import sync_playwright, Error

try:
    from openpyxl import Workbook, load_workbook
except ImportError:
    print("openpyxl install kora nai. Age ei command ta run koro:")
    print("    pip install openpyxl --break-system-packages")
    sys.exit(1)


LOGIN_URL = "https://seller-admin.cartup.com/auth/signin?callbackUrl=%2F"
APP_HOME = "https://seller-admin.cartup.com/"

# Mapping file khunjar search directory + keyword (filename fuzzy match).
# Ei week file er name/dated thakleo auto-detect hobe - hardcoded path nai.
DEFAULT_MAPPING = "Data/Image Links Mapping (13.09.2026) Bakaul.xlsx"  # legacy default, ar use hoy na
MAPPING_KEYWORDS = ("mapping", "maping", "mapng", "bakaul", "image", "link")
SCRIPT_DIR = Path(__file__).resolve().parent
CONFIG_FILE = SCRIPT_DIR / "export_automation_config.json"


def log(msg):
    print(f"[{datetime.now().strftime('%H:%M:%S')}] {msg}", flush=True)


def nowstamp():
    return datetime.now().strftime("%Y-%m-%d_%H-%M-%S")


# ---------- mapping file auto-detection (fuzzy) ----------

def _name_tokens(fname):
    return set(re.findall(r"[a-z0-9&]+", str(fname).lower()))


def _squashed(fname):
    return re.sub(r"[^a-z0-9&]", "", str(fname).lower())


def _mtime_safe(p):
    try:
        return Path(p).stat().st_mtime
    except OSError:
        return 0.0


def load_last_mapping():
    """Sesh bar je mapping file use hoise seta config file theke pore -
    jate bar bar path dite na hoy. File ta ar na thakle None."""
    try:
        with open(CONFIG_FILE, "r", encoding="utf-8") as f:
            data = json.load(f)
        p = data.get("last_mapping")
        if p and Path(p).is_file():
            return p
    except Exception:
        pass
    return None


def save_last_mapping(path):
    try:
        data = {}
        try:
            with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                data = json.load(f)
        except Exception:
            data = {}
        data["last_mapping"] = str(path)
        with open(CONFIG_FILE, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
    except Exception as e:
        log(f"[warn] config save korte parlam na: {e}")


def find_mapping_candidates(search_dirs):
    """Folder gulo scan kore mapping-file candidates ber kore, ranked.
    Match rule: filename e (case-insensitive, token match) ei keyword
    gulo er kono ekta thaklei candidate: mapping/maping/mapng (common
    typo variant), bakaul, image+link. Score: specific keyword beshi,
    generic kom; tie hole newest age."""
    SPECIFIC = ("mapping", "maping", "mapng", "bakaul")
    GENERIC = ("image", "link")
    seen = set()
    found = []
    for d in search_dirs:
        d = Path(d)
        if not d.is_dir():
            continue
        for p in sorted(d.glob("*.xlsx")):
            if p.name.startswith("~$") or p.name in seen:
                continue
            seen.add(p.name)
            toks = _name_tokens(p.name)
            sq = _squashed(p.name)
            hit = any(k in toks for k in MAPPING_KEYWORDS)
            if not hit and "image" in toks and "link" in toks:
                hit = True
            if not hit:
                continue
            score = 10 * sum(1 for k in SPECIFIC if k in toks) + sum(1 for k in GENERIC if k in toks)
            found.append((-score, -_mtime_safe(p), str(p), p.name))
    found.sort()
    return [(full, name) for _, _, full, name in found]


def resolve_mapping_path(cli_value):
    """Mapping file path ber kora, priority order e:
    1. --mapping CLI flag e valid path dile seta (kintu na paoa gele
       fallback e auto-detect, crash na)
    2. Config file er 'last_mapping' (ager bar use kora file)
    3. Data/ ar script folder e fuzzy auto-detect
    Shob case e ekta valid Path return korar chesta kore; na pele
    None (caller clear error dekhabe, hardcoded default crash na hobe)."""
    script_dir = Path(__file__).resolve().parent
    data_dir = script_dir / "Data"
    search_dirs = [data_dir, script_dir, Path.cwd()]

    # 1. CLI flag
    if cli_value:
        p = Path(cli_value)
        if p.is_file():
            return p, "CLI flag (--mapping)"
        # CLI e default hardcoded string eshe thakleo bhul path hoile
        # fallback korbo - crash korbo na.
        if cli_value != DEFAULT_MAPPING:
            log(f"[warn] --mapping er file pawa jay ni: {cli_value}")

    # 2. Last-used (config memory)
    last = load_last_mapping()
    if last:
        name = Path(last).name
        ans = input(f"Ager bar use kora mapping file: {name}\nEi file ta use korbo? [Enter = Yes, 'n' dile onno khunje dekhbo]: ").strip().lower()
        if ans in ("", "y", "yes"):
            return Path(last), "config memory (ager bar use kora)"

    # 3. Fuzzy auto-detect + numbered chooser
    cands = find_mapping_candidates(search_dirs)
    if not cands:
        p = input("Kono mapping file pawa jay ni. Full path likhe din: ").strip()
        return (Path(p) if p else None), "manual path"
    if len(cands) == 1:
        full, name = cands[0]
        ans = input(f"Auto-detected mapping file: {name}\nEi file ta use korbo? [Enter = Yes, 'n' dile onno path likhun]: ").strip().lower()
        if ans in ("", "y", "yes"):
            return Path(full), "auto-detected"
        p = input("Notun full path: ").strip()
        return (Path(p) if p else None), "manual path"
    print(f"\n{len(cands)} ta matching file pawa geche - kon ta use korbo? (newest first):")
    for i, (full, name) in enumerate(cands, start=1):
        mt = datetime.fromtimestamp(_mtime_safe(full)).strftime("%Y-%m-%d %H:%M")
        print(f"  {i:2}. {name}   (modified: {mt})")
    ans = input("Number dao (ba full path likhun) [Enter = 1 no.]: ").strip()
    if not ans:
        return Path(cands[0][0]), "auto-detected (newest)"
    if ans.isdigit() and 1 <= int(ans) <= len(cands):
        return Path(cands[int(ans) - 1][0]), "auto-detected (choosen)"
    return (Path(ans) if ans else None), "manual path"


def pause(page, ms, msg=""):
    if msg:
        log(msg)
    if ms > 0:
        page.wait_for_timeout(ms)


# ---------- mapping file reading ----------

def read_seller_codes_from_mapping(mapping_path: Path, sheet_name: str = "Seller Info",
                                    code_col_name: str = "Seller Code"):
    """'Image Links Mapping (...).xlsx' file er 'Seller Info' sheet theke
    shob Seller Code poRe. Header naam e space/capital thakuk ba na thakuk
    - shob e match kore (fuzzy match)."""

    def normalize(s):
        if s is None:
            return ""
        return re.sub(r"[^a-z0-9]", "", str(s).strip().lower())

    df = pd.read_excel(mapping_path, sheet_name=sheet_name, dtype=str, keep_default_na=False)

    norm_to_actual = {normalize(c): c for c in df.columns}
    actual_col = norm_to_actual.get(normalize(code_col_name))
    if actual_col is None:
        raise ValueError(
            f"'{code_col_name}' column pawa jayni '{sheet_name}' sheet e. "
            f"Available columns: {list(df.columns)}"
        )

    codes = []
    for val in df[actual_col]:
        code = (val or "").strip()
        if code:
            codes.append(code)
    return codes


def batch_codes(codes, batch_size):
    for i in range(0, len(codes), batch_size):
        yield codes[i:i + batch_size]


CREDENTIALS_FILE = SCRIPT_DIR / "credentials.json"


def load_credentials():
    """credentials.json theke username/password load kore.
    File na thakle ba parse error hole terminal e jiggesh kore (fallback).
    credentials.json format:
    {
      "username": "nazmul_limon",
      "password": "your_password_here"
    }
    """
    if CREDENTIALS_FILE.exists():
        try:
            with open(CREDENTIALS_FILE, "r", encoding="utf-8") as f:
                data = json.load(f)
            username = data.get("username", "").strip()
            password = data.get("password", "").strip()
            if username and password:
                log(f"[info] Credentials loaded from credentials.json (user: {username})")
                return username, password
            log("[warn] credentials.json e username/password khali - terminal e jiggesh korbo.")
        except Exception as e:
            log(f"[warn] credentials.json porate error: {e} - terminal e jiggesh korbo.")
    # Fallback: terminal e jiggesh
    log("[info] credentials.json nai - terminal e credential dite hobe.")
    username = input("Seller-admin username dao: ").strip()
    password = getpass.getpass("Seller-admin password dao (typing hidden thakbe): ").strip()
    return username, password


def scrape_display_name(page) -> str:
    """Login er por browser er top-right corner theke logged-in user er
    naam automatically poRe - kono manual input lage na.
    Prothome credentials.json theke 'display_name' nebar try korbe,
    na pele DOM theke scrape korbe."""
    
    # Try reading from credentials.json first
    if CREDENTIALS_FILE.exists():
        try:
            with open(CREDENTIALS_FILE, "r", encoding="utf-8") as f:
                data = json.load(f)
            display_name = data.get("display_name", "").strip()
            if display_name:
                log(f"[info] Display name loaded from credentials.json: '{display_name}'")
                return display_name
        except Exception as e:
            log(f"[warn] credentials.json theke display_name porate error: {e}")

    log("[info] credentials.json e display_name nai, DOM theke scrape korar chesta korchi...")

    selectors = [
        # Top-right user button (screenshot e 'Nazmul Limon' dekhay)
        "button.p-button.p-component >> nth=-1",
        "[class*='user'] span",

        "[class*='profile'] span",
        "header button span",
        # Generic: shobcheye last button text (top-right e thake)
        "header >> button >> nth=-1",
        "nav button >> nth=-1",
    ]
    for sel in selectors:
        try:
            loc = page.locator(sel).last
            if loc.count() and loc.is_visible(timeout=800):
                txt = loc.inner_text(timeout=800).strip()
                # Naam: khali na, 2 char er beshi, prothom char letter (icon/emoji filter)
                if txt and len(txt) > 2 and txt[0].isalpha() and "\n" not in txt:
                    log(f"[info] Display name scraped from DOM: '{txt}' (selector: {sel})")
                    return txt
        except Exception:
            continue

    # Last resort: page title ba URL theke kichhu ber kora jay na,
    # tai blank return korbo - filter_history_by_me skip hobe.
    log("[warn] Top-right theke naam scrape kora jacche na - 'Downloaded By' filter skip hobe.")
    return ""


# ---------- generic helpers (PrimeReact-based UI) ----------

def _wait_overlay_clear(page, loops=60, sleep_ms=250):
    overlay = page.locator("div.p-datatable-loading-overlay").first
    for _ in range(loops):
        try:
            if overlay.count() == 0 or not overlay.is_visible():
                return True
        except Exception:
            return True
        page.wait_for_timeout(sleep_ms)
    return False


def _resilient_click(locator, page, desc="element", total_timeout_ms=4000):
    end = time.time() + total_timeout_ms / 1000.0
    last_err = None
    while time.time() < end:
        try:
            if locator.count() == 0:
                last_err = Error("missing")
                page.wait_for_timeout(200)
                continue
            if not locator.first.is_visible():
                page.wait_for_timeout(150)
                continue
            try:
                locator.first.click(timeout=600)
                return True
            except Exception as e:
                last_err = e
                try:
                    locator.first.click(timeout=600, force=True)
                    return True
                except Exception as e2:
                    last_err = e2
                    try:
                        locator.first.evaluate("el=>el.click()")
                        return True
                    except Exception as e3:
                        last_err = e3
                        page.wait_for_timeout(220)
        except Exception as e:
            last_err = e
            page.wait_for_timeout(220)
    log(f"[warn] _resilient_click failed on {desc}: {last_err}")
    return False


# ---------- login ----------

def is_logged_in(page) -> bool:
    u = (page.url or "").lower()
    return ("seller-admin.cartup.com" in u) and ("/auth/" not in u) and ("/auth/signin" not in u)


def confirm_logged_in(page) -> bool:
    try:
        page.goto(APP_HOME, wait_until="domcontentloaded", timeout=15000)
    except Exception:
        try:
            page.goto(APP_HOME)
        except Exception:
            pass
    page.wait_for_timeout(400)
    return is_logged_in(page)


def press_enter_after_otp_and_confirm(page, max_rounds=8) -> None:
    for i in range(1, max_rounds + 1):
        input(f"\n[MANUAL] Browser e OTP dao, tarpor ekhane ENTER chapo... (attempt {i}/{max_rounds})\n")
        if confirm_logged_in(page):
            log("[info] Login confirmed ✅")
            return
        log("[warn] Ekhono login hoyni (signin e redirect hocche). Browser e login shesh kore abar Enter chapo.")
    raise RuntimeError("Login not confirmed after multiple attempts.")


def do_login(page, auth_path: Path, force_login: bool):
    has_saved_auth = auth_path.exists() and (not force_login)

    log("[info] Opening APP_HOME...")
    page.goto(APP_HOME, wait_until="domcontentloaded")

    if is_logged_in(page):
        log("[info] Already logged in (saved session).")
        return

    log("[info] Session not valid. Going to LOGIN page...")
    page.goto(LOGIN_URL, wait_until="domcontentloaded")

    username, password = load_credentials()
    if username and password:
        log("[info] Filling username/password...")
        page.fill('input[name="username"]', username)
        page.fill('input[name="password"]', password)
        page.click('button:has-text("Sign In")')
        log("[info] Sign In clicked. OTP/manual step ashte pare browser e.")
    else:
        log("[info] Username/password deya hoyni. Browser e manually login koro (OTP shoho).")

    press_enter_after_otp_and_confirm(page)


# ---------- Multi Seller Bulk Update page ----------

def navigate_to_bulk_update(page):
    log("Products menu e jacchi...")
    _resilient_click(page.locator("text=Products").first, page, "Products menu", 4000)
    page.wait_for_timeout(400)
    log("'Multi Seller Bulk...' page e jacchi...")
    _resilient_click(page.locator("text=Multi Seller Bulk").first, page, "Multi Seller Bulk menu item", 4000)
    page.wait_for_load_state("domcontentloaded")
    page.wait_for_timeout(500)

    # Nishchit hocchi 'Export Excel File' tab e achi
    export_tab = page.locator("text=Export Excel File").first
    if export_tab.count():
        _resilient_click(export_tab, page, "Export Excel File tab", 3000)
        page.wait_for_timeout(300)


def submit_batch(page, seller_codes):
    codes_str = ",".join(seller_codes)
    log(f"Batch submit korchi: {codes_str}")

    field = page.get_by_placeholder("Enter Seller Code(s)")
    field.fill("")
    field.fill(codes_str)

    # Status: All, Edit: Basic Information - nishchit kore click kora
    _resilient_click(page.locator("text=All").first, page, "Status: All", 2000)
    _resilient_click(page.locator("text=Basic Information").first, page, "Edit: Basic Information", 2000)

    # Submit button - <button> tag, table er <a> Download link theke alada
    _resilient_click(page.locator("button:has-text('Download')").first, page, "Submit/Download button", 3000)
    log("Batch queue e joma hoyeche.")


def filter_history_by_me(page, display_name):
    if not display_name:
        return
    field = page.get_by_placeholder("Downloaded By")
    if field.count():
        field.fill("")
        field.fill(display_name)
        _resilient_click(page.locator("button:has-text('Search')").first, page, "Search history", 2500)
        page.wait_for_timeout(400)


def _unique_target(out_dir: Path, base: str) -> Path:
    p = out_dir / base
    if not p.exists():
        return p
    stem, suf = p.stem, p.suffix
    for i in range(1, 500):
        alt = out_dir / f"{stem}({i}){suf}"
        if not alt.exists():
            return alt
    return out_dir / f"{stem}_{nowstamp()}{suf}"


def wait_and_download_latest(page, out_dir: Path, batch_label: str, poll_interval_s=10, timeout_s=300):
    """Export History er PROTHOM (shobcheye notun) row-er status 'completed'
    hobar jonno wait kore, tarpor download link e click kore file save kore."""
    start = time.time()

    def click_refresh():
        btn = page.locator("button:has-text('Refresh')").first
        if btn.count():
            _resilient_click(btn, page, "Refresh history", 3000)

    while time.time() - start < timeout_s:
        _wait_overlay_clear(page)
        row = page.locator("table tbody tr").first
        if not row.count():
            click_refresh()
            page.wait_for_timeout(poll_interval_s * 1000)
            continue

        try:
            row_text = row.inner_text(timeout=1500).lower()
        except Exception:
            row_text = ""

        if "failed" in row_text:
            log(f"[warn] {batch_label}: export FAILED (site e 'failed' dekhacche)")
            return False, "failed", ""

        if "completed" in row_text:
            link = row.locator("a:has-text('Download')").first
            if link.count():
                try:
                    with page.expect_download(timeout=120000) as d:
                        _resilient_click(link, page, "Row Download link", 2500)
                    dl = d.value
                    fname = f"{batch_label}_{dl.suggested_filename or 'export.xlsx'}"
                    target = _unique_target(out_dir, fname)
                    dl.save_as(str(target))
                    log(f"Downloaded: {target}")
                    return True, "completed", str(target)
                except Exception as e:
                    log(f"[warn] download click e error: {e}")

        log(f"  ... status ekhono ready na, {poll_interval_s}s por abar check korbo")
        click_refresh()
        page.wait_for_timeout(poll_interval_s * 1000)

    log(f"[warn] {batch_label}: timeout - export ready hoyni {timeout_s}s er moddhe")
    return False, "timeout", ""


# ---------- combine step ----------

def normalize_pid(value):
    """Product ID ke consistent string e convert kore, jate
    2261883 ar 2261883.0 ar '2261883' shob match kore."""
    if value is None:
        return None
    s = str(value).strip()
    if s.endswith(".0"):
        s = s[:-2]
    return s


def find_pid_column(headers, pid_col_name="Product ID"):
    """Header list theke Product ID column er (1-indexed) position ber
    kore - prothom match take rakha hoy (duplicate header thakleo)."""
    target = re.sub(r"[^a-z0-9]", "", pid_col_name.strip().lower())
    for i, h in enumerate(headers):
        if h is None:
            continue
        norm = re.sub(r"[^a-z0-9]", "", str(h).strip().lower())
        if norm == target:
            return i + 1
    return None


def read_valid_product_ids(mapping_path: Path, sheet_name: str = "Data Mapping",
                            pid_col_name: str = "Product ID"):
    """Mapping file er 'Data Mapping' sheet theke shob Product ID poRe,
    normalize kore ekta set banay - Workings sheet filter korার jonno."""
    wb = load_workbook(mapping_path, data_only=True)
    if sheet_name not in wb.sheetnames:
        raise ValueError(f"'{sheet_name}' sheet pawa jayni mapping file e. Available: {wb.sheetnames}")
    ws = wb[sheet_name]
    headers = [c.value for c in ws[1]]
    col = find_pid_column(headers, pid_col_name)
    if col is None:
        raise ValueError(f"'{pid_col_name}' column pawa jayni '{sheet_name}' sheet e. Headers: {headers}")

    valid_ids = set()
    for row in range(2, ws.max_row + 1):
        pid = normalize_pid(ws.cell(row=row, column=col).value)
        if pid:
            valid_ids.add(pid)
    return valid_ids


def combine_downloaded_files(out_dir: Path, output_path: Path, mapping_path: Path = None,
                              mapping_sheet: str = "Data Mapping", mapping_pid_col: str = "Product ID"):
    """Shob downloaded seller export file (xlsx) ekshathe merge kore.
    - 'Combined' sheet e SHOB row thake (kono filter chhara)
    - mapping_path deya thakle, 'Workings' sheet e SHUDHU shei row thake
      jader Product ID mapping file e ache (XLOOKUP-equivalent filter).
      Workings-i active sheet hoy, karon pipeline script gula wb.active
      poRe."""
    output_path = Path(output_path)  # str path dile o kaj korbe
    files = sorted(Path(out_dir).glob("*.xlsx"))
    if not files:
        log(f"[warn] {out_dir} e kono xlsx file pawa jayni - combine skip kora hocche.")
        return None

    log(f"Combining {len(files)} file(s) ...")

    new_wb = Workbook()
    combined_ws = new_wb.active
    combined_ws.title = "Combined"

    header_row = None
    total_rows = 0

    for f in files:
        try:
            wb = load_workbook(f, data_only=True)
            ws = wb.active
            rows = list(ws.iter_rows(values_only=True))
        except Exception as e:
            log(f"[warn] {f.name} porte parlam na, skip kora hocche: {e}")
            continue
        if len(rows) < 2:
            log(f"[warn] {f.name} e 2 ta row o nai (group header + column header thakar kotha), skip kora hocche.")
            continue

        # Row 1 (index 0) = group header (Basic Information, Product Description, ...) - EI ROW BAD DEWA HOCHE.
        # Row 2 (index 1) = actual column header (Seller Code, Product ID, ...) - EITAI real header.
        actual_header = rows[1]
        data_rows = rows[2:]

        if header_row is None:
            header_row = actual_header
            combined_ws.append(header_row)

        for row in data_rows:
            combined_ws.append(row)
            total_rows += 1

    log(f"'Combined' sheet: {total_rows} ta row, {len(files)} ta file theke.")

    if mapping_path is not None and header_row is not None:
        log("Mapping file theke valid Product ID gula porচি (XLOOKUP-equivalent filter)...")
        valid_ids = read_valid_product_ids(mapping_path, mapping_sheet, mapping_pid_col)
        log(f"Mapping file e {len(valid_ids)} ta unique Product ID pawa geche.")

        pid_col = find_pid_column(list(header_row), mapping_pid_col)
        if pid_col is None:
            log(f"[warn] Combined sheet e '{mapping_pid_col}' column pawa jayni - Workings sheet banano jacche na.")
        else:
            workings_ws = new_wb.create_sheet("Workings")
            workings_ws.append(header_row)
            kept = 0
            for row in combined_ws.iter_rows(min_row=2, values_only=True):
                pid = normalize_pid(row[pid_col - 1])
                if pid in valid_ids:
                    workings_ws.append(row)
                    kept += 1
            log(f"'Workings' sheet: {kept} / {total_rows} row rakha hoyeche (mapping file er sathe match kore).")
            new_wb.active = new_wb.sheetnames.index("Workings")

    new_wb.save(output_path)
    log(f"Done! Save hoyeche: {output_path.resolve()}")
    return output_path


# ---------- MAIN ----------

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--mapping", default=None,
                     help="'Image Links Mapping (...).xlsx' file er path. Na dile: ager bar use kora file (memory) > Data/ folder e fuzzy auto-detect > manual path.")
    ap.add_argument("--batch-size", type=int, default=4, help="Ekbar e koyta seller code ekshathe submit kora hobe.")
    ap.add_argument("--outdir", default="downloads")
    ap.add_argument("--headless", action="store_true")
    ap.add_argument("--slowmo", type=int, default=150)
    ap.add_argument("--auth-state", default="auth_state.json")
    ap.add_argument("--force-login", action="store_true")
    ap.add_argument("--poll-interval", type=int, default=10, help="Koto second por por export status check hobe.")
    ap.add_argument("--poll-timeout", type=int, default=300, help="Ekta batch er jonno max koto second wait korbe.")
    ap.add_argument("--combine-output", default="Basic combine.xlsx")
    ap.add_argument("--no-combine", action="store_true")
    args = ap.parse_args()

    log("=== Script started ===")

    mapping_path, how = resolve_mapping_path(args.mapping)
    if mapping_path is None or not mapping_path.is_file():
        print("FATAL_ERROR: Mapping file pawa jay ni / path thik nai. 'Image Links Mapping (...).xlsx' file ta Data/ folder e rakhun.")
        sys.exit(1)
    log(f"Mapping file ({how}): {mapping_path.name}")
    save_last_mapping(mapping_path)  # porer bar ei file default hisebe memory-te thakbe

    codes = read_seller_codes_from_mapping(mapping_path)
    log(f"Mapping file read OK. Total seller codes: {len(codes)}")
    batches = list(batch_codes(codes, args.batch_size))
    log(f"{len(batches)} ta batch e vag kora hobe ({args.batch_size} seller/batch).")

    outdir = Path(args.outdir)
    outdir.mkdir(parents=True, exist_ok=True)
    auth_path = Path(args.auth_state).resolve()

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=args.headless, slow_mo=args.slowmo)
        ctx_kwargs = dict(accept_downloads=True, viewport={"width": 1400, "height": 900})
        if auth_path.exists() and not args.force_login:
            ctx_kwargs["storage_state"] = str(auth_path)
        context = browser.new_context(**ctx_kwargs)
        page = context.new_page()

        do_login(page, auth_path, args.force_login)
        context.storage_state(path=str(auth_path))
        log(f"[info] Auth state saved: {auth_path}")

        if not confirm_logged_in(page):
            raise RuntimeError("Login confirm kora jayni. Thamiye dicchi.")

        display_name = scrape_display_name(page)
        if display_name:
            log(f"[info] Export history '{display_name}' naam diye filter hobe.")

        navigate_to_bulk_update(page)

        results = []
        for i, batch in enumerate(batches, start=1):
            log(f"\n--- Batch {i}/{len(batches)} ---")
            submit_batch(page, batch)
            page.wait_for_timeout(1500)
            filter_history_by_me(page, display_name)
            ok, status, path = wait_and_download_latest(
                page, outdir, f"batch{i}",
                poll_interval_s=args.poll_interval,
                timeout_s=args.poll_timeout,
            )
            results.append({"batch": i, "seller_codes": ",".join(batch), "status": status, "file": path})

        # Age main project folder e "seller_export_report_*.xlsx" clutter hoye
        # thakto - ekhon output/downloads folder er pashe rakha hocche, jate
        # main directory clean thake.
        report_dir = outdir.parent if outdir.parent.exists() else outdir
        report = report_dir / f"seller_export_report_{nowstamp()}.xlsx"
        pd.DataFrame(results).to_excel(report, index=False)
        log(f"Report saved: {report.resolve()}")

        browser.close()

    if not args.no_combine:
        combine_downloaded_files(
            outdir, Path(args.combine_output),
            mapping_path=mapping_path,
            mapping_sheet="Data Mapping",
            mapping_pid_col="Product ID",
        )

    log("=== Script finished ===")


if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        print("FATAL ERROR:", e, flush=True)
        traceback.print_exc()
        sys.exit(1)
