# Catalogue Prep Agent: Full Documentation

This describes the project **as it works today** (the "old" version).
For the proposed new version see [REDESIGN.md](REDESIGN.md).

---

## 1. Purpose

Sellers on cartup.com copied product content from other marketplaces.
Their Highlights/Descriptions contain `<img src="...">` tags pointing to
**external images** (e.g. Daraz CDN). The images were re-uploaded to our S3
bucket, and the content team (Bakaul) sends an **image mapping file** that
lists the new image key for every product.

This project:

1. Turns image keys into full S3 URLs.
2. Downloads the current product content of every affected seller.
3. Replaces the old image URLs inside the HTML with the new ones.
4. Splits the work so Claude can clean/improve the text.
5. Merges Claude's output back.
6. Builds upload files in the **official bulk-update template** plus a QC
   **approval file**, ready to upload in seller-admin.

---

## 2. Files in the project

### Code

| File | Role | Lines |
|---|---|---|
| `agent.py` | **Main program.** Menu + `--auto` mode, file auto-detection, status tracking, Steps 1–6, merge-back. Imports the 4 step scripts. | 1816 |
| `add_http_to_image_key.py` | Step 1 engine: prefixes image keys with the base URL. | 166 |
| `seller_bulk_export_automation.py` | Step 2 engine: Playwright browser robot that exports each seller's products and combines them. Run by `agent.py` as a separate process. | 737 |
| `trim_columns.py` | Step 3 engine: keeps 7 columns. | 131 |
| `merge_images_into_html.py` | Step 4 engine: replaces `<img src>` URLs in HTML. | 353 |
| `split_excel.py` | Step 5 engine: splits a sheet into N-row files. | 88 |
| `run_pipeline.py` | **Old/legacy** runner for steps 1–5 (manual step 2). Replaced by `agent.py`; not used anymore. | 115 |

### Launchers

| File | What it runs |
|---|---|
| `run_auto.bat` | `python agent.py --auto`, which runs Steps 1→5 |
| `Merge_Claude_Output.bat` | `python agent.py --auto --step 9` (merge back), then `--step 6` (build upload files) |

Both `.bat` files hard-code the folder `e:\FreeBuff\Image URL work agent`.

### Settings, state and secrets

| File | Purpose | In Git? |
|---|---|---|
| `agent_config.json` | Auto-mode settings: `base_url`, `overwrite_existing`, `header_rows`, `chunk_size` (optional `step6.max_rows_per_file`, `step6.min_files`) | yes |
| `pipeline_status.json` | Last result of every step (written by the agent) | yes (could be ignored, see §8) |
| `agent_log.txt` | One line per finished step (history) | no |
| `credentials.json` | seller-admin username, password, display name | **never** |
| `auth_state.json` | Saved browser login session (cookies) | **never** |
| `export_automation_config.json` | Remembers the last mapping file used by the export robot | never |

### Folders

| Folder | Contents |
|---|---|
| `Data/` | Input mapping file, e.g. `image_links_mapping (28.09.2026) Bakaul.xlsx` |
| `H&D_Processed_Input/` | Claude-cleaned parts (`Part-1.xlsx`, `Part-2.xlsx`, ...) |
| `templates/` | `Product-approval-template.xlsx` (QC approval layout) |
| `_archive_old_files/` | Old stray files moved here automatically |
| `D:\Image URL change\<Month Year>\<DD-MM-YYYY>\` | **All outputs** of a run (see §5) |

---

## 3. Input files

### 3.1 Image mapping file (`Data/`)

| Sheet | Columns | Rows (28.09 file) |
|---|---|---|
| `Data Mapping` | `Product ID`, `image1` … `image35` (image keys like `dcf2879f-….webp`) | 11,444 |
| `Seller Info` | `Seller ID`, `Seller Code`, `Shop Name`, `Shop URL` | 25 sellers |

The **date in the file name** (`28.09.2026`) decides the output folder
(`D:\Image URL change\September 2026\28-09-2026\`).

How the file is found automatically: the newest `.xlsx` in `Data/` (or the
project folder) whose name contains **all three**: `image`, `link`, and
(`mapping` or `maping` or `bakaul`).

### 3.2 Seller export (downloaded by Step 2)

Downloaded from seller-admin → Products → Multi Seller Bulk Update →
Export Excel File (Status = All, Edit = Basic Information). Each file has
**2 header rows**: row 1 = group names, row 2 = real column names, data
from row 3. Sheet `Product Update` plus helper sheets with dropdown lists.
One of these files is reused as the **upload template** in Step 6.

### 3.3 Approval template (`templates/`)

`Product-approval-template.xlsx`, sheet `Sheet1`: `Seller ID`, `Product ID`,
`Approval Status`, `Product Tags`, `Reject Reason`.

---

## 4. The steps in detail

Column names are matched loosely everywhere: case, spaces and brackets are
ignored, so `Highlights(English)` = `Highlights (English)`.
Product IDs are normalised (`2261883.0` → `2261883`).

### Step 1: Add image URL (`add_http_to_image_key.py`)

- **In:** mapping file, base URL (`https://sl-dev-s3.s3.amazonaws.com/product/`)
- **Does:** in every sheet, every column whose header contains `image`:
  `key.webp` → `base_url + key.webp`. Cells that already start with
  `http://`/`https://` and empty cells are skipped.
- **Out:** `image_http_url_add.xlsx` (same sheets + a `Report` sheet with counts)

### Step 2: Export seller products (`seller_bulk_export_automation.py`)

- **In:** `Seller Info` sheet → list of Seller Codes; `credentials.json`
- **Does:**
  1. Opens Chrome (visible). Re-uses `auth_state.json` if the session is still valid.
  2. Otherwise fills username/password; **you type the OTP in the browser and press Enter in the console**.
  3. Goes to Multi Seller Bulk Update → Export Excel File.
  4. Submits seller codes in **batches of 4**, waits (poll every 10 s, max 300 s) until the newest Export History row says `completed`, downloads it.
  5. Combines all downloads (drops header row 1, keeps row 2 as header):
     - sheet `Combined`: every row
     - sheet `Workings`: only rows whose Product ID is in the mapping file's `Data Mapping` sheet (active sheet)
- **Out:** `downloads/batchN_*.xlsx`, `Basic combine.xlsx`, `seller_export_report_<time>.xlsx`
- Step 2 can also be done **manually** (make `Basic combine.xlsx` yourself); the agent then just checks the 7 columns exist.

### Step 3: Trim columns (`trim_columns.py`)

- **In:** `Basic combine.xlsx` (active sheet = `Workings`)
- **Keeps:** `Product ID`, `Seller Code`, `Name (English)`, `Highlights(English)`,
  `Highlights(Bengali)`, `Description (English)`, `Description (Bengali)`
- **Out:** `Basic combine H & D.xlsx`

### Step 4: Merge new images into HTML (`merge_images_into_html.py`)

- **In:** `Basic combine H & D.xlsx` + `image_http_url_add.xlsx`
- **Does:** for each product, in each of the 4 HTML columns, the 1st
  `<img src>` → image1, 2nd → image2, … (each column starts again at image1).
  Extra tags with no new image are left unchanged.
- **Moves to a `Needs Review` sheet** (with an `Issue` column):
  - `No Mapping Found`: Product ID not in the mapping file
  - `Insufficient Mapped Images`: more `<img>` tags than new images
- **Out:** `Basic combine H & D_image_replaced.xlsx` (main sheet = good rows,
  `Needs Review`, `Merge Report`)

### Step 5: Split for Claude (`split_excel.py`)

- **In:** `..._image_replaced.xlsx` (first sheet only, so Needs Review rows are **not** included)
- **Does:** splits into files of `chunk_size` rows (3,500), header repeated.
- **Out:** folder `..._image_replaced_split/` with `..._part1.xlsx`, `..._part2.xlsx`, …

### Manual: Claude cleaning

Upload each part to Claude with the prompt file **`H&D_improved_prompt.md`**
(expected in the project folder; **not part of the files received**).
Save the cleaned results (same 7 columns) into `H&D_Processed_Input/`.
Menu 8 ("handoff check") writes `handoff_manifest.md` with row counts and a checklist.

### Merge back (agent menu 9 / `--step 9`)

- **In:** all `.xlsx` in `H&D_Processed_Input/`
- **Does:** checks every file has the same columns, joins them, forces
  Product ID to text.
- **Out:** `Basic combine cleaned.xlsx` (sheet `Data`)

### Step 6: Build upload files (agent menu 12 / `--step 6`)

- **In:** `Basic combine.xlsx` (`Workings`, all columns), `Basic combine cleaned.xlsx`,
  the first file in `downloads/` (template), `templates/Product-approval-template.xlsx`
- **Does:**
  1. For each row of `Workings`, replaces the 4 content columns with the cleaned
     version (matched by Product ID). Rows not found keep their **original** content.
  2. Re-orders all columns to the template's order (template row 2).
  3. Splits rows evenly: at least **2 files**, at most **7,500 rows** per file.
  4. For each part: opens the template, clears old data, writes rows from row 3,
     keeps formatting and dropdowns (extends dropdown ranges if needed).
  5. Writes `Product-approval.xlsx`: every Product ID with `Approval Status = 1`.
- **Out:** `update/update_part1.xlsx`, `update/update_part2.xlsx`, …, `update/Product-approval.xlsx`

---

## 5. Output folder of one run

```
D:\Image URL change\September 2026\28-09-2026\
├── image_http_url_add.xlsx                      Step 1
├── downloads\batch1_…xlsx, batch2_…              Step 2 raw exports (also the template)
├── seller_export_report_<time>.xlsx             Step 2 batch report
├── Basic combine.xlsx                           Step 2
├── Basic combine H & D.xlsx                     Step 3
├── Basic combine H & D_image_replaced.xlsx      Step 4
├── Basic combine H & D_image_replaced_split\    Step 5 (+ handoff_manifest.md)
├── Basic combine cleaned.xlsx                   Merge back
└── update\                                      Step 6 → upload these
    ├── update_part1.xlsx
    ├── update_part2.xlsx
    └── Product-approval.xlsx
```

Real numbers from the 28-09-2026 run: 11,433 products in Basic combine →
1,267 to Needs Review → 10,166 cleaned by Claude (3 parts) → 2 upload files.

---

## 6. How to run

### Auto mode

```
python agent.py --auto              # Steps 1→5 (same as --step 7)
python agent.py --auto --step 9     # merge back cleaned parts
python agent.py --auto --step 6     # build upload files
```

`--step` accepts 1, 2, 3, 4, 5, 6, 7, 9, 11. In auto mode every question is
answered from `agent_config.json`; only the **OTP** still needs a person.

### Interactive menu (`python agent.py`)

| Menu | Action | Same as `--step` |
|---|---|---|
| 1 | Validate a file (read-only), cross-check Product IDs of two files | – |
| 2 | Step 1: add image URL | 1 |
| 3 | Step 2: confirm Basic combine (or start export) | 2 |
| 4 | Step 3: trim columns | 3 |
| 5 | Step 4: merge images into HTML | 4 |
| 6 | Step 5: split for Claude | 5 |
| 7 | Run steps 1–5 | 7 |
| 8 | Claude handoff check | – |
| 9 | Merge back cleaned files | 9 |
| 10 | Show status | – |
| 11 | Step 2 automated export | 11 |
| 12 | Step 6: build upload files | 6 |

### Status and log

- `pipeline_status.json`: last state/input/output per step.
- `agent_log.txt`: append-only history.

---

## 7. Configuration (`agent_config.json`)

```json
{
  "auto_mode": {
    "base_url": "https://sl-dev-s3.s3.amazonaws.com/product/",
    "overwrite_existing": true,
    "header_rows": 1,
    "chunk_size": 3500
  }
}
```

Optional: `"step6": {"max_rows_per_file": 7500, "min_files": 2}`.
Export robot options (command-line only): `--batch-size 4`, `--poll-interval 10`,
`--poll-timeout 300`, `--headless`, `--force-login`.

**Hard-coded (not configurable):** output root `D:\Image URL change`, project
path inside the `.bat` files, sheet names `Data Mapping` / `Seller Info` /
`Workings` / `Product Update`, the 7 / 4 column lists, site URLs.

---

## 8. Known problems and risks

### Behaviour that may be wrong (please confirm)

1. **Needs Review rows are uploaded unchanged and approved.** Step 6 writes
   *every* row of `Workings` into the upload files, including the 1,267 rows that
   were sent to Needs Review. They still have the **old external images**, and
   `Product-approval.xlsx` approves them too. Nothing in the pipeline ever fixes
   Needs Review rows.
2. **Rows Claude dropped are silently kept old.** If a cleaned part is missing
   rows, Step 6 only prints a warning.

### Bugs and weak spots

3. **"Auto" mode is not fully automatic.** The OTP needs a person. With
   `--auto --step 11` the export robot also asks which mapping file to use.
4. **Hard-coded paths.** `D:\Image URL change` and `e:\FreeBuff\...` in the
   `.bat` files break on another PC or drive.
5. **Two different mapping-file rules.** `agent.py` requires image + link +
   mapping. The export robot accepts *any one* keyword. They can pick
   different files.
6. **Slow Step 4.** Rows are deleted one by one (`delete_rows`), which is very
   slow with thousands of rows.
7. **Status view hides Step 6.** `show_status()` only lists steps 1–5.
8. **Confusing numbers.** Menu numbers, step numbers and `--step` numbers
   don't match (e.g. Step 6 = menu 12; merge back = `--step 9`).
9. **One answer for every yes/no.** In auto mode, all yes/no questions (even
   "start the browser?") use `overwrite_existing`.
10. **Template is picked as "first file in downloads/".** If that folder has
    an unrelated or broken file, Step 6 uses it.

### Code health

11. The same helpers (`normalize_pid`, header normalising, file finding) are
    copied in 4 files with small differences.
12. Dead code: `run_pipeline.py`, `detect_template_file()` (duplicate of
    `find_downloads_template()`), unused `STEP1_DEFAULT_OUT` /
    `STEP3_DEFAULT_OUT`, unused imports, legacy `DEFAULT_MAPPING`.
13. `agent.py` is 1,800 lines in one file: menu, detection, steps and Excel
    template handling are all mixed together.
14. Comments and messages are in Banglish. That's fine for you, harder for
    anyone else.
15. No tests.
16. `credentials.json` stores the password in plain text.
17. `H&D_improved_prompt.md` (the Claude prompt) is not in the project files.
