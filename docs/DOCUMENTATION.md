# Catalogue Prep v2: Full Documentation

## 1. Purpose

Sellers on cartup.com copied product content from other marketplaces, so their
Highlights/Descriptions contain `<img src="...">` tags that point to **external
images** (e.g. Daraz CDN). The images were re-uploaded to our S3 bucket. The
content team (Bakaul) sends an **image mapping file** with the new image keys for
every product.

The tool:

1. turns image keys into full S3 URLs
2. downloads the current content of every affected seller from seller-admin
3. replaces the old image URLs inside the HTML with the new ones
4. splits the work into parts for Claude to clean (prompt: `templates/H&D_improved_prompt.md`)
5. merges Claude's output back and checks nothing is missing
6. builds upload files in the **official bulk-update template** and a QC **approval file**

---

## 2. Project layout

```
catprep.py                 the program (commands below)
config.toml                all settings
1_Prepare.bat              prepare (steps 1-5)
1_Prepare_Resume.bat       continue a stopped prepare
2_Finish.bat               finish (steps 6-7)
Status.bat                 show progress of the newest run
requirements.txt           openpyxl, playwright
catprep/
  common.py                shared helpers (header/ID matching, Excel, console)
  config.py                loads config.toml + defaults
  run_folder.py            run folder, run.json status, run.log
  steps/
    s1_urls.py  s2_export.py  s3_trim.py  s4_images.py
    s5_split.py s6_merge.py   s7_upload.py
templates/
  H&D_improved_prompt.md           Claude cleaning prompt
  Product-approval-template.xlsx   QC approval layout
tests/test_pipeline.py     automatic tests (small fake files)
Data/                      put the mapping file here (not in Git)
legacy/                    the old version, for reference only
credentials.json           seller-admin login (private, not in Git)
auth_state.json            saved browser session (private, not in Git)
```

Every step file has one function `run(ctx)`. It reads its input from the run
folder, writes its output there, and returns a short summary that is saved in
`run.json`. To change one step, you edit one small file.

---

## 3. Commands

| Command | What it does |
|---|---|
| `python catprep.py prepare` | New run from the newest mapping file in `Data\`: steps 1-5 |
| `python catprep.py prepare --resume` | Continue the newest run from its first unfinished step |
| `python catprep.py prepare --manual-export` | No browser: you put export files in `2_export\downloads` yourself |
| `python catprep.py prepare --mapping "<file>"` | Use a specific mapping file |
| `python catprep.py finish` | Steps 6-7 on the newest run |
| `python catprep.py finish --allow-missing` | Continue even if Claude's output misses some rows |
| `python catprep.py step <name>` | Re-run one step: `urls`, `export`, `trim`, `images`, `split`, `merge`, `upload` |
| `python catprep.py status` | Progress of the newest run |
| `python catprep.py runs` | List all runs |
| `python catprep.py check <file> [file2]` | Show sheets/headers/ID count; with 2 files, compare Product IDs |

Add `--run "<run folder>"` to `finish`, `step`, `status` or `prepare --resume`
to work on an older run.

The program asks questions only when it really must: the OTP, and (in a
terminal) whether to continue when Claude's output is missing rows.

---

## 4. Run folder

```
D:\Image URL change\September 2026_28-09-2026_15-37\
├── 0_input\         copy of the mapping file
├── 1_urls\          image_http_url_add.xlsx
├── 2_export\        downloads\batchN_*.xlsx, Basic combine.xlsx, export_report.xlsx
├── 3_trim\          Basic combine H & D.xlsx
├── 4_images\        Basic combine H & D_image_replaced.xlsx  (Data / Needs Review / Merge Report)
├── 5_for_claude\    Part-1.xlsx ..., H&D_improved_prompt.md, README - what to do next.txt
├── 6_claude_done\   ← you put Claude's cleaned files here
├── 7_merged\        Basic combine cleaned.xlsx
├── 8_upload\        update_part1.xlsx ..., Product-approval.xlsx, not_uploaded.xlsx
├── run.json         status + numbers of every step
└── run.log          history
```

The name is `<Month Year>_<DD-MM-YYYY>_<HH-MM>`. The date comes from the mapping
file name; the time is when the run started. If there is no date in the name,
today's date is used.

---

## 5. Input files

### Mapping file (`Data\`)

| Sheet | Columns |
|---|---|
| `Data Mapping` | `Product ID`, `image1`, `image2`, … (keys like `dcf2879f-….webp`) |
| `Seller Info` | `Seller ID`, `Seller Code`, `Shop Name`, `Shop URL` |

The newest `.xlsx` whose name contains `image`, `link` and `mapping`/`maping`/`bakaul` is used.

### Seller export files (step 2)

seller-admin → Products → Multi Seller Bulk Update → Export Excel File
(Status = All, Edit = Basic Information). These files have 2 header rows (row 1 =
groups, row 2 = column names) in sheet `Product Update`, plus dropdown helper
sheets. One of them is also used as the **upload template**.

---

## 6. Steps

Column names are matched loosely everywhere: case, spaces and brackets are
ignored, so `Highlights (English)` = `Highlights(English)`. Product IDs are
compared as text (`2261883.0` = `2261883`).

| # | Name | Input → Output | Rules |
|---|---|---|---|
| 1 | `urls` | mapping file → `image_http_url_add.xlsx` | Every column with "image" in the header: `key` → `base_url + key`. Cells already starting with http(s) and empty cells are skipped. A `Report` sheet is added. |
| 2 | `export` | Seller Codes → `downloads\`, `Basic combine.xlsx` | Browser robot, batches of `batch_size` sellers, waits for Export History "completed", downloads. Already-downloaded batches are skipped on resume. Combine: `Combined` = all rows; `Workings` = only Product IDs in the mapping file. |
| 3 | `trim` | `Workings` → `Basic combine H & D.xlsx` | Keeps `columns.keep` in that order. |
| 4 | `images` | trim + urls → `..._image_replaced.xlsx` | In each HTML column, 1st `<img src>` → image1, 2nd → image2, … Rows go to **Needs Review** when the Product ID has no mapping, or there are more `<img>` tags than new images. |
| 5 | `split` | `Data` sheet → `Part-N.xlsx` | `rows_per_part` rows each; the prompt and an instructions file are copied in. |
| – | Claude | Part files → cleaned files | Manual. Save results in `6_claude_done\`. |
| 6 | `merge` | `6_claude_done\*.xlsx` → `Basic combine cleaned.xlsx` | If `6_claude_done` is empty, `fallback_done_folder` is used. Reports Product IDs missing from Claude's output, extra, and duplicates. |
| 7 | `upload` | Workings + cleaned + template → `update_partN.xlsx`, approval | See below. |

### Step 7 in detail

1. All columns of `Workings` are used. The 4 HTML columns get Claude's content.
2. A product is uploaded **only if** it was sent to Claude **and** Claude returned it.
   Everything else goes into `not_uploaded.xlsx` with a reason:
   - `Needs Review: No Mapping Found` / `Needs Review: Insufficient Mapped Images`
   - `Missing from Claude output`: this **stops** the step unless you allow it
     (answer yes, `--allow-missing`, or `allow_missing_cleaned_rows = true`)
   - `Not sent to Claude` (e.g. blank Product ID)
3. Columns are put in the template's order (template row 2). Columns not in the
   template are left out (listed on screen).
4. Rows are split evenly: at least `min_files` files, at most `max_rows_per_file`
   rows each. Each file is a copy of the template: old rows cleared, new rows
   from row 3, formatting copied and dropdowns stretched to the last row.
5. `Product-approval.xlsx`: every uploaded Product ID, `Approval Status = 1`.

---

## 7. Settings (`config.toml`)

| Setting | Default | Meaning |
|---|---|---|
| `output_root` | `D:\Image URL change` | where run folders are made |
| `base_url` | `https://sl-dev-s3.s3.amazonaws.com/product/` | added before image keys |
| `mapping.folder` | `Data` | where the mapping file is |
| `mapping.data_sheet` / `seller_sheet` | `Data Mapping` / `Seller Info` | sheet names |
| `export.batch_size` | 4 | sellers per export request |
| `export.poll_seconds` / `timeout_seconds` | 10 / 300 | Export History checks |
| `export.headless` | false | hide the browser |
| `export.credentials_file` / `auth_state_file` | `credentials.json` / `auth_state.json` | login files |
| `claude.rows_per_part` | 3500 | rows per Part file |
| `claude.prompt_file` | `templates\H&D_improved_prompt.md` | copied next to the parts |
| `claude.fallback_done_folder` | `H&D_Processed_Input` | used if `6_claude_done` is empty |
| `upload.max_rows_per_file` / `min_files` | 7500 / 2 | upload file split |
| `upload.allow_missing_cleaned_rows` | false | see step 7 |
| `upload.approval_template` | `templates\Product-approval-template.xlsx` | approval layout |
| `columns.keep` | 7 columns | columns sent to Claude |
| `columns.html` | 4 columns | columns with images / cleaned content |

Paths may use `\` or `/`. Relative paths are taken from the project folder.

---

## 8. Changes from the old version

| Old | v2 |
|---|---|
| Needs Review rows uploaded with **old images** and approved | Left out, listed in `not_uploaded.xlsx` |
| Rows missing from Claude's output: warning only | Stops, shows which ones |
| Hard-coded `D:\…` and `e:\FreeBuff\…` paths | `output_root` in config; `.bat` files work from any folder |
| 12-item menu, numbers that don't match | 3 commands; numbered folders |
| Outputs mixed in one folder per day | One folder per run (date + time) with sub-folders |
| Two different mapping-file rules | One rule |
| Step 4 deleted rows one by one (slow) | Builds sheets in one pass (2× faster even with few review rows) |
| Restart = run everything again | `--resume`; downloaded batches are skipped |
| Same helpers copied in 4 files, dead code, 1,800-line `agent.py` | One `common.py`; one small file per step |
| Banglish messages | English |
| pandas + openpyxl | openpyxl only |
| No tests | `tests/` (13 tests) |

The results are otherwise **identical**. Each step was run old-vs-new on the real
28-09-2026 mapping file and the real Claude parts. Every sheet matched cell by
cell, and every uploaded row was identical to the old version's row for the same
product.

---

## 9. Troubleshooting

| Message | What to do |
|---|---|
| `No mapping file in …\Data` | Put the file in `Data\`; check its name has image + link + mapping |
| `Login was not confirmed` | Finish login/OTP in the browser, press Enter. Delete `auth_state.json` if the session is broken |
| `N batch(es) did not download` | Run `1_Prepare_Resume.bat` (only failed batches are retried) |
| `No cleaned files found` | Put Claude's files in the run's `6_claude_done\` |
| `… sent to Claude are missing from its output` | Re-do those rows in Claude, or allow and continue |
| `No export template found` | `2_export\downloads` needs at least one real export file |
| `Playwright is not installed` | `pip install playwright` and `python -m playwright install chromium` |
| Something unexpected | `run.log` in the run folder has the history; `Status.bat` shows where it stopped |
