# Redesign Plan: Catalogue Prep v2

Goal: **same result, fewer clicks, easier to change.**

## 1. What changes for the user

Today: 2 `.bat` files, a 12-item menu, mixed step numbers, and a hard-coded `D:` drive.

v2 has **3 commands, one for each real stage of the work**:

| Stage | Command | `.bat` | What happens |
|---|---|---|---|
| A. Prepare | `python catprep.py prepare` | `1_Prepare.bat` | URLs → export → trim → swap images → split for Claude |
| B. Claude | *(manual)* | – | Clean parts, drop them into the run's `claude_done\` folder |
| C. Finish | `python catprep.py finish` | `2_Finish.bat` | Merge back → check → upload files + approval file |

Extras: `catprep.py status`, `catprep.py check <file>`, and
`catprep.py step <name>` to re-run a single step.

Each run gets **one self-contained folder** that holds its inputs, outputs and status:

```
<output_root>\2026-09-28\
├── 0_input\        mapping file copy
├── 1_urls\         image_http_url_add.xlsx
├── 2_export\       downloads\ + Basic combine.xlsx
├── 3_trim\         Basic combine H & D.xlsx
├── 4_images\       ..._image_replaced.xlsx, needs_review.xlsx
├── 5_for_claude\   part1.xlsx … + prompt copy
├── 6_claude_done\  ← put Claude's output here
├── 7_upload\       update_part1.xlsx … + Product-approval.xlsx
├── run.json        status of every step
└── run.log
```

The folder names show the order, so nobody has to remember it.

## 2. New code layout

```
catprep.py                  entry point (commands above), about 100 lines
config.yaml                 ALL settings in one readable file
catprep/
  common.py                 normalize_pid, match_header, find_file, Excel helpers (ONE copy)
  run_folder.py             creates/opens the dated run folder, run.json, log
  steps/
    s1_urls.py              (was add_http_to_image_key.py)
    s2_export.py            (was seller_bulk_export_automation.py)
    s3_trim.py              (was trim_columns.py)
    s4_images.py            (was merge_images_into_html.py, faster)
    s5_split.py             (was split_excel.py)
    s6_merge_back.py        (was inside agent.py)
    s7_upload.py            (was Step 6 inside agent.py)
templates/
  Product-approval-template.xlsx
  H&D_improved_prompt.md
tests/                      small sample files + checks for each step
```

Each step is **one function** with clear inputs and outputs:
`run(run_folder, config) -> StepResult`. Adding or changing a step means
editing one small file.

## 3. One settings file (`config.yaml`)

```yaml
output_root: "D:/Image URL change"     # change drive/folder here
base_url: "https://sl-dev-s3.s3.amazonaws.com/product/"

mapping_file:
  folder: "Data"
  name_must_contain: ["image", "link", "mapping"]

export:
  batch_size: 4
  poll_seconds: 10
  timeout_seconds: 300
  headless: false

claude_split_rows: 3500
upload:
  max_rows_per_file: 7500
  min_files: 2
  include_needs_review_rows: false     # see decision 1 below
  approve_needs_review_rows: false

columns:
  keep: [Product ID, Seller Code, Name (English), Highlights(English),
         Highlights(Bengali), Description (English), Description (Bengali)]
  html: [Highlights(English), Highlights(Bengali),
         Description (English), Description (Bengali)]
```

## 4. Fixes included

| # | Problem today | v2 |
|---|---|---|
| 1 | Needs Review rows uploaded unchanged and approved | Off by default (setting). Exported separately as `needs_review.xlsx` |
| 2 | Missing cleaned rows only warn | `finish` stops with a clear list unless you allow it |
| 3 | Hard-coded `D:` / `e:` paths | `output_root` in config; `.bat` uses its own folder (`%~dp0`) |
| 4 | Two different mapping-file rules | One `find_file()` used everywhere |
| 5 | Slow Step 4 | Build the new sheets in one pass; no row-by-row delete |
| 6 | Step 6 missing from status | Status table lists every step from `run.json` |
| 7 | Confusing menu/step numbers | Named commands; folders numbered 0–7 |
| 8 | One yes/no answer for everything | Separate `overwrite` and `confirm` settings |
| 9 | Template = "first file in downloads" | Checks for a `Product Update` sheet with 2 header rows before using it |
| 10 | Copied helper code, dead code | One `common.py`; `run_pipeline.py` and unused code removed |
| 11 | No tests | Tests with small fake files for steps 1, 3, 4, 5, 6, 7 |
| 12 | Plain-text password | Optional: password read from Windows Credential Manager (`keyring`) |

## 5. What stays the same

- The results: the same Excel files, sheets, columns, template formatting and dropdowns.
- The browser robot's login/OTP flow, selectors and batch logic. It is only
  moved and tidied, because it can't be tested here without the live site.
- Banglish messages in the console, if you prefer them; or English.

## 6. Decisions needed before building

1. **Needs Review rows:** should they be left out of the upload/approval files? (Recommended: yes.)
2. **Claude prompt:** please send `H&D_improved_prompt.md` so it can live in `templates/`.
3. **Console language:** Banglish (like now) or English?
4. **Old run folders:** keep the current `D:\Image URL change\<Month Year>\<DD-MM-YYYY>\` layout, or switch to the numbered sub-folders above?

## 7. Build order

1. `common.py` + `run_folder.py` + config loading, with tests
2. Move steps 1, 3, 4, 5 (pure Excel); compare output with the old scripts on the real files
3. Merge back + upload builder; compare with the old Step 6 output
4. Move the export robot (step 2). This needs one live test on your PC.
5. `catprep.py` commands, the two `.bat` files, README update
6. Delete the old scripts once v2 gives the same results
