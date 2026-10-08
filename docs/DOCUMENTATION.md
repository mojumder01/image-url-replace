# Catalogue Prep v2: Full Documentation

## 1. Purpose

Sellers on cartup.com copied product content from other marketplaces, so their
Highlights/Descriptions contain `<img src="...">` tags that point to **external
images** (e.g. Daraz CDN). The images were re-uploaded to our S3 bucket. The
content team (Bakaul) sends an **image mapping file** with the new image keys for
every product.

The tool, without anyone at the PC:

1. turns image keys into full S3 URLs
2. downloads the current content of every affected seller from seller-admin (OTP read from email)
3. replaces the old image URLs inside the HTML with the new ones
4. cleans the HTML with the rules of `templates/H&D_improved_prompt.md` (Python, no Claude needed)
5. builds upload files in the **official bulk-update template** and a QC **approval file**
6. uploads them to seller-admin and emails a report

---

## 2. Project layout

```
catprep.py                 the program (commands below)
config.toml                all settings
credentials.example.json   copy to credentials.json (private) and fill in
Run_All.bat                full automatic run          Run_All_Resume.bat  continue it
Auto_Scheduled.bat         what the scheduled task runs (writes scheduler.log)
Schedule_Watch.bat / Schedule_Daily.bat / Schedule_Remove.bat
Status.bat                 progress of the newest run
manual_claude\             1_Prepare.bat, 1_Prepare_Resume.bat, 2_Finish.bat (Claude by hand)
catprep/
  common.py                shared helpers (header/ID matching, Excel, console)
  config.py                loads config.toml + defaults
  run_folder.py            run folder, run.json status, run.log
  cleaner.py               the prompt's cleaning rules + its 6 self-checks
  site.py                  seller-admin browser session + login
  otp_email.py             reads the login OTP from the mailbox (IMAP)
  notify.py                report email (SMTP)
  steps/  s1_urls  s2_export  s3_trim  s4_images  s5_clean  s5_split  s6_merge  s7_upload  s8_submit
templates/                 H&D_improved_prompt.md, Product-approval-template.xlsx
tests/                     automatic tests (python -m unittest discover tests)
Data/                      put the mapping file here (not in Git)
legacy/                    the old version, for reference only
```

Every step file has one function `run(ctx)`. It reads its input from the run
folder, writes its output there, and returns a short summary that is saved in
`run.json`.

---

## 3. The two routes

| Route | Steps | Started by |
|---|---|---|
| **Automatic** (normal) | urls → export → trim → images → **clean** → upload → submit | `auto` |
| Manual Claude | urls → export → trim → images → **split** · *Claude by hand* · **merge** → upload → submit | `prepare`, then `finish` |

---

## 4. Commands

| Command | What it does |
|---|---|
| `python catprep.py auto` | New automatic run from the newest mapping file in `Data\` |
| `python catprep.py auto --if-new` | Same, only if that mapping file was not done yet (used by the schedule) |
| `python catprep.py auto --resume` | Continue the newest run from its first unfinished step |
| `python catprep.py schedule --watch` / `--daily [--at 09:00]` / `--remove` | Windows scheduled tasks |
| `python catprep.py prepare [--resume] [--manual-export]` | Manual route, up to the Part files |
| `python catprep.py finish [--allow-missing]` | Manual route, after Claude |
| `python catprep.py step <name>` | Re-run one step: `urls export trim images clean split merge upload submit` |
| `python catprep.py status` / `runs` | Progress of the newest run / list all runs |
| `python catprep.py check <file> [file2]` | Show sheets/headers/ID count; with 2 files, compare Product IDs |

Add `--run "<run folder>"` to `auto --resume`, `finish`, `step` or `status` to use an older run.

### Automatic-run safety

- **Only once per file:** `--if-new` compares the mapping file's name, size and date with earlier runs.
- **Retries:** an unfinished automatic run is retried by the schedule up to `[auto] max_attempts`
  times (default 3), then left for you (`Run_All_Resume.bat` after fixing).
- **No overlap:** a lock file (`.catprep.lock` in the output folder) stops a second run.
  A lock older than 6 hours is treated as stale.
- **Scheduled runs** hide the browser (`headless_when_scheduled`) and log to `scheduler.log`.
  Windows Task Scheduler runs them only while you are logged in to Windows.
- **Report email** after every automatic run, DONE or FAILED, with the numbers of each step (`[notify]`).

---

## 5. Run folder

```
D:\Image URL change\September 2026_28-09-2026_15-37\
├── 0_input\         copy of the mapping file
├── 1_urls\          image_http_url_add.xlsx
├── 2_export\        downloads\batchN_*.xlsx, Basic combine.xlsx, export_report.xlsx
├── 3_trim\          Basic combine H & D.xlsx
├── 4_images\        Basic combine H & D_image_replaced.xlsx  (Data / Needs Review / Merge Report)
├── 5_for_claude\    manual route only: Part-N.xlsx + prompt + instructions
├── 6_claude_done\   manual route only: Claude's cleaned files
├── 7_merged\        Basic combine cleaned.xlsx  (+ Cleaning Report, Check Problems)
├── 8_upload\        update_partN.xlsx, Product-approval.xlsx, not_uploaded.xlsx, site_*.png
├── run.json         status + numbers of every step
└── run.log          history
```

The name is `<Month Year>_<DD-MM-YYYY>_<HH-MM>`. The date comes from the mapping file
name (today's date if there is none); the time is when the run started.

---

## 6. Input files

**Mapping file** (`Data\`, newest `.xlsx` whose name contains `image`, `link` and
`mapping`/`maping`/`bakaul`):

| Sheet | Columns |
|---|---|
| `Data Mapping` | `Product ID`, `image1`, `image2`, … (keys like `dcf2879f-….webp`) |
| `Seller Info` | `Seller ID`, `Seller Code`, `Shop Name`, `Shop URL` |

**Seller export files** (step 2): seller-admin → Products → Multi Seller Bulk Update →
Export Excel File (Status = All, Edit = Basic Information). They have 2 header rows in
sheet `Product Update` plus dropdown helper sheets. One of them is the **upload template**.

---

## 7. Steps

Column names are matched loosely (case, spaces, brackets ignored). Product IDs are
compared as text (`2261883.0` = `2261883`).

| # | Name | Input → Output | Rules |
|---|---|---|---|
| 1 | `urls` | mapping → `image_http_url_add.xlsx` | Columns with "image" in the header: `key` → `base_url + key`; http(s) and empty cells skipped |
| 2 | `export` | Seller Codes → `downloads\`, `Basic combine.xlsx` | Robot exports `batch_size` sellers at a time; already-downloaded batches are skipped. `Workings` = only Product IDs in the mapping file |
| 3 | `trim` | `Workings` → `Basic combine H & D.xlsx` | Keeps `columns.keep` |
| 4 | `images` | trim + urls → `..._image_replaced.xlsx` | 1st `<img src>` → image1, 2nd → image2, … per HTML column. **Needs Review**: no mapping, or more `<img>` tags than new images |
| 5 | `clean` | `Data` sheet → `Basic combine cleaned.xlsx` | Python cleaning rules (section 8) + the prompt's 6 self-checks |
| 6 | `upload` | Workings + cleaned + template → `update_partN.xlsx`, approval | Section 9 |
| 7 | `submit` | `8_upload` → seller-admin | Section 10 |

Manual route instead of step 5: `split` (Part files of `rows_per_part` rows) → Claude →
`merge` (joins `6_claude_done\*.xlsx`, reports missing/extra/duplicate Product IDs).

---

## 8. Cleaning rules (`catprep/cleaner.py`)

These are the rules of `H&D_improved_prompt.md`, written as code. The output format copies
Claude's real output:

- **Highlights:** `<ul><li>…</li></ul>`
- **Description:** `<h2>title</h2>`, `<ul>` bullets, `<p>` text, `<h3>` + spec lists/tables, `<img src="…"/>` images

| Rule | How |
|---|---|
| HTML cleanup | Only `h2 h3 ul li p table tr td img` are written; other tags are unwrapped; every attribute except `img src` is removed |
| Corrupted markup | Leaked text like ` style="...">` is removed; stray `<` `>` removed (real `< 20 mA` kept) |
| Remove | Hebrew, emojis/dingbats, broken symbols, zero-width characters, empty tags and table cells, extra spaces |
| "Empty" | fewer than 3 readable characters (a description with a real title is not empty) |
| Missing highlights | both empty → the description's bullets (same language); one empty → copy of the other |
| Non-bullet highlights | paragraphs, `*` `•` `$` separated text → bullets (`54 * 45` sizes are not split) |
| Junk bullets | dropped when they have under 3 readable characters, or no 3-letter word and no number (`200 g` stays) |
| Near-duplicate bullets | exact repeats, "same bullet with a word missing", and bullets that are other bullets glued together are removed |
| Missing description | both empty → built from the highlights; one empty → copy of the other |
| Repeated paragraph | dropped when 85% of its meaningful words are already in the bullets (`REPEAT_SHARE`) |
| Images | every image URL of the row is kept, at the bottom of the descriptions |
| Generic highlights | highlights used by 2+ products are replaced by the product's own description bullets (if it has 4+ specific ones) |
| Language | Bangla and English are never translated or mixed |

**Self-checks** (every row): no empty cell, no `style=`/`class=`/`<div`/`<span`/`<a`, no empty
tags, no junk bullets, no lost images, no duplicated paragraph. Rows failing one are listed in
the **Check Problems** sheet and are **not uploaded** (`skip_rows_failing_checks`).

**How close it is to Claude:** I ran the cleaner on Claude's own cleaned output for the
28-09-2026 run (10,166 products). 99.5% of highlights and 98.9% of descriptions came back
unchanged (ignoring spaces/invisible characters). My output passes all 6 checks; Claude's own
output fails about 100 of them. The final comparison against the **raw** input files is still
to be done.

---

## 9. Building the upload files (step `upload`)

1. All columns of `Workings` are used; the 4 HTML columns get the cleaned content.
2. A product is uploaded **only if** it was cleaned and passed the checks. Everything else
   goes into `not_uploaded.xlsx` with a reason:
   - `Needs Review: No Mapping Found` / `Needs Review: Insufficient Mapped Images`
   - `Cleaning check failed: …`
   - `Missing from Claude output` (manual route; stops unless allowed)
3. Columns are put in the template's order (template row 2).
4. Rows are split evenly: at least `min_files`, at most `max_rows_per_file` rows per file. Each
   file is a copy of the template, with formats copied and dropdowns stretched.
5. `Product-approval.xlsx`: every uploaded Product ID, `Approval Status = 1`.

---

## 10. Uploading to seller-admin (step `submit`)

For each `update_partN.xlsx`, and then `Product-approval.xlsx`, the robot:
1. clicks the menu items in `[site_upload.update].menu` / `[site_upload.approval].menu`
2. puts the file into the page's file box (`file_input`) and clicks `submit_button`
3. waits until the page shows a `success_text` or `error_text` word (max `wait_seconds`)
4. saves a screenshot `8_upload\site_<file>.png`

If one update file fails, it stops and the approval file is **not** uploaded.

**Status:** switched off (`enabled = false`) until the menu, button and message names are
confirmed from screenshots of the real upload pages. Until then the step lists the files to
upload by hand.

---

## 11. Login and OTP

1. **Saved session:** `auth_state.json`. While it is valid, no login is needed.
2. **Username/password:** from `credentials.json`.
3. **OTP:**
   - With `[otp] enabled = true`, the program logs in to the mailbox (IMAP, Gmail app password).
     It waits up to `wait_seconds` for an OTP email that arrived **after** the Sign In click (an
     old code is never reused), from `sender` and with `subject_contains`. It reads the digits
     next to "OTP"/"code" (or uses `code_pattern`) and types them into the OTP page.
   - Otherwise, and as a fallback when someone is at the PC: you type the OTP in the browser.

---

## 12. Settings (`config.toml`)

| Section | Key settings |
|---|---|
| top | `output_root` (run folders), `base_url` |
| `[mapping]` | `folder`, `data_sheet`, `seller_sheet` |
| `[export]` | `batch_size` 4, `poll_seconds` 10, `timeout_seconds` 300, `headless`, login file names |
| `[claude]` | manual route: `rows_per_part` 3500, `prompt_file`, `fallback_done_folder` |
| `[upload]` | `max_rows_per_file` 7500, `min_files` 2, `allow_missing_cleaned_rows`, `skip_rows_failing_checks`, `approval_template` |
| `[auto]` | `watch_minutes` 10, `daily_time` "09:00", `max_attempts` 3, `headless_when_scheduled` |
| `[otp]` | `enabled`, `imap_host`, `sender`, `subject_contains`, `code_pattern`, `wait_seconds`, `input_selector`, `submit_selector` |
| `[notify]` | `enabled`, `to` |
| `[site_upload]` | `enabled`, `upload_approval`, `wait_seconds`, and `.update` / `.approval`: `menu`, `file_input`, `submit_button`, `success_text`, `error_text` |
| `[columns]` | `keep` (7 columns), `html` (4 columns) |

Paths may use `\` or `/`. Relative paths are taken from the project folder.

---

## 13. Changes from the old version

| Old | v2 |
|---|---|
| Claude cleaning by hand (upload parts, download results) | Python rules, ~25 s for 10,000 products |
| OTP typed by a person | Read from the email inbox |
| Upload by hand | Upload robot (after setup) |
| Started by double-clicking 2 files, with steps in between | One command, folder watch, or daily schedule |
| Needs Review rows uploaded with **old images** and approved | Left out, listed in `not_uploaded.xlsx` |
| Hard-coded `D:\…` / `e:\FreeBuff\…` paths | `output_root` in config; `.bat` files work from any folder |
| 12-item menu, numbers that don't match | Named commands, numbered folders |
| Step 4 deleted rows one by one | 2× faster or more |
| One 1,800-line file, copied helpers, Banglish | Small files per step, one helpers module, English |
| No tests | 30 tests |

Data results of steps 1–4 and of the upload files are **identical** to the old version,
checked old-vs-new on the real 28-09-2026 files.

---

## 14. Troubleshooting

| Message | What to do |
|---|---|
| `No mapping file in …\Data` | Put the file in `Data\`; its name needs image + link + mapping |
| `Email OTP: cannot log in` | Check `email_address` / `email_app_password`; Gmail needs an app password, not the normal one |
| `Email OTP: no OTP email arrived` | Check `[otp] sender` / `subject_contains`; check that the email really arrives in that inbox |
| `OTP page has no input box` | Set `[otp] input_selector` |
| `Login was not confirmed` | Delete `auth_state.json` and run `Run_All.bat` while watching |
| `N batch(es) did not download` | `Run_All_Resume.bat` (only failed batches are retried) |
| `Another run is still working` | Wait; if no run is active, delete `.catprep.lock` in the output folder |
| `… failed 3 times - not retrying automatically` | Read `run.log`, fix, then `Run_All_Resume.bat` |
| Rows in `Check Problems` | They were not uploaded; check them in `7_merged\Basic combine cleaned.xlsx` |
| `No export template found` | `2_export\downloads` needs at least one real export file |
| `Playwright is not installed` | `pip install playwright` and `python -m playwright install chromium` |
| Scheduled run did nothing | See `scheduler.log`; the PC must be on and you logged in to Windows |
