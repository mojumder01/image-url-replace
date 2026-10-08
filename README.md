# Catalogue Prep (Image URL Replace) v2

## Goal

Replace the **old external images** (Daraz CDN etc.) inside product
Highlights/Descriptions on **seller-admin.cartup.com** with **our own S3 image
URLs**, clean the HTML, and produce (and upload) the **bulk update files** and the
**QC approval file**, fully automatically.

## The automatic run

```
 new mapping file in Data\   (Task Scheduler notices it, or you double-click Run_All.bat)
        │
 1 urls     add the base URL to the image keys
 2 export   robot downloads every seller's products   (login OTP read from your email)
 3 trim     keep the 7 needed columns
 4 images   put the new image URLs into the HTML     → "Needs Review" rows set aside
 5 clean    clean the HTML with the prompt's rules (Python, no Claude needed)
 6 upload   fill the official template + approval file
 7 submit   robot uploads them to seller-admin       (switched on after setup)
        │
 email report: DONE / FAILED with the numbers
```

## Ways to start it

| How | What to do |
|---|---|
| **Double-click** | `Run_All.bat` (continue a stopped run: `Run_All_Resume.bat`) |
| **Watch the Data folder** | run `Schedule_Watch.bat` once. Every 10 minutes Windows checks `Data\`; a new mapping file is processed automatically |
| **Every day at a set time** | run `Schedule_Daily.bat` once (time in `config.toml` → `[auto] daily_time`) |
| Stop the schedule | `Schedule_Remove.bat` |

A mapping file is processed only once. A failed run is retried up to 3 times and then left for
you, and two runs never overlap. Scheduled runs write to `scheduler.log`.
`Status.bat` shows the newest run.

## One-time setup

1. `pip install -r requirements.txt` and `python -m playwright install chromium`
2. Copy `credentials.example.json` to `credentials.json` and fill it in. For Gmail, create an
   **app password**: Google Account → Security → 2-Step Verification → App passwords.
3. In `config.toml`:
   - `[otp] enabled = true`, plus `sender` / `subject_contains` of the OTP email
   - `[notify] enabled = true` to get the report email
   - `[site_upload] enabled = true` once the upload pages are configured
4. Run `Run_All.bat` once while watching, then turn on a schedule.

## Where the files go

One folder per run, named after the mapping file date and the start time:

```
D:\Image URL change\September 2026_28-09-2026_15-37\
  0_input  1_urls  2_export  3_trim  4_images  7_merged  8_upload  run.json  run.log
```

- `8_upload\not_uploaded.xlsx` lists every product that was **not** uploaded, and why.
- `7_merged\Basic combine cleaned.xlsx` has a **Cleaning Report** sheet and a **Check Problems** sheet.

## Manual Claude route (optional)

To have Claude clean the HTML instead of the Python rules: `manual_claude\1_Prepare.bat`, clean
the `5_for_claude` parts in Claude, put the results in `6_claude_done`, then
`manual_claude\2_Finish.bat`.

## More

- [Full documentation](docs/DOCUMENTATION.md): every step, setting, command and file.
- Tests: `python -m unittest discover tests`
- The old version is in [`legacy/`](legacy/) for reference.
