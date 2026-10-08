# Catalogue Prep (Image URL Replace) v2

## Goal

Replace the **old external images** (Daraz CDN etc.) inside product
Highlights/Descriptions on **seller-admin.cartup.com** with **our own S3 image
URLs**. Get the text cleaned by Claude, then produce **upload-ready bulk update
files** and a **QC approval file**.

## The work in 3 stages

```
 1_Prepare.bat         you + Claude              2_Finish.bat
┌──────────────────┐  ┌──────────────────────┐  ┌──────────────────────┐
│ 1 urls           │  │ Upload each Part +   │  │ 6 merge  Claude files│
│ 2 export (OTP)   │→ │ the prompt to Claude │→ │ 7 upload files +     │
│ 3 trim           │  │ Save results in      │  │   approval file      │
│ 4 images         │  │ 6_claude_done\       │  │                      │
│ 5 split          │  │                      │  │ → upload them        │
└──────────────────┘  └──────────────────────┘  └──────────────────────┘
```

## How to use

1. Put the new mapping file in `Data\`. Its name must contain `image`, `link` and
   `mapping` (or `bakaul`), e.g. `image_links_mapping (28.09.2026) Bakaul.xlsx`.
2. Double-click **`1_Prepare.bat`**. A browser opens: type the **OTP**, then
   press Enter in the black window. When it finishes, the `5_for_claude` folder opens.
3. For each `Part-N.xlsx`: upload it with `H&D_improved_prompt.md` to Claude, and save
   the cleaned file into the run's **`6_claude_done`** folder.
4. Double-click **`2_Finish.bat`**. The `8_upload` folder opens. Upload the
   `update_part*.xlsx` files, then `Product-approval.xlsx`.

If prepare stops half-way (OTP, network, ...), fix it and double-click
**`1_Prepare_Resume.bat`**. It continues where it stopped, and already-downloaded
batches are not downloaded again. **`Status.bat`** shows the progress.

## Where the files go

One folder per run, named after the mapping file date and the start time:

```
D:\Image URL change\September 2026_28-09-2026_15-37\
  0_input  1_urls  2_export  3_trim  4_images  5_for_claude
  6_claude_done  7_merged  8_upload  run.json  run.log
```

`8_upload\not_uploaded.xlsx` lists every product that was **not** uploaded and why
(Needs Review rows, rows Claude did not return).

## Setup (once)

```
pip install -r requirements.txt
python -m playwright install chromium
```

Put `credentials.json` (username, password, display_name) in this folder. It
and `auth_state.json` are private and never committed.

## More

- [Full documentation](docs/DOCUMENTATION.md): every step, setting, command and file.
- Settings: [`config.toml`](config.toml)
- Tests: `python -m unittest discover tests`
- The old version is in [`legacy/`](legacy/) for reference.
