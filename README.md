# Catalogue Prep Agent (Image URL Replace)

## Goal in one line

Replace the **old external images** (Daraz CDN etc.) inside product
Highlights/Descriptions on **seller-admin.cartup.com** with **our own S3
image URLs**, clean the text with Claude, and produce **upload-ready bulk
update files** + a **QC approval file**.

## The big picture

```
 Image mapping file (from Bakaul)            seller-admin.cartup.com
  - Data Mapping: Product ID -> image keys     (Multi Seller Bulk Update
  - Seller Info : Seller Codes                  -> Export Excel File)
          │                                              │
          ▼                                              ▼
 [1] Add base URL to image keys            [2] Export every seller's products
     -> image_http_url_add.xlsx                 -> Basic combine.xlsx
          │                                              │
          │                               [3] Keep 7 columns
          │                                   -> Basic combine H & D.xlsx
          └──────────────┬───────────────────────────────┘
                         ▼
          [4] Swap <img src> in the 4 HTML columns with new URLs
              -> ..._image_replaced.xlsx  (+ "Needs Review" sheet)
                         ▼
          [5] Split into 3,500-row parts  -> ..._split/
                         ▼
          ===== MANUAL: clean each part in Claude (H&D prompt) =====
                         ▼
          [9] Merge cleaned parts back   -> Basic combine cleaned.xlsx
                         ▼
          [6] Put cleaned content into the official upload template,
              split into upload files + Product-approval.xlsx
              -> update/update_part1.xlsx, update_part2.xlsx, ...
                         ▼
          ===== MANUAL: upload in seller-admin =====
```

## How it is used today

| What you do | How |
|---|---|
| Put the new mapping file in `Data/` | name must contain `image`, `link`, `mapping` |
| Run steps 1–5 | double-click `run_auto.bat` (OTP must be typed during step 2) |
| Clean parts in Claude | upload each part + `H&D_improved_prompt.md` |
| Put cleaned parts in `H&D_Processed_Input/` | |
| Build upload files | double-click `Merge_Claude_Output.bat` (runs merge-back, then step 6) |
| Upload | files in `D:\Image URL change\<Month Year>\<DD-MM-YYYY>\update\` |

## Documents

- [Full documentation](docs/DOCUMENTATION.md): every step, file, setting and known problem.
- [Redesign plan](docs/REDESIGN.md): the proposed simpler version.

## Setup

```
pip install pandas openpyxl playwright
python -m playwright install chromium
```

Keep `credentials.json` and `auth_state.json` private. They are in `.gitignore`.
