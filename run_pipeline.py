"""
Full Catalogue Prep Pipeline - Step 1 theke Step 5 (Claude cleaning er AGE porjonto)
--------------------------------------------------------------------------------------
Ei script apnar pura workflow-er shuru theke Claude-e deyar age porjonto shob
kaj ekta interactive run-e kore dey:

  Step 1  - "Image Links Mapping (...).xlsx" file e full URL bosano
            (add_http_to_image_key.py) -> image_http_url_add.xlsx

  Step 2  - [MANUAL] Ei mapping file e jei seller-der product ache, shei
            seller-der shob product export file download kore combine kore
            "Basic combine.xlsx" banano. EI STEP TA SCRIPT AUTOMATE KORE
            NA - apnake nijei korte hobe (karon eta seller system theke
            manual export).

  Step 3  - "Basic combine.xlsx" theke shudhu dorkari 7 ta column rekhe
            "Basic combine H & D.xlsx" banano (trim_columns.py)

  Step 4  - image_http_url_add.xlsx er URL gula "Basic combine H & D.xlsx"
            er HTML column e merge kora (merge_images_into_html.py)
            -> ..._image_replaced.xlsx

  Step 5  - Result file ke choto choto part e split kora, jate Claude e
            upload kora jay (split_excel.py) -> ekta "..._split" folder

Step 6 (Claude AI diye H&D_improved_prompt.md onujayi clean kora) EI SCRIPT
KORE NA - oita apnake nijei Claude e file + prompt diye korte hobe, karon
oi step-e manual review dorkar. Clean kora file gula ekta folder e joma
korle, porer ধাপ (merge back + upload template e bosano) er jonno alada
script ase.

Kivabe run korbe:
    python run_pipeline.py

Shorto: ei file, add_http_to_image_key.py, trim_columns.py, ar
merge_images_into_html.py, split_excel.py - shob ekই folder e thakte hobe.
"""

import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from add_http_to_image_key import process_file as add_http_urls
from trim_columns import trim_file
from merge_images_into_html import process as merge_images
from split_excel import split_excel


def ask(prompt_text, default=None):
    val = input(prompt_text).strip().strip('"')
    if not val and default is not None:
        return default
    return val


def step_header(n, title):
    print(f"\n{'=' * 60}")
    print(f"STEP {n}: {title}")
    print(f"{'=' * 60}")


def main():
    print("=== Full Catalogue Prep Pipeline (Step 1-5) ===")

    # ---------------------------------------------------------------
    step_header(1, "Image Links Mapping file e URL bosano")
    mapping_path = ask("'Image Links Mapping (...).xlsx' file er full path dao: ")
    base_url = ask(
        "Base URL dao (e.g. https://sl-dev-s3.s3.amazonaws.com/product/): "
    )
    image_url_file = add_http_urls(mapping_path, base_url)

    # ---------------------------------------------------------------
    step_header(2, "Seller Export Combine (MANUAL STEP)")
    print("Ekhon apnake manually shei mapping file-e thaka seller-der")
    print("shob product export file download kore combine kore")
    print("'Basic combine.xlsx' banate hobe (ei column format e):")
    print("  Product ID | Seller Code | ... | Highlights(English) | ")
    print("  Highlights(Bengali) | Description (English) | Description (Bengali) | ...")
    input("\nKaj shesh hole Enter chapun continue korte...")
    basic_combine_path = ask("'Basic combine.xlsx' file er full path dao: ")

    # ---------------------------------------------------------------
    step_header(3, "Dorkari column diye 'Basic combine H & D.xlsx' banano")
    trimmed_path = trim_file(basic_combine_path)

    # ---------------------------------------------------------------
    step_header(4, "Image URL gula HTML column e merge kora")
    merged_path = merge_images(trimmed_path, image_url_file)

    # ---------------------------------------------------------------
    step_header(5, "Choto choto part e split kora (Claude-e deyar jonno)")
    header_rows = int(ask("Koyta header row? (usually 1) [Enter = 1]: ", default="1"))
    chunk_size = int(
        ask("Protita split file e koyta row thakbe? (e.g. 3500) [Enter = 3500]: ", default="3500")
    )
    split_excel(merged_path, header_rows, chunk_size)

    base_no_ext = os.path.splitext(merged_path)[0]
    split_folder = f"{base_no_ext}_split"

    print("\n" + "=" * 60)
    print("STEP 1-5 COMPLETE!")
    print("=" * 60)
    print(f"\nSplit file gula ekhane pawa jabe:\n    {split_folder}\n")
    print("Ekhon apni:")
    print("  1. Protita split file Claude-e upload korun, sathe H&D_improved_prompt.md")
    print("  2. Claude-er deya output diye file thik kore nin")
    print("  3. Shob cleaned/updated file ekta alada folder e joma korun")
    print("  4. Tarpor merge-back + upload template e bosanor jonno agent-ke dite paren")


if __name__ == "__main__":
    main()
