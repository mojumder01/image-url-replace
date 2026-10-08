"""
Excel Splitter (Interactive)
-----------------------------
Run this script directly. It will ask you:
  1. The path of the Excel file
  2. How many header rows it has
  3. How many data rows per split file

It then creates a folder next to your file and saves all split workbooks
inside it, each with the header row(s) repeated at the top.

Usage:
    python split_excel.py
"""

import math
import os
import pandas as pd


def get_file_path():
    while True:
        path = input("Enter the full path of your Excel file: ").strip().strip('"')
        if os.path.isfile(path):
            return path
        print(f"File not found: {path}\nTry again.\n")


def get_header_rows():
    while True:
        val = input("How many header rows does the file have? (usually 1): ").strip()
        if val.isdigit() and int(val) > 0:
            return int(val)
        print("Please enter a valid positive number.\n")


def get_chunk_size():
    while True:
        val = input("How many rows per split file?: ").strip()
        if val.isdigit() and int(val) > 0:
            return int(val)
        print("Please enter a valid positive number.\n")


def split_excel(input_path, header_rows, chunk_size, output_dir=None):
    base_name = os.path.splitext(os.path.basename(input_path))[0]
    if output_dir is None:
        outdir = os.path.join(os.path.dirname(os.path.abspath(input_path)), f"{base_name}_split")
    else:
        outdir = os.path.join(output_dir, f"{base_name}_split")
    os.makedirs(outdir, exist_ok=True)

    # header_rows=1 -> header=0 ; header_rows=2 -> header=[0,1], etc.
    header_arg = 0 if header_rows == 1 else list(range(header_rows))
    df = pd.read_excel(input_path, header=header_arg)

    total_rows = len(df)
    if total_rows == 0:
        print("No data rows found. Nothing to split.")
        return

    num_parts = math.ceil(total_rows / chunk_size)

    for i in range(num_parts):
        start = i * chunk_size
        end = min(start + chunk_size, total_rows)
        chunk_df = df.iloc[start:end]

        out_path = os.path.join(outdir, f"{base_name}_part{i + 1}.xlsx")
        chunk_df.to_excel(out_path, index=False)
        print(f"Wrote {out_path}  ({end - start} rows + header)")

    print(f"\nDone. {total_rows} rows split into {num_parts} file(s) of up to {chunk_size} rows each.")
    print(f"Output folder: {outdir}")


def main():
    print("=== Excel Splitter ===\n")
    input_path = get_file_path()
    header_rows = get_header_rows()
    chunk_size = get_chunk_size()
    print()
    split_excel(input_path, header_rows, chunk_size)
    input("\nPress Enter to exit...")


if __name__ == "__main__":
    main()
