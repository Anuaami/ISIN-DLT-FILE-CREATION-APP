import argparse
import os
import re
import sys
import pandas as pd
from core_processor import (
    process_dataframe,
    split_by_isin,
    dataframe_to_styled_excel,
    create_isin_zip_archive,
    TARGET_COLUMNS,
)

# Ensure UTF-8 output encoding on Windows console
if sys.stdout.encoding and sys.stdout.encoding.lower() != "utf-8":
    try:
        sys.stdout.reconfigure(encoding="utf-8")
    except Exception:
        pass


def process_excel_file(
    input_path: str,
    output_dir: str,
    sheet_name: str = None,
    create_zip: bool = True,
    fill_missing: bool = True,
):
    if not os.path.exists(input_path):
        raise FileNotFoundError(f"Input file not found: {input_path}")

    os.makedirs(output_dir, exist_ok=True)

    print(f"\n[INFO] Reading input Excel file: {input_path}")
    excel_file = pd.ExcelFile(input_path)

    if sheet_name is None:
        sheet_name = excel_file.sheet_names[0]
    print(f"[INFO] Processing Sheet: '{sheet_name}'")

    raw_df = excel_file.parse(sheet_name)
    print(f"[INFO] Loaded {len(raw_df)} rows and {len(raw_df.columns)} columns.")

    filtered_df, col_mapping, missing_targets, dropped_columns = process_dataframe(
        raw_df, fill_missing_cols=fill_missing
    )

    print(f"[INFO] Dropped {len(dropped_columns)} unwanted columns: {dropped_columns}")
    if missing_targets:
        print(f"[WARN] Missing target columns ({len(missing_targets)}): {missing_targets}")
    print(f"[INFO] Filtered to {len(filtered_df.columns)} target columns.")

    isin_groups = split_by_isin(filtered_df)
    print(f"[INFO] Found {len(isin_groups)} ISIN groups: {list(isin_groups.keys())}\n")

    # Save individual Excel files
    for isin, sub_df in isin_groups.items():
        safe_isin = re.sub(r'[\\/*?:"<>|]', "_", str(isin))
        out_filename = f"{safe_isin}.xlsx"
        out_filepath = os.path.join(output_dir, out_filename)
        
        excel_bytes = dataframe_to_styled_excel(sub_df, sheet_name=safe_isin[:31])
        with open(out_filepath, "wb") as f:
            f.write(excel_bytes)
        print(f"  -> Saved: {out_filepath} ({len(sub_df)} rows)")

    # Optionally save zip archive
    if create_zip:
        zip_filepath = os.path.join(output_dir, "all_isin_files.zip")
        zip_bytes = create_isin_zip_archive(isin_groups)
        with open(zip_filepath, "wb") as f:
            f.write(zip_bytes)
        print(f"\n[INFO] Bundle created: {zip_filepath}")

    print(f"\n[SUCCESS] All done! Output generated in: {os.path.abspath(output_dir)}")


def main():
    parser = argparse.ArgumentParser(
        description="Filter Excel file to 30 target columns and split records by isin_code."
    )
    parser.add_argument("input", help="Path to input Excel file (.xlsx, .xls)")
    parser.add_argument(
        "-o",
        "--output-dir",
        default="./output",
        help="Directory where output Excel files will be saved (default: ./output)",
    )
    parser.add_argument(
        "-s",
        "--sheet",
        default=None,
        help="Name of sheet to process (defaults to the first sheet)",
    )
    parser.add_argument(
        "--no-zip",
        action="store_true",
        help="Skip creating the bundled all_isin_files.zip archive",
    )
    parser.add_argument(
        "--no-fill-missing",
        action="store_true",
        help="Do not fill missing target columns with blanks",
    )

    args = parser.parse_args()
    process_excel_file(
        input_path=args.input,
        output_dir=args.output_dir,
        sheet_name=args.sheet,
        create_zip=not args.no_zip,
        fill_missing=not args.no_fill_missing,
    )


if __name__ == "__main__":
    main()
