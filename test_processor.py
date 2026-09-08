import io
import os
import zipfile
import pandas as pd
from core_processor import (
    TARGET_COLUMNS,
    detect_header_row,
    process_dataframe,
    split_by_isin,
    dataframe_to_styled_excel,
    create_isin_zip_archive,
)
from generate_sample import create_sample_excel


def test_standard_sample():
    sample_path = "sample_test.xlsx"
    create_sample_excel(sample_path)

    # 1. Read input Excel
    xl = pd.ExcelFile(sample_path)
    preview = xl.parse("Sheet1", header=None, nrows=20)
    best_row_idx, match_count = detect_header_row(preview)
    assert best_row_idx == 0, f"Expected header row 0, got {best_row_idx}"
    assert match_count == 30, f"Expected 30 matches, got {match_count}"

    df_raw = xl.parse("Sheet1", header=best_row_idx)
    assert len(df_raw.columns) == 35, f"Expected 35 columns, got {len(df_raw.columns)}"

    # 2. Process dataframe
    df_filtered, col_mapping, missing_targets, dropped_columns, dropped_summary = process_dataframe(df_raw)

    print("\n--- STANDARD SAMPLE TEST REPORT ---")
    print(f"Original Columns Count: {len(df_raw.columns)}")
    print(f"Filtered Columns Count: {len(df_filtered.columns)}")
    print(f"Dropped Columns ({len(dropped_columns)}): {dropped_columns}")
    print(f"Missing Target Columns ({len(missing_targets)}): {missing_targets}")

    assert len(df_filtered.columns) == 30
    assert list(df_filtered.columns) == TARGET_COLUMNS
    assert len(dropped_columns) == 5

    # 3. Split by ISIN
    groups = split_by_isin(df_filtered)
    assert "INE002A01018" in groups
    assert "INE040A01034" in groups
    assert "INE758T01015" in groups
    assert "UNASSIGNED_ISIN" in groups

    print("Standard sample test passed!")


def test_real_ccac_file():
    real_file = r"C:\Users\USER\Desktop\ncd public issue\INTEREST FILE\ISSUE III\CCAC_109918- 3.xlsx"
    if not os.path.exists(real_file):
        print("Real CCAC file not found at expected path, skipping.")
        return

    xl = pd.ExcelFile(real_file)
    assert "InterestRegister" in xl.sheet_names

    # Auto-detect header row
    preview = xl.parse("InterestRegister", header=None, nrows=30)
    best_row_idx, match_count = detect_header_row(preview)
    print(f"\n--- REAL CCAC FILE TEST ---")
    print(f"Auto-detected header row (0-indexed): {best_row_idx} (Excel row {best_row_idx + 1})")
    print(f"Matched target columns: {match_count} / {len(TARGET_COLUMNS)}")

    assert best_row_idx == 2, f"Expected header row 2 (row 3 in Excel), got {best_row_idx}"
    assert match_count == 30, f"Expected 30 matches, got {match_count}"

    df_raw = xl.parse("InterestRegister", header=best_row_idx)
    df_filtered, col_mapping, missing_targets, dropped_columns, dropped_summary = process_dataframe(df_raw)

    print(f"Total rows in raw: {len(df_raw)}")
    print(f"Filtered rows: {len(df_filtered)}")
    print(f"Summary rows dropped: {dropped_summary}")
    print(f"Target columns kept: {len(df_filtered.columns)}")
    print(f"Dropped unwanted columns: {len(dropped_columns)}")

    assert len(df_filtered.columns) == 30
    assert list(df_filtered.columns) == TARGET_COLUMNS
    assert len(missing_targets) == 0
    assert dropped_summary == 1, f"Expected 1 summary footer row dropped, got {dropped_summary}"

    groups = split_by_isin(df_filtered)
    print(f"ISIN groups found: {list(groups.keys())}")
    for isin, g_df in groups.items():
        print(f"  ISIN {isin}: {len(g_df)} records")

    assert "INE051307AT0" in groups
    assert "INE051307AU8" in groups
    assert len(groups["INE051307AT0"]) == 425
    assert len(groups["INE051307AU8"]) == 274

    # Test zip export
    zip_bytes = create_isin_zip_archive(groups)
    assert len(zip_bytes) > 0
    with zipfile.ZipFile(io.BytesIO(zip_bytes)) as z:
        assert "INE051307AT0.xlsx" in z.namelist()
        assert "INE051307AU8.xlsx" in z.namelist()

    print("REAL CCAC FILE TEST PASSED 100%!")


if __name__ == "__main__":
    test_standard_sample()
    test_real_ccac_file()
