import io
import zipfile
import pandas as pd
from core_processor import (
    TARGET_COLUMNS,
    process_dataframe,
    split_by_isin,
    dataframe_to_styled_excel,
    create_isin_zip_archive,
)
from generate_sample import create_sample_excel


def test_end_to_end_processor():
    sample_path = "sample_test.xlsx"
    create_sample_excel(sample_path)

    # 1. Read input Excel
    df_raw = pd.read_excel(sample_path)
    assert len(df_raw.columns) == 35, f"Expected 35 columns, got {len(df_raw.columns)}"

    # 2. Process dataframe
    df_filtered, col_mapping, missing_targets, dropped_columns = process_dataframe(df_raw)

    print("\n--- TEST REPORT ---")
    print(f"Original Columns Count: {len(df_raw.columns)}")
    print(f"Filtered Columns Count: {len(df_filtered.columns)}")
    print(f"Dropped Columns ({len(dropped_columns)}): {dropped_columns}")
    print(f"Missing Target Columns ({len(missing_targets)}): {missing_targets}")

    # Verify column count and exact names
    assert len(df_filtered.columns) == 30, f"Expected 30 columns, got {len(df_filtered.columns)}"
    assert list(df_filtered.columns) == TARGET_COLUMNS, "Columns do not match TARGET_COLUMNS order"
    assert len(dropped_columns) == 5, f"Expected 5 dropped columns, got {len(dropped_columns)}"
    assert "dummy_col_xyz" in dropped_columns
    assert "internal_tracking_id" in dropped_columns

    # 3. Split by ISIN
    groups = split_by_isin(df_filtered)
    print(f"ISIN Groups found ({len(groups)}): {list(groups.keys())}")

    assert "INE002A01018" in groups
    assert "INE040A01034" in groups
    assert "INE758T01015" in groups
    assert "UNASSIGNED_ISIN" in groups

    assert len(groups["INE002A01018"]) == 5
    assert len(groups["INE040A01034"]) == 5
    assert len(groups["INE758T01015"]) == 4
    assert len(groups["UNASSIGNED_ISIN"]) == 1

    # 4. Test single excel styling and round-trip
    excel_bytes = dataframe_to_styled_excel(groups["INE002A01018"])
    assert len(excel_bytes) > 0

    df_read_back = pd.read_excel(io.BytesIO(excel_bytes))
    assert list(df_read_back.columns) == TARGET_COLUMNS
    assert len(df_read_back) == 5

    # 5. Test ZIP archive generation
    zip_bytes = create_isin_zip_archive(groups)
    assert len(zip_bytes) > 0

    with zipfile.ZipFile(io.BytesIO(zip_bytes)) as z:
        namelist = z.namelist()
        print(f"Files in ZIP ({len(namelist)}): {namelist}")
        assert "INE002A01018.xlsx" in namelist
        assert "INE040A01034.xlsx" in namelist
        assert "INE758T01015.xlsx" in namelist
        assert "UNASSIGNED_ISIN.xlsx" in namelist

    print("\nALL TESTS PASSED SUCCESSFULLY!\n")


if __name__ == "__main__":
    test_end_to_end_processor()
