import io
import os
import pandas as pd
import streamlit as st

from core_processor import (
    TARGET_COLUMNS,
    detect_header_row,
    process_dataframe,
    split_by_isin,
    dataframe_to_styled_excel,
    create_isin_zip_archive,
)
from generate_sample import create_sample_excel

# Page configuration
st.set_page_config(
    page_title="ISIN Excel Splitter & Column Reducer",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

# Custom CSS for styling
st.markdown(
    """
    <style>
    .main-header {
        font-size: 2.2rem;
        font-weight: 700;
        color: #1E3A8A;
        margin-bottom: 0.2rem;
    }
    .sub-header {
        font-size: 1.05rem;
        color: #4B5563;
        margin-bottom: 1.5rem;
    }
    .metric-card {
        background-color: #F8FAFC;
        border: 1px solid #E2E8F0;
        border-radius: 10px;
        padding: 16px;
        text-align: center;
        box-shadow: 0 1px 3px rgba(0,0,0,0.05);
    }
    .metric-value {
        font-size: 1.8rem;
        font-weight: 700;
        color: #1D4ED8;
    }
    .metric-label {
        font-size: 0.85rem;
        font-weight: 500;
        color: #64748B;
        text-transform: uppercase;
        letter-spacing: 0.05em;
    }
    .badge {
        display: inline-block;
        padding: 3px 8px;
        margin: 2px;
        border-radius: 6px;
        font-size: 0.78rem;
        font-family: monospace;
    }
    .badge-success {
        background-color: #DCFCE7;
        color: #166534;
        border: 1px solid #BBF7D0;
    }
    .badge-warning {
        background-color: #FEF3C7;
        color: #92400E;
        border: 1px solid #FDE68A;
    }
    .badge-danger {
        background-color: #FEE2E2;
        color: #991B1B;
        border: 1px solid #FECACA;
    }
    .info-banner {
        background-color: #EFF6FF;
        border-left: 4px solid #3B82F6;
        padding: 10px 14px;
        border-radius: 4px;
        margin-bottom: 15px;
        font-size: 0.95rem;
        color: #1E40AF;
    }
    </style>
    """,
    unsafe_allow_html=True,
)

# Sidebar
with st.sidebar:
    st.image("https://img.icons8.com/fluency/96/microsoft-excel-2019.png", width=64)
    st.title("Settings & Info")
    st.markdown("This app filters input Excel sheets to only **30 mandated columns** and splits records by **`isin_code`**.")

    fill_missing = st.checkbox(
        "Fill Missing Target Columns",
        value=True,
        help="If any of the 30 required columns are missing, keep them with blank values to preserve exact schema format.",
    )

    drop_summary = st.checkbox(
        "Exclude Summary/Total Footer Rows",
        value=True,
        help="Automatically detects and removes trailing grand-total / summary rows so output files contain pure record data.",
    )

    st.markdown("---")
    st.subheader(f"Mandated Columns ({len(TARGET_COLUMNS)})")
    with st.expander("View Full Target Column List"):
        for i, col in enumerate(TARGET_COLUMNS, 1):
            st.markdown(f"`{i:02d}.` **{col}**")

    st.markdown("---")
    st.caption("Excel ISIN Filter & Splitter • Built with Streamlit & OpenPyXL")

# Main Header
st.markdown('<div class="main-header">📊 Excel ISIN Filter & Splitter</div>', unsafe_allow_html=True)
st.markdown(
    '<div class="sub-header">Upload an Excel file to automatically strip unwanted columns, retain the 30 specified attributes, and generate individual ISIN-wise workbooks.</div>',
    unsafe_allow_html=True,
)

# File uploader & sample load
col_upload, col_sample = st.columns([3, 1])

with col_upload:
    uploaded_file = st.file_uploader(
        "Choose an Excel file (.xlsx or .xls)",
        type=["xlsx", "xls"],
        help="Upload the master Excel file containing interest/bond records.",
    )

with col_sample:
    st.markdown("<br>", unsafe_allow_html=True)
    use_sample = st.button("🧪 Load Sample Test File", help="Loads a realistic 35-column sample dataset with 3 ISINs.")

file_source = None
if uploaded_file is not None:
    file_source = uploaded_file
elif use_sample:
    sample_path = os.path.join(os.path.dirname(__file__), "sample_test.xlsx")
    if not os.path.exists(sample_path):
        create_sample_excel(sample_path)
    with open(sample_path, "rb") as f:
        file_source = io.BytesIO(f.read())
    st.info("Loaded sample dataset with 15 rows, 35 columns, and 3 distinct ISINs.")

if file_source is not None:
    try:
        # Read Excel workbook sheets
        excel_file = pd.ExcelFile(file_source)
        sheet_names = excel_file.sheet_names

        col_sheet, col_header_row = st.columns([2, 1])

        with col_sheet:
            selected_sheet = sheet_names[0]
            if len(sheet_names) > 1:
                selected_sheet = st.selectbox(
                    "Select Sheet to Process",
                    sheet_names,
                    index=0,
                    help="This workbook has multiple sheets. Choose the one to process.",
                )

        # Pre-scan top 30 rows to auto-detect header row
        preview_raw = excel_file.parse(selected_sheet, header=None, nrows=30)
        auto_detected_idx, match_count = detect_header_row(preview_raw)
        auto_detected_row_1indexed = auto_detected_idx + 1

        with col_header_row:
            header_row_choice = st.number_input(
                "Header Row in Excel",
                min_value=1,
                max_value=max(len(preview_raw), 1),
                value=auto_detected_row_1indexed,
                step=1,
                help=f"Row containing the column headers. Auto-detected as Row {auto_detected_row_1indexed}.",
            )

        if match_count > 0:
            if auto_detected_row_1indexed > 1:
                st.markdown(
                    f'<div class="info-banner">✨ <b>Header row auto-detected at Row {auto_detected_row_1indexed}</b> ({match_count} of 30 target columns found). Title lines on Rows 1 to {auto_detected_row_1indexed - 1} are automatically skipped.</div>',
                    unsafe_allow_html=True,
                )
            else:
                st.markdown(
                    f'<div class="info-banner">✨ <b>Header row auto-detected at Row 1</b> ({match_count} of 30 target columns found).</div>',
                    unsafe_allow_html=True,
                )
        else:
            st.warning("⚠️ No standard target column names matched in the first rows. Please ensure the correct Header Row is selected above.")

        # Load data using chosen header row
        df_raw = excel_file.parse(selected_sheet, header=header_row_choice - 1)

        if df_raw.empty:
            st.error("The selected sheet is empty under the specified header row.")
            st.stop()

        # Process columns and rows
        filtered_df, col_mapping, missing_targets, dropped_columns, dropped_summary_rows = process_dataframe(
            df_raw, fill_missing_cols=fill_missing, drop_summary_rows=drop_summary
        )

        # Check if isin_code is present
        if "isin_code" not in filtered_df.columns or filtered_df["isin_code"].isnull().all():
            st.warning("⚠️ Warning: 'isin_code' column could not be found or is completely empty. Please verify the Header Row setting above.")

        # Split by ISIN
        isin_groups = split_by_isin(filtered_df)
        unique_isin_count = len([k for k in isin_groups.keys() if k != "UNASSIGNED_ISIN"])

        # Metric cards
        m1, m2, m3, m4 = st.columns(4)
        with m1:
            st.markdown(
                f"""
                <div class="metric-card">
                    <div class="metric-value">{len(filtered_df):,}</div>
                    <div class="metric-label">Data Records</div>
                </div>
                """,
                unsafe_allow_html=True,
            )
        with m2:
            st.markdown(
                f"""
                <div class="metric-card">
                    <div class="metric-value">{unique_isin_count}</div>
                    <div class="metric-label">Unique ISINs</div>
                </div>
                """,
                unsafe_allow_html=True,
            )
        with m3:
            st.markdown(
                f"""
                <div class="metric-card">
                    <div class="metric-value">{len(TARGET_COLUMNS) - len(missing_targets)} / {len(TARGET_COLUMNS)}</div>
                    <div class="metric-label">Target Columns Kept</div>
                </div>
                """,
                unsafe_allow_html=True,
            )
        with m4:
            st.markdown(
                f"""
                <div class="metric-card">
                    <div class="metric-value" style="color: #DC2626;">{len(dropped_columns)}</div>
                    <div class="metric-label">Unwanted Columns Dropped</div>
                </div>
                """,
                unsafe_allow_html=True,
            )

        if dropped_summary_rows > 0:
            st.caption(f"ℹ️ Excluded {dropped_summary_rows} footer/summary row(s) (e.g. Grand Total) to keep output files clean.")

        st.markdown("<br>", unsafe_allow_html=True)

        # Column Audit Details
        with st.expander("🔍 Column Audit Report (Kept vs Dropped Columns)", expanded=False):
            c_audit1, c_audit2 = st.columns(2)
            with c_audit1:
                st.markdown(f"**Retained Target Columns ({len(TARGET_COLUMNS) - len(missing_targets)}):**")
                matched_badges = "".join(
                    [f'<span class="badge badge-success">✓ {col}</span>' for col in TARGET_COLUMNS if col not in missing_targets]
                )
                st.markdown(matched_badges, unsafe_allow_html=True)

                if missing_targets:
                    st.markdown(f"<br>**Missing Target Columns ({len(missing_targets)}):**", unsafe_allow_html=True)
                    missing_badges = "".join(
                        [f'<span class="badge badge-danger">✗ {col}</span>' for col in missing_targets]
                    )
                    st.markdown(missing_badges, unsafe_allow_html=True)
                    st.caption("These columns were not found in the source file and are filled with blank values.")

            with c_audit2:
                st.markdown(f"**Dropped Unwanted Columns ({len(dropped_columns)}):**")
                if dropped_columns:
                    dropped_badges = "".join(
                        [f'<span class="badge badge-warning">✕ {col}</span>' for col in dropped_columns]
                    )
                    st.markdown(dropped_badges, unsafe_allow_html=True)
                else:
                    st.markdown("*None — the input file contained no extraneous columns.*")

        st.markdown("---")

        # Download All Section
        down_col1, down_col2 = st.columns([2, 1])
        with down_col1:
            st.subheader("📦 Download Processed Files")
            st.markdown("Download all ISIN-wise Excel files packaged into a single ZIP archive, or download individually below.")
        with down_col2:
            zip_data = create_isin_zip_archive(isin_groups)
            st.download_button(
                label=f"⬇️ Download All ({len(isin_groups)} ISIN Files) as ZIP",
                data=zip_data,
                file_name="isin_filtered_excel_files.zip",
                mime="application/zip",
                type="primary",
                use_container_width=True,
            )

        # Tabs for Overview vs Detailed Preview
        tab_summary, tab_inspect = st.tabs(["📋 Summary by ISIN", "🔎 Inspect & Download Individual ISIN"])

        with tab_summary:
            summary_data = []
            for isin, group_df in isin_groups.items():
                record_count = len(group_df)
                gross_total = group_df["gross_amt"].sum() if "gross_amt" in group_df.columns and pd.api.types.is_numeric_dtype(group_df["gross_amt"]) else "N/A"
                princ_total = group_df["princ_amt"].sum() if "princ_amt" in group_df.columns and pd.api.types.is_numeric_dtype(group_df["princ_amt"]) else "N/A"
                
                summary_data.append({
                    "ISIN Code": isin,
                    "Total Records": record_count,
                    "Total Gross Amount": f"₹ {gross_total:,.2f}" if isinstance(gross_total, (int, float)) else str(gross_total),
                    "Total Principal Amount": f"₹ {princ_total:,.2f}" if isinstance(princ_total, (int, float)) else str(princ_total),
                })
            
            summary_df = pd.DataFrame(summary_data)
            st.dataframe(summary_df, use_container_width=True, hide_index=True)

        with tab_inspect:
            isin_keys = list(isin_groups.keys())
            chosen_isin = st.selectbox("Select ISIN Code to inspect:", isin_keys)

            target_group_df = isin_groups[chosen_isin]

            col_sub_info, col_sub_btn = st.columns([3, 1])
            with col_sub_info:
                st.markdown(f"**Showing {len(target_group_df)} records for `{chosen_isin}`** (all 30 filtered columns in exact order):")
            with col_sub_btn:
                single_excel_bytes = dataframe_to_styled_excel(target_group_df, sheet_name=chosen_isin[:31])
                st.download_button(
                    label=f"⬇️ Download {chosen_isin}.xlsx",
                    data=single_excel_bytes,
                    file_name=f"{chosen_isin}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                )

            st.dataframe(target_group_df, use_container_width=True, hide_index=True)

    except Exception as e:
        st.error(f"An error occurred while processing the Excel file: {str(e)}")
        st.exception(e)
else:
    st.info("👆 Please upload an Excel file or click 'Load Sample Test File' to begin.")
