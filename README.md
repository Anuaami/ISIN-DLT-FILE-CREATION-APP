# 📊 Excel ISIN Filter & Splitter Application

A high-performance Python and Streamlit application that accepts an Excel workbook, discards all extraneous columns, retains only the **30 required columns** in specified order, and splits the data into separate Excel files by **`isin_code`**.

---

## 🎯 Retained Columns (30 in exact order)

1. `interest_id`
2. `unit_code`
3. `security_code`
4. `int_type`
5. `isin_code`
6. `dpid`
7. `holder_folio`
8. `holder`
9. `holder_addr1`
10. `holder_pin`
11. `bank_accno`
12. `bank_name`
13. `ifsc_code`
14. `ecs_actype`
15. `bfitpan`
16. `hold_minor`
17. `bonds`
18. `princ_amt`
19. `gross_amt`
20. `int_per`
21. `face_value`
22. `fr_date`
23. `to_date`
24. `days`
25. `wardate`
26. `mode_pay`
27. `warno`
28. `war_acno`
29. `type`
30. `benpos_date`

---

## 🚀 How to Run the Application

### Method 1: Web Interface (Streamlit)
Double-click `run.bat` or run:
```bash
streamlit run app.py
```
This opens the browser UI where you can:
- Drag-and-drop your `.xlsx` or `.xls` file.
- Inspect total rows, unique ISINs, and retained/dropped columns.
- Preview each ISIN group individually.
- Download all files at once as a single **ZIP bundle** (`isin_filtered_excel_files.zip`).
- Download individual `<isin_code>.xlsx` files.

### Method 2: Command-Line Interface (CLI / Batch Mode)
To process files directly from the terminal or script:
```bash
python cli.py "C:\path\to\your_input.xlsx" --output-dir "./output"
```
**CLI Options**:
- `-o, --output-dir`: Target folder for output files (default: `./output`).
- `-s, --sheet`: Specify sheet name (defaults to first sheet).
- `--header-row`: 1-indexed row number for column headers (auto-detected if omitted).
- `--keep-summary`: Do not filter out trailing summary/grand-total rows.
- `--no-zip`: Skip generating the consolidated `.zip` file.
- `--no-fill-missing`: Do not create blank columns for missing target fields.

---

## 🧪 Testing

Run the automated test suite:
```bash
python test_processor.py
```
Generate test sample data:
```bash
python generate_sample.py
```

---

## 📁 Project Structure

```
excel-isin-splitter/
├── app.py               # Streamlit web application
├── core_processor.py    # Core column pruning, ISIN splitting, and openpyxl styling logic
├── cli.py               # Command-line interface for batch processing
├── generate_sample.py   # Test dataset generator
├── test_processor.py    # Verification test suite
├── run.bat              # 1-click Windows batch launcher
└── README.md            # Documentation
```
