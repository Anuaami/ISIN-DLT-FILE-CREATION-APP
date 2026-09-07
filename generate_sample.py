import pandas as pd
from core_processor import TARGET_COLUMNS

def create_sample_excel(filepath: str = "sample_data.xlsx"):
    # Realistic dummy data
    data = {
        # Target columns with varying case and spaces to test normalization
        "interest_id": [f"INT-{1000+i}" for i in range(15)],
        "UNIT_CODE": ["U01", "U01", "U02", "U02", "U03"] * 3,
        "security_code": ["SEC101", "SEC102", "SEC103", "SEC101", "SEC102"] * 3,
        "int_type": ["Annual", "Semi-Annual", "Annual", "Quarterly", "Annual"] * 3,
        "isin_code": [
            "INE002A01018", "INE002A01018", "INE002A01018", "INE002A01018", "INE002A01018",
            "INE040A01034", "INE040A01034", "INE040A01034", "INE040A01034", "INE040A01034",
            "INE758T01015", "INE758T01015", "INE758T01015", "INE758T01015", None
        ],
        "dpid": ["IN300123", "IN300456", "12010900", "IN300789", "12020000"] * 3,
        "holder_folio": [f"FOLIO_{500+i}" for i in range(15)],
        "holder": [
            "Aarav Sharma", "Priya Patel", "Vikram Malhotra", "Sunita Rao", "Rajesh Gupta",
            "Ananya Sen", "Kavita Iyer", "Amitabh Verma", "Neha Joshi", "Rohan Mehta",
            "Deepak Chopra", "Kiran Bedi", "Suresh Raina", "Meena Kumari", "Ajay Dev"
        ],
        "holder_addr1": [f"Flat {100+i}, MG Road" for i in range(15)],
        "holder_pin": [400001, 110001, 560001, 600001, 700001] * 3,
        "bank_accno": [f"9182736450{i:02d}" for i in range(15)],
        "bank_name": ["HDFC Bank", "State Bank of India", "ICICI Bank", "Axis Bank", "Punjab National Bank"] * 3,
        "ifsc_code": ["HDFC0001234", "SBIN0004321", "ICIC0009876", "UTIB0005544", "PUNB0001122"] * 3,
        "ecs_actype": ["Savings", "Current", "Savings", "Savings", "Current"] * 3,
        "bfitpan": ["ABCDE1234F", "BCDEF2345G", "CDEFG3456H", "DEFGH4567I", "EFGHI5678J"] * 3,
        "hold_minor": ["N", "N", "Y", "N", "N"] * 3,
        "bonds": [10, 25, 50, 100, 15] * 3,
        "princ_amt": [10000.0, 25000.0, 50000.0, 100000.0, 15000.0] * 3,
        "gross_amt": [850.0, 2125.0, 4250.0, 8500.0, 1275.0] * 3,
        "int_per": [8.5, 8.5, 8.5, 8.5, 8.5] * 3,
        "face_value": [1000, 1000, 1000, 1000, 1000] * 3,
        "fr_date": ["2025-04-01"] * 15,
        "to_date": ["2026-03-31"] * 15,
        "days": [365] * 15,
        "wardate": ["2026-04-15"] * 15,
        "mode_pay": ["NEFT", "RTGS", "NACH", "NEFT", "Warrant"] * 3,
        "warno": [f"WAR_{8000+i}" for i in range(15)],
        "war_acno": [f"WA_{9000+i}" for i in range(15)],
        "type": ["Regular"] * 15,
        "benpos_date": ["2026-03-15"] * 15,

        # Extra unwanted columns that should be automatically removed by the app
        "internal_tracking_id": [f"TRK_{i}" for i in range(15)],
        "temp_notes": ["Verified", "Pending review", "Audited", "Flagged", "Ok"] * 3,
        "system_timestamp": ["2026-09-07 10:00:00"] * 15,
        "operator_name": ["Admin1", "Admin2", "Admin1", "Admin3", "Admin2"] * 3,
        "dummy_col_xyz": [999] * 15,
    }

    df = pd.DataFrame(data)
    df.to_excel(filepath, index=False, engine="openpyxl")
    print(f"Sample Excel file created successfully at: {filepath}")
    print(f"Total rows: {len(df)}, Total columns: {len(df.columns)}")

if __name__ == "__main__":
    create_sample_excel()
