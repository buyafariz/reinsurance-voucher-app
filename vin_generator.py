import os
import pandas as pd
from datetime import datetime
import portalocker

LOG_COLUMNS = [
    "Seq No",
    "VIN No",
    "Account With",
    "PIC",
    "Product",
    "CBY",
    "CBM",
    "OBY",
    "OBM",
    "COB",
    "MOP",
    "Total Contribution",
    "Commission",
    "Tabarru",
    "Ujrah",
    "Overiding",
    "Claim",
    "Balance",
    "REMARKS",
    "STATUS",
    "CREATED_AT",
    "CREATED_BY",
    "CANCELLED_AT",
    "CANCELLED_BY",
    "CANCEL_REASON",
    "ENTRY_TYPE",
    "CANCEL_OF_VIN",
]


def get_log_path(base_path, year, month):
    period = f"{year}_{month:02d}"
    period_path = os.path.join(base_path, period)
    os.makedirs(period_path, exist_ok=True)
    return os.path.join(period_path, "log_produksi.xlsx")


def generate_vin(base_path, year, month):
    """
    Generate VIN dengan file lock
    Aman untuk multi-user (Streamlit Cloud)
    """
    log_path = get_log_path(base_path, year, month)

    # pastikan file ada
    if not os.path.exists(log_path):
        pd.DataFrame(columns=LOG_COLUMNS).to_excel(log_path, index=False)

    # 🔒 LOCK FILE
    with open(log_path, "rb+") as f:
        portalocker.lock(f, portalocker.LOCK_EX)

        log_df = pd.read_excel(log_path)

        if log_df.empty:
            next_seq = 1
        else:
            next_seq = int(log_df["Seq No"].max()) + 1

        vin = f"VIN{year}{month:02d}{next_seq:04d}"

        portalocker.unlock(f)

    return vin, next_seq, log_path


def create_cancel_row(original_row, new_vin, seq_no, user, reason):
    cancel = original_row.copy()

    cancel["Seq No"] = seq_no
    cancel["VIN No"] = new_vin
    cancel["ENTRY_TYPE"] = "CANCEL"
    cancel["CANCEL_OF_VIN"] = original_row["VIN No"]
    cancel["STATUS"] = "POSTED"
    cancel["CREATED_AT"] = datetime.now()
    cancel["CREATED_BY"] = user
    cancel["CANCEL_REASON"] = reason

    numeric_cols = [
        "Total Contribution",
        "Commission",
        "Tabarru",
        "Ujrah",
        "Overiding",
        "Claim",
        "Balance",
    ]

    for col in numeric_cols:
        cancel[col] = -1 * float(original_row.get(col, 0))

    cancel["REMARKS"] = f"Cancel voucher {original_row['VIN No']}"

    return cancel
