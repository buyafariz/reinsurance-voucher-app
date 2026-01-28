import os
import pandas as pd
from datetime import datetime

LOG_COLUMNS = [
    "Seq No",
    "Department",
    "Biz Type",
    "Voucher No",
    "Account With",
    "Cedant Company",
    "PIC",
    "Product",
    "CBY",
    "CBM",
    "OBY",
    "OBM",
    "KOB",
    "COB",
    "MOP",
    "Curr",
    "Total Contribution",
    "Commission",
    "Overiding",
    "Total Commission",
    "Gross Premium Income",
    "Tabarru",
    "Ujrah",
    "Claim",
    "Balance",
    "Rate Exchange",
    "Kontribusi (IDR)",
    "Commission (IDR)",
    "Overiding (IDR)",
    "Total Commission (IDR)",
    "Gross Premium Income (IDR)",
    "Tabarru (IDR)",
    "Ujrah (IDR)",
    "Claim (IDR)",
    "REMARKS",
    "STATUS",
    "CREATED_AT",
    "CREATED_BY"
]

def get_log_path(base_path, year, month):
    period = f"{year}_{month:02d}"
    period_path = os.path.join(base_path, period)
    os.makedirs(period_path, exist_ok=True)
    return os.path.join(period_path, "log_produksi.xlsx")


def load_or_create_log(log_path):
    if os.path.exists(log_path):
        return pd.read_excel(log_path)
    return pd.DataFrame(columns=LOG_COLUMNS)


def generate_vin(base_path, year, month):
    log_path = get_log_path(base_path, year, month)
    log_df = load_or_create_log(log_path)

    if log_df.empty:
        next_seq = 1
    else:
        next_seq = int(log_df["Seq No"].max()) + 1

    vin = f"VIN{year}{month:02d}LST{next_seq:04d}"
    return vin, next_seq, log_path


def create_cancel_row(original_row, new_voucher, seq_no, user, reason):
    cancel = original_row.copy()

    cancel["Biz Type"] = "CANCEL"
    cancel["Seq No"] = seq_no
    cancel["Voucher No"] = new_voucher
    #cancel["ENTRY_TYPE"] = "CANCEL"
    cancel["CANCEL_OF_VIN"] = original_row["Voucher No"]
    cancel["STATUS"] = "CANCELED"
    cancel["CREATED_AT"] = datetime.now()
    cancel["CREATED_BY"] = user
    cancel["CANCEL_REASON"] = reason

    numeric_cols = [
        "Total Contribution",
        "Commission",
        "Overiding",
        "Total Commission",
        "Gross Premium Income",
        "Tabarru",
        "Ujrah",
        "Claim",
        "Balance",
        "Kontribusi (IDR)",
        "Commission (IDR)",
        "Overiding (IDR)",
        "Total Commission (IDR)",
        "Gross Premium Income (IDR)",
        "Tabarru (IDR)",
        "Ujrah (IDR)",
        "Claim (IDR)"
        ]

    for col in numeric_cols:
        cancel[col] = -1 * float(original_row.get(col, 0))

    cancel["REMARKS"] = f"Cancel voucher {original_row['Voucher No']}"

    return cancel
