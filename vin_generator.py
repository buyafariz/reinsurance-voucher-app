import os
import pandas as pd
from datetime import datetime


LOG_COLUMNS = [
    "Seq No",
    "VIN No",

    # Relasi & tipe entry
    "ENTRY_TYPE",        # NORMAL / CANCEL
    "CANCEL_OF_VIN",

    # Informasi bisnis
    "Account With",
    "PIC",
    "PRODUCT",
    "CBY",
    "CBM",
    "OBY",
    "OBM",
    "COB",
    "MOP",

    # Finansial
    "Total Contribution",
    "Commission",
    "Tabarru",
    "Ujrah",
    "Overiding",
    "Claim",
    "Balance",

    # Audit
    "REMARKS",
    "STATUS",            # POSTED / CANCELLED
    "CREATED_AT",
    "CREATED_BY",
    "CANCELLED_AT",
    "CANCELLED_BY",
    "CANCEL_REASON",
]


def generate_vin(base_path, year, month, prefix="VIN", lst="LST"):
    """
    Generate VIN berdasarkan LOG (bukan file voucher)
    VIN TIDAK BOLEH TURUN
    """

    folder = f"{year}_{month:02d}"
    folder_path = os.path.join(base_path, folder)
    os.makedirs(folder_path, exist_ok=True)

    log_path = os.path.join(folder_path, "log_produksi.xlsx")

    # ==========================
    # Buat log jika belum ada
    # ==========================
    if not os.path.exists(log_path):
        log_df = pd.DataFrame(columns=LOG_COLUMNS)
        log_df.to_excel(log_path, index=False)
        next_seq = 1
    else:
        log_df = pd.read_excel(log_path)
        if log_df.empty:
            next_seq = 1
        else:
            next_seq = int(log_df["Seq No"].max()) + 1

    vin = f"{prefix}{year}{month:02d}{lst}{next_seq:03d}"

    return vin, next_seq, log_path


def create_cancel_row(original_row, new_vin, seq_no, user, reason=""):
    """
    Membuat baris CANCEL (negative posting)
    """

    cancel = original_row.copy()

    cancel["Seq No"] = seq_no
    cancel["VIN No"] = new_vin
    cancel["ENTRY_TYPE"] = "CANCEL"
    cancel["CANCEL_OF_VIN"] = original_row["VIN No"]

    cancel["STATUS"] = "POSTED"
    cancel["CREATED_AT"] = datetime.now()
    cancel["CREATED_BY"] = user

    cancel["CANCELLED_AT"] = datetime.now()
    cancel["CANCELLED_BY"] = user
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
        cancel[col] = -1 * float(original_row[col])

    cancel["REMARKS"] = f"Cancel voucher {original_row['VIN No']}"

    return cancel
