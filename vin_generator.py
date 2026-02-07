import io
import os
import pandas as pd
from datetime import datetime
from googleapiclient.http import MediaIoBaseDownload
from googleapiclient.http import MediaIoBaseUpload



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

def generate_vin_from_drive(service, period_folder_id, year, month, find_drive_file):
    """
    Generate voucher number berdasarkan log yang ada di Google Drive.
    Tidak tergantung file lokal.
    """

    filename = "log_produksi.xlsx"

    # 🔎 Cek apakah log sudah ada di Drive
    file_id = find_drive_file(
        service=service,
        filename=filename,
        parent_id=period_folder_id
    )

    # ==========================
    # Jika belum ada log sama sekali
    # ==========================
    if not file_id:
        next_seq = 1

    else:
        # ==========================
        # Download log dari Drive ke memory
        # ==========================
        request = service.files().get_media(fileId=file_id)
        file_buffer = io.BytesIO()

        downloader = MediaIoBaseDownload(file_buffer, request)

        done = False
        while not done:
            _, done = downloader.next_chunk()

        file_buffer.seek(0)

        log_df = pd.read_excel(file_buffer)

        if log_df.empty:
            next_seq = 1
        else:
            next_seq = int(log_df["Seq No"].max()) + 1

    voucher = f"VIN{year}{month:02d}LST{next_seq:04d}"

    return voucher, next_seq


def generate_vin_from_drive_log(log_df, year, month):
    if log_df.empty:
        next_seq = 1
    else:
        next_seq = int(log_df["Seq No"].max()) + 1

    voucher = f"VIN{year}{month:02d}LST{next_seq:04d}"
    return voucher, next_seq


def create_negative_excel(original_df):

    negative_cols = [
        "reins premium",
        "reins em premium",
        "reins er premium",
        "reins oth. premium",
        "reins total premium",
        "reins comm",
        "reins em comm",
        "reins er comm",
        "reins oth. comm",
        "reins profit share",
        "reins overriding",
        "reins broker fee",
        "reins total comm",
        "reins tabarru",
        "reins ujrah",
        "reins nett premium"
    ]

    df_negative = original_df.copy()

    for col in negative_cols:
        if col in df_negative.columns:
            df_negative[col] = -1 * df_negative[col]

    return df_negative


def dataframe_to_excel_bytes(df):
    output = io.BytesIO()
    df.to_excel(output, index=False)
    output.seek(0)
    return output


def upload_excel_bytes(service, file_bytes, filename, parent_id):

    file_metadata = {
        "name": filename,
        "parents": [parent_id]
    }

    media = MediaIoBaseUpload(
        file_bytes,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

    service.files().create(
        body=file_metadata,
        media_body=media,
        supportsAllDrives=True
    ).execute()



def create_cancel_row(original_row, new_voucher, seq_no, user, reason):
    cancel = original_row.copy()

    cancel["Biz Type"] = "Cancel"
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
