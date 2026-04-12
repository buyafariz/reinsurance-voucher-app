import io
import os
import pandas as pd
from datetime import datetime
from googleapiclient.http import MediaIoBaseDownload
from googleapiclient.http import MediaIoBaseUpload
from drive_utils import load_log_from_gsheet, find_drive_file, append_gsheet, upload_dataframe_to_drive
from zoneinfo import ZoneInfo

MONTH_ID = [
    "", "Januari", "Februari", "Maret", "April",
    "Mei", "Juni", "Juli", "Agustus",
    "September", "Oktober", "November", "Desember"
]

def get_log_filename(year, month):
    return f"Log Produksi {MONTH_ID[month]} {year}"

def get_log_pml_filename(year, month):
    return f"Log PML Produksi {MONTH_ID[month]} {year}"

def get_log_filename_outward(year, month):
    return f"Log Produksi {MONTH_ID[month]} {year} - Outward"

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
    return os.path.join(period_path, get_log_filename(year,month))


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


def generate_vin_from_drive(
    service,
    period_folder_id,
    year,
    month,
    find_drive_file,
    biz_type
):
    filename = get_log_filename(year,month)

    file_id = find_drive_file(
        service=service,
        filename=filename,
        parent_id=period_folder_id,
        mime_type="application/vnd.google-apps.spreadsheet"
    )

    # ==========================
    # Jika belum ada log
    # ==========================
    if not file_id:
        next_seq = 1

    else:
        log_df = load_log_from_gsheet(
            service=service,
            spreadsheet_id=file_id
        )

        if log_df.empty or "Seq No" not in log_df.columns:
            next_seq = 1
        else:
            seq_series = pd.to_numeric(log_df["Seq No"], errors="coerce")
            seq_series = seq_series.dropna()

            if seq_series.empty:
                next_seq = 1
            else:
                next_seq = int(seq_series.max()) + 1

    # ==========================
    # Format Voucher
    # ==========================
    if biz_type in ["Kontribusi", "Refund", "Alteration", "Retur", "Revise", "Batal", "Cancel"]:
        voucher = f"VIN{year}{month:02d}LST{next_seq:04d}"

    elif biz_type == "Claim":
        voucher = f"VCL{year}{month:02d}LSC{next_seq:04d}"

    return voucher, next_seq

def generate_pml_from_drive(
    service,
    period_folder_id,
    year,
    month,
    find_drive_file,
    biz_type
):
    filename = get_log_pml_filename(year,month)

    file_id = find_drive_file(
        service=service,
        filename=filename,
        parent_id=period_folder_id,
        mime_type="application/vnd.google-apps.spreadsheet"
    )

    # ==========================
    # Jika belum ada log
    # ==========================
    if not file_id:
        next_seq = 1

    else:
        log_df = load_log_from_gsheet(
            service=service,
            spreadsheet_id=file_id
        )

        if log_df.empty or "Seq No" not in log_df.columns:
            next_seq = 1
        else:
            seq_series = pd.to_numeric(log_df["Seq No"], errors="coerce")
            seq_series = seq_series.dropna()

            if seq_series.empty:
                next_seq = 1
            else:
                next_seq = int(seq_series.max()) + 1

    # ==========================
    # Format Voucher
    # ==========================
    if biz_type in ["Kontribusi", "Refund", "Alteration", "Retur", "Revise", "Batal", "Cancel"]:
        voucher = f"PML{year}{month:02d}LIS{next_seq:04d}"

    elif biz_type == "Claim":
        voucher = f"PLA{year}{month:02d}LSC{next_seq:04d}"

    return voucher, next_seq


def generate_vou_from_drive(
    service,
    outward_folder_id,
    year,
    month,
    find_drive_file,
    biz_type
):
    filename = get_log_filename_outward(year,month)

    file_id = find_drive_file(
        service=service,
        filename=filename,
        parent_id=outward_folder_id,
        mime_type="application/vnd.google-apps.spreadsheet"
    )

    # ==========================
    # Jika belum ada log
    # ==========================
    if not file_id:
        next_seq = 1

    else:
        log_df = load_log_from_gsheet(
            service=service,
            spreadsheet_id=file_id
        )

        if log_df.empty or "Seq No" not in log_df.columns:
            next_seq = 1
        else:
            seq_series = pd.to_numeric(log_df["Seq No"], errors="coerce")
            seq_series = seq_series.dropna()

            if seq_series.empty:
                next_seq = 1
            else:
                next_seq = int(seq_series.max()) + 1

    # ==========================
    # Format Voucher
    # ==========================
    if biz_type in ["Kontribusi", "Refund", "Alteration", "Retur", "Revise", "Batal", "Cancel"]:
        voucher = f"VOU{year}{month:02d}LST{next_seq:04d}"

    elif biz_type == "Claim":
        voucher = f"VCR{year}{month:02d}LSC{next_seq:04d}"

    return voucher, next_seq


def generate_vin_from_drive_log(log_df, year, month, biz_type):
    if log_df.empty:
        next_seq = 1
    else:
        next_seq = int(log_df["Seq No"].max()) + 1

    if biz_type in ["Kontribusi", "Refund", "Alteration", "Retur", "Revise", "Batal", "Cancel"]:
        voucher = f"VIN{year}{month:02d}LST{next_seq:04d}"

    elif biz_type == "Claim":
        voucher = f"VCL{year}{month:02d}LSC{next_seq:04d}"

    return voucher, next_seq


def create_negative_excel(original_df, original_voucher_id, voucher_id):

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
    
    if "voucher id" in df_negative.columns:
            df_negative["voucher id"] = voucher_id

    if "ref voucher id" in df_negative.columns:
            df_negative["ref voucher id"] = f"Cancel of {original_voucher_id}"

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



def create_cancel_row(original_row, new_voucher, seq_no, year, month, user, reason):
    cancel = original_row.copy()

    cancel["Biz Type"] = "Cancel"
    cancel["Seq No"] = seq_no
    cancel["Voucher No"] = new_voucher
    cancel["OBY"] = year
    cancel["OBM"] = month
    #cancel["ENTRY_TYPE"] = "CANCEL"
    cancel["CANCEL OF VOUCHER"] = original_row["Voucher No"]
    cancel["STATUS"] = "CANCELED"
    cancel["CREATED AT"] = datetime.now(ZoneInfo("Asia/Jakarta")).strftime("%Y-%m-%d %H:%M:%S")
    cancel["CREATED BY"] = user
    cancel["CANCEL REASON"] = reason

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
        "Claim (IDR)",
        "Balance (IDR)"
        ]

    for col in numeric_cols:
        value = original_row.get(col, 0)
        try:
            cancel[col] = -1 * float(value)
        except:
            cancel[col] = 0

    cancel["REMARKS"] = f"Cancel voucher {original_row['Voucher No']}"

    return cancel

def now_wib_naive():
    return datetime.now(ZoneInfo("Asia/Jakarta")).replace(tzinfo=None)

def get_last_seq_no(sheets_service, spreadsheet_id):
    result = sheets_service.spreadsheets().values().get(
        spreadsheetId=spreadsheet_id,
        range="A:Z"
    ).execute()

    values = result.get("values", [])

    if len(values) <= 1:
        return 0

    headers = values[0]
    seq_col = headers.index("Seq No")

    seq_numbers = []
    for row in values[1:]:
        if len(row) > seq_col:
            try:
                seq_numbers.append(int(row[seq_col]))
            except:
                continue

    return max(seq_numbers) if seq_numbers else 0


def generate_pml_id(seq_no, year, month):
    new_seq = seq_no
    pml_id = f"PML{year}{str(month).zfill(2)}LIS{str(new_seq).zfill(4)}"
    return pml_id, new_seq

def split_upload_with_log(
    service,
    sheets_service,
    df,
    split_column,
    period_drive_id,
    pml_folder_id,
    log_pml_drive_id,
    year,
    month,
    biz_type,
    base_info,
    columns_template,
    progress_bar=None,
    status_text=None
):
    results = []

    df.columns = df.columns.str.strip()
    df = df.dropna(subset=[split_column])

    grouped = list(df.groupby(split_column))
    total = len(grouped)

    # 🔥 ambil sequence SEKALI
    current_seq = get_last_seq_no(sheets_service, log_pml_drive_id)

    for i, (key, group) in enumerate(grouped):

        if group.empty:
            continue

        # 🔥 UPDATE UI
        if status_text:
            status_text.text(f"Processing {i+1}/{total} → {split_column} = {key}")

        if progress_bar:
            progress_bar.progress((i + 1) / total)

        # ==========================
        # GENERATE PML (BENAR)
        # ==========================
        pml_id, current_seq = generate_pml_id(
            current_seq,
            year,
            month
        )

        # ==========================
        # HITUNG NILAI
        # ==========================
        total_contribution = group["Reins Total Premium"].sum()
        commission = group["Reins Total Comm"].sum()
        overriding = group["Reins Overriding"].sum() if "Reins Overriding" in group.columns else 0
        total_commission = commission + overriding
        tabarru = group["Reins Tabarru"].sum()
        ujrah = group["Reins Ujrah"].sum()

        # ==========================
        # LOG
        # ==========================
        log_pml = {
            "Seq No": current_seq,
            "Department": base_info["department"],
            "Biz Type": biz_type,
            "PML ID": pml_id,
            "Account With": base_info["account_with"],
            "Cedant Company": base_info["cedant_company"],
            "PIC": base_info["pic"],
            "Curr": base_info["curr"],
            "Total Contribution": total_contribution,
            "Commission": commission,
            "Overriding": overriding,
            "Total Commission": total_commission,
            "Gross Premium Income": total_contribution - total_commission,
            "Tabarru": tabarru,
            "Ujrah": ujrah,
            "Claim": 0,
            "Balance": total_contribution - total_commission,
            "REMARKS": f"Split from {base_info['source_pml']} ({split_column}={key})",
            "STATUS": "POSTED",
            "CREATED AT": now_wib_naive(),
            "CREATED BY": base_info["pic"],
            "Subject Email": base_info["subject_email"],
            "Email Date": base_info["email_date"],
            "CANCELED AT": "-",
            "CANCELED BY": "-",
            "CANCEL OF VOUCHER": "-",
            "CANCEL REASON": "-"
        }

        # ==========================
        # APPEND LOG
        # ==========================
        append_gsheet(
            service=sheets_service,
            spreadsheet_id=log_pml_drive_id,
            row_dict=log_pml
        )

        # ==========================
        # UPLOAD FILE
        # ==========================
        upload_dataframe_to_drive(
            service=service,
            df=group,
            template_columns=columns_template,
            voucher_id=pml_id,
            filename=f"{pml_id}.xlsx",
            folder_id=pml_folder_id,
            file_type="PML"
        )

        results.append({
            "pml_id": pml_id,
            "rows": len(group),
            "split_value": key
        })

    return results