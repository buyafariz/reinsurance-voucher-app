import streamlit as st
import io
import pandas as pd
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload
from io import BytesIO
from googleapiclient.http import MediaIoBaseUpload
from googleapiclient.http import MediaIoBaseDownload
import calendar
from datetime import datetime
from googleapiclient.errors import HttpError


SCOPES = [
    "https://www.googleapis.com/auth/drive",
    "https://www.googleapis.com/auth/spreadsheets"
]
CONFIG_FOLDER_ID = st.secrets["config_folder_id"]


def get_drive_service():
    credentials = service_account.Credentials.from_service_account_info(
        dict(st.secrets["gcp_service_account"]),
        scopes=SCOPES
    )
    return build("drive", "v3", credentials=credentials)


def upload_or_update_drive_file(
    file_path: str,
    filename: str,
    folder_id: str,
    file_id: str | None = None
):
    """
    - Jika file_id None  → CREATE file baru
    - Jika file_id ada   → UPDATE file existing
    """

    service = get_drive_service()

    media = MediaFileUpload(file_path, resumable=True)

    # UPDATE
    if file_id:
        updated = service.files().update(
            fileId=file_id,
            media_body=media,
            supportsAllDrives=True
        ).execute()
        return updated["id"]

    # CREATE
    file_metadata = {
        "name": filename,
        "parents": [folder_id]
    }

    created = service.files().create(
        body=file_metadata,
        media_body=media,
        fields="id",
        supportsAllDrives=True
    ).execute()

    return created["id"]


def get_or_create_folder(service, folder_name, parent_id):
    query = (
        f"name='{folder_name}' "
        f"and mimeType='application/vnd.google-apps.folder' "
        f"and '{parent_id}' in parents "
        f"and trashed=false"
    )

    results = service.files().list(
        q=query,
        spaces="drive",
        fields="files(id, name)",
        supportsAllDrives=True,
        includeItemsFromAllDrives=True,
    ).execute()

    files = results.get("files", [])
    if files:
        return files[0]["id"]

    folder_metadata = {
        "name": folder_name,
        "mimeType": "application/vnd.google-apps.folder",
        "parents": [parent_id],
    }

    folder = service.files().create(
        body=folder_metadata,
        fields="id",
        supportsAllDrives=True
    ).execute()

    return folder["id"]


def get_period_drive_folders(year, month, root_folder_id):
    service = get_drive_service()

    period_name = f"{year}_{month:02d}"

    period_id = get_or_create_folder(
        service,
        folder_name=period_name,
        parent_id=root_folder_id
    )

    return {
        "period_id": period_id,
    }


def get_or_create_ceding_folders(
    service,
    period_folder_id: str,
    ceding_name: str
):
    ceding_id = get_or_create_folder(
        service,
        folder_name=ceding_name,
        parent_id=period_folder_id
    )

    # voucher_id = get_or_create_folder(
    #     service,
    #     folder_name="vouchers",
    #     parent_id=ceding_id
    # )

    return {
        "ceding_id": ceding_id
        #"voucher_id": voucher_id
    }

def get_or_create_outward_folders(
    service,
    period_folder_id: str
):
    outward_id = get_or_create_folder(
        service,
        folder_name="OUTWARD",
        parent_id=period_folder_id
    )

    # voucher_id = get_or_create_folder(
    #     service,
    #     folder_name="vouchers",
    #     parent_id=ceding_id
    # )

    return {
        "outward_id": outward_id
        #"voucher_id": voucher_id
    }



def find_drive_file(service, filename, parent_id, mime_type=None):
    query = (
        f"name='{filename}' "
        f"and '{parent_id}' in parents "
        f"and trashed=false"
    )

    if mime_type:
        query += f" and mimeType='{mime_type}'"

    results = service.files().list(
        q=query,
        spaces="drive",
        fields="files(id, name, mimeType)",
        supportsAllDrives=True,
        includeItemsFromAllDrives=True,
    ).execute()

    files = results.get("files", [])
    return files[0]["id"] if files else None


def delete_drive_file(file_id: str):
    service = get_drive_service()

    service.files().delete(
        fileId=file_id,
        supportsAllDrives=True
    ).execute()


def acquire_drive_lock(service, parent_id, lock_name="log_produksi.lock"):
    query = (
        f"name='{lock_name}' "
        f"and '{parent_id}' in parents "
        f"and trashed=false"
    )

    result = service.files().list(
        q=query,
        fields="files(id)",
        supportsAllDrives=True,
        includeItemsFromAllDrives=True,
    ).execute()

    if result.get("files"):
        raise RuntimeError("LOG SEDANG DIGUNAKAN USER LAIN")

    service.files().create(
        body={
            "name": lock_name,
            "parents": [parent_id]
        },
        supportsAllDrives=True
    ).execute()


def release_drive_lock(service, parent_id, lock_name="log_produksi.lock"):
    # 1. VALIDASI AWAL: Cegah ID kosong yang menyebabkan Error 400
    if not parent_id or parent_id == "" or parent_id is None:
        print("Log: parent_id kosong, tidak ada kunci yang perlu dilepas.")
        return

    try:
        # 2. QUERY: Mencari file lock di dalam folder tertentu
        query = (
            f"name='{lock_name}' "
            f"and '{parent_id}' in parents "
            f"and trashed=false"
        )

        result = service.files().list(
            q=query,
            fields="files(id)",
            supportsAllDrives=True,
            includeItemsFromAllDrives=True,
        ).execute()

        files_to_delete = result.get("files", [])
        
        # 3. PENGHAPUSAN: Hapus semua file lock yang ditemukan
        for f in files_to_delete:
            try:
                service.files().delete(
                    fileId=f["id"],
                    supportsAllDrives=True
                ).execute()
            except Exception as e:
                # Jika file sudah dihapus proses lain, jangan hentikan aplikasi
                print(f"Gagal hapus file gembok {f['id']}: {e}")

    except Exception as e:
        # Menangkap error API lainnya agar tidak merusak UI Streamlit
        print(f"Error pada release_drive_lock: {e}")


def upload_dataframe_to_drive(service, df, template_columns, voucher_id, filename, folder_id, type):
    buffer = BytesIO()

    # 1. Buat mapping (case-insensitive)
    mapping_lower_to_template = {col.strip().lower(): col for col in template_columns}

    # 2. Inisialisasi DataFrame hasil dengan kolom sesuai template
    final_df = pd.DataFrame(columns=template_columns)

    # 3. Pindahkan data dari df ke final_df berdasarkan kecocokan nama kolom
    for col_original in df.columns:
        col_lower = col_original.strip().lower()
        if col_lower in mapping_lower_to_template:
            target_col = mapping_lower_to_template[col_lower]
            final_df[target_col] = df[col_original]

    # 4. Pastikan Voucher ID terisi di final_df
    # Cari nama asli kolom Voucher ID di template (misal: "Voucher ID" atau "VOUCHER ID")
    if type == "PML":
        pml_id_col = mapping_lower_to_template.get("pml id")
        if pml_id_col:
            final_df[pml_id_col] = voucher_id

    else:
        voucher_id_col = mapping_lower_to_template.get("voucher id")
        if voucher_id_col:
            final_df[voucher_id_col] = voucher_id

    # 5. Tulis ke Excel dengan Format BOLD pada Header
    with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
        final_df.to_excel(writer, index=False, sheet_name='Sheet1')
        
        # Ambil objek workbook dan worksheet
        workbook  = writer.book
        worksheet = writer.sheets['Sheet1']

        # Definisikan format Bold
        header_format = workbook.add_format({
            'bold': True,
            'text_wrap': False,
            'valign': 'vcenter',
            'border': 1
        })

        # Tulis ulang header dengan format bold
        for col_num, value in enumerate(final_df.columns.values):
            worksheet.write(0, col_num, value, header_format)
            
    buffer.seek(0)

    # 6. Upload ke Google Drive
    media = MediaIoBaseUpload(
        buffer,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        resumable=True
    )

    file_metadata = {
        "name": filename,
        "parents": [folder_id]
    }

    file = service.files().create(
        body=file_metadata,
        media_body=media,
        fields="id",
        supportsAllDrives=True
    ).execute()

    return file.get("id")

def upload_dataframe_to_drive_outward(service, df, original_columns, voucher_id, filename, folder_id, biz_type):
    buffer = BytesIO()

    if biz_type in ["Kontribusi", "Refund", "Alteration", "Retur", "Revise", "Batal", "Cancel"]:
        df["out vouc id"] = voucher_id
    
    elif biz_type == "Claim":
        df["voucher id"] = voucher_id
 
    df.columns = original_columns

    df.to_excel(buffer, index=False)
    buffer.seek(0)

    media = MediaIoBaseUpload(
        buffer,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        resumable=True
    )

    file_metadata = {
        "name": filename,
        "parents": [folder_id]
    }

    file = service.files().create(
        body=file_metadata,
        media_body=media,
        fields="id",
        supportsAllDrives=True
    ).execute()

    return file.get("id")


def load_log_from_drive(service, filename, parent_id):
    file_id = find_drive_file(service, filename, parent_id)

    if not file_id:
        return pd.DataFrame()

    request = service.files().get_media(fileId=file_id)
    fh = io.BytesIO()
    downloader = MediaIoBaseDownload(fh, request)
    done = False

    while not done:
        _, done = downloader.next_chunk()

    fh.seek(0)
    return pd.read_excel(fh)


def load_log_from_gsheet(service, spreadsheet_id):
    sheets_service = build(
        "sheets",
        "v4",
        credentials=service._http.credentials
    )

    result = sheets_service.spreadsheets().values().get(
        spreadsheetId=spreadsheet_id,
        range="Log Produksi"
    ).execute()

    values = result.get("values", [])

    if not values:
        return pd.DataFrame()

    df = pd.DataFrame(values[1:], columns=values[0])
    return df

def update_gsheet(service, spreadsheet_id, df):
    sheets_service = build(
        "sheets",
        "v4",
        credentials=service._http.credentials
    )

    df = df.copy()
    df = df.where(pd.notnull(df), None)

    values = [df.columns.tolist()] + df.values.tolist()

    # 1️⃣ CLEAR DULU
    sheets_service.spreadsheets().values().clear(
        spreadsheetId=spreadsheet_id,
        range="Log Produksi"
    ).execute()

    # 2️⃣ UPDATE ULANG
    try:
        sheets_service.spreadsheets().values().update(
            spreadsheetId=spreadsheet_id,
            range="'Log Produksi'!A1:ZZ",
            valueInputOption="USER_ENTERED",
            body={"values": values}
        ).execute()

    except Exception as e:
        st.error(e)

#Update
import httplib2
from google_auth_httplib2 import AuthorizedHttp
from googleapiclient.discovery import build

def init_sheets_service(creds):
    http = httplib2.Http(timeout=60)
    authed_http = AuthorizedHttp(creds, http=http)

    return build("sheets", "v4", http=authed_http)

@st.cache_data(ttl=600)
def get_headers(_service, spreadsheet_id):
    result = _service.spreadsheets().values().get(
        spreadsheetId=spreadsheet_id,
        range="Sheet1!1:1"
    ).execute()
    return result.get("values", [[]])[0]

import time

def execute_with_retry(request, max_retries=3):
    for i in range(max_retries):
        try:
            return request.execute()
        except Exception as e:
            if i == max_retries - 1:
                raise e
            time.sleep(2 ** i)  # exponential backoff

def append_gsheet(service, spreadsheet_id, row_dict):
    import pandas as pd
    import numpy as np
    from datetime import datetime, date
    from decimal import Decimal

    def clean_value(value):
        if value is None or pd.isna(value):
            return None

        if isinstance(value, (datetime, pd.Timestamp)):
            return value.strftime("%Y-%m-%d %H:%M:%S")

        if isinstance(value, date):
            return value.strftime("%Y-%m-%d")

        if isinstance(value, Decimal):
            return float(value)

        if isinstance(value, (np.integer,)):
            return int(value)

        if isinstance(value, (np.floating,)):
            return float(value)

        if isinstance(value, (np.bool_,)):
            return bool(value)

        if not isinstance(value, (str, int, float, bool)):
            return str(value)

        return value

    headers = get_headers(service, spreadsheet_id)

    cleaned_row = [
        clean_value(row_dict.get(col, None))
        for col in headers
    ]

    request = service.spreadsheets().values().append(
        spreadsheetId=spreadsheet_id,
        range="Sheet1!A1",
        valueInputOption="USER_ENTERED",
        insertDataOption="INSERT_ROWS",
        body={"values": [cleaned_row]}
    )

    return execute_with_retry(request)

# def append_gsheet(service, spreadsheet_id, row_dict):
#     from googleapiclient.discovery import build
#     import pandas as pd
#     import numpy as np
#     from datetime import datetime, date
#     from decimal import Decimal

#     sheets_service = build(
#         "sheets",
#         "v4",
#         credentials=service._http.credentials
#     )

    # 🔹 Ambil header dari sheet
    # header_response = sheets_service.spreadsheets().values().get(
    #     spreadsheetId=spreadsheet_id,
    #     range="Log Produksi!1:1"
    # ).execute()

    # headers = header_response.get("values", [[]])[0]

    # def clean_value(value):

    #     if value is None:
    #         return None

    #     if pd.isna(value):
    #         return None

    #     if isinstance(value, (datetime, pd.Timestamp)):
    #         return value.strftime("%Y-%m-%d %H:%M:%S")

    #     if isinstance(value, date):
    #         return value.strftime("%Y-%m-%d")

    #     if isinstance(value, Decimal):
    #         return float(value)

    #     if isinstance(value, (np.integer,)):
    #         return int(value)

    #     if isinstance(value, (np.floating,)):
    #         return float(value)

    #     if isinstance(value, (np.bool_,)):
    #         return bool(value)

    #     if not isinstance(value, (str, int, float, bool)):
    #         return str(value)

    #     return value

    # # 🔹 Susun row mengikuti urutan header
    # cleaned_row = [
    #     clean_value(row_dict.get(col, None))
    #     for col in headers
    # ]

    # sheets_service.spreadsheets().values().append(
    #     spreadsheetId=spreadsheet_id,
    #     range="Log Produksi!A1",
    #     valueInputOption="USER_ENTERED",
    #     insertDataOption="INSERT_ROWS",
    #     body={"values": [cleaned_row]}
    # ).execute()


template_id = "1FbnbPq8fitRRRCSXeo4WakUr4QQLgAyXsVHbxSeXBhw"
def create_log_gsheet(service, parent_id, filename, columns=None):
    try:
        # 1️⃣ BUAT SPREADSHEET BARU
        file_metadata = {
            "name": filename,
            "mimeType": "application/vnd.google-apps.spreadsheet",
            "parents": [parent_id]
        }

        # Tambahkan supportsAllDrives=True jika folder tujuan ada di Shared Drive
        file = service.files().create(
            body=file_metadata,
            supportsAllDrives=True, 
            fields="id"
        ).execute()

        spreadsheet_id = file["id"]

        # 2️⃣ ISI HEADER
        if columns:
            sheets_service = build(
                "sheets", 
                "v4", 
                credentials=service._http.credentials,
                cache_discovery=False
            )

            # File baru secara default memiliki sheet bernama "Sheet1"
            sheets_service.spreadsheets().values().update(
                spreadsheetId=spreadsheet_id,
                range="Sheet1!A1",
                valueInputOption="RAW",
                body={"values": [columns]}
            ).execute()

        return spreadsheet_id

    except HttpError as error:
        # Ini akan memunculkan pesan error detail di log Streamlit Anda
        print(f"Detail Error Google API: {error.resp.status} - {error.content}")
        raise error


def upload_log_dataframe(service, df, filename, folder_id, file_id=None):
    output = io.BytesIO()
    df.to_excel(output, index=False)
    output.seek(0)

    media = MediaIoBaseUpload(
        output,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        resumable=True
    )

    if file_id:
        service.files().update(
            fileId=file_id,
            media_body=media,
            supportsAllDrives=True
        ).execute()
    else:
        service.files().create(
            body={
                "name": filename,
                "parents": [folder_id]
            },
            media_body=media,
            supportsAllDrives=True
        ).execute()


def load_voucher_excel_from_drive(service, voucher_no, ceding_folder_id):

    filename = f"{voucher_no}.xlsx"

    # 🔎 Cari file di dalam folder ceding
    query = (
        f"name='{filename}' "
        f"and '{ceding_folder_id}' in parents "
        f"and trashed=false"
    )

    results = service.files().list(
        q=query,
        spaces="drive",
        fields="files(id, name)",
        supportsAllDrives=True,
        includeItemsFromAllDrives=True,
    ).execute()

    files = results.get("files", [])

    if not files:
        raise FileNotFoundError(
            f"Voucher {filename} tidak ditemukan di folder ceding"
        )

    file_id = files[0]["id"]

    # 📥 Download file ke memory (tanpa simpan local)
    request = service.files().get_media(fileId=file_id)
    file_stream = io.BytesIO()

    downloader = MediaIoBaseDownload(file_stream, request)

    done = False
    while not done:
        status, done = downloader.next_chunk()

    file_stream.seek(0)

    # 📊 Convert ke DataFrame
    df = pd.read_excel(file_stream)

    return df


def load_due_mapping(service):
    file_id = find_drive_file(
        service=service,
        filename="Due Date Mapping.xlsx",
        parent_id=CONFIG_FOLDER_ID
    )

    if not file_id:
        return pd.DataFrame(columns=["Account With", "Days"])

    file_df = load_voucher_excel_from_drive(
        service=service,
        voucher_no="Due Date Mapping",
        ceding_folder_id=CONFIG_FOLDER_ID
    )

    return file_df



def get_quarter_end(year, month):
    if month in [1,2,3]:
        q_month = 3
    elif month in [4,5,6]:
        q_month = 6
    elif month in [7,8,9]:
        q_month = 9
    else:
        q_month = 12

    last_day = calendar.monthrange(year, q_month)[1]
    return datetime(year, q_month, last_day)


def calculate_due_date(account_with, year, month, service):
    config_df = load_due_mapping(service)

    quarter_end = get_quarter_end(year, month)

    row = config_df[
        config_df["Account With"] == account_with
    ]

    if row.empty:
        return quarter_end  # default tanpa tambahan hari

    days = int(row.iloc[0]["Days"])
    return quarter_end + pd.Timedelta(days=days)


def load_exchange_rate_config(service, config_folder_id):
    file_id = find_drive_file(
        service=service,
        filename="Rate Change.xlsx",
        parent_id=config_folder_id
    )

    if not file_id:
        return pd.DataFrame()

    request = service.files().get_media(
        fileId=file_id,
        supportsAllDrives=True
    )

    file_bytes = request.execute()

    return pd.read_excel(io.BytesIO(file_bytes))


def get_exchange_rate(service, config_folder_id, currency, month):
    df_rate = load_exchange_rate_config(service, config_folder_id)

    if df_rate.empty:
        return 1

    df_rate.columns = df_rate.columns.map(str)

    row = df_rate[df_rate["CcyID"] == currency]

    if row.empty:
        return 1

    month_col = str(month)

    if month_col not in df_rate.columns:
        return 1

    return float(row.iloc[0][month_col])

