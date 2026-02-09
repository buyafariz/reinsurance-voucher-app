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

SCOPES = ["https://www.googleapis.com/auth/drive"]
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


def find_drive_file(service, filename, parent_id):
    query = (
        f"name='{filename}' "
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

    for f in result.get("files", []):
        service.files().delete(
            fileId=f["id"],
            supportsAllDrives=True
        ).execute()



def upload_dataframe_to_drive(service, df, filename, folder_id):
    buffer = BytesIO()
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
        filename="due_date_config.xlsx",
        parent_id=CONFIG_FOLDER_ID
    )

    if not file_id:
        return pd.DataFrame(columns=["Account With", "Days"])

    file_df = load_voucher_excel_from_drive(
        service=service,
        voucher_no="due_date_config",
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
