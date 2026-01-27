import streamlit as st
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload

SCOPES = ["https://www.googleapis.com/auth/drive"]


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

    voucher_folder_id = get_or_create_folder(
        service,
        folder_name="vouchers",
        parent_id=period_id
    )

    return {
        "period_id": period_id,
        #"voucher_folder_id": voucher_folder_id
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

    voucher_id = get_or_create_folder(
        service,
        folder_name="vouchers",
        parent_id=ceding_id
    )

    return {
        "ceding_id": ceding_id,
        "voucher_id": voucher_id
    }

