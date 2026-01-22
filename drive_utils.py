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
