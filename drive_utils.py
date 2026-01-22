import json
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


def upload_or_update_drive_file(file_path, filename, folder_id, file_id=None):
    service = get_drive_service()
    media = MediaFileUpload(file_path, resumable=True)

    # 🔁 UPDATE FILE (JIKA SUDAH ADA)
    if file_id:
        updated = service.files().update(
            fileId=file_id,
            media_body=media
        ).execute()
        return updated.get("id")

    # 🆕 CREATE FILE (PERTAMA KALI)
    file_metadata = {
        "name": filename,
        "parents": [folder_id]
    }

    created = service.files().create(
        body=file_metadata,
        media_body=media,
        fields="id"
    ).execute()

    return created.get("id")
