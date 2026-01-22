import json
import streamlit as st
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload

SCOPES = ["https://www.googleapis.com/auth/drive"]


def get_drive_service():
    service_account_info = dict(st.secrets["gcp_service_account"])

    credentials = service_account.Credentials.from_service_account_info(
        service_account_info,
        scopes=SCOPES,
    )

    return build("drive", "v3", credentials=credentials)


def find_file_in_folder(service, filename, folder_id):
    query = (
        f"name = '{filename}' "
        f"and '{folder_id}' in parents "
        f"and trashed = false"
    )

    result = service.files().list(
        q=query,
        spaces="drive",
        fields="files(id, name)",
    ).execute()

    files = result.get("files", [])
    return files[0]["id"] if files else None


def upload_or_update_drive_file(file_path, filename, folder_id):
    service = get_drive_service()

    media = MediaFileUpload(file_path, resumable=True)

    existing_file_id = find_file_in_folder(
        service, filename, folder_id
    )

    if existing_file_id:
        # 🔁 UPDATE FILE
        updated = service.files().update(
            fileId=existing_file_id,
            media_body=media,
        ).execute()
        return existing_file_id

    else:
        # ➕ CREATE FILE
        file_metadata = {
            "name": filename,
            "parents": [folder_id],
        }

        created = service.files().create(
            body=file_metadata,
            media_body=media,
            fields="id",
        ).execute()
        return created.get("id")
