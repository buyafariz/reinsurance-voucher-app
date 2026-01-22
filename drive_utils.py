from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload
import streamlit as st

SCOPES = ["https://www.googleapis.com/auth/drive"]


def get_drive_service():
    # ⬇️ INI SUDAH DI-PARSE OLEH STREAMLIT
    service_account_info = dict(st.secrets["gcp_service_account"])

    credentials = service_account.Credentials.from_service_account_info(
        service_account_info,
        scopes=SCOPES
    )

    return build("drive", "v3", credentials=credentials)


def upload_to_drive(file_path, filename, folder_id):
    try:
        service = get_drive_service()

        file_metadata = {
            "name": filename,
            "parents": [folder_id]
        }

        media = MediaFileUpload(file_path, resumable=True)

        uploaded = service.files().create(
            body=file_metadata,
            media_body=media,
            fields="id"
        ).execute()

        return uploaded.get("id")

    except Exception as e:
        st.error("❌ Gagal upload ke Google Drive")
        st.exception(e)
        return None
