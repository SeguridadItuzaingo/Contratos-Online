# drive_util.py
import os, io
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload, MediaIoBaseUpload

_DRIVE = None

def _get_drive():
    global _DRIVE
    if _DRIVE:
        return _DRIVE
    creds = service_account.Credentials.from_service_account_file(
        os.getenv("GOOGLE_APPLICATION_CREDENTIALS"),
        scopes=["https://www.googleapis.com/auth/drive"]
    )
    _DRIVE = build("drive", "v3", credentials=creds, cache_discovery=False)
    return _DRIVE

def upload_path_to_drive(path: str, filename: str, mimetype: str, folder_id: str = None) -> str:
    """Sube un archivo del disco a la carpeta indicada (o a GOOGLE_DRIVE_FOLDER_ID). Devuelve fileId."""
    folder_id = folder_id or os.getenv("GOOGLE_DRIVE_FOLDER_ID")
    meta = {"name": filename, "parents": [folder_id]}
    media = MediaFileUpload(path, mimetype=mimetype, resumable=True)
    created = _get_drive().files().create(
        body=meta, media_body=media, fields="id", supportsAllDrives=True
    ).execute()
    return created["id"]

def upload_bytes_to_drive(data: bytes, filename: str, mimetype: str, folder_id: str = None) -> str:
    """Sube bytes (BytesIO) como archivo a Drive. Devuelve fileId."""
    folder_id = folder_id or os.getenv("GOOGLE_DRIVE_FOLDER_ID")
    meta = {"name": filename, "parents": [folder_id]}
    media = MediaIoBaseUpload(io.BytesIO(data), mimetype=mimetype, resumable=True)
    created = _get_drive().files().create(
        body=meta, media_body=media, fields="id", supportsAllDrives=True
    ).execute()
    return created["id"]
