"""
google_drive_auth.py — Autenticación y acceso a Google Drive con Service Account.
Mismo patrón que Panel_Campari/google_drive_auth.py.
"""

import io
import time
import streamlit as st
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload
import urllib3
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)


def _con_reintentos(func, intentos=3, espera=1.5):
    """Reintenta una llamada a la API de Drive ante errores de red
    transitorios (BrokenPipeError, conexión cortada a mitad de request,
    timeouts). No reintenta errores de permisos/API (esos no se arreglan solos)."""
    ultimo_error = None
    for intento in range(intentos):
        try:
            return func()
        except (BrokenPipeError, ConnectionError, TimeoutError, OSError) as e:
            ultimo_error = e
            if intento < intentos - 1:
                time.sleep(espera * (intento + 1))
    raise ultimo_error


@st.cache_resource
def get_credentials():
    return service_account.Credentials.from_service_account_info(
        st.secrets["gcp_service_account"],
        scopes=["https://www.googleapis.com/auth/drive"]
    )


@st.cache_resource
def get_drive_service():
    creds = get_credentials()
    return build("drive", "v3", credentials=creds)


_XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
_GOOGLE_SHEET_MIME = "application/vnd.google-apps.spreadsheet"


def download_file(file_id):
    """Descarga un archivo de Drive como bytes de xlsx. Soporta tanto un
    .xlsx binario real (get_media) como una Hoja de cálculo de Google nativa
    (export_media) — ver Panel_Campari/CLAUDE.md para el detalle de por qué
    hace falta soportar los dos casos."""
    service = get_drive_service()
    meta = _con_reintentos(lambda: service.files().get(fileId=file_id, fields="mimeType").execute())
    if meta.get("mimeType") == _GOOGLE_SHEET_MIME:
        request = service.files().export_media(fileId=file_id, mimeType=_XLSX_MIME)
    else:
        request = service.files().get_media(fileId=file_id)
    file_stream = io.BytesIO()
    downloader = MediaIoBaseDownload(file_stream, request)
    done = False
    while not done:
        status, done = _con_reintentos(lambda: downloader.next_chunk())
    file_stream.seek(0)
    return file_stream


def find_file_id(folder_id, filename):
    service = get_drive_service()
    q = f"'{folder_id}' in parents and name = '{filename}' and trashed = false"
    res = _con_reintentos(lambda: service.files().list(q=q, spaces="drive", fields="files(id)").execute())
    files = res.get("files", [])
    return files[0]["id"] if files else None
