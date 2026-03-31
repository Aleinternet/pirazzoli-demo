import os
import shutil
from pathlib import Path

import requests
from dotenv import load_dotenv
from flask import Flask, jsonify, send_from_directory

BASE_DIR = Path(__file__).resolve().parent
load_dotenv(BASE_DIR / "credenciales.env")

TENANT_ID = os.getenv("TENANT_ID")
CLIENT_ID = os.getenv("CLIENT_ID")
CLIENT_SECRET = os.getenv("CLIENT_SECRET")
SHAREPOINT_HOSTNAME = os.getenv("SHAREPOINT_HOSTNAME")
SHAREPOINT_SITE_PATH = os.getenv("SHAREPOINT_SITE_PATH")

TARGET_DRIVE_NAME = "Documentos"
TARGET_FILE_NAME = "piloto_datos_minerva.xlsx"

EXCELS_DIR = BASE_DIR / "excels"
EXCELS_DIR.mkdir(exist_ok=True)

BACKUP_DIR = BASE_DIR / "excels_backup"
BACKUP_DIR.mkdir(exist_ok=True)

app = Flask(__name__, static_folder=str(BASE_DIR), static_url_path="")

def get_token():
    url = f"https://login.microsoftonline.com/{TENANT_ID}/oauth2/v2.0/token"
    data = {
        "client_id": CLIENT_ID,
        "client_secret": CLIENT_SECRET,
        "grant_type": "client_credentials",
        "scope": "https://graph.microsoft.com/.default",
    }
    res = requests.post(url, data=data, timeout=30)
    res.raise_for_status()
    return res.json()["access_token"]

def graph_get(url, token):
    headers = {"Authorization": f"Bearer {token}"}
    res = requests.get(url, headers=headers, timeout=60)
    res.raise_for_status()
    return res.json()

def graph_get_bytes(url, token):
    headers = {"Authorization": f"Bearer {token}"}
    res = requests.get(url, headers=headers, timeout=120)
    res.raise_for_status()
    return res.content

def download_sharepoint_excel():
    token = get_token()

    site = graph_get(
        f"https://graph.microsoft.com/v1.0/sites/{SHAREPOINT_HOSTNAME}:{SHAREPOINT_SITE_PATH}",
        token
    )
    site_id = site["id"]

    drives = graph_get(
        f"https://graph.microsoft.com/v1.0/sites/{site_id}/drives",
        token
    ).get("value", [])

    target_drive = next((d for d in drives if d.get("name") == TARGET_DRIVE_NAME), None)
    if not target_drive:
        raise RuntimeError(f"No se encontró el drive '{TARGET_DRIVE_NAME}'")

    drive_id = target_drive["id"]

    children = graph_get(
        f"https://graph.microsoft.com/v1.0/drives/{drive_id}/root/children",
        token
    ).get("value", [])

    target_file = next((i for i in children if i.get("name") == TARGET_FILE_NAME), None)
    if not target_file:
        raise RuntimeError(f"No se encontró el archivo '{TARGET_FILE_NAME}' en la raíz del drive.")

    file_id = target_file["id"]
    last_modified = target_file.get("lastModifiedDateTime")
    web_url = target_file.get("webUrl")

    content = graph_get_bytes(
        f"https://graph.microsoft.com/v1.0/drives/{drive_id}/items/{file_id}/content",
        token
    )

    destination = EXCELS_DIR / TARGET_FILE_NAME

    if destination.exists():
        backup_name = f"{destination.stem}_backup{destination.suffix}"
        shutil.copy2(destination, BACKUP_DIR / backup_name)

    destination.write_bytes(content)

    return {
        "site_id": site_id,
        "drive_id": drive_id,
        "file_id": file_id,
        "file_name": TARGET_FILE_NAME,
        "saved_to": str(destination),
        "last_modified": last_modified,
        "web_url": web_url,
    }

@app.route("/")
def serve_index():
    return send_from_directory(BASE_DIR, "index.html")

@app.route("/api/refresh-sharepoint", methods=["POST"])
def refresh_sharepoint():
    try:
        result = download_sharepoint_excel()
        return jsonify({"ok": True, "message": "Excel actualizado desde SharePoint", "result": result})
    except Exception as e:
        return jsonify({"ok": False, "error": str(e)}), 500

@app.route("/api/health", methods=["GET"])
def health():
    return jsonify({"ok": True})

if __name__ == "__main__":
    app.run(host="127.0.0.1", port=5500, debug=True)