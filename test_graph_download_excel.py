import os
import requests
from dotenv import load_dotenv
from pathlib import Path

load_dotenv(Path(__file__).with_name("credenciales.env"))

TENANT_ID = os.getenv("TENANT_ID")
CLIENT_ID = os.getenv("CLIENT_ID")
CLIENT_SECRET = os.getenv("CLIENT_SECRET")
SHAREPOINT_HOSTNAME = os.getenv("SHAREPOINT_HOSTNAME")
SHAREPOINT_SITE_PATH = os.getenv("SHAREPOINT_SITE_PATH")

# Ya detectado por ti
TARGET_DRIVE_NAME = "Documentos"
TARGET_FILE_NAME = "piloto_datos_minerva.xlsx"

BASE_DIR = Path(__file__).resolve().parent
OUT_DIR = BASE_DIR / "excels_sharepoint"
OUT_DIR.mkdir(exist_ok=True)

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

def main():
    token = get_token()

    site = graph_get(
        f"https://graph.microsoft.com/v1.0/sites/{SHAREPOINT_HOSTNAME}:{SHAREPOINT_SITE_PATH}",
        token
    )
    site_id = site["id"]
    print("SITE_ID:", site_id)

    drives = graph_get(
        f"https://graph.microsoft.com/v1.0/sites/{site_id}/drives",
        token
    ).get("value", [])

    target_drive = next((d for d in drives if d.get("name") == TARGET_DRIVE_NAME), None)
    if not target_drive:
        raise RuntimeError(f"No se encontró el drive '{TARGET_DRIVE_NAME}'")

    drive_id = target_drive["id"]
    print("DRIVE_ID:", drive_id)

    children = graph_get(
        f"https://graph.microsoft.com/v1.0/drives/{drive_id}/root/children",
        token
    ).get("value", [])

    target_file = next((i for i in children if i.get("name") == TARGET_FILE_NAME), None)
    if not target_file:
        raise RuntimeError(f"No se encontró el archivo '{TARGET_FILE_NAME}' en la raíz del drive.")

    file_id = target_file["id"]
    print("FILE_ID:", file_id)

    content = graph_get_bytes(
        f"https://graph.microsoft.com/v1.0/drives/{drive_id}/items/{file_id}/content",
        token
    )

    out_path = OUT_DIR / TARGET_FILE_NAME
    out_path.write_bytes(content)

    print("DESCARGADO OK:", out_path)

if __name__ == "__main__":
    main()