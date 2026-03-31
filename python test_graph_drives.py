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
    res = requests.get(url, headers=headers, timeout=30)
    res.raise_for_status()
    return res.json()

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

for drive in drives:
    drive_id = drive["id"]
    drive_name = drive["name"]

    print(f"\n===== DRIVE: {drive_name} =====")

    children = graph_get(
        f"https://graph.microsoft.com/v1.0/drives/{drive_id}/root/children",
        token
    ).get("value", [])

    if not children:
        print("Sin elementos en raíz.")
        continue

    for item in children:
        name = item.get("name", "")
        item_id = item.get("id", "")
        is_folder = "folder" in item

        print(
            f"- {name} | "
            f"{'CARPETA' if is_folder else 'ARCHIVO'} | "
            f"ID: {item_id}"
        )

    excels = [
        item for item in children
        if item.get("name", "").lower().endswith((".xlsx", ".xlsm", ".xls"))
    ]

    print("\nEXCELS DETECTADOS:")
    if excels:
        for item in excels:
            print(f"  * {item.get('name')} | ID: {item.get('id')}")
    else:
        print("  Ninguno en la raíz.")