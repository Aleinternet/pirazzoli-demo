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
print("SITE_ID:", site_id)

drives = graph_get(
    f"https://graph.microsoft.com/v1.0/sites/{site_id}/drives",
    token
)

print("DRIVES ENCONTRADOS:")
for d in drives.get("value", []):
    print("- NOMBRE:", d.get("name"))
    print("  ID:", d.get("id"))
    print("  URL:", d.get("webUrl"))
    print("  TIPO:", d.get("driveType"))
    print()