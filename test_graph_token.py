import os
import requests
from dotenv import load_dotenv
from pathlib import Path

load_dotenv(Path(__file__).with_name("credenciales.env"))

TENANT_ID = os.getenv("TENANT_ID")
CLIENT_ID = os.getenv("CLIENT_ID")
CLIENT_SECRET = os.getenv("CLIENT_SECRET")

print("TENANT_ID:", "OK" if TENANT_ID else "VACIO")
print("CLIENT_ID:", "OK" if CLIENT_ID else "VACIO")
print("CLIENT_SECRET:", "OK" if CLIENT_SECRET else "VACIO")

url = f"https://login.microsoftonline.com/{TENANT_ID}/oauth2/v2.0/token"

data = {
    "client_id": CLIENT_ID,
    "client_secret": CLIENT_SECRET,
    "grant_type": "client_credentials",
    "scope": "https://graph.microsoft.com/.default",
}

res = requests.post(url, data=data, timeout=30)
print("STATUS:", res.status_code)
print(res.text[:2000])