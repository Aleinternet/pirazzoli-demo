import os
import io
import requests
from flask import Flask, jsonify
from openpyxl import load_workbook

app = Flask(__name__)

TENANT_ID = os.getenv("TENANT_ID")
CLIENT_ID = os.getenv("CLIENT_ID")
CLIENT_SECRET = os.getenv("CLIENT_SECRET")
SHAREPOINT_HOSTNAME = os.getenv("SHAREPOINT_HOSTNAME")
SHAREPOINT_SITE_PATH = os.getenv("SHAREPOINT_SITE_PATH")
TARGET_DRIVE_NAME = os.getenv("TARGET_DRIVE_NAME", "Documentos")
TARGET_FILE_NAME = os.getenv("TARGET_FILE_NAME", "piloto_datos_minerva.xlsx")


def normalize_text(value):
    if value is None:
        return ""
    return " ".join(str(value).strip().split())


def detect_transport_key(sheet_name: str):
    s = normalize_text(sheet_name).upper()
    if "TERR" in s:
        return "terrestre"
    if "AERE" in s:
        return "aereo"
    if "MAR" in s:
        return "maritimo"
    return None


def detect_operation_type(sheet_name: str):
    s = normalize_text(sheet_name).upper()
    if "IMPO" in s:
        return "importacion"
    if "EXPO" in s:
        return "exportacion"
    return None


def prettify_company_name(file_name: str):
    base = file_name.rsplit(".", 1)[0]
    base = base.replace("_", " ").replace("-", " ").strip()
    lowered = base.lower()
    for prefix in ["piloto datos ", "datos ", "piloto "]:
        if lowered.startswith(prefix):
            base = base[len(prefix):].strip()
            break
    return base.title() if base else "Empresa"


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


def fetch_excel_bytes():
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

    available_drive_names = [d.get("name") for d in drives]

    target_drive = next((d for d in drives if d.get("name") == TARGET_DRIVE_NAME), None)
    if not target_drive:
        raise RuntimeError(
            f"No se encontró el drive '{TARGET_DRIVE_NAME}'. Disponibles: {available_drive_names}"
        )

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

    content = graph_get_bytes(
        f"https://graph.microsoft.com/v1.0/drives/{drive_id}/items/{file_id}/content",
        token
    )

    return content, last_modified


def parse_workbook_to_payload(content: bytes):
    wb = load_workbook(io.BytesIO(content), data_only=False)
    file_name = TARGET_FILE_NAME
    company_name = prettify_company_name(file_name)

    files = [{
        "file": file_name,
        "baseName": file_name.rsplit(".", 1)[0],
        "companyName": company_name,
        "sheets": []
    }]

    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]

        if sheet_name.strip().upper() in {"MATRIZ DATOS", "HOJA1"}:
            continue

        transport_key = detect_transport_key(sheet_name)
        operation_type = detect_operation_type(sheet_name)

        headers = []
        header_meta = []
        visible_columns = []

        for col_idx, cell in enumerate(ws[1], start=1):
            header = normalize_text(cell.value)
            if not header:
                header = f"Columna {col_idx}"
            headers.append(header)
            header_meta.append({
                "header": header,
                "colIndex": col_idx - 1,
                "excelColLetter": cell.column_letter
            })
            visible_columns.append((col_idx, header, cell.column_letter))

        rows = []
        for row_idx in range(2, ws.max_row + 1):
            values = {}
            has_any = False

            for col_idx, header, _letter in visible_columns:
                cell = ws.cell(row=row_idx, column=col_idx)
                raw = cell.value

                if cell.hyperlink and cell.hyperlink.target:
                    text = normalize_text(raw) or header
                    values[header] = {
                        "text": text,
                        "url": cell.hyperlink.target,
                        "tooltip": ""
                    }
                    has_any = True
                else:
                    text = normalize_text(raw)
                    values[header] = text
                    if text not in ("", "-", "—"):
                        has_any = True

            if not has_any:
                continue

            ref_header = next((h for h in headers if "REFERENCIA CLIENTE" in h.upper()), headers[0] if headers else "")
            dispatch_header = next((h for h in headers if "NUMERO DE DESPACHO" in h.upper() or h.upper() == "DESPACHO"), headers[1] if len(headers) > 1 else "")
            base_id = "|".join([
                file_name,
                sheet_name,
                str(values.get(ref_header, "—")),
                str(values.get(dispatch_header, "—")),
                str(row_idx)
            ])

            rows.append({
                "_id": base_id,
                "_sheet": sheet_name,
                "_file": file_name,
                "_company": company_name,
                "_excelRowNumber": row_idx,
                "values": values
            })

        files[0]["sheets"].append({
            "name": sheet_name,
            "headers": headers,
            "headerMeta": header_meta,
            "rows": rows,
            "transportKey": transport_key,
            "operationType": operation_type,
            "companyName": company_name
        })

    return files


@app.route("/api", methods=["GET"])
def api_root():
    try:
        content, last_modified = fetch_excel_bytes()
        files = parse_workbook_to_payload(content)
        return jsonify({
            "ok": True,
            "source": "sharepoint",
            "lastModified": last_modified,
            "files": files
        })
    except Exception as e:
        return jsonify({
            "ok": False,
            "error": str(e)
        }), 500