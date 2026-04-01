import os
import io
import re
import base64
import requests
from datetime import datetime, date
from flask import Flask, jsonify
from openpyxl import load_workbook
from pypdf import PdfReader
import unicodedata
import time

app = Flask(__name__)

TENANT_ID = os.getenv("TENANT_ID")
CLIENT_ID = os.getenv("CLIENT_ID")
CLIENT_SECRET = os.getenv("CLIENT_SECRET")
SHAREPOINT_HOSTNAME = os.getenv("SHAREPOINT_HOSTNAME")
SHAREPOINT_SITE_PATH = os.getenv("SHAREPOINT_SITE_PATH")
TARGET_DRIVE_NAME = os.getenv("TARGET_DRIVE_NAME", "Documentos")
TARGET_FILE_NAME = os.getenv("TARGET_FILE_NAME", "piloto_datos_minerva.xlsx")
CMF_API_KEY = os.getenv("CMF_API_KEY")



API_CACHE = {
    "payload": None,
    "last_build_ts": 0,
    "last_modified": None,
}
CACHE_SECONDS = 300

def build_payload(force_refresh=False):
    now = time.time()

    if (
        not force_refresh
        and API_CACHE["payload"] is not None
        and (now - API_CACHE["last_build_ts"]) < CACHE_SECONDS
    ):
        return API_CACHE["payload"]

    token = get_token()
    content, last_modified = fetch_excel_bytes(token)
    files = parse_workbook_to_payload(content, token)

    payload = {
        "ok": True,
        "source": "sharepoint",
        "lastModified": last_modified,
        "files": files
    }

    API_CACHE["payload"] = payload
    API_CACHE["last_build_ts"] = now
    API_CACHE["last_modified"] = last_modified
    return payload

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



def normalize_header(value):
    text = normalize_text(value)
    text = unicodedata.normalize("NFD", text)
    text = "".join(ch for ch in text if unicodedata.category(ch) != "Mn")
    return text.upper()


def is_blank_tc_value(value):
    if value is None:
        return True
    if isinstance(value, (int, float)):
        return float(value) == 0.0
    text = normalize_text(value)
    return text in {"", "-", "—", "0", "0.0", "0,0", "0,00"}


def parse_excel_like_date(value):
    if value is None:
        return None

    if isinstance(value, datetime):
        return value.date()

    if isinstance(value, date):
        return value

    text = str(value).strip()
    if not text:
        return None

    # toma solo la parte fecha si viene con hora
    match = re.search(r"(\d{1,2}[/-]\d{1,2}[/-]\d{4})", text)
    if match:
        date_part = match.group(1).replace("-", "/")
        try:
            return datetime.strptime(date_part, "%d/%m/%Y").date()
        except Exception:
            pass

    patterns = [
        "%d-%m-%Y %H:%M:%S",
        "%d-%m-%Y %H:%M",
        "%d-%m-%Y",
        "%Y-%m-%d %H:%M:%S",
        "%Y-%m-%d %H:%M",
        "%Y-%m-%d",
        "%d/%m/%Y %H:%M:%S",
        "%d/%m/%Y %H:%M",
        "%d/%m/%Y",
    ]

    for fmt in patterns:
        try:
            return datetime.strptime(text, fmt).date()
        except Exception:
            pass

    return None


def format_clp_dollar(value_float):
    text = f"{value_float:,.2f}"
    text = text.replace(",", "X").replace(".", ",").replace("X", ".")
    return text


def parse_cmf_value(raw):
    if raw is None:
        return None
    text = str(raw).strip()
    if not text:
        return None
    return float(text.replace(".", "").replace(",", "."))


def get_share_token_from_url(share_url: str):
    encoded = base64.b64encode(share_url.encode("utf-8")).decode("utf-8")
    encoded = encoded.rstrip("=").replace("/", "_").replace("+", "-")
    return f"u!{encoded}"


def download_shared_pdf_bytes(share_url: str, token: str):
    share_token = get_share_token_from_url(share_url)

    meta_url = f"https://graph.microsoft.com/v1.0/shares/{share_token}/driveItem"
    headers = {"Authorization": f"Bearer {token}"}
    meta_res = requests.get(meta_url, headers=headers, timeout=60)
    meta_res.raise_for_status()
    meta = meta_res.json()

    download_url = meta.get("@microsoft.graph.downloadUrl")
    if download_url:
        file_res = requests.get(download_url, timeout=120)
        file_res.raise_for_status()
        return file_res.content

    item_id = meta.get("id")
    parent_ref = meta.get("parentReference", {})
    drive_id = parent_ref.get("driveId")
    if not drive_id or not item_id:
        raise RuntimeError("No se pudo resolver el PDF compartido desde SharePoint.")

    content_url = f"https://graph.microsoft.com/v1.0/drives/{drive_id}/items/{item_id}/content"
    file_res = requests.get(content_url, headers=headers, timeout=120)
    file_res.raise_for_status()
    return file_res.content


def extract_fecha_pago_from_pdf_bytes(pdf_bytes: bytes):
    reader = PdfReader(io.BytesIO(pdf_bytes))
    text = "\n".join((page.extract_text() or "") for page in reader.pages)

    match = re.search(r"Fecha\s+Pago\s+(\d{2}-\d{2}-\d{4})(?:\s+\d{2}:\d{2}:\d{2})?", text, re.IGNORECASE)
    if not match:
        return None

    try:
        return datetime.strptime(match.group(1), "%d-%m-%Y").date()
    except Exception:
        return None


def fetch_cmf_dollar_for_date(target_date: date, cmf_cache: dict):
    if target_date is None:
        return None

    cache_key = target_date.isoformat()
    if cache_key in cmf_cache:
        return cmf_cache[cache_key]

    if not CMF_API_KEY:
        cmf_cache[cache_key] = None
        return None

    headers = {"User-Agent": "PirazzoliDemo/1.0"}

    # intento exacto
    exact_url = (
        f"https://api.cmfchile.cl/api-sbifv3/recursos_api/dolar/"
        f"{target_date.year}/{target_date.month:02d}/dias/{target_date.day:02d}"
        f"?apikey={CMF_API_KEY}&formato=json"
    )
    exact_res = requests.get(exact_url, headers=headers, timeout=60)

    if exact_res.ok:
        data = exact_res.json()
        dolares = data.get("Dolares", []) or data.get("Dolar", [])
        if dolares:
            value = parse_cmf_value(dolares[0].get("Valor"))
            if value is not None:
                formatted = format_clp_dollar(value)
                cmf_cache[cache_key] = formatted
                return formatted

    # fallback: fecha anterior disponible
    prev_url = (
        f"https://api.cmfchile.cl/api-sbifv3/recursos_api/dolar/anteriores/"
        f"{target_date.year}/{target_date.month:02d}/dias/{target_date.day:02d}"
        f"?apikey={CMF_API_KEY}&formato=json"
    )
    prev_res = requests.get(prev_url, headers=headers, timeout=60)
    prev_res.raise_for_status()
    data = prev_res.json()
    dolares = data.get("Dolares", []) or data.get("Dolar", [])

    best_date = None
    best_value = None

    for item in dolares:
        fecha_txt = item.get("Fecha")
        valor_txt = item.get("Valor")
        if not fecha_txt:
            continue
        try:
            item_date = datetime.strptime(fecha_txt, "%Y-%m-%d").date()
        except Exception:
            continue

        if item_date <= target_date:
            if best_date is None or item_date > best_date:
                best_date = item_date
                best_value = parse_cmf_value(valor_txt)

    if best_value is None:
        cmf_cache[cache_key] = None
        return None

    formatted = format_clp_dollar(best_value)
    cmf_cache[cache_key] = formatted
    return formatted

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


def fetch_excel_bytes(token=None):
    token = token or get_token()

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

    available_root_items = [i.get("name") for i in children]

    target_file = next((i for i in children if i.get("name") == TARGET_FILE_NAME), None)
    if not target_file:
        raise RuntimeError(
            f"No se encontró el archivo '{TARGET_FILE_NAME}' en la raíz del drive. Elementos en raíz: {available_root_items}"
        )

    file_id = target_file["id"]
    last_modified = target_file.get("lastModifiedDateTime")

    content = graph_get_bytes(
        f"https://graph.microsoft.com/v1.0/drives/{drive_id}/items/{file_id}/content",
        token
    )

    return content, last_modified

def enrich_exchange_rate_columns(rows, headers, token):
    normalized_map = {normalize_header(h): h for h in headers}

    hdr_pago_tgr = next((h for n, h in normalized_map.items() if "COMPROBANTE PAGO TGR" in n), None)
    hdr_fecha_liberacion = next((h for n, h in normalized_map.items() if "FECHA DE LIBERACION" in n), None)
    hdr_tc_legalizacion = next((h for n, h in normalized_map.items() if "T/C DOLAR LEGALIZACION" in n), None)
    hdr_tc_liberacion = next((h for n, h in normalized_map.items() if "T/C LIBERACION CARGA" in n), None)

    if not hdr_tc_legalizacion and not hdr_tc_liberacion:
        return rows

    pdf_date_cache = {}
    cmf_cache = {}

    for row in rows:
        values = row["values"]

        # AK - T/C DOLAR legalizacion
        if hdr_tc_legalizacion and is_blank_tc_value(values.get(hdr_tc_legalizacion)):
            source_date = None

            pago_val = values.get(hdr_pago_tgr) if hdr_pago_tgr else None

            # 1) primero intenta usar la fecha visible en AF
            if isinstance(pago_val, dict):
                source_date = parse_excel_like_date(pago_val.get("text"))
            else:
                source_date = parse_excel_like_date(pago_val)

            # 2) si no hay fecha visible, intenta leer el PDF
            if source_date is None and isinstance(pago_val, dict) and pago_val.get("url"):
                pdf_url = pago_val["url"]

                if pdf_url in pdf_date_cache:
                    source_date = pdf_date_cache[pdf_url]
                else:
                    try:
                        pdf_bytes = download_shared_pdf_bytes(pdf_url, token)
                        source_date = extract_fecha_pago_from_pdf_bytes(pdf_bytes)
                    except Exception:
                        source_date = None

                    pdf_date_cache[pdf_url] = source_date

            if source_date is not None:
                dolar_val = fetch_cmf_dollar_for_date(source_date, cmf_cache)
                if dolar_val is not None:
                    values[hdr_tc_legalizacion] = dolar_val

        # AL - T/C liberacion carga
        if hdr_tc_liberacion and is_blank_tc_value(values.get(hdr_tc_liberacion)):
            liberacion_val = values.get(hdr_fecha_liberacion) if hdr_fecha_liberacion else None
            source_date = parse_excel_like_date(liberacion_val)

            if source_date is not None:
                dolar_val = fetch_cmf_dollar_for_date(source_date, cmf_cache)
                if dolar_val is not None:
                    values[hdr_tc_liberacion] = dolar_val

    return rows

def parse_workbook_to_payload(content: bytes, token: str):
    wb_values = load_workbook(io.BytesIO(content), data_only=True)
    wb_raw = load_workbook(io.BytesIO(content), data_only=False)

    file_name = TARGET_FILE_NAME
    company_name = prettify_company_name(file_name)

    files = [{
        "file": file_name,
        "baseName": file_name.rsplit(".", 1)[0],
        "companyName": company_name,
        "sheets": []
    }]

    for sheet_name in wb_values.sheetnames:
        ws_values = wb_values[sheet_name]
        ws_raw = wb_raw[sheet_name]

        if sheet_name.strip().upper() in {"MATRIZ DATOS", "HOJA1"}:
            continue

        transport_key = detect_transport_key(sheet_name)
        operation_type = detect_operation_type(sheet_name)

        headers = []
        header_meta = []
        visible_columns = []

        for col_idx, cell in enumerate(ws_values[1], start=1):
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
        for row_idx in range(2, ws_values.max_row + 1):
            values = {}
            has_any = False

            for col_idx, header, _letter in visible_columns:
                cell_value = ws_values.cell(row=row_idx, column=col_idx)
                cell_raw = ws_raw.cell(row=row_idx, column=col_idx)

                raw_display = cell_value.value
                text = normalize_text(raw_display)

                hyperlink = None
                if cell_raw.hyperlink and cell_raw.hyperlink.target:
                    hyperlink = cell_raw.hyperlink.target

                if hyperlink:
                    values[header] = {
                        "text": text or header,
                        "url": hyperlink,
                        "tooltip": ""
                    }
                    has_any = True
                else:
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

        rows = enrich_exchange_rate_columns(rows, headers, token)
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
        force_refresh = str(request.args.get("refresh", "")).lower() in {"1", "true", "yes"}
        return jsonify(build_payload(force_refresh=force_refresh))
    except Exception as e:
        return jsonify({
            "ok": False,
            "error": str(e)
        }), 500