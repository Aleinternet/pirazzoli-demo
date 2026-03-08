import os
import json
import openpyxl
from urllib.parse import urlparse, parse_qs, unquote

BASE = os.path.dirname(os.path.dirname(__file__))

excel_path = None
for f in os.listdir(BASE):
    if f.lower().endswith(".xlsx") and not f.startswith("~$"):
        excel_path = os.path.join(BASE, f)
        break

if not excel_path:
    raise Exception("No se encontró un archivo Excel .xlsx en la carpeta del proyecto")

print("Excel encontrado:", excel_path)

wb = openpyxl.load_workbook(excel_path, data_only=False)
ws = wb["Hoja1"]

HEADER_ROW = 4
DATA_START_ROW = 5

headers = {}
for col in range(1, ws.max_column + 1):
    val = ws.cell(HEADER_ROW, col).value
    if val is not None:
        headers[str(val).strip()] = col

def clean_text(value):
    if value is None:
        return ""
    return str(value).strip()

def get_link_info(cell):
    text = clean_text(cell.value)
    url = None

    if cell.hyperlink and cell.hyperlink.target:
        url = cell.hyperlink.target
    elif isinstance(cell.value, str) and cell.value.startswith("http"):
        url = cell.value

    local_url = None
    filename = None

    if url:
        try:
            parsed = urlparse(url)
            qs = parse_qs(parsed.query)

            if "id" in qs and qs["id"]:
                path_value = unquote(qs["id"][0])
                filename = os.path.basename(path_value)
            else:
                filename = os.path.basename(parsed.path)

            if filename:
                local_path = os.path.join(BASE, filename)
                if os.path.exists(local_path):
                    local_url = f"./{filename}"
        except Exception:
            pass

    return {
        "text": text,
        "url": url,
        "localUrl": local_url,
        "filename": filename
    }

def estado_visual(estado_operacion, estado_proceso, instruccion_descarga, observaciones):
    eo = clean_text(estado_operacion).upper()
    ep = clean_text(estado_proceso).upper()
    ed = clean_text(instruccion_descarga).upper()
    obs = clean_text(observaciones).upper()

    if "URGENTE" in eo or "URGENTE" in ep:
        return "Urgente", "yellow"

    if "LIBERADO" in ep or "FINAL" in eo:
        return "Finalizado", "green"

    if obs not in ("", "NAN") and (
        "OBS" in obs or
        "PENDIENTE" in obs or
        "CORREGIR" in obs or
        "OBJET" in obs or
        "ERROR" in obs or
        "VALIDAR" in obs
    ):
        return "Observación", "red"

    if "EMITIDA" in ed or "PENDIENTE" in ed or "ALMACÉN" in ep or "ALMACEN" in ep or "CURSO" in eo:
        return "En curso", "white"

    return "En curso", "white"

despachos = []

for row in range(DATA_START_ROW, ws.max_row + 1):
    referencia_cliente = clean_text(ws.cell(row, headers["Referencia de cliente."]).value) if "Referencia de cliente." in headers else ""
    numero_despacho = clean_text(ws.cell(row, headers["Número de despacho."]).value) if "Número de despacho." in headers else ""

    codigo = referencia_cliente if referencia_cliente else numero_despacho
    if not codigo:
        continue

    instruccion_descarga = clean_text(ws.cell(row, headers["Estado: INSTRUCCIÓN de DESCARGA"]).value) if "Estado: INSTRUCCIÓN de DESCARGA" in headers else ""
    fecha_liberacion = clean_text(ws.cell(row, headers["Fecha de aprobación/liberación."]).value) if "Fecha de aprobación/liberación." in headers else ""
    observaciones = clean_text(ws.cell(row, headers["Observaciones"]).value) if "Observaciones" in headers else ""
    estado_proceso = clean_text(ws.cell(row, headers["Estado actual del proceso."]).value) if "Estado actual del proceso." in headers else ""
    estado_operacion = clean_text(ws.cell(row, headers["ESTADO OPERACION"]).value) if "ESTADO OPERACION" in headers else clean_text(ws.cell(row, headers["ESTADO OPERACION "]).value) if "ESTADO OPERACION " in headers else ""

    estado, clase = estado_visual(
        estado_operacion=estado_operacion,
        estado_proceso=estado_proceso,
        instruccion_descarga=instruccion_descarga,
        observaciones=observaciones
    )

    pdfs = []

    if "Resolucion Sanitaria UYD" in headers:
        info = get_link_info(ws.cell(row, headers["Resolucion Sanitaria UYD"]))
        if info["url"] or info["text"]:
            pdfs.append({
                "tipo": "Resolución Sanitaria",
                "text": info["text"],
                "url": info["url"],
                "localUrl": info["localUrl"],
                "filename": info["filename"]
            })

    if "DIN" in headers:
        info = get_link_info(ws.cell(row, headers["DIN"]))
        if info["url"] or info["text"]:
            pdfs.append({
                "tipo": "DIN",
                "text": info["text"],
                "url": info["url"],
                "localUrl": info["localUrl"],
                "filename": info["filename"]
            })

    comp_key = None
    for key in headers.keys():
        if key.strip().upper() == "COMPROBANTE PAGO TGR":
            comp_key = key
            break

    if comp_key:
        info = get_link_info(ws.cell(row, headers[comp_key]))
        if info["url"] or info["text"]:
            pdfs.append({
                "tipo": "Comprobante Pago TGR",
                "text": info["text"],
                "url": info["url"],
                "localUrl": info["localUrl"],
                "filename": info["filename"]
            })

    despachos.append({
        "codigo": codigo,
        "numeroDespacho": numero_despacho,
        "referenciaCliente": referencia_cliente,
        "estado": estado,
        "clase": clase,
        "estadoOperacion": estado_operacion,
        "estadoProceso": estado_proceso,
        "instruccionDescarga": instruccion_descarga,
        "fechaLiberacion": fecha_liberacion,
        "observaciones": observaciones,
        "pdfs": pdfs
    })

out_path = os.path.join(BASE, "data", "despachos.json")
os.makedirs(os.path.dirname(out_path), exist_ok=True)

with open(out_path, "w", encoding="utf-8") as f:
    json.dump(despachos, f, indent=2, ensure_ascii=False)

print("JSON generado en:", out_path)
print("Despachos encontrados:", len(despachos))
