import os
import requests
from flask import Flask, jsonify, request

app = Flask(__name__)

SUPABASE_URL = os.getenv("SUPABASE_URL")
SUPABASE_SERVICE_ROLE_KEY = os.getenv("SUPABASE_SERVICE_ROLE_KEY")

def supabase_headers():
    return {
        "apikey": SUPABASE_SERVICE_ROLE_KEY,
        "Authorization": f"Bearer {SUPABASE_SERVICE_ROLE_KEY}",
        "Content-Type": "application/json",
        "Prefer": "return=representation"
    }

@app.route("/api/row-history", methods=["GET"])
def api_row_history():
    try:
        file_key = request.args.get("file_key", "").strip()
        sheet_name = request.args.get("sheet_name", "").strip()
        row_identity = request.args.get("row_identity", "").strip()

        if not file_key or not sheet_name or not row_identity:
            return jsonify({
                "ok": False,
                "error": "Faltan parámetros obligatorios."
            }), 400

        if not SUPABASE_URL or not SUPABASE_SERVICE_ROLE_KEY:
            return jsonify({
                "ok": False,
                "error": "Faltan SUPABASE_URL o SUPABASE_SERVICE_ROLE_KEY en Vercel."
            }), 500

        url = (
            f"{SUPABASE_URL}/rest/v1/row_audit_log"
            f"?file_key=eq.{requests.utils.quote(file_key)}"
            f"&sheet_name=eq.{requests.utils.quote(sheet_name)}"
            f"&row_identity=eq.{requests.utils.quote(row_identity)}"
            f"&order=changed_at.desc"
        )

        res = requests.get(url, headers=supabase_headers(), timeout=60)
        res.raise_for_status()

        return jsonify({
            "ok": True,
            "items": res.json() or []
        })

    except Exception as e:
        return jsonify({
            "ok": False,
            "error": str(e)
        }), 500