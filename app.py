import os
import re
import json
import time
import threading
import gspread
from google.oauth2.service_account import Credentials
import openpyxl
from datetime import datetime
from flask import Flask, jsonify

app = Flask(__name__)

SHEET_ID = os.environ.get("SHEET_ID", "1FLznJQ0PBxqnMNRPI_JEgv_QD7o7RoiodZLZmDzGE6Y")
UPLOAD_DIR = os.environ.get("UPLOAD_DIR", os.path.join(os.path.expanduser("~"), "Desktop"))
SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive"
]

os.makedirs(UPLOAD_DIR, exist_ok=True)

def get_credentials():
    creds_json = os.environ.get("GOOGLE_CREDENTIALS")
    if creds_json:
        creds_info = json.loads(creds_json)
        return Credentials.from_service_account_info(creds_info, scopes=SCOPES)
    creds_file = os.environ.get("GOOGLE_CREDENTIALS_FILE", "./credentials.json")
    return Credentials.from_service_account_file(creds_file, scopes=SCOPES)

def get_sheet():
    creds = get_credentials()
    client = gspread.authorize(creds)
    return client.open_by_key(SHEET_ID)

ENCABEZADOS_REQUERIDOS = ["articulo", "descripcion", "ubicacion"]

def normalizar(texto):
    import unicodedata
    texto = str(texto).lower().strip()
    texto = unicodedata.normalize("NFKD", texto)
    return "".join(c for c in texto if not unicodedata.combining(c))

def es_archivo_boss(filepath):
    try:
        wb = openpyxl.load_workbook(filepath, data_only=True, read_only=True)
        ws = wb.active
        rows = []
        for row in ws.iter_rows(max_row=10, values_only=True):
            rows.append([str(c).strip() if c is not None else "" for c in row])

        tiene_titulo = any("maquinas fer" in normalizar(r[0]) for r in rows[:3] if r[0])
        tiene_inventarios = any("inventarios" in normalizar(r[0]) for r in rows[:3] if r[0])
        tiene_fecha = any(__import__("re").search(r'\d{1,2}/\w+/\d{2}', r[0]) for r in rows[:4] if r[0])

        tiene_encabezados = False
        for r in rows:
            vals = [normalizar(v) for v in r if v]
            if any(h in v for v in vals for h in ENCABEZADOS_REQUERIDOS):
                tiene_encabezados = True
                break

        wb.close()
        return tiene_titulo and tiene_inventarios and tiene_fecha and tiene_encabezados
    except Exception:
        pass
    return False

def read_excel(filepath):
    wb = openpyxl.load_workbook(filepath, data_only=True)
    ws = wb.active
    data = []
    for row in ws.iter_rows(values_only=True):
        if any(cell is not None for cell in row):
            data.append([str(cell) if cell is not None else "" for cell in row])
    wb.close()
    return data

def upload_to_sheet(data, sheet, filename):
    try:
        worksheets = sheet.worksheets()
        if worksheets:
            ws = worksheets[0]
            ws.clear()
            ws.update(range_name='A1', values=data)
        else:
            ws = sheet.add_worksheet(title='BOSS', rows=len(data)+10, cols=len(data[0]) if data else 10)
            ws.update(range_name='A1', values=data)
        return True, len(data)
    except Exception as e:
        return False, str(e)

def process_file(filepath, filename, sheet):
    if es_archivo_boss(filepath):
        data = read_excel(filepath)
        if data:
            success, result = upload_to_sheet(data, sheet, filename)
            return {
                "file": filename,
                "uploaded": success,
                "rows": result if success else None,
                "error": result if not success else None,
                "timestamp": datetime.now().isoformat()
            }
    return None

upload_log = []

def monitor_loop():
    processed = []
    while True:
        try:
            sheet = get_sheet()
            files = [f for f in os.listdir(UPLOAD_DIR) if f.endswith('.xlsx') and not f.startswith('~$')]

            for filename in files:
                if filename not in processed:
                    filepath = os.path.join(UPLOAD_DIR, filename)
                    result = process_file(filepath, filename, sheet)
                    if result:
                        upload_log.append(result)
                        processed.append(filename)
                        print(f"[UPLOAD] {filename}: {result}")

            time.sleep(10)
        except Exception as e:
            print(f"[ERROR] {e}")
            time.sleep(30)

@app.route("/")
def index():
    return jsonify({
        "status": "running",
        "service": "BOSS Sheets Sync",
        "upload_dir": UPLOAD_DIR,
        "sheet_id": SHEET_ID,
        "log_count": len(upload_log)
    })

@app.route("/status")
def status():
    return jsonify({
        "status": "active",
        "uploads": upload_log[-20:]
    })

@app.route("/health")
def health():
    return jsonify({"status": "ok"}), 200

if __name__ == "__main__":
    monitor_thread = threading.Thread(target=monitor_loop, daemon=True)
    monitor_thread.start()
    print("[STARTED] Monitor BOSS Sheets Sync")
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 5000)))
