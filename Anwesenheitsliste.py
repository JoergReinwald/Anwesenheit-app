from flask import Flask, request, jsonify, send_file, render_template, Response
from flask_cors import CORS
import sqlite3
from datetime import datetime
from openpyxl import Workbook
from openpyxl.styles import Alignment
from openpyxl.utils import get_column_letter
import os
import secrets
import traceback

from functools import wraps
from flask import Response

app = Flask(__name__)
CORS(app)
app.secret_key = secrets.token_hex(16)
DATABASE = 'anwesenheit.db'

# 🔒 Zugangsschutz für Excel-Download
def check_auth(username, password):
    return username == 'Admin' and password == 'Admin'

def authenticate():
    return Response(
        'Login erforderlich', 401,
        {'WWW-Authenticate': 'Basic realm="Login erforderlich"'}
    )

def requires_auth(f):
    @wraps(f)
    def decorated(*args, **kwargs):
        auth = request.authorization
        if not auth or not check_auth(auth.username, auth.password):
            return authenticate()
        return f(*args, **kwargs)
    return decorated

# 📥 Excel-Download (mit breiten Spalten & Querformat)
@app.route('/excel')
@requires_auth
def download_liste():
    datum = request.args.get('date', datetime.now().strftime('%Y-%m-%d'))

    conn = sqlite3.connect(DATABASE)
    c = conn.cursor()
    c.execute('SELECT name, datum, uhrzeit, kurs FROM teilnahmen WHERE datum = ?', (datum,))
    daten = c.fetchall()
    conn.close()

    wb = Workbook()
    ws = wb.active
    ws.title = f"Teilnehmer {datum}"

    headers = ["Name", "Datum", "Uhrzeit", "Kurs"]
    ws.append(headers)

    for row in daten:
        ws.append(row)

    # Spaltenbreite automatisch anpassen
    for col in ws.columns:
        max_length = max(len(str(cell.value)) for cell in col)
        col_letter = get_column_letter(col[0].column)
        ws.column_dimensions[col_letter].width = max_length + 5

    # Querformat (optional)
    ws.page_setup.orientation = ws.ORIENTATION_LANDSCAPE

    dateiname = f"anwesenheit_{datum}.xlsx"
    wb.save(dateiname)

    return send_file(
        dateiname,
        as_attachment=True,
        download_name=dateiname,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

