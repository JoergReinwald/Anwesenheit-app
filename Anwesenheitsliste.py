from flask import Flask, request, jsonify, send_file, render_template, Response
import sqlite3
from datetime import datetime
from openpyxl import Workbook
from openpyxl.utils import get_column_letter
from functools import wraps
import os

app = Flask(__name__)
DATABASE = 'anwesenheit.db'

# Zugangsdaten für Admin
ADMIN_USER = "Admin"
ADMIN_PASS = "Admin"

def init_db():
    conn = sqlite3.connect(DATABASE)
    c = conn.cursor()
    c.execute('''
        CREATE TABLE IF NOT EXISTS teilnahmen (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT NOT NULL,
            datum TEXT NOT NULL,
            uhrzeit TEXT NOT NULL,
            kurs TEXT NOT NULL
        )
    ''')
    conn.commit()
    conn.close()

def check_auth(username, password):
    return username == ADMIN_USER and password == ADMIN_PASS

def authenticate():
    return Response('Login erforderlich.', 401, {'WWW-Authenticate': 'Basic realm="Login erforderlich"'})

def requires_auth(f):
    @wraps(f)
    def decorated(*args, **kwargs):
        auth = request.authorization
        if not auth or not check_auth(auth.username, auth.password):
            return authenticate()
        return f(*args, **kwargs)
    return decorated

@app.route('/')
def index():
    return render_template("index.html")

@app.route('/api/teilnahme', methods=['POST'])
def teilnahme():
    data = request.get_json()
    name = data.get('name')
    kurs = data.get('kurs', 'Allgemeiner Lehrgang')
    if not name or not kurs:
        return jsonify({'message': 'Name und Kurs erforderlich'}), 400

    datum = datetime.now().strftime('%Y-%m-%d')
    uhrzeit = datetime.now().strftime('%H:%M:%S')

    conn = sqlite3.connect(DATABASE)
    c = conn.cursor()
    c.execute('INSERT INTO teilnahmen (name, datum, uhrzeit, kurs) VALUES (?, ?, ?, ?)',
              (name, datum, uhrzeit, kurs))
    conn.commit()
    conn.close()

    return jsonify({'message': f'{name} eingetragen ({kurs}) – {datum}, {uhrzeit}'})

@app.route('/api/teilnehmerliste', methods=['GET'])
def teilnehmer_liste():
    kurs = request.args.get('kurs', '')
    datum = datetime.now().strftime('%Y-%m-%d')
    conn = sqlite3.connect(DATABASE)
    c = conn.cursor()
    if kurs:
        c.execute('SELECT name FROM teilnahmen WHERE datum = ? AND kurs = ?', (datum, kurs))
    else:
        c.execute('SELECT name FROM teilnahmen WHERE datum = ?', (datum,))
    eintraege = c.fetchall()
    conn.close()
    return jsonify([x[0] for x in eintraege])

@app.route('/excel')
@requires_auth
def excel_download():
    datum = datetime.now().strftime('%Y-%m-%d')
    conn = sqlite3.connect(DATABASE)
    c = conn.cursor()
    c.execute('SELECT name, datum, uhrzeit, kurs FROM teilnahmen WHERE datum = ?', (datum,))
    daten = c.fetchall()
    conn.close()

    wb = Workbook()
    ws = wb.active
    ws.title = f"Teilnehmer {datum}"
    ws.append(['Name', 'Datum', 'Uhrzeit', 'Kurs'])

    for eintrag in daten:
        ws.append(eintrag)

    for col in range(1, 5):
        ws.column_dimensions[get_column_letter(col)].width = 25

    filename = f"anwesenheit_{datum}.xlsx"
    wb.save(filename)

    return send_file(filename, as_attachment=True)

if __name__ == '__main__':
    init_db()
    app.run(debug=True)

