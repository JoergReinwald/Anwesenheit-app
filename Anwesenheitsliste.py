from flask import Flask, request, jsonify, send_file, render_template
from flask_cors import CORS
import sqlite3
from datetime import datetime
from openpyxl import Workbook
import os
import secrets
import traceback

app = Flask(__name__)
CORS(app)
app.secret_key = secrets.token_hex(16)
DATABASE = 'anwesenheit.db'

ZULASSIGE_TEILNEHMER = [
    "Anna Müller", "Tom Schneider", "Lisa Becker", "Jonas Schmidt", "Fatima Yilmaz",
    "Paul Meier", "Mia Weber", "Lukas Hoffmann", "Sophie Braun", "Maximilian König"
]

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

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/api/teilnahme', methods=['POST'])
def teilnahme_bestaetigen():
    try:
        if not request.is_json:
            return jsonify({'message': 'Anfrage ist kein JSON!'}), 400

        data = request.get_json()
        name = data.get('name')
        kurs = data.get('kurs', 'Allgemeiner Lehrgang')

        if not name or not kurs:
            return jsonify({'message': 'Name oder Kurs fehlen'}), 400

        if name not in ZULASSIGE_TEILNEHMER:
            return jsonify({'message': f'Ungültiger Teilnehmername: {name}'}), 403

        datum = datetime.now().strftime('%Y-%m-%d')
        uhrzeit = datetime.now().strftime('%H:%M:%S')

        conn = sqlite3.connect(DATABASE)
        c = conn.cursor()
        c.execute('INSERT INTO teilnahmen (name, datum, uhrzeit, kurs) VALUES (?, ?, ?, ?)',
                  (name, datum, uhrzeit, kurs))
        conn.commit()
        conn.close()

        return jsonify({'message': f'Teilnahme von {name} am {datum} um {uhrzeit} gespeichert (Kurs: {kurs}).'})

    except Exception as e:
        traceback.print_exc()
        return jsonify({'message': f'Fehler im Server: {str(e)}'}), 500

@app.route('/download', methods=['GET'])
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
    ws.append(["Name", "Datum", "Uhrzeit", "Kurs"])
    for eintrag in daten:
        ws.append(list(eintrag))

    dateiname = f"anwesenheit_{datum}.xlsx"
    wb.save(dateiname)

    return send_file(
        dateiname,
        as_attachment=True,
        download_name=dateiname,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

@app.route('/api/teilnehmerliste', methods=['GET'])
def get_teilnehmerliste():
    datum = datetime.now().strftime('%Y-%m-%d')
    kurs = request.args.get('kurs', 'Allgemeiner Lehrgang')
    conn = sqlite3.connect(DATABASE)
    c = conn.cursor()
    c.execute('SELECT name FROM teilnahmen WHERE datum = ? AND kurs = ?', (datum, kurs))
    daten = c.fetchall()
    conn.close()
    return jsonify([eintrag[0] for eintrag in daten])

if __name__ == '__main__':
    init_db()
    port = int(os.environ.get('PORT', 5000))
    app.run(debug=False, host='0.0.0.0', port=port)
