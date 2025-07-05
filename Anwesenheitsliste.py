from flask import Flask, request, jsonify, send_file, render_template
from flask_cors import CORS
from datetime import datetime
from openpyxl import Workbook
import sqlite3
import os

app = Flask(__name__)
CORS(app)

DATABASE = 'anwesenheit.db'

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
def teilnahme():
    data = request.get_json()
    name = data.get('name', '').strip()
    kurs = data.get('kurs', '').strip()

    if not name or not kurs:
        return jsonify({'message': 'Name und Kurs sind erforderlich'}), 400

    datum = datetime.now().strftime('%Y-%m-%d')
    uhrzeit = datetime.now().strftime('%H:%M:%S')

    conn = sqlite3.connect(DATABASE)
    c = conn.cursor()
    c.execute('INSERT INTO teilnahmen (name, datum, uhrzeit, kurs) VALUES (?, ?, ?, ?)',
              (name, datum, uhrzeit, kurs))
    conn.commit()
    conn.close()

    return jsonify({'message': f'Teilnahme von {name} am {datum} um {uhrzeit} gespeichert (Kurs: {kurs}).'})

@app.route('/api/liste', methods=['GET'])
def liste():
    datum = datetime.now().strftime('%Y-%m-%d')
    conn = sqlite3.connect(DATABASE)
    c = conn.cursor()
    c.execute('SELECT name FROM teilnahmen WHERE datum = ?', (datum,))
    namen = [row[0] for row in c.fetchall()]
    conn.close()
    return jsonify(namen)

@app.route('/download', methods=['GET'])
def download():
    datum = datetime.now().strftime('%Y-%m-%d')
    conn = sqlite3.connect(DATABASE)
    c = conn.cursor()
    c.execute('SELECT name, datum, uhrzeit, kurs FROM teilnahmen WHERE datum = ?', (datum,))
    daten = c.fetchall()
    conn.close()

    wb = Workbook()
    ws = wb.active
    ws.append(['Name', 'Datum', 'Uhrzeit', 'Kurs'])
    for eintrag in daten:
        ws.append(list(eintrag))

    filename = f'anwesenheit_{datum}.xlsx'
    wb.save(filename)
    return send_file(filename, as_attachment=True)

if __name__ == '__main__':
    init_db()
    port = int(os.environ.get('PORT', 10000))
    app.run(host='0.0.0.0', port=port)
