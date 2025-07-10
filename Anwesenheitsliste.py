from flask import Flask, render_template, request, redirect, send_file, session, url_for
from datetime import datetime, date
import sqlite3
import pandas as pd
import os

app = Flask(__name__)
app.secret_key = 'dein_geheimer_schluessel'

DB_NAME = 'anwesenheit.db'

# Datenbank initialisieren
def init_db():
    with sqlite3.connect(DB_NAME) as conn:
        cursor = conn.cursor()
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS teilnahmen (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                name TEXT NOT NULL,
                kurs TEXT NOT NULL,
                timestamp TEXT NOT NULL
            )
        ''')
        conn.commit()

init_db()

@app.route('/', methods=['GET', 'POST'])
def index():
    meldung = ''
    if request.method == 'POST':
        name = request.form['name'].strip()
        kurs = request.form['kurs'].strip()
        zeitpunkt = datetime.now().strftime('%Y-%m-%d %H:%M:%S')

        with sqlite3.connect(DB_NAME) as conn:
            cursor = conn.cursor()
            cursor.execute("INSERT INTO teilnahmen (name, kurs, timestamp) VALUES (?, ?, ?)", (name, kurs, zeitpunkt))
            conn.commit()
        meldung = f"Teilnahme von {name} am {zeitpunkt} gespeichert (Kurs: {kurs})."

    # Teilnehmer für heute abrufen
    heute = date.today().isoformat()
    with sqlite3.connect(DB_NAME) as conn:
        cursor = conn.cursor()
        cursor.execute("SELECT name, kurs FROM teilnahmen WHERE date(timestamp) = ?", (heute,))
        teilnehmer = cursor.fetchall()

    return render_template('index.html', meldung=meldung, teilnehmer=teilnehmer)

@app.route('/export')
def export():
    if session.get('admin'):
        with sqlite3.connect(DB_NAME) as conn:
            df = pd.read_sql_query("SELECT * FROM teilnahmen", conn)

        # Excel speichern mit Querformat und breiteren Spalten
        output_path = os.path.join('static', 'Teilnehmerliste_breit.xlsx')
        with pd.ExcelWriter(output_path, engine='xlsxwriter') as writer:
            df.to_excel(writer, index=False, sheet_name='Teilnehmer')
            workbook = writer.book
            worksheet = writer.sheets['Teilnehmer']
            worksheet.set_landscape()
            for i, col in enumerate(df.columns):
                worksheet.set_column(i, i, 25)

        return send_file(output_path, as_attachment=True)
    else:
        return redirect(url_for('login'))

@app.route('/login', methods=['GET', 'POST'])
def login():
    fehler = ''
    if request.method == 'POST':
        benutzer = request.form['benutzer']
        passwort = request.form['passwort']
        if benutzer == 'Admin' and passwort == 'Admin':
            session['admin'] = True
            return redirect('/')
        else:
            fehler = 'Zugangsdaten falsch'
    return render_template('login.html', fehler=fehler)

@app.route('/logout')
def logout():
    session.pop('admin', None)
    return redirect('/')

if __name__ == '__main__':
    app.run(debug=True)

