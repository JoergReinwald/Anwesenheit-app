from flask import Flask, request, jsonify, render_template, send_file
import sqlite3
import os
from datetime import datetime, date
import pandas as pd

app = Flask(__name__)

DB_FILE = "anwesenheit.db"

def init_db():
    with sqlite3.connect(DB_FILE) as conn:
        conn.execute('''CREATE TABLE IF NOT EXISTS anwesenheit (
                            id INTEGER PRIMARY KEY AUTOINCREMENT,
                            name TEXT NOT NULL,
                            kurs TEXT NOT NULL,
                            zeit TIMESTAMP NOT NULL
                        )''')
init_db()

@app.route("/", methods=["GET"])
def index():
    today = date.today().isoformat()
    with sqlite3.connect(DB_FILE) as conn:
        cur = conn.cursor()
        cur.execute("""
            SELECT name, kurs FROM anwesenheit
            WHERE DATE(zeit) = ?
            GROUP BY name
        """, (today,))
        teilnehmer = cur.fetchall()
    return render_template("index.html", teilnehmer=teilnehmer)

@app.route("/api/teilnahme", methods=["POST"])
def teilnahme():
    data = request.get_json()
    name = data.get("name", "").strip()
    kurs = data.get("kurs", "").strip()

    if not name or not kurs:
        return jsonify({"message": "Name und Kurs sind erforderlich!"}), 400

    jetzt = datetime.now()
    with sqlite3.connect(DB_FILE) as conn:
        conn.execute("""
            INSERT INTO anwesenheit (name, kurs, zeit)
            VALUES (?, ?, ?)
        """, (name, kurs, jetzt))
    return jsonify({
        "message": f"Teilnahme von {name} am {jetzt.strftime('%Y-%m-%d um %H:%M:%S')} gespeichert (Kurs: {kurs})."
    })

@app.route("/excel")
def excel_download():
    today = date.today().isoformat()
    with sqlite3.connect(DB_FILE) as conn:
        df = pd.read_sql_query("SELECT * FROM anwesenheit WHERE DATE(zeit) = ?", conn, params=(today,))
    filename = f"anwesenheit_{today}.xlsx"
    df.to_excel(filename, index=False)
    return send_file(filename, as_attachment=True)

