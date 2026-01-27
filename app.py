from flask import Flask, render_template, request, send_file, jsonify
import os
import pandas as pd
from datetime import datetime
from num2words import num2words
import zipfile

from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.units import mm
from reportlab.lib.utils import ImageReader

import psycopg2
from psycopg2.extras import RealDictCursor


app = Flask(__name__)

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
UPLOAD_FOLDER = os.path.join(BASE_DIR, "uploads")
OUTPUT_FOLDER = os.path.join(BASE_DIR, "output")
TEMPLATE_FOLDER = os.path.join(BASE_DIR, "templates")
DATA_FOLDER = os.path.join(BASE_DIR, "data")

os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)
os.makedirs(DATA_FOLDER, exist_ok=True)

app.config["UPLOAD_FOLDER"] = UPLOAD_FOLDER


# ==========================
# ✅ FIXED DETAILS (SAME ALWAYS)
# ==========================
FIXED_PARTY = {
    "PartyName": "Grivaa Springs Private Ltd.",
    "PartyAddress": "Khasra no 135, Tansipur, Roorkee",
    "PartyCity": "Roorkee",
    "PartyState": "Uttarakhand",
    "PartyPincode": "247656",
    "PartyGSTIN": "05AAICG4793P1ZV",
}

FIXED_STC_BANK = {
    "PANNo": "BSSPG9414K",
    "STCGSTIN": "05BSSPG9414K1ZA",
    "STCStateCode": "5",
    "AccountName": "South Transport Company",
    "AccountNo": "364205500142",
    "IFSCode": "ICIC0003642",
}

# Excel headers required
REQUIRED_HEADERS = [
    "FreightBillNo","InvoiceDate","DueDate","FromLocation",
    "ShipmentDate","LRNo","Destination","CNNumber","TruckNo","InvoiceNo",
    "Pkgs","WeightKgs","DateArrival","DateDelivery","TruckType",
    "FreightAmt","ToPointCharges","UnloadingCharge","SourceDetention","DestinationDetention"
]


# ---------------- HELPERS ----------------
def safe_str(v):
    if pd.isna(v):
        return ""
    return str(v).strip()

def safe_float(v):
    try:
        if pd.isna(v):
            return 0.0
        if str(v).strip() == "":
            return 0.0
        return float(v)
    except:
        return 0.0

def format_date(v):
    s = safe_str(v)
    if not s:
        return ""
    try:
        dt = pd.to_datetime(s, dayfirst=True, errors="coerce")
        if pd.isna(dt):
            return s
        return dt.strftime("%d %b %Y")
    except:
        return s

def money(v):
    return f"{safe_float(v):.2f}"

def calc_total(row):
    return (
        safe_float(row.get("FreightAmt")) +
        safe_float(row.get("ToPointCharges")) +
        safe_float(row.get("UnloadingCharge")) +
        safe_float(row.get("SourceDetention")) +
        safe_float(row.get("DestinationDetention"))
    )


# ==========================
# ✅ DATABASE (SAFE FIXED)
# ==========================
def get_db_conn():
    DATABASE_URL = os.environ.get("DATABASE_URL")
    if not DATABASE_URL:
        print("⚠️ DATABASE_URL not set. DB disabled.")
        return None

    if DATABASE_URL.startswith("postgresql://"):
        DATABASE_URL = DATABASE_URL.replace("postgresql://", "postgres://", 1)

    try:
        conn = psycopg2.connect(DATABASE_URL, sslmode="require")
        return conn
    except Exception as e:
        print("❌ DB CONNECT ERROR:", e)
        return None


def init_db():
    conn = get_db_conn()
    if not conn:
        return

    with conn.cursor() as cur:
        cur.execute("""
            CREATE TABLE IF NOT EXISTS bill_history (
                id SERIAL PRIMARY KEY,
                created_at TIMESTAMP DEFAULT NOW(),
                source_excel VARCHAR(255),
                bill_no VARCHAR(100),
                lr_no VARCHAR(100),
                invoice_date VARCHAR(50),
                due_date VARCHAR(50),
                destination VARCHAR(200),
                total_amount NUMERIC(12,2),
                zip_name VARCHAR(255)
            );
        """)
        conn.commit()

    conn.close()
    print("✅ DB ready")


def add_history_entry_db(source_excel, df, zip_name):
    conn = get_db_conn()
    if not conn:
        return

    with conn.cursor() as cur:
        for _, r in df.iterrows():
            row = r.to_dict()
            cur.execute("""
                INSERT INTO bill_history
                (source_excel, bill_no, lr_no, invoice_date, due_date, destination, total_amount, zip_name)
                VALUES (%s,%s,%s,%s,%s,%s,%s,%s)
            """, (
                source_excel,
                safe_str(row.get("FreightBillNo")),
                safe_str(row.get("LRNo")),
                format_date(row.get("InvoiceDate")),
                format_date(row.get("DueDate")),
                safe_str(row.get("Destination")),
                float(calc_total(row)),
                zip_name
            ))

        conn.commit()

    conn.close()


def get_history_db(limit=10):
    conn = get_db_conn()
    if not conn:
        return []

    with conn.cursor(cursor_factory=RealDictCursor) as cur:
        cur.execute("""
            SELECT * FROM bill_history
            ORDER BY created_at DESC
            LIMIT %s
        """, (limit,))
        rows = cur.fetchall()

    conn.close()

    for r in rows:
        if r.get("created_at"):
            r["created_at"] = r["created_at"].strftime("%d %b %Y %I:%M %p")

    return rows


# ---------------- PDF GENERATOR ----------------
# ❗❗❗ बिल्कुल वही है – कुछ भी change नहीं किया ❗❗❗
# (PDF CODE AS IT IS)
