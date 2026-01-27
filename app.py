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

os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)

app.config["UPLOAD_FOLDER"] = UPLOAD_FOLDER


# ==========================
# FIXED DETAILS
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
        return float(v)
    except:
        return 0.0

def format_date(v):
    try:
        return pd.to_datetime(v, dayfirst=True).strftime("%d %b %Y")
    except:
        return ""

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
# DATABASE (SAFE)
# ==========================
def get_db_conn():
    DATABASE_URL = os.environ.get("DATABASE_URL")
    if not DATABASE_URL:
        return None

    if DATABASE_URL.startswith("postgresql://"):
        DATABASE_URL = DATABASE_URL.replace("postgresql://", "postgres://", 1)

    return psycopg2.connect(DATABASE_URL)


def init_db():
    conn = get_db_conn()
    if not conn:
        print("⚠️ DATABASE_URL not set. DB disabled.")
        return

    with conn.cursor() as cur:
        cur.execute("""
            CREATE TABLE IF NOT EXISTS bill_history (
                id SERIAL PRIMARY KEY,
                created_at TIMESTAMP DEFAULT NOW(),
                source_excel VARCHAR(255),
                bill_no VARCHAR(100),
                lr_no VARCHAR(100),
                destination VARCHAR(200),
                total_amount NUMERIC(12,2)
            );
        """)
        conn.commit()

    conn.close()
    print("✅ DB ready")


# ---------------- PDF GENERATOR ----------------
def generate_invoice_pdf(row: dict, pdf_path: str):
    row = {**FIXED_PARTY, **FIXED_STC_BANK, **row}

    W, H = landscape(A4)
    c = canvas.Canvas(pdf_path, pagesize=(W, H))

    c.setFont("Helvetica-Bold", 14)
    c.drawCentredString(W/2, H-40, "SOUTH TRANSPORT COMPANY")

    c.setFont("Helvetica", 9)
    c.drawString(40, H-80, f"Freight Bill No: {safe_str(row.get('FreightBillNo'))}")
    c.drawString(40, H-100, f"LR No: {safe_str(row.get('LRNo'))}")
    c.drawString(40, H-120, f"Destination: {safe_str(row.get('Destination'))}")

    total = calc_total(row)
    c.drawString(40, H-160, f"Total Amount: Rs {money(total)}")

    c.showPage()
    c.save()


# ---------------- ROUTES ----------------
@app.route("/", methods=["GET", "POST"])
def index():
    if request.method == "POST":
        file = request.files.get("file")
        if not file:
            return "No file uploaded", 400

        path = os.path.join(UPLOAD_FOLDER, file.filename)
        file.save(path)

        df = pd.read_excel(path)
        df.columns = [c.strip() for c in df.columns]

        for h in REQUIRED_HEADERS:
            if h not in df.columns:
                return f"Missing column: {h}", 400

        pdfs = []
        for _, r in df.iterrows():
            row = r.to_dict()
            name = f"{safe_str(row.get('FreightBillNo'))}_{safe_str(row.get('LRNo'))}.pdf"
            pdf_path = os.path.join(OUTPUT_FOLDER, name)
            generate_invoice_pdf(row, pdf_path)
            pdfs.append(pdf_path)

        zip_path = os.path.join(OUTPUT_FOLDER, "Bills.zip")
        with zipfile.ZipFile(zip_path, "w") as z:
            for p in pdfs:
                z.write(p, arcname=os.path.basename(p))

        return send_file(zip_path, as_attachment=True)

    return render_template("index.html")


# ---------------- START ----------------
init_db()

if __name__ == "__main__":
    app.run(debug=True)
