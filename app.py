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
import re

app = Flask(__name__)

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
UPLOAD_FOLDER = os.path.join(BASE_DIR, "uploads")
OUTPUT_FOLDER = os.path.join(BASE_DIR, "output")
DATA_FOLDER = os.path.join(BASE_DIR, "data")

os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)
os.makedirs(DATA_FOLDER, exist_ok=True)

app.config["UPLOAD_FOLDER"] = UPLOAD_FOLDER

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
    "ShipmentDate","LRNo","Destination","CNNumber","TruckNo",
    "InvoiceNo","Pkgs","WeightKgs","DateArrival","DateDelivery",
    "TruckType","FreightAmt","ToPointCharges","UnloadingCharge",
    "SourceDetention","DestinationDetention"
]

def safe_str(v):
    if pd.isna(v):
        return ""
    return str(v).strip()

def safe_float(v):
    try:
        if pd.isna(v) or str(v).strip() == "":
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

def clean_filename(v):
    v = safe_str(v)
    v = re.sub(r'[\\/:*?"<>|]', '_', v)
    return v if v else "NA"

def get_db_conn():
    db_url = os.environ.get("DATABASE_URL")
    if not db_url:
        return None
    return psycopg2.connect(db_url, sslmode="require")

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

def add_history_entry_db(source_excel, df, zip_name):
    conn = get_db_conn()
    if not conn:
        return
    with conn.cursor() as cur:
        for _, r in df.iterrows():
            row = r.to_dict()
            cur.execute("""
                INSERT INTO bill_history
                (source_excel,bill_no,lr_no,invoice_date,due_date,destination,total_amount,zip_name)
                VALUES (%s,%s,%s,%s,%s,%s,%s,%s)
            """, (
                source_excel,
                safe_str(row.get("FreightBillNo")),
                safe_str(row.get("LRNo")),
                format_date(row.get("InvoiceDate")),
                format_date(row.get("DueDate")),
                safe_str(row.get("Destination")),
                calc_total(row),
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
            ORDER BY created_at DESC LIMIT %s
        """, (limit,))
        rows = cur.fetchall()
    conn.close()
    for r in rows:
        r["created_at"] = r["created_at"].strftime("%d %b %Y %I:%M %p")
    return rows

def generate_invoice_pdf(row, pdf_path):
    row = {**FIXED_PARTY, **FIXED_STC_BANK, **row}
    os.makedirs(os.path.dirname(pdf_path), exist_ok=True)

    W, H = landscape(A4)
    c = canvas.Canvas(pdf_path, pagesize=(W, H))

    LM = RM = TM = BM = 10 * mm
    c.setLineWidth(1)
    c.rect(LM, BM, W - LM - RM, H - TM - BM)

    c.setFont("Helvetica-Bold", 14)
    c.drawCentredString(W/2, H-18, "SOUTH TRANSPORT COMPANY")
    c.setFont("Helvetica", 8)
    c.drawCentredString(W/2, H-30, "Dehradun Road Near power Grid Bhagwanpur")
    c.drawCentredString(W/2, H-40, "Roorkee,Haridwar, U.K. 247661, India")
    c.setFont("Helvetica-Bold", 10)
    c.drawCentredString(W/2, H-55, "INVOICE")

    logo_path = os.path.join(BASE_DIR, "logo.png")
    if os.path.exists(logo_path):
        img = ImageReader(logo_path)
        c.drawImage(img, LM+5, H-95, 200, 70, mask="auto")

    c.setFont("Helvetica", 7)
    c.drawString(LM+5, H-130, "To,")
    c.setFont("Helvetica-Bold", 8)
    c.drawString(LM+5, H-145, row["PartyName"])
    c.setFont("Helvetica", 7)
    c.drawString(LM+5, H-158, row["PartyAddress"])
    c.drawString(LM+5, H-170, f"{row['PartyCity']} {row['PartyState']} {row['PartyPincode']}")
    c.setFont("Helvetica-Bold", 7)
    c.drawString(LM+5, H-183, f"GSTIN: {row['PartyGSTIN']}")

    c.setFont("Helvetica", 7)
    c.drawRightString(W-50, H-145, f"Freight Bill No: {safe_str(row.get('FreightBillNo'))}")
    c.drawRightString(W-50, H-160, f"Invoice Date: {format_date(row.get('InvoiceDate'))}")
    c.drawRightString(W-50, H-175, f"Due Date: {format_date(row.get('DueDate'))}")

    c.showPage()
    c.save()

@app.route("/", methods=["GET","POST"])
def index():
    if request.method == "POST":
        file = request.files.get("file")
        if not file:
            return "No file", 400

        path = os.path.join(UPLOAD_FOLDER, file.filename)
        file.save(path)

        df = pd.read_excel(path)
        df.columns = [str(c).strip() for c in df.columns]

        missing = [h for h in REQUIRED_HEADERS if h not in df.columns]
        if missing:
            return f"Missing columns {missing}", 400

        generated = []

        for _, r in df.iterrows():
            row = r.to_dict()
            bill = clean_filename(row.get("FreightBillNo"))
            lr = clean_filename(row.get("LRNo"))
            ts = datetime.now().strftime("%H%M%S")
            pdf_name = f"{bill}_{lr}_{ts}.pdf"
            pdf_path = os.path.join(OUTPUT_FOLDER, pdf_name)
            generate_invoice_pdf(row, pdf_path)
            generated.append(pdf_path)

        zip_name = f"BILLS_{datetime.now().strftime('%Y%m%d_%H%M%S')}.zip"
        zip_path = os.path.join(OUTPUT_FOLDER, zip_name)

        with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as z:
            for p in generated:
                z.write(p, os.path.basename(p))

        add_history_entry_db(file.filename, df, zip_name)
        return send_file(zip_path, as_attachment=True)

    return render_template("index.html")

@app.route("/api/history")
def history():
    return jsonify(get_history_db())

init_db()

if __name__ == "__main__":
    app.run(debug=True)
