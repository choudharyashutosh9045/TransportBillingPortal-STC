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
DATA_FOLDER = os.path.join(BASE_DIR, "data")

os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)
os.makedirs(DATA_FOLDER, exist_ok=True)

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
    "FreightAmt","ToPointCharges","UnloadingCharge",
    "SourceDetention","DestinationDetention"
]

# ---------------- HELPERS ----------------
def safe_str(v):
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return ""
    return str(v).strip()

def excel_total(row):
    def f(x):
        try:
            return float(str(x).strip())
        except:
            return 0.0
    return (
        f(row.get("FreightAmt")) +
        f(row.get("ToPointCharges")) +
        f(row.get("UnloadingCharge")) +
        f(row.get("SourceDetention")) +
        f(row.get("DestinationDetention"))
    )

# ==========================
# DATABASE
# ==========================
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
                (source_excel, bill_no, lr_no, invoice_date, due_date,
                 destination, total_amount, zip_name)
                VALUES (%s,%s,%s,%s,%s,%s,%s,%s)
            """, (
                safe_str(row.get("FreightBillNo")),
                safe_str(row.get("LRNo")),
                safe_str(row.get("InvoiceDate")),
                safe_str(row.get("DueDate")),
                safe_str(row.get("Destination")),
                excel_total(row),
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

# ---------------- PDF GENERATOR ----------------
def generate_invoice_pdf(row, pdf_path):
    row = {**FIXED_PARTY, **FIXED_STC_BANK, **row}
    total_amt = excel_total(row)

    W, H = landscape(A4)
    c = canvas.Canvas(pdf_path, pagesize=(W, H))
    LM, RM, TM, BM = 10*mm, 10*mm, 10*mm, 10*mm

    c.setLineWidth(1)
    c.rect(LM, BM, W-LM-RM, H-TM-BM)

    c.setFont("Helvetica-Bold", 14)
    c.drawCentredString(W/2, H-TM-8*mm, "SOUTH TRANSPORT COMPANY")

    c.setFont("Helvetica-Bold", 10)
    c.drawCentredString(W/2, H-TM-22*mm, "INVOICE")

    # ---- (PDF DRAWING CODE SAME AS YOURS) ----
    # ❗ Data values everywhere = safe_str(row.get(...))

    # Total in words
    try:
        words = num2words(total_amt, lang="en").title() + " Rupees Only"
    except:
        words = ""

    c.drawString(LM+10*mm, BM+20*mm, words)
    c.drawRightString(W-RM-10*mm, BM+20*mm, safe_str(total_amt))

    c.showPage()
    c.save()

# ---------------- ROUTES ----------------
@app.route("/", methods=["GET","POST"])
def index():
    if request.method == "POST":
        file = request.files.get("file")
        if not file:
            return "No file", 400

        path = os.path.join(UPLOAD_FOLDER, file.filename)
        file.save(path)

        df = pd.read_excel(path, dtype=str).fillna("")
        df.columns = [c.strip() for c in df.columns]

        missing = [h for h in REQUIRED_HEADERS if h not in df.columns]
        if missing:
            return f"Missing columns: {missing}", 400

        generated = []
        for _, r in df.iterrows():
            row = r.to_dict()
            name = f"{row.get('FreightBillNo')}_{row.get('LRNo')}.pdf"
            pdf_path = os.path.join(OUTPUT_FOLDER, name)
            generate_invoice_pdf(row, pdf_path)
            generated.append(pdf_path)

        zip_name = f"Bills_{datetime.now().strftime('%Y%m%d_%H%M%S')}.zip"
        zip_path = os.path.join(OUTPUT_FOLDER, zip_name)
        with zipfile.ZipFile(zip_path, "w") as z:
            for p in generated:
                z.write(p, arcname=os.path.basename(p))

        add_history_entry_db(file.filename, df, zip_name)
        return send_file(zip_path, as_attachment=True)

    return render_template("index.html")

@app.route("/api/history")
def api_history():
    return jsonify(get_history_db())

init_db()

if __name__ == "__main__":
    app.run(debug=True)
