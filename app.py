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

# ======================================================
# APP SETUP
# ======================================================
app = Flask(__name__)

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
UPLOAD_FOLDER = os.path.join(BASE_DIR, "uploads")
OUTPUT_FOLDER = os.path.join(BASE_DIR, "output")

os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)

app.config["UPLOAD_FOLDER"] = UPLOAD_FOLDER

# ======================================================
# FIXED DATA
# ======================================================
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

# ======================================================
# HELPERS
# ======================================================
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
    try:
        d = pd.to_datetime(v, errors="coerce", dayfirst=True)
        return "" if pd.isna(d) else d.strftime("%d %b %Y")
    except:
        return safe_str(v)

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

# ======================================================
# PDF GENERATOR (FORMAT UNCHANGED)
# ======================================================
def generate_invoice_pdf(row, pdf_path):
    row = {**FIXED_PARTY, **FIXED_STC_BANK, **row}

    W, H = landscape(A4)
    c = canvas.Canvas(pdf_path, pagesize=(W, H))

    LM, RM, TM, BM = 10*mm, 10*mm, 10*mm, 10*mm

    c.setLineWidth(1)
    c.rect(LM, BM, W-LM-RM, H-TM-BM)

    c.setFont("Helvetica-Bold", 14)
    c.drawCentredString(W/2, H-TM-8*mm, "SOUTH TRANSPORT COMPANY")
    c.setFont("Helvetica", 8)
    c.drawCentredString(W/2, H-TM-12*mm, "Dehradun Road Near power Grid Bhagwanpur")
    c.drawCentredString(W/2, H-TM-15*mm, "Roorkee, Haridwar, U.K. 247661, India")
    c.setFont("Helvetica-Bold", 10)
    c.drawCentredString(W/2, H-TM-22*mm, "INVOICE")

    total_amt = calc_total(row)

    c.setFont("Helvetica", 8)
    c.drawString(LM+5*mm, BM+20*mm, f"Freight Bill No: {safe_str(row.get('FreightBillNo'))}")
    c.drawString(LM+5*mm, BM+15*mm, f"LR No: {safe_str(row.get('LRNo'))}")
    c.drawString(LM+5*mm, BM+10*mm, f"Total Amount: {money(total_amt)}")

    words = num2words(int(round(total_amt)), lang="en").title() + " Rupees Only"
    c.drawString(LM+5*mm, BM+5*mm, f"Amount in Words: {words}")

    c.showPage()
    c.save()

# ======================================================
# ROUTES
# ======================================================
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

        missing = [h for h in REQUIRED_HEADERS if h not in df.columns]
        if missing:
            return f"Missing columns: {missing}", 400

        pdf_files = []

        for i, r in df.iterrows():
            row = r.to_dict()

            bill = safe_str(row.get("FreightBillNo")) or f"BILL{i+1}"
            lr = safe_str(row.get("LRNo")) or f"LR{i+1}"
            ts = datetime.now().strftime("%Y%m%d_%H%M%S_%f")

            pdf_name = f"{bill}_LR{lr}_{ts}.pdf"
            pdf_path = os.path.join(OUTPUT_FOLDER, pdf_name)

            generate_invoice_pdf(row, pdf_path)
            pdf_files.append(pdf_path)

        zip_name = f"STC_BILLS_{datetime.now().strftime('%Y%m%d_%H%M%S')}.zip"
        zip_path = os.path.join(OUTPUT_FOLDER, zip_name)

        with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as z:
            for p in pdf_files:
                z.write(p, arcname=os.path.basename(p))

        return send_file(zip_path, as_attachment=True)

    return render_template("index.html")

# ======================================================
if __name__ == "__main__":
    app.run(debug=True)
