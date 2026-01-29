import re

def safe_filename(name):
    """
    Replace characters not allowed in Windows filenames
    """
    if not name:
        return "FILE"
    name = str(name)
    name = name.replace("/", "_").replace("\\", "_")
    name = re.sub(r'[^A-Za-z0-9._-]', '_', name)
    return name
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

# ---------------- APP INIT ----------------
app = Flask(__name__)

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
UPLOAD_FOLDER = os.path.join(BASE_DIR, "uploads")
OUTPUT_FOLDER = os.path.join(BASE_DIR, "output")

os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)

app.config["UPLOAD_FOLDER"] = UPLOAD_FOLDER

# ---------------- FIXED DETAILS ----------------
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

# ---------------- SAFE HELPERS ----------------
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
        dt = pd.to_datetime(v, dayfirst=True, errors="coerce")
        if pd.isna(dt):
            return ""
        return dt.strftime("%d %b %Y")
    except:
        return ""

def money(v):
    try:
        return f"{float(v):.2f}"
    except:
        return "0.00"

def calc_total(row):
    total = 0.0
    for k in ["FreightAmt","ToPointCharges","UnloadingCharge","SourceDetention","DestinationDetention"]:
        total += safe_float(row.get(k))
    return round(total, 2)

# ---------------- PDF GENERATOR ----------------
def generate_invoice_pdf(row, pdf_path):
    row = {**FIXED_PARTY, **FIXED_STC_BANK, **row}

    W, H = landscape(A4)
    c = canvas.Canvas(pdf_path, pagesize=(W, H))

    LM = 10 * mm
    RM = 10 * mm
    TM = 10 * mm
    BM = 10 * mm

    c.setLineWidth(1)
    c.rect(LM, BM, W - LM - RM, H - TM - BM)

    # HEADER
    c.setFont("Helvetica-Bold", 14)
    c.drawCentredString(W/2, H - TM - 8*mm, "SOUTH TRANSPORT COMPANY")
    c.setFont("Helvetica", 8)
    c.drawCentredString(W/2, H - TM - 14*mm, "Dehradun Road Near Power Grid Bhagwanpur, Roorkee")

    # BILL BOX
    c.setFont("Helvetica-Bold", 9)
    c.drawString(LM + 5*mm, H - TM - 30*mm, f"Freight Bill No: {safe_str(row.get('FreightBillNo'))}")
    c.drawString(LM + 5*mm, H - TM - 36*mm, f"Invoice Date: {format_date(row.get('InvoiceDate'))}")
    c.drawString(LM + 5*mm, H - TM - 42*mm, f"Due Date: {format_date(row.get('DueDate'))}")

    # TABLE
    y = H - 80*mm
    c.setFont("Helvetica-Bold", 7)

    headers = [
        "Shipment Date","LR No","Destination","Truck No",
        "Pkgs","Weight","Freight","To Point","Unloading",
        "Src Det","Dest Det","Total"
    ]

    x = LM + 5*mm
    col_w = 22*mm

    for h in headers:
        c.drawString(x, y, h)
        x += col_w

    c.setFont("Helvetica", 7)
    y -= 8*mm
    x = LM + 5*mm

    total_amt = calc_total(row)

    values = [
        format_date(row.get("ShipmentDate")),
        safe_str(row.get("LRNo")),
        safe_str(row.get("Destination")),
        safe_str(row.get("TruckNo")),
        safe_str(row.get("Pkgs")),
        safe_str(row.get("WeightKgs")),
        money(row.get("FreightAmt")),
        money(row.get("ToPointCharges")),
        money(row.get("UnloadingCharge")),
        money(row.get("SourceDetention")),
        money(row.get("DestinationDetention")),
        money(total_amt),
    ]

    for v in values:
        c.drawString(x, y, v)
        x += col_w

    # TOTAL IN WORDS (CRASH SAFE)
    try:
        amt_words = int(round(total_amt))
    except:
        amt_words = 0

    words = num2words(amt_words, lang="en").title() + " Rupees Only"

    c.setFont("Helvetica-Bold", 8)
    c.drawString(LM + 5*mm, BM + 20*mm, f"Total (in words): {words}")

    c.showPage()
    c.save()

# ---------------- ROUTES ----------------
@app.route("/", methods=["GET","POST"])
def index():
    if request.method == "POST":
        file = request.files.get("file")
        if not file:
            return "No file uploaded", 400

        path = os.path.join(UPLOAD_FOLDER, file.filename)
        file.save(path)

        df = pd.read_excel(path)
        df.columns = [str(c).strip() for c in df.columns]

        missing = [h for h in REQUIRED_HEADERS if h not in df.columns]
        if missing:
            return f"Missing columns: {missing}", 400

        pdfs = []

        for _, r in df.iterrows():
            row = r.to_dict()
            bill_raw = row.get("FreightBillNo", "BILL")
            lr_raw = row.get("LRNo", "LR")
            
            bill = safe_filename(bill_raw)
            lr = safe_filename(lr_raw)
            pdf_name = f"{bill}_{lr}_{ts}.pdf"
            pdf_path = os.path.join(OUTPUT_FOLDER, pdf_name)

            generate_invoice_pdf(row, pdf_path)
            pdfs.append(pdf_path)

        zip_name = f"INVOICES_{datetime.now().strftime('%Y%m%d_%H%M%S')}.zip"
        zip_path = os.path.join(OUTPUT_FOLDER, zip_name)

        with zipfile.ZipFile(zip_path,"w",zipfile.ZIP_DEFLATED) as z:
            for p in pdfs:
                z.write(p, os.path.basename(p))

        return send_file(zip_path, as_attachment=True)

    return render_template("index.html")

# ---------------- RUN ----------------
if __name__ == "__main__":
    print("✅ APP RUNNING - PDF SAFE MODE ENABLED")
    app.run(debug=True)
