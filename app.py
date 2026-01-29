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

app = Flask(__name__)

# ---------------- PATHS ----------------
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
UPLOAD_FOLDER = os.path.join(BASE_DIR, "uploads")
OUTPUT_FOLDER = os.path.join(BASE_DIR, "output")

os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)

app.config["UPLOAD_FOLDER"] = UPLOAD_FOLDER

# ---------------- FIXED DATA ----------------
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

# ---------------- HELPERS ----------------
def safe_str(v):
    return "" if pd.isna(v) else str(v).strip()

def safe_float(v):
    try:
        return float(v)
    except:
        return 0.0

def format_date(v):
    try:
        return pd.to_datetime(v, dayfirst=True).strftime("%d %b %Y")
    except:
        return safe_str(v)

def money(v):
    return f"{safe_float(v):.2f}"

def calc_total(r):
    return (
        safe_float(r.get("FreightAmt")) +
        safe_float(r.get("ToPointCharges")) +
        safe_float(r.get("UnloadingCharge")) +
        safe_float(r.get("SourceDetention")) +
        safe_float(r.get("DestinationDetention"))
    )

def clean_filename(name):
    for ch in ['/', '\\', ':', '*', '?', '"', '<', '>', '|']:
        name = name.replace(ch, "_")
    return name

# ---------------- PDF ----------------
def generate_invoice_pdf(row, pdf_path):
    row = {**FIXED_PARTY, **FIXED_STC_BANK, **row}
    W, H = landscape(A4)
    c = canvas.Canvas(pdf_path, pagesize=(W, H))

    LM, RM, TM, BM = 10*mm, 10*mm, 10*mm, 10*mm

    c.rect(LM, BM, W-LM-RM, H-TM-BM)

    c.setFont("Helvetica-Bold", 14)
    c.drawCentredString(W/2, H-20, "SOUTH TRANSPORT COMPANY")

    logo_path = os.path.join(BASE_DIR, "logo.png")
    if os.path.exists(logo_path):
        c.drawImage(logo_path, LM+5, H-90, 200, 60, mask="auto")

    total_amt = calc_total(row)

    c.setFont("Helvetica", 9)
    c.drawString(LM+10, H-110, f"Freight Bill No : {row.get('FreightBillNo')}")
    c.drawString(LM+10, H-125, f"LR No : {row.get('LRNo')}")
    c.drawString(LM+10, H-140, f"Destination : {row.get('Destination')}")
    c.drawString(LM+10, H-155, f"Total Amount : Rs. {money(total_amt)}")

    words = num2words(int(round(total_amt)), lang="en").title() + " Rupees Only"
    c.drawString(LM+10, H-170, f"In Words : {words}")

    c.showPage()
    c.save()

# ---------------- ROUTE ----------------
@app.route("/", methods=["GET","POST"])
def index():
    if request.method == "POST":
        file = request.files.get("file")
        if not file:
            return "No file", 400

        path = os.path.join(UPLOAD_FOLDER, file.filename)
        file.save(path)

        df = pd.read_excel(path)
        df.columns = [c.strip() for c in df.columns]

        missing = [h for h in REQUIRED_HEADERS if h not in df.columns]
        if missing:
            return f"Missing columns {missing}", 400

        pdfs = []

        for _, r in df.iterrows():
            row = r.to_dict()

            bill = clean_filename(safe_str(row.get("FreightBillNo", "BILL")))
            lr = clean_filename(safe_str(row.get("LRNo", "LR")))
            ts = datetime.now().strftime("%H%M%S")

            pdf_name = f"{bill}_LR{lr}_{ts}.pdf"
            pdf_path = os.path.join(OUTPUT_FOLDER, pdf_name)

            generate_invoice_pdf(row, pdf_path)
            pdfs.append(pdf_path)

        zip_name = f"BILLS_{datetime.now().strftime('%Y%m%d_%H%M%S')}.zip"
        zip_path = os.path.join(OUTPUT_FOLDER, zip_name)

        with zipfile.ZipFile(zip_path, "w") as z:
            for p in pdfs:
                z.write(p, os.path.basename(p))

        return send_file(zip_path, as_attachment=True)

    return render_template("index.html")

# ---------------- RUN ----------------
if __name__ == "__main__":
    print("✅ APP RUNNING — PDF FORMAT SAFE")
    app.run(debug=True)
