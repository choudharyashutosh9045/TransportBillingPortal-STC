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

def generate_invoice_pdf(row, pdf_path):
    row = {**FIXED_PARTY, **FIXED_STC_BANK, **row}

    W, H = landscape(A4)
    c = canvas.Canvas(pdf_path, pagesize=(W, H))

    LM = RM = TM = BM = 10 * mm

    c.rect(LM, BM, W - LM - RM, H - TM - BM)

    c.setFont("Helvetica-Bold", 14)
    c.drawCentredString(W/2, H - 18, "SOUTH TRANSPORT COMPANY")
    c.setFont("Helvetica", 8)
    c.drawCentredString(W/2, H - 30, "Dehradun Road Near Power Grid Bhagwanpur")
    c.drawCentredString(W/2, H - 42, "Roorkee, Haridwar, U.K. 247661, India")
    c.setFont("Helvetica-Bold", 10)
    c.drawCentredString(W/2, H - 60, "INVOICE")

    logo = os.path.join(BASE_DIR, "logo.png")
    if os.path.exists(logo):
        c.drawImage(logo, LM+5, H-90, width=140, height=60, mask="auto")

    c.rect(LM+5, H-200, 300, 90)
    c.setFont("Helvetica-Bold", 8)
    c.drawString(LM+10, H-185, "To,")
    c.drawString(LM+10, H-200+65, row["PartyName"])
    c.setFont("Helvetica", 8)
    c.drawString(LM+10, H-200+50, row["PartyAddress"])
    c.drawString(LM+10, H-200+35, f'{row["PartyCity"]}, {row["PartyState"]} {row["PartyPincode"]}')
    c.drawString(LM+10, H-200+20, f'GSTIN: {row["PartyGSTIN"]}')
    c.drawString(LM+10, H-215, f'From location: {row.get("FromLocation")}')

    c.rect(W-300, H-200, 260, 90)
    c.setFont("Helvetica-Bold", 8)
    c.drawString(W-290, H-170, f'Freight Bill No: {row.get("FreightBillNo")}')
    c.drawString(W-290, H-190, f'Invoice Date: {format_date(row.get("InvoiceDate"))}')
    c.drawString(W-290, H-210, f'Due Date: {format_date(row.get("DueDate"))}')

    table_y = H - 260
    row_h = 28

    cols = [
        ("S.No", 30), ("Shipment Date", 80), ("LR No", 50),
        ("Destination", 90), ("CN No", 60), ("Truck No", 80),
        ("Invoice No", 140), ("Pkgs", 40), ("Weight", 60),
        ("Arrival", 70), ("Delivery", 70), ("Truck Type", 100),
        ("Freight", 70), ("To Point", 60), ("Unload", 60),
        ("Src Det", 60), ("Dest Det", 60), ("Total", 80)
    ]

    x = LM+5
    c.setFont("Helvetica-Bold", 7)
    for h, w in cols:
        c.rect(x, table_y, w, row_h)
        c.drawCentredString(x+w/2, table_y+10, h)
        x += w

    total = calc_total(row)

    data = [
        "1", format_date(row["ShipmentDate"]), row["LRNo"], row["Destination"],
        row["CNNumber"], row["TruckNo"], row["InvoiceNo"], row["Pkgs"],
        row["WeightKgs"], format_date(row["DateArrival"]),
        format_date(row["DateDelivery"]), row["TruckType"],
        money(row["FreightAmt"]), money(row["ToPointCharges"]),
        money(row["UnloadingCharge"]), money(row["SourceDetention"]),
        money(row["DestinationDetention"]), money(total)
    ]

    x = LM+5
    c.setFont("Helvetica", 7)
    for (h,w), val in zip(cols, data):
        c.rect(x, table_y-row_h, w, row_h)
        c.drawCentredString(x+w/2, table_y-row_h+10, safe_str(val))
        x += w

    words = num2words(int(total), lang="en").title() + " Rupees Only"
    c.drawString(LM+10, table_y-row_h-25, f"Total in words (Rs.): {words}")
    c.drawRightString(W-20, table_y-row_h-25, money(total))

    c.rect(LM+10, BM+30, 320, 130)
    y = BM+140
    for k,v in FIXED_STC_BANK.items():
        c.drawString(LM+20, y, f"{k} : {v}")
        y -= 18

    c.drawString(W-260, BM+120, "For SOUTH TRANSPORT COMPANY")
    c.line(W-260, BM+80, W-60, BM+80)
    c.drawRightString(W-60, BM+60, "(Authorized Signatory)")

    c.showPage()
    c.save()

@app.route("/", methods=["GET","POST"])
def index():
    if request.method == "POST":
        file = request.files["file"]
        path = os.path.join(UPLOAD_FOLDER, file.filename)
        file.save(path)

        df = pd.read_excel(path)
        df.columns = [c.strip() for c in df.columns]

        for h in REQUIRED_HEADERS:
            if h not in df.columns:
                return f"Missing column {h}", 400

        pdfs = []
        for _, r in df.iterrows():
            name = f'{r["FreightBillNo"]}_LR{r["LRNo"]}.pdf'
            pdf_path = os.path.join(OUTPUT_FOLDER, name)
            generate_invoice_pdf(r.to_dict(), pdf_path)
            pdfs.append(pdf_path)

        zip_name = "Invoices.zip"
        zip_path = os.path.join(OUTPUT_FOLDER, zip_name)
        with zipfile.ZipFile(zip_path, "w") as z:
            for p in pdfs:
                z.write(p, os.path.basename(p))

        return send_file(zip_path, as_attachment=True)

    return render_template("index.html")

if __name__ == "__main__":
    app.run(debug=True)
