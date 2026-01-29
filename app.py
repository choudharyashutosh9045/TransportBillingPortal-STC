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

FIXED_PARTY = {
    "PartyName": "Grivaa Springs Private Ltd.",
    "PartyAddress": "Khasra no 135, Tansipur, Roorkee",
    "PartyCity": "Roorkee",
    "PartyState": "Uttarakhand",
    "PartyPincode": "247656",
    "PartyGSTIN": "05AAICG4793P1ZV"
}

FIXED_STC_BANK = {
    "PANNo": "BSSPG9414K",
    "STCGSTIN": "05BSSPG9414K1ZA",
    "STCStateCode": "5",
    "AccountName": "South Transport Company",
    "AccountNo": "364205500142",
    "IFSCode": "ICIC0003642"
}

REQUIRED_HEADERS = [
    "FreightBillNo","InvoiceDate","DueDate","FromLocation",
    "ShipmentDate","LRNo","Destination","CNNumber","TruckNo","InvoiceNo",
    "Pkgs","WeightKgs","DateArrival","DateDelivery","TruckType",
    "FreightAmt","ToPointCharges","UnloadingCharge","SourceDetention","DestinationDetention"
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
    try:
        dt = pd.to_datetime(v, dayfirst=True, errors="coerce")
        if pd.isna(dt):
            return safe_str(v)
        return dt.strftime("%d %b %Y")
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

    c.setFont("Helvetica-Bold", 14)
    c.drawCentredString(W / 2, H - TM - 8 * mm, "SOUTH TRANSPORT COMPANY")
    c.setFont("Helvetica", 8)
    c.drawCentredString(W / 2, H - TM - 12 * mm, "Dehradun Road Near power Grid Bhagwanpur")
    c.drawCentredString(W / 2, H - TM - 15 * mm, "Roorkee, Haridwar, U.K. 247661, India")
    c.setFont("Helvetica-Bold", 10)
    c.drawCentredString(W / 2, H - TM - 22 * mm, "INVOICE")

    logo_path = os.path.join(BASE_DIR, "logo.png")
    if os.path.exists(logo_path):
        img = ImageReader(logo_path)
        c.drawImage(img, LM + 5 * mm, H - TM - 40 * mm, 75 * mm, 38 * mm, preserveAspectRatio=True)

    left_x = LM + 2 * mm
    left_y = H - TM - 62 * mm
    left_w = 110 * mm
    left_h = 28 * mm
    c.rect(left_x, left_y, left_w, left_h)

    c.setFont("Helvetica-Bold", 8)
    c.drawString(left_x + 2 * mm, left_y + left_h - 6 * mm, "To,")
    c.drawString(left_x + 2 * mm, left_y + left_h - 11 * mm, row["PartyName"])
    c.setFont("Helvetica", 7.5)
    c.drawString(left_x + 2 * mm, left_y + left_h - 15 * mm, row["PartyAddress"])
    c.drawString(left_x + 2 * mm, left_y + left_h - 19 * mm, f'{row["PartyCity"]}, {row["PartyState"]} {row["PartyPincode"]}')
    c.setFont("Helvetica-Bold", 7.5)
    c.drawString(left_x + 2 * mm, left_y + left_h - 23 * mm, f'GSTIN: {row["PartyGSTIN"]}')

    c.setFont("Helvetica", 7.5)
    c.drawString(left_x + 2 * mm, left_y - 5 * mm, f'From location: {safe_str(row.get("FromLocation"))}')

    right_w = 85 * mm
    right_h = 28 * mm
    right_x = W - RM - right_w - 2 * mm
    right_y = left_y
    c.rect(right_x, right_y, right_w, right_h)

    c.setFont("Helvetica-Bold", 8)
    c.drawString(right_x + 4 * mm, right_y + right_h - 8 * mm, f'Freight Bill No: {safe_str(row.get("FreightBillNo"))}')
    c.drawString(right_x + 4 * mm, right_y + right_h - 14 * mm, f'Invoice Date: {format_date(row.get("InvoiceDate"))}')
    c.drawString(right_x + 4 * mm, right_y + right_h - 20 * mm, f'Due Date: {format_date(row.get("DueDate"))}')

    table_x = LM + 2 * mm
    table_top = left_y - 18 * mm
    table_w = (W - LM - RM) - 4 * mm
    header_h = 12 * mm
    row_h = 10 * mm

    cols = [
        ("S.\nno.",10),("Shipment\nDate",20),("LR No.",14),("Destination",22),
        ("CN\nNumber",18),("Truck No",18),("Invoice No",18),
        ("Pkgs",10),("Weight\n(kgs)",14),("Date of\nArrival",16),
        ("Date of\nDelivery",16),("Truck\nType",14),
        ("Freight\nAmt (Rs.)",16),("To Point\nCharges(Rs.)",16),
        ("Unloading\nCharge (Rs.)",16),("Source\nDetention\n(Rs.)",16),
        ("Destination\nDetention\n(Rs.)",16),("Total\nAmount (Rs.)",18)
    ]

    scale = table_w / sum(w for _, w in cols)
    cols = [(n, w * scale) for n, w in cols]

    c.rect(table_x, table_top - header_h, table_w, header_h)
    c.setFont("Helvetica-Bold", 6.5)

    x = table_x
    for name, w in cols:
        c.line(x, table_top - header_h, x, table_top)
        cx = x + w / 2
        yy = table_top - 4 * mm
        for p in name.split("\n"):
            c.drawCentredString(cx, yy, p)
            yy -= 3 * mm
        x += w
    c.line(table_x + table_w, table_top - header_h, table_x + table_w, table_top)

    data_top = table_top - header_h
    data_bottom = data_top - row_h
    c.rect(table_x, data_bottom, table_w, row_h)

    total_amt = calc_total(row)

    data = [
        "1",format_date(row.get("ShipmentDate")),safe_str(row.get("LRNo")),
        safe_str(row.get("Destination")),safe_str(row.get("CNNumber")),
        safe_str(row.get("TruckNo")),safe_str(row.get("InvoiceNo")),
        safe_str(row.get("Pkgs")),safe_str(row.get("WeightKgs")),
        format_date(row.get("DateArrival")),format_date(row.get("DateDelivery")),
        safe_str(row.get("TruckType")),money(row.get("FreightAmt")),
        money(row.get("ToPointCharges")),money(row.get("UnloadingCharge")),
        money(row.get("SourceDetention")),money(row.get("DestinationDetention")),
        money(total_amt)
    ]

    c.setFont("Helvetica", 7)
    x = table_x
    for (_, w), txt in zip(cols, data):
        c.line(x, data_bottom, x, data_top)
        c.drawCentredString(x + w / 2, data_bottom + 3.5 * mm, txt)
        x += w
    c.line(table_x + table_w, data_bottom, table_x + table_w, data_top)

    c.showPage()
    c.save()

@app.route("/", methods=["GET", "POST"])
def index():
    if request.method == "POST":
        file = request.files.get("file")
        filepath = os.path.join(UPLOAD_FOLDER, file.filename)
        file.save(filepath)
        df = pd.read_excel(filepath)
        df.columns = [str(c).strip() for c in df.columns]

        generated = []
        for _, r in df.iterrows():
            row = r.to_dict()
            pdf_name = f'{row.get("FreightBillNo")}_{datetime.now().strftime("%H%M%S")}.pdf'
            pdf_path = os.path.join(OUTPUT_FOLDER, pdf_name)
            generate_invoice_pdf(row, pdf_path)
            generated.append(pdf_path)

        zip_path = os.path.join(OUTPUT_FOLDER, "Bills.zip")
        with zipfile.ZipFile(zip_path, "w") as z:
            for p in generated:
                z.write(p, os.path.basename(p))

        return send_file(zip_path, as_attachment=True)

    return render_template("index.html")

if __name__ == "__main__":
    app.run(debug=True)
