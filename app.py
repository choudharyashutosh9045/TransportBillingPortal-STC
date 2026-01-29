from flask import Flask, render_template, request, send_file, jsonify
import os
import pandas as pd
from datetime import datetime
from num2words import num2words
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.units import mm
from reportlab.lib.utils import ImageReader
import zipfile

app = Flask(__name__)

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
UPLOAD_FOLDER = os.path.join(BASE_DIR, "uploads")
OUTPUT_FOLDER = os.path.join(BASE_DIR, "output")

os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)

FIXED_PARTY = {
    "PartyName": "Grivaa Springs Private Ltd.",
    "PartyAddress": "Khasra no 135, Tansipur, Roorkee",
    "PartyCity": "Roorkee",
    "PartyState": "Uttarakhand",
    "PartyPincode": "247656",
    "PartyGSTIN": "05AAICG4793P1ZV"
}

FIXED_BANK = {
    "PANNo": "BSSPG9414K",
    "GSTIN": "05BSSPG9414K1ZA",
    "StateCode": "5",
    "AccountName": "South Transport Company",
    "AccountNo": "364205500142",
    "IFSCode": "ICIC0003642"
}

REQUIRED_HEADERS = [
    "FreightBillNo","InvoiceDate","DueDate","FromLocation","ShipmentDate",
    "LRNo","Destination","CNNumber","TruckNo","InvoiceNo","Pkgs","WeightKgs",
    "DateArrival","DateDelivery","TruckType","FreightAmt","ToPointCharges",
    "UnloadingCharge","SourceDetention","DestinationDetention"
]

def s(v):
    return "" if pd.isna(v) else str(v)

def f(v):
    try:
        return float(v)
    except:
        return 0.0

def d(v):
    try:
        return pd.to_datetime(v).strftime("%d %b %Y")
    except:
        return s(v)

def generate_invoice_pdf(row, pdf_path):
    W, H = landscape(A4)
    c = canvas.Canvas(pdf_path, pagesize=(W, H))

    LM = 10 * mm
    BM = 10 * mm

    c.rect(LM, BM, W - 20*mm, H - 20*mm)

    c.setFont("Helvetica-Bold", 14)
    c.drawCentredString(W/2, H-20*mm, "SOUTH TRANSPORT COMPANY")
    c.setFont("Helvetica", 8)
    c.drawCentredString(W/2, H-25*mm, "Dehradun Road Near power Grid Bhagwanpur")
    c.drawCentredString(W/2, H-29*mm, "Roorkee, Haridwar, U.K. 247661, India")
    c.setFont("Helvetica-Bold", 10)
    c.drawCentredString(W/2, H-36*mm, "INVOICE")

    logo = os.path.join(BASE_DIR, "logo.png")
    if os.path.exists(logo):
        c.drawImage(ImageReader(logo), LM+5*mm, H-60*mm, 70*mm, 35*mm, mask='auto')

    c.rect(LM+5*mm, H-95*mm, 110*mm, 30*mm)
    c.setFont("Helvetica-Bold", 8)
    c.drawString(LM+8*mm, H-75*mm, "To,")
    c.drawString(LM+8*mm, H-80*mm, FIXED_PARTY["PartyName"])
    c.setFont("Helvetica", 7)
    c.drawString(LM+8*mm, H-85*mm, FIXED_PARTY["PartyAddress"])
    c.drawString(LM+8*mm, H-89*mm, f'{FIXED_PARTY["PartyCity"]}, {FIXED_PARTY["PartyState"]} {FIXED_PARTY["PartyPincode"]}')
    c.setFont("Helvetica-Bold", 7)
    c.drawString(LM+8*mm, H-93*mm, f'GSTIN: {FIXED_PARTY["PartyGSTIN"]}')

    c.setFont("Helvetica", 7)
    c.drawString(LM+8*mm, H-100*mm, f'From location: {s(row["FromLocation"])}')

    c.rect(W-100*mm, H-95*mm, 85*mm, 30*mm)
    c.setFont("Helvetica-Bold", 8)
    c.drawString(W-96*mm, H-80*mm, f'Freight Bill No: {s(row["FreightBillNo"])}')
    c.drawString(W-96*mm, H-86*mm, f'Invoice Date: {d(row["InvoiceDate"])}')
    c.drawString(W-96*mm, H-92*mm, f'Due Date: {d(row["DueDate"])}')

    table_y = H-115*mm
    row_h = 10*mm

    cols = [
        ("S.No",15),("Shipment Date",30),("LR No",20),("Destination",30),
        ("CN No",25),("Truck No",30),("Invoice No",55),("Pkgs",18),
        ("Weight",22),("Arrival",28),("Delivery",28),("Truck Type",35),
        ("Freight",28),("To Point",28),("Unloading",28),
        ("Src Det.",28),("Dest Det.",28),("Total",30)
    ]

    x = LM+5*mm
    c.setFont("Helvetica-Bold", 7)
    for h,w in cols:
        c.rect(x, table_y, w, row_h)
        c.drawCentredString(x+w/2, table_y+4*mm, h)
        x += w

    total = (
        f(row["FreightAmt"]) + f(row["ToPointCharges"]) +
        f(row["UnloadingCharge"]) + f(row["SourceDetention"]) +
        f(row["DestinationDetention"])
    )

    data = [
        "1", d(row["ShipmentDate"]), s(row["LRNo"]), s(row["Destination"]),
        s(row["CNNumber"]), s(row["TruckNo"]), s(row["InvoiceNo"]),
        s(row["Pkgs"]), s(row["WeightKgs"]), d(row["DateArrival"]),
        d(row["DateDelivery"]), s(row["TruckType"]),
        f'{f(row["FreightAmt"]):.2f}', f'{f(row["ToPointCharges"]):.2f}',
        f'{f(row["UnloadingCharge"]):.2f}', f'{f(row["SourceDetention"]):.2f}',
        f'{f(row["DestinationDetention"]):.2f}', f'{total:.2f}'
    ]

    x = LM+5*mm
    y = table_y-row_h
    c.setFont("Helvetica", 7)
    for (h,w),v in zip(cols,data):
        c.rect(x, y, w, row_h)
        c.drawCentredString(x+w/2, y+4*mm, v)
        x += w

    c.rect(LM+5*mm, y-row_h, sum(w for _,w in cols), row_h)
    words = num2words(int(total), lang="en").title()+" Rupees Only"
    c.drawString(LM+8*mm, y-row_h+4*mm, f"Total in words (Rs.): {words}")
    c.drawRightString(W-15*mm, y-row_h+4*mm, f"{total:.2f}")

    bank_y = BM+20*mm
    c.rect(LM+5*mm, bank_y, 90*mm, 35*mm)
    c.setFont("Helvetica", 7)
    lines = [
        ("Our PAN No", FIXED_BANK["PANNo"]),
        ("STC GSTIN", FIXED_BANK["GSTIN"]),
        ("STC State Code", FIXED_BANK["StateCode"]),
        ("Account Name", FIXED_BANK["AccountName"]),
        ("Account No", FIXED_BANK["AccountNo"]),
        ("IFS Code", FIXED_BANK["IFSCode"])
    ]
    yy = bank_y+28*mm
    for k,v in lines:
        c.drawString(LM+8*mm, yy, k)
        c.drawString(LM+45*mm, yy, v)
        yy -= 5*mm

    c.setFont("Helvetica-Bold", 8)
    c.drawString(W-80*mm, BM+45*mm, "For SOUTH TRANSPORT COMPANY")
    c.line(W-80*mm, BM+35*mm, W-20*mm, BM+35*mm)
    c.setFont("Helvetica", 7)
    c.drawRightString(W-20*mm, BM+30*mm, "(Authorized Signatory)")

    c.save()

@app.route("/", methods=["GET","POST"])
def index():
    if request.method == "POST":
        file = request.files["file"]
        path = os.path.join(UPLOAD_FOLDER, file.filename)
        file.save(path)

        df = pd.read_excel(path)
        df.columns = df.columns.str.strip()

        missing = [h for h in REQUIRED_HEADERS if h not in df.columns]
        if missing:
            return f"Missing columns: {missing}", 400

        pdfs = []
        for _,row in df.iterrows():
            name = f'{row["FreightBillNo"]}_{row["LRNo"]}.pdf'
            out = os.path.join(OUTPUT_FOLDER, name)
            generate_invoice_pdf(row, out)
            pdfs.append(out)

        zip_path = os.path.join(OUTPUT_FOLDER, "Invoices.zip")
        with zipfile.ZipFile(zip_path, "w") as z:
            for p in pdfs:
                z.write(p, os.path.basename(p))

        return send_file(zip_path, as_attachment=True)

    return "<h2>Upload Excel</h2><form method='post' enctype='multipart/form-data'><input type='file' name='file'><button>Upload</button></form>"

if __name__ == "__main__":
    app.run(debug=True)
