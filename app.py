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
    "FreightBillNo","InvoiceDate","DueDate","FromLocation","ShipmentDate",
    "LRNo","Destination","CNNumber","TruckNo","InvoiceNo","Pkgs","WeightKgs",
    "DateArrival","DateDelivery","TruckType","FreightAmt","ToPointCharges",
    "UnloadingCharge","SourceDetention","DestinationDetention"
]

def safe(v):
    if pd.isna(v):
        return ""
    return str(v)

def fdate(v):
    try:
        return pd.to_datetime(v).strftime("%d %b %Y")
    except:
        return ""

def fnum(v):
    try:
        return float(v)
    except:
        return 0.0

def generate_invoice_pdf(row, pdf_path):
    row = {**FIXED_PARTY, **FIXED_STC_BANK, **row}

    W, H = landscape(A4)
    c = canvas.Canvas(pdf_path, pagesize=(W, H))

    LM = 10*mm; RM = 10*mm; TM = 10*mm; BM = 10*mm

    c.rect(LM, BM, W-LM-RM, H-TM-BM)

    c.setFont("Helvetica-Bold", 14)
    c.drawCentredString(W/2, H-20, "SOUTH TRANSPORT COMPANY")

    c.setFont("Helvetica", 8)
    c.drawCentredString(W/2, H-35, "Dehradun Road Near power Grid Bhagwanpur")
    c.drawCentredString(W/2, H-48, "Roorkee, Haridwar, U.K. 247661, India")

    c.setFont("Helvetica-Bold", 10)
    c.drawCentredString(W/2, H-65, "INVOICE")

    logo = os.path.join(BASE_DIR, "logo.png")
    if os.path.exists(logo):
        c.drawImage(ImageReader(logo), LM+10, H-120, 90, 45, mask="auto")

    c.rect(LM+10, H-180, 300, 80)
    c.setFont("Helvetica-Bold", 9)
    c.drawString(LM+15, H-155, "To,")
    c.drawString(LM+15, H-170, row["PartyName"])
    c.setFont("Helvetica", 8)
    c.drawString(LM+15, H-185, row["PartyAddress"])
    c.drawString(LM+15, H-198, f'{row["PartyCity"]}, {row["PartyState"]} {row["PartyPincode"]}')
    c.setFont("Helvetica-Bold", 8)
    c.drawString(LM+15, H-212, f'GSTIN: {row["PartyGSTIN"]}')

    c.rect(W-320, H-180, 260, 80)
    c.setFont("Helvetica-Bold", 9)
    c.drawString(W-300, H-155, f'Freight Bill No: {row["FreightBillNo"]}')
    c.drawString(W-300, H-175, f'Invoice Date: {fdate(row["InvoiceDate"])}')
    c.drawString(W-300, H-195, f'Due Date: {fdate(row["DueDate"])}')

    y = H-260
    c.rect(LM+10, y, W-40, 70)

    headers = [
        "S.No","Shipment Date","LR No","Destination","CN No","Truck No",
        "Invoice No","Pkgs","Weight","Arrival","Delivery","Truck Type",
        "Freight","To Point","Unloading","Src Det","Dest Det","Total"
    ]

    x = LM+10
    col_w = (W-40)/len(headers)

    c.setFont("Helvetica-Bold", 7)
    for h in headers:
        c.drawCentredString(x+col_w/2, y+50, h)
        c.line(x, y, x, y+70)
        x += col_w

    total = (
        fnum(row["FreightAmt"]) +
        fnum(row["ToPointCharges"]) +
        fnum(row["UnloadingCharge"]) +
        fnum(row["SourceDetention"]) +
        fnum(row["DestinationDetention"])
    )

    values = [
        "1", fdate(row["ShipmentDate"]), row["LRNo"], row["Destination"],
        row["CNNumber"], row["TruckNo"], row["InvoiceNo"], row["Pkgs"],
        row["WeightKgs"], fdate(row["DateArrival"]), fdate(row["DateDelivery"]),
        row["TruckType"], fnum(row["FreightAmt"]),
        fnum(row["ToPointCharges"]), fnum(row["UnloadingCharge"]),
        fnum(row["SourceDetention"]), fnum(row["DestinationDetention"]), total
    ]

    c.setFont("Helvetica", 7)
    x = LM+10
    for v in values:
        c.drawCentredString(x+col_w/2, y+20, safe(v))
        x += col_w

    c.rect(LM+10, y-30, W-40, 30)
    words = num2words(int(total)).title()+" Rupees Only"
    c.drawString(LM+15, y-20, "Total in words (Rs.): "+words)
    c.drawRightString(W-50, y-20, f"{total:.2f}")

    c.rect(LM+10, BM+40, 300, 90)
    bank = [
        ("PAN No",row["PANNo"]),
        ("GSTIN",row["STCGSTIN"]),
        ("State Code",row["STCStateCode"]),
        ("Account Name",row["AccountName"]),
        ("Account No",row["AccountNo"]),
        ("IFSC",row["IFSCode"]),
    ]

    yy = BM+110
    for k,v in bank:
        c.drawString(LM+20, yy, k)
        c.drawString(LM+150, yy, v)
        yy -= 15

    c.drawString(W-250, BM+80, "For SOUTH TRANSPORT COMPANY")
    c.line(W-250, BM+55, W-70, BM+55)
    c.drawString(W-170, BM+40, "(Authorized Signatory)")

    c.showPage()
    c.save()

@app.route("/", methods=["POST"])
def upload():
    file = request.files["file"]
    path = os.path.join(UPLOAD_FOLDER, file.filename)
    file.save(path)

    df = pd.read_excel(path)
    df.columns = [c.strip() for c in df.columns]

    pdfs = []
    for _, r in df.iterrows():
        name = f'{r["FreightBillNo"]}.pdf'
        out = os.path.join(OUTPUT_FOLDER, name)
        generate_invoice_pdf(r.to_dict(), out)
        pdfs.append(out)

    zip_name = os.path.join(OUTPUT_FOLDER, "Bills.zip")
    with zipfile.ZipFile(zip_name, "w") as z:
        for p in pdfs:
            z.write(p, os.path.basename(p))

    return send_file(zip_name, as_attachment=True)

if __name__ == "__main__":
    app.run(debug=True)
