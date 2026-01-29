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

import os
from reportlab.lib.pagesizes import A4, landscape
from reportlab.pdfgen import canvas
from reportlab.lib.units import mm
from reportlab.lib.utils import ImageReader
from num2words import num2words
from datetime import datetime

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

def safe(v):
    return "" if v is None else str(v)

def fmt_date(d):
    if not d:
        return ""
    if isinstance(d, datetime):
        return d.strftime("%d %b %Y")
    return str(d)

def money(v):
    try:
        return f"{float(v):.2f}"
    except:
        return "0.00"

def total_amt(r):
    return (
        float(r.get("FreightAmt", 0)) +
        float(r.get("ToPointCharges", 0)) +
        float(r.get("UnloadingCharge", 0)) +
        float(r.get("SourceDetention", 0)) +
        float(r.get("DestinationDetention", 0))
    )

def generate_invoice_pdf(row, pdf_path):

    W, H = landscape(A4)
    c = canvas.Canvas(pdf_path, pagesize=(W, H))

    LM, RM, TM, BM = 12*mm, 12*mm, 12*mm, 12*mm

    # ================= BORDER =================
    c.setLineWidth(1)
    c.rect(LM, BM, W-LM-RM, H-TM-BM)

    # ================= HEADER =================
    c.setFont("Helvetica-Bold", 14)
    c.drawCentredString(W/2, H-TM-8*mm, "SOUTH TRANSPORT COMPANY")

    c.setFont("Helvetica", 8)
    c.drawCentredString(W/2, H-TM-13*mm, "Dehradun Road Near Power Grid Bhagwanpur")
    c.drawCentredString(W/2, H-TM-17*mm, "Roorkee, Haridwar, U.K. 247661, India")

    c.setFont("Helvetica-Bold", 10)
    c.drawCentredString(W/2, H-TM-24*mm, "INVOICE")

    # ================= LOGO =================
    logo = os.path.join(BASE_DIR, "logo.png")
    if os.path.exists(logo):
        c.drawImage(
            ImageReader(logo),
            LM+5*mm,
            H-TM-45*mm,
            width=65*mm,
            height=32*mm,
            preserveAspectRatio=True,
            mask="auto"
        )

    # ================= LEFT BOX =================
    lbx, lby = LM+2*mm, H-TM-72*mm
    lbw, lbh = 110*mm, 32*mm
    c.rect(lbx, lby, lbw, lbh)

    c.setFont("Helvetica-Bold", 8)
    c.drawString(lbx+3*mm, lby+lbh-6*mm, "To,")

    c.drawString(lbx+3*mm, lby+lbh-11*mm, safe(row["PartyName"]))
    c.setFont("Helvetica", 7.5)
    c.drawString(lbx+3*mm, lby+lbh-16*mm, safe(row["PartyAddress"]))
    c.drawString(lbx+3*mm, lby+lbh-20*mm,
        f"{safe(row['PartyCity'])}, {safe(row['PartyState'])} {safe(row['PartyPincode'])}"
    )
    c.setFont("Helvetica-Bold", 7.5)
    c.drawString(lbx+3*mm, lby+lbh-25*mm, f"GSTIN: {safe(row['PartyGSTIN'])}")

    c.setFont("Helvetica", 7.5)
    c.drawString(lbx+3*mm, lby-5*mm, f"From location: {safe(row['FromLocation'])}")

    # ================= RIGHT BOX =================
    rbx, rby = W-RM-90*mm, lby
    rbw, rbh = 90*mm, 32*mm
    c.rect(rbx, rby, rbw, rbh)

    c.setFont("Helvetica-Bold", 8)
    c.drawString(rbx+6*mm, rby+rbh-10*mm, f"Freight Bill No: {safe(row['FreightBillNo'])}")
    c.drawString(rbx+6*mm, rby+rbh-17*mm, f"Invoice Date: {fmt_date(row['InvoiceDate'])}")
    c.drawString(rbx+6*mm, rby+rbh-24*mm, f"Due Date: {fmt_date(row['DueDate'])}")

    # ================= TABLE =================
    tx = LM+2*mm
    top = lby-20*mm

    cols = [
        ("S.\nNo", 12),
        ("Shipment\nDate", 22),
        ("LR No.", 16),
        ("Destination", 22),
        ("CN\nNumber", 20),
        ("Truck No", 28),
        ("Invoice No", 36),
        ("Pkgs", 12),
        ("Weight\n(kgs)", 16),
        ("Date of\nArrival", 18),
        ("Date of\nDelivery", 18),
        ("Truck\nType", 18),
        ("Freight\nAmt", 18),
        ("To Point\nCharges", 18),
        ("Unloading\nCharge", 18),
        ("Source\nDetention", 18),
        ("Destination\nDetention", 18),
        ("Total\nAmount", 20),
    ]

    h_h, r_h = 12*mm, 10*mm
    total_w = sum(w for _, w in cols)

    # Header
    x = tx
    c.rect(tx, top-h_h, total_w*mm, h_h)
    c.setFont("Helvetica-Bold", 6.5)
    for name, w in cols:
        c.line(x*mm, top-h_h, x*mm, top)
        yy = top-4*mm
        for p in name.split("\n"):
            c.drawCentredString(x*mm + (w*mm)/2, yy, p)
            yy -= 3*mm
        x += w
    c.line((tx+total_w)*mm, top-h_h, (tx+total_w)*mm, top)

    # Data
    data_top = top-h_h
    data_bot = data_top-r_h
    c.rect(tx*mm, data_bot, total_w*mm, r_h)

    t_amt = total_amt(row)

    data = [
        "1",
        fmt_date(row["ShipmentDate"]),
        safe(row["LRNo"]),
        safe(row["Destination"]),
        safe(row["CNNumber"]),
        safe(row["TruckNo"]),
        safe(row["InvoiceNo"]),
        safe(row["Pkgs"]),
        safe(row["WeightKgs"]),
        fmt_date(row["DateArrival"]),
        fmt_date(row["DateDelivery"]),
        safe(row["TruckType"]),
        money(row["FreightAmt"]),
        money(row["ToPointCharges"]),
        money(row["UnloadingCharge"]),
        money(row["SourceDetention"]),
        money(row["DestinationDetention"]),
        money(t_amt),
    ]

    c.setFont("Helvetica", 7)
    x = tx
    for (_, w), v in zip(cols, data):
        c.line(x*mm, data_bot, x*mm, data_top)
        c.drawCentredString(x*mm + (w*mm)/2, data_bot+3.5*mm, v)
        x += w
    c.line((tx+total_w)*mm, data_bot, (tx+total_w)*mm, data_top)

    # ================= TOTAL IN WORDS =================
    wb = data_bot-7*mm
    c.rect(tx*mm, wb, total_w*mm, 7*mm)
    c.setFont("Helvetica-Bold", 7)
    c.drawString(tx*mm+3*mm, wb+2.2*mm, "Total in words (Rs.):")
    c.setFont("Helvetica", 7)
    c.drawString(tx*mm+35*mm, wb+2.2*mm,
        num2words(int(t_amt), lang="en").title()+" Rupees Only"
    )
    c.drawRightString((tx+total_w)*mm-2*mm, wb+2.2*mm, money(t_amt))

    # ================= NOTE =================
    c.setFont("Helvetica", 7)
    c.drawString(
        tx*mm,
        wb-10*mm,
        'Any changes or discrepancies should be highlighted within 5 working days else it will be considered final.'
    )

    # ================= BANK =================
    bx, by = tx*mm, BM+8*mm
    bw, bh = 90*mm, 34*mm
    c.rect(bx, by, bw, bh)

    bank = [
        ("Our PAN No.", row["PANNo"]),
        ("STC GSTIN", row["STCGSTIN"]),
        ("STC State Code", row["STCStateCode"]),
        ("Account Name", row["AccountName"]),
        ("Account No", row["AccountNo"]),
        ("IFS Code", row["IFSCode"]),
    ]

    rh = bh/len(bank)
    c.setFont("Helvetica-Bold", 7)
    for i,(k,v) in enumerate(bank):
        y = by+bh-(i+1)*rh
        c.line(bx, y, bx+bw, y)
        c.drawString(bx+3*mm, y+2*mm, k)
        c.drawString(bx+35*mm, y+2*mm, safe(v))
    c.line(bx+33*mm, by, bx+33*mm, by+bh)

    # ================= SIGN =================
    sx = W-RM-80*mm
    sy = BM+12*mm
    c.setFont("Helvetica-Bold", 8)
    c.drawString(sx, sy+18*mm, "For SOUTH TRANSPORT COMPANY")
    c.line(sx+10*mm, sy+8*mm, sx+70*mm, sy+8*mm)
    c.setFont("Helvetica", 7)
    c.drawRightString(sx+75*mm, sy+2*mm, "(Authorized Signatory)")

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
