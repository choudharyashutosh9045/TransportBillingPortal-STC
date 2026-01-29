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

# ================= PATHS =================
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
UPLOAD_FOLDER = os.path.join(BASE_DIR, "uploads")
OUTPUT_FOLDER = os.path.join(BASE_DIR, "output")
TEMPLATE_FOLDER = os.path.join(BASE_DIR, "templates")

os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)

app.config["UPLOAD_FOLDER"] = UPLOAD_FOLDER

# ================= FIXED DATA =================
FIXED_PARTY = {
    "PartyName": "Grivaa Springs Private Ltd.",
    "PartyAddress": "Khasra no 135, Tansipur, Roorkee",
    "PartyCity": "Roorkee",
    "PartyState": "Uttarakhand",
    "PartyPincode": "247656",
    "PartyGSTIN": "05AAICG4793P1ZV",
}

FIXED_BANK = {
    "PANNo": "BSSPG9414K",
    "STCGSTIN": "05BSSPG9414K1ZA",
    "STCStateCode": "5",
    "AccountName": "South Transport Company",
    "AccountNo": "364205500142",
    "IFSCode": "ICIC0003642",
}

# ================= REQUIRED EXCEL HEADERS =================
REQUIRED_HEADERS = [
    "FreightBillNo","InvoiceDate","DueDate","FromLocation",
    "ShipmentDate","LRNo","Destination","CNNumber","TruckNo",
    "InvoiceNo","Pkgs","WeightKgs","DateArrival","DateDelivery",
    "TruckType","FreightAmt","ToPointCharges","UnloadingCharge",
    "SourceDetention","DestinationDetention"
]

# ================= HELPERS =================
def s(v): 
    return "" if pd.isna(v) else str(v)

def f(v):
    try:
        return float(v)
    except:
        return 0.0

def d(v):
    try:
        return pd.to_datetime(v, dayfirst=True).strftime("%d %b %Y")
    except:
        return s(v)

def money(v):
    return f"{f(v):.2f}"

def total(row):
    return (
        f(row["FreightAmt"]) +
        f(row["ToPointCharges"]) +
        f(row["UnloadingCharge"]) +
        f(row["SourceDetention"]) +
        f(row["DestinationDetention"])
    )

# ================= PDF GENERATOR =================
def generate_invoice_pdf(row, pdf_path):
    row = {**FIXED_PARTY, **FIXED_BANK, **row}

    W, H = landscape(A4)
    c = canvas.Canvas(pdf_path, pagesize=(W, H))
    LM, RM, TM, BM = 10*mm, 10*mm, 10*mm, 10*mm

    # OUTER BORDER
    c.rect(LM, BM, W-LM-RM, H-TM-BM)

    # HEADER
    c.setFont("Helvetica-Bold", 14)
    c.drawCentredString(W/2, H-15*mm, "SOUTH TRANSPORT COMPANY")
    c.setFont("Helvetica", 8)
    c.drawCentredString(W/2, H-20*mm, "Dehradun Road Near Power Grid Bhagwanpur")
    c.drawCentredString(W/2, H-24*mm, "Roorkee, Haridwar, U.K. 247661, India")
    c.setFont("Helvetica-Bold", 10)
    c.drawCentredString(W/2, H-30*mm, "INVOICE")

    # LOGO
    logo = os.path.join(BASE_DIR, "logo.png")
    if os.path.exists(logo):
        c.drawImage(logo, LM+5*mm, H-45*mm, 60*mm, 30*mm, preserveAspectRatio=True)

    # TO BOX
    c.rect(LM+5*mm, H-80*mm, 110*mm, 30*mm)
    c.setFont("Helvetica-Bold", 8)
    c.drawString(LM+8*mm, H-55*mm, "To,")
    c.drawString(LM+8*mm, H-60*mm, row["PartyName"])
    c.setFont("Helvetica", 7.5)
    c.drawString(LM+8*mm, H-65*mm, row["PartyAddress"])
    c.drawString(LM+8*mm, H-70*mm, f"{row['PartyCity']} {row['PartyPincode']}")
    c.setFont("Helvetica-Bold", 7.5)
    c.drawString(LM+8*mm, H-75*mm, f"GSTIN: {row['PartyGSTIN']}")

    c.setFont("Helvetica", 7.5)
    c.drawString(LM+8*mm, H-85*mm, f"From location: {row['FromLocation']}")

    # RIGHT BOX
    rx = W-RM-90*mm
    c.rect(rx, H-80*mm, 85*mm, 30*mm)
    c.setFont("Helvetica-Bold", 8)
    c.drawString(rx+5*mm, H-60*mm, f"Freight Bill No: {row['FreightBillNo']}")
    c.drawString(rx+5*mm, H-66*mm, f"Invoice Date: {d(row['InvoiceDate'])}")
    c.drawString(rx+5*mm, H-72*mm, f"Due Date: {d(row['DueDate'])}")

    # TABLE HEADER
    table_y = H-105*mm
    headers = [
        "S.No","Shipment Date","LR No","Destination","CN No","Truck No",
        "Invoice No","Pkgs","Weight","Arrival","Delivery","Truck Type",
        "Freight","To Point","Unloading","Src Det","Dst Det","Total"
    ]
    widths = [10,20,14,22,16,18,18,10,14,16,16,14,16,16,16,16,16,18]
    scale = (W-LM-RM-10*mm)/sum(widths)
    widths = [w*scale for w in widths]

    x = LM+5*mm
    c.setFont("Helvetica-Bold", 6.5)
    for h,w in zip(headers,widths):
        c.rect(x, table_y, w, 10*mm)
        c.drawCentredString(x+w/2, table_y+4*mm, h)
        x += w

    # DATA ROW
    vals = [
        "1", d(row["ShipmentDate"]), row["LRNo"], row["Destination"],
        row["CNNumber"], row["TruckNo"], row["InvoiceNo"], row["Pkgs"],
        row["WeightKgs"], d(row["DateArrival"]), d(row["DateDelivery"]),
        row["TruckType"], money(row["FreightAmt"]), money(row["ToPointCharges"]),
        money(row["UnloadingCharge"]), money(row["SourceDetention"]),
        money(row["DestinationDetention"]), money(total(row))
    ]

    x = LM+5*mm
    c.setFont("Helvetica", 7)
    for v,w in zip(vals,widths):
        c.rect(x, table_y-10*mm, w, 10*mm)
        c.drawCentredString(x+w/2, table_y-6*mm, s(v))
        x += w

    # TOTAL WORDS
    words = num2words(int(total(row)), lang="en").title()+" Rupees Only"
    c.rect(LM+5*mm, table_y-20*mm, W-LM-RM-10*mm, 8*mm)
    c.setFont("Helvetica-Bold", 7)
    c.drawString(LM+8*mm, table_y-17*mm, "Total in words (Rs.):")
    c.setFont("Helvetica", 7)
    c.drawString(LM+40*mm, table_y-17*mm, words)
    c.drawRightString(W-RM-8*mm, table_y-17*mm, money(total(row)))

    # BANK DETAILS
    bx, by = LM+5*mm, BM+15*mm
    bw, bh = 90*mm, 35*mm
    c.rect(bx, by, bw, bh)

    bank = [
        ("Our PAN No.", row["PANNo"]),
        ("STC GSTIN", row["STCGSTIN"]),
        ("STC State Code", row["STCStateCode"]),
        ("Account name", row["AccountName"]),
        ("Account no", row["AccountNo"]),
        ("IFS Code", row["IFSCode"]),
    ]
    rh = bh/len(bank)
    for i,(k,v) in enumerate(bank):
        y = by+bh-(i+1)*rh
        c.line(bx,y,bx+bw,y)
        c.drawString(bx+3*mm,y+2*mm,k)
        c.drawString(bx+35*mm,y+2*mm,v)

    # SIGN
    c.setFont("Helvetica-Bold", 8)
    c.drawString(W-RM-80*mm, BM+40*mm, "For SOUTH TRANSPORT COMPANY")
    c.line(W-RM-80*mm, BM+30*mm, W-RM-20*mm, BM+30*mm)
    c.setFont("Helvetica", 7)
    c.drawRightString(W-RM-20*mm, BM+24*mm, "(Authorized Signatory)")

    c.showPage()
    c.save()

# ================= ROUTES =================
@app.route("/", methods=["GET","POST"])
def index():
    if request.method == "POST":
        file = request.files.get("file")
        if not file:
            return "No file", 400

        path = os.path.join(UPLOAD_FOLDER, file.filename)
        file.save(path)

        df = pd.read_excel(path)
        df.columns = df.columns.str.strip()

        missing = [h for h in REQUIRED_HEADERS if h not in df.columns]
        if missing:
            return f"Missing columns: {missing}", 400

        pdfs = []
        for _,r in df.iterrows():
            name = f"{r['FreightBillNo']}_{r['LRNo']}.pdf"
            out = os.path.join(OUTPUT_FOLDER, name)
            generate_invoice_pdf(r.to_dict(), out)
            pdfs.append(out)

        zip_path = os.path.join(OUTPUT_FOLDER,"Bills.zip")
        with zipfile.ZipFile(zip_path,"w") as z:
            for p in pdfs:
                z.write(p, os.path.basename(p))

        return send_file(zip_path, as_attachment=True)

    return render_template("index.html")


if __name__ == "__main__":
    app.run(debug=True)
