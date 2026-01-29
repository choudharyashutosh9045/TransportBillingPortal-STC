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
TEMPLATE_FOLDER = os.path.join(BASE_DIR, "templates")

os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)
os.makedirs(TEMPLATE_FOLDER, exist_ok=True)

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
    "ShipmentDate","LRNo","Destination","CNNumber",
    "TruckNo","InvoiceNo","Pkgs","WeightKgs",
    "DateArrival","DateDelivery","TruckType",
    "FreightAmt","ToPointCharges","UnloadingCharge",
    "SourceDetention","DestinationDetention"
]

def s(v):
    if pd.isna(v):
        return ""
    return str(v).strip()

def f(v):
    try:
        if pd.isna(v) or str(v).strip()=="":
            return 0.0
        return float(v)
    except:
        return 0.0

def d(v):
    try:
        return pd.to_datetime(v, dayfirst=True).strftime("%d %b %Y")
    except:
        return ""

def money(v):
    return f"{f(v):.2f}"

def total(r):
    return (
        f(r.get("FreightAmt")) +
        f(r.get("ToPointCharges")) +
        f(r.get("UnloadingCharge")) +
        f(r.get("SourceDetention")) +
        f(r.get("DestinationDetention"))
    )

def generate_invoice_pdf(row, pdf_path):
    row = {**FIXED_PARTY, **FIXED_STC_BANK, **row}
    W, H = landscape(A4)
    c = canvas.Canvas(pdf_path, pagesize=(W, H))

    LM = RM = TM = BM = 10 * mm
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
        c.drawImage(logo, LM+5, H-100, width=75*mm, height=40*mm, preserveAspectRatio=True)

    c.rect(LM+5, H-160, 110*mm, 30*mm)
    c.setFont("Helvetica-Bold", 8)
    c.drawString(LM+10, H-145, "To,")
    c.drawString(LM+10, H-158, s(row["PartyName"]))
    c.setFont("Helvetica", 7)
    c.drawString(LM+10, H-170, s(row["PartyAddress"]))
    c.drawString(LM+10, H-182, f'{row["PartyCity"]}, {row["PartyState"]} {row["PartyPincode"]}')
    c.setFont("Helvetica-Bold", 7)
    c.drawString(LM+10, H-194, f'GSTIN: {row["PartyGSTIN"]}')
    c.setFont("Helvetica", 7)
    c.drawString(LM+10, H-208, f'From location: {s(row["FromLocation"])}')

    rx = W-RM-90*mm
    c.rect(rx, H-160, 85*mm, 30*mm)
    c.setFont("Helvetica-Bold", 8)
    c.drawString(rx+8, H-145, f'Freight Bill No: {s(row["FreightBillNo"])}')
    c.drawString(rx+8, H-160, f'Invoice Date: {d(row["InvoiceDate"])}')
    c.drawString(rx+8, H-175, f'Due Date: {d(row["DueDate"])}')

    tx = LM+5
    ty = H-235
    tw = W-LM-RM-10
    th = 12
    rh = 10

    headers = [
        "S.No","Shipment Date","LR No","Destination","CN No","Truck No",
        "Invoice No","Pkgs","Weight","Arrival","Delivery","Truck Type",
        "Freight","To Point","Unloading","Source Det.","Dest. Det.","Total"
    ]

    colw = tw/len(headers)
    c.setFont("Helvetica-Bold", 6)

    for i,h in enumerate(headers):
        c.rect(tx+i*colw, ty, colw, th)
        c.drawCentredString(tx+i*colw+colw/2, ty+4, h)

    vals = [
        "1", d(row["ShipmentDate"]), s(row["LRNo"]), s(row["Destination"]),
        s(row["CNNumber"]), s(row["TruckNo"]), s(row["InvoiceNo"]),
        s(row["Pkgs"]), s(row["WeightKgs"]),
        d(row["DateArrival"]), d(row["DateDelivery"]), s(row["TruckType"]),
        money(row["FreightAmt"]), money(row["ToPointCharges"]),
        money(row["UnloadingCharge"]), money(row["SourceDetention"]),
        money(row["DestinationDetention"]), money(total(row))
    ]

    c.setFont("Helvetica", 6)
    for i,v in enumerate(vals):
        c.rect(tx+i*colw, ty-rh, colw, rh)
        c.drawCentredString(tx+i*colw+colw/2, ty-rh+3, v)

    words = num2words(int(round(total(row))), lang="en").title()+" Rupees Only"
    c.rect(tx, ty-rh-8, tw, 8)
    c.setFont("Helvetica-Bold", 7)
    c.drawString(tx+5, ty-rh-5, f"Total in words (Rs.): {words}")
    c.drawRightString(tx+tw-5, ty-rh-5, money(total(row)))

    bx = tx
    by = BM+20
    bw = 85*mm
    bh = 35*mm
    c.rect(bx, by, bw, bh)
    bank = [
        ("Our PAN No.", row["PANNo"]),
        ("STC GSTIN", row["STCGSTIN"]),
        ("STC State Code", row["STCStateCode"]),
        ("Account name", row["AccountName"]),
        ("Account no", row["AccountNo"]),
        ("IFS Code", row["IFSCode"])
    ]
    h = bh/len(bank)
    c.setFont("Helvetica", 7)
    for i,(k,v) in enumerate(bank):
        c.drawString(bx+5, by+bh-(i+1)*h+3, k)
        c.drawString(bx+45, by+bh-(i+1)*h+3, v)

    sx = W-RM-90
    c.setFont("Helvetica-Bold", 8)
    c.drawString(sx, by+25, "For SOUTH TRANSPORT COMPANY")
    c.line(sx, by+10, sx+120, by+10)
    c.setFont("Helvetica", 7)
    c.drawRightString(sx+120, by+2, "(Authorized Signatory)")

    c.showPage()
    c.save()

@app.route("/", methods=["GET","POST"])
def index():
    if request.method=="POST":
        file = request.files["file"]
        path = os.path.join(UPLOAD_FOLDER, file.filename)
        file.save(path)
        df = pd.read_excel(path)
        df.columns = [c.strip() for c in df.columns]
        for h in REQUIRED_HEADERS:
            if h not in df.columns:
                return f"Missing column: {h}"
        files=[]
        for _,r in df.iterrows():
            name = f'{r["FreightBillNo"]}_{r["LRNo"]}.pdf'
            p = os.path.join(OUTPUT_FOLDER,name)
            generate_invoice_pdf(r.to_dict(),p)
            files.append(p)
        zname = os.path.join(OUTPUT_FOLDER,"invoices.zip")
        with zipfile.ZipFile(zname,"w") as z:
            for f in files:
                z.write(f,os.path.basename(f))
        return send_file(zname,as_attachment=True)
    return render_template("index.html")

if __name__=="__main__":
    app.run(debug=True)
