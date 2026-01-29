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

app.config["UPLOAD_FOLDER"] = UPLOAD_FOLDER

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

REQUIRED_HEADERS = [
    "FreightBillNo","InvoiceDate","DueDate","FromLocation","ShipmentDate",
    "LRNo","Destination","CNNumber","TruckNo","InvoiceNo","Pkgs","WeightKgs",
    "DateArrival","DateDelivery","TruckType","FreightAmt","ToPointCharges",
    "UnloadingCharge","SourceDetention","DestinationDetention"
]

def s(v):
    if pd.isna(v):
        return ""
    return str(v).strip()

def f(v):
    try:
        if pd.isna(v) or str(v).strip() == "":
            return 0.0
        return float(v)
    except:
        return 0.0

def d(v):
    try:
        dt = pd.to_datetime(v, dayfirst=True, errors="coerce")
        if pd.isna(dt):
            return s(v)
        return dt.strftime("%d %b %Y")
    except:
        return s(v)

def total(row):
    return (
        f(row.get("FreightAmt")) +
        f(row.get("ToPointCharges")) +
        f(row.get("UnloadingCharge")) +
        f(row.get("SourceDetention")) +
        f(row.get("DestinationDetention"))
    )

def generate_pdf(row, pdf_path):
    os.makedirs(os.path.dirname(pdf_path), exist_ok=True)

    row = {**FIXED_PARTY, **FIXED_BANK, **row}

    W, H = landscape(A4)
    c = canvas.Canvas(pdf_path, pagesize=(W, H))

    LM = RM = TM = BM = 10 * mm

    c.setLineWidth(1)
    c.rect(LM, BM, W-LM-RM, H-TM-BM)

    c.setFont("Helvetica-Bold", 14)
    c.drawCentredString(W/2, H-18, "SOUTH TRANSPORT COMPANY")
    c.setFont("Helvetica", 8)
    c.drawCentredString(W/2, H-32, "Dehradun Road Near power Grid Bhagwanpur")
    c.drawCentredString(W/2, H-44, "Roorkee,Haridwar, U.K. 247661, India")
    c.setFont("Helvetica-Bold", 10)
    c.drawCentredString(W/2, H-60, "INVOICE")

    logo = os.path.join(BASE_DIR, "logo.png")
    if os.path.exists(logo):
        c.drawImage(ImageReader(logo), LM+6, H-120, 75*mm, 38*mm, mask="auto")

    c.rect(LM+6, H-210, 110*mm, 55*mm)
    c.setFont("Helvetica-Bold", 8)
    c.drawString(LM+10, H-180, "To,")
    c.drawString(LM+10, H-195, row["PartyName"])
    c.setFont("Helvetica", 8)
    c.drawString(LM+10, H-210+20, row["PartyAddress"])
    c.drawString(LM+10, H-210+10, f'{row["PartyCity"]}, {row["PartyState"]} {row["PartyPincode"]}')
    c.setFont("Helvetica-Bold", 8)
    c.drawString(LM+10, H-210-5, f'GSTIN: {row["PartyGSTIN"]}')

    c.setFont("Helvetica", 8)
    c.drawString(LM+10, H-235, f'From location: {s(row.get("FromLocation"))}')

    rx = W-RM-95*mm
    c.rect(rx, H-210, 90*mm, 55*mm)
    c.setFont("Helvetica-Bold", 8)
    c.drawString(rx+10, H-180, f'Freight Bill No: {s(row.get("FreightBillNo"))}')
    c.drawString(rx+10, H-195, f'Invoice Date: {d(row.get("InvoiceDate"))}')
    c.drawString(rx+10, H-210+10, f'Due Date: {d(row.get("DueDate"))}')

    tx = LM+6
    ty = H-280
    th = 26
    cols = [
        ("S.\nno.",20),("Shipment\nDate",50),("LR No.",35),("Destination",60),
        ("CN\nNumber",45),("Truck No",45),("Invoice No",45),("Pkgs",25),
        ("Weight\n(kgs)",35),("Date of\nArrival",40),("Date of\nDelivery",40),
        ("Truck\nType",35),("Freight\nAmt",40),("To Point\nCharges",40),
        ("Unloading\nCharge",40),("Source\nDetention",40),
        ("Destination\nDetention",40),("Total\nAmount",45)
    ]

    tw = sum(w for _,w in cols)
    scale = (W-LM-RM-12)/tw
    cols = [(n,w*scale) for n,w in cols]

    x = tx
    c.rect(tx, ty-th, sum(w for _,w in cols), th)
    c.setFont("Helvetica-Bold",7)
    for n,w in cols:
        c.line(x, ty-th, x, ty)
        yy = ty-10
        for p in n.split("\n"):
            c.drawCentredString(x+w/2, yy, p)
            yy -= 9
        x += w
    c.line(x, ty-th, x, ty)

    row_y = ty-th-24
    c.rect(tx, row_y, sum(w for _,w in cols), 24)

    data = [
        "1", d(row["ShipmentDate"]), s(row["LRNo"]), s(row["Destination"]),
        s(row["CNNumber"]), s(row["TruckNo"]), s(row["InvoiceNo"]),
        s(row["Pkgs"]), s(row["WeightKgs"]), d(row["DateArrival"]),
        d(row["DateDelivery"]), s(row["TruckType"]),
        f(row["FreightAmt"]), f(row["ToPointCharges"]),
        f(row["UnloadingCharge"]), f(row["SourceDetention"]),
        f(row["DestinationDetention"]), total(row)
    ]

    x = tx
    c.setFont("Helvetica",7)
    for (n,w),v in zip(cols,data):
        c.line(x, row_y, x, row_y+24)
        c.drawCentredString(x+w/2, row_y+7, str(v))
        x += w
    c.line(x, row_y, x, row_y+24)

    words = num2words(int(round(total(row))), lang="en").title()+" Rupees Only"
    c.rect(tx, row_y-20, sum(w for _,w in cols), 20)
    c.drawString(tx+5, row_y-14, f"Total in words (Rs.) : {words}")
    c.drawRightString(tx+sum(w for _,w in cols)-5, row_y-14, f"{total(row):.2f}")

    c.rect(tx, BM+50, 85*mm, 70)
    bank = [
        ("Our PAN No.",row["PANNo"]),("STC GSTIN",row["STCGSTIN"]),
        ("STC State Code",row["STCStateCode"]),("Account name",row["AccountName"]),
        ("Account no",row["AccountNo"]),("IFS Code",row["IFSCode"])
    ]
    y = BM+50+70
    for k,v in bank:
        y -= 12
        c.line(tx, y, tx+85*mm, y)
        c.drawString(tx+5, y+3, k)
        c.drawString(tx+120, y+3, v)

    c.drawString(W-260, BM+90, "For SOUTH TRANSPORT COMPANY")
    c.line(W-260, BM+60, W-120, BM+60)
    c.drawString(W-190, BM+45, "(Authorized Signatory)")

    c.showPage()
    c.save()

@app.route("/", methods=["GET","POST"])
def index():
    if request.method == "POST":
        file = request.files.get("file")
        if not file:
            return "No file",400

        path = os.path.join(UPLOAD_FOLDER, file.filename)
        file.save(path)

        df = pd.read_excel(path)
        df.columns = [c.strip() for c in df.columns]

        for h in REQUIRED_HEADERS:
            if h not in df.columns:
                return f"Missing column {h}",400

        pdfs = []

        for _,r in df.iterrows():
            row = r.to_dict()
            bill = s(row["FreightBillNo"])
            ts = datetime.now().strftime("%H%M%S")
            pdf_path = os.path.join(OUTPUT_FOLDER, bill, f"{bill}_{ts}.pdf")
            generate_pdf(row, pdf_path)
            pdfs.append(pdf_path)

        zip_name = f"PDF_{datetime.now().strftime('%Y%m%d_%H%M%S')}.zip"
        zip_path = os.path.join(OUTPUT_FOLDER, zip_name)

        with zipfile.ZipFile(zip_path,"w",zipfile.ZIP_DEFLATED) as z:
            for p in pdfs:
                z.write(p, os.path.basename(p))

        return send_file(zip_path, as_attachment=True)

    return render_template("index.html")

if __name__ == "__main__":
    app.run(debug=True)
