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
    "ShipmentDate","LRNo","Destination","CNNumber","TruckNo","InvoiceNo",
    "Pkgs","WeightKgs","DateArrival","DateDelivery","TruckType",
    "FreightAmt","ToPointCharges","UnloadingCharge",
    "SourceDetention","DestinationDetention"
]

# ---------------- HELPERS ----------------
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

def money(v):
    return f"{safe_float(v):.2f}"

# ---------------- PDF GENERATOR (UNCHANGED LAYOUT) ----------------
def generate_invoice_pdf(row: dict, pdf_path: str):

    row = {**FIXED_PARTY, **FIXED_STC_BANK, **row}

    W, H = landscape(A4)
    c = canvas.Canvas(pdf_path, pagesize=(W, H))

    LM = 10 * mm
    RM = 10 * mm
    TM = 10 * mm
    BM = 10 * mm

    # Outer Border
    c.setLineWidth(1)
    c.rect(LM, BM, W - LM - RM, H - TM - BM)

    # Header
    c.setFont("Helvetica-Bold", 14)
    c.drawCentredString(W / 2, H - TM - 8 * mm, "SOUTH TRANSPORT COMPANY")

    c.setFont("Helvetica", 8)
    c.drawCentredString(W / 2, H - TM - 12 * mm, "Dehradun Road Near power Grid Bhagwanpur")
    c.drawCentredString(W / 2, H - TM - 15 * mm, "Roorkee,Haridwar, U.K. 247661, India")

    c.setFont("Helvetica-Bold", 10)
    c.drawCentredString(W / 2, H - TM - 22 * mm, "INVOICE")

    # Logo
    logo_path = os.path.join(BASE_DIR, "logo.png")
    if os.path.exists(logo_path):
        img = ImageReader(logo_path)
        c.drawImage(img, LM + 6 * mm, H - TM - 36 * mm, width=75 * mm, height=38 * mm, mask="auto")

    # Left box
    left_x = LM + 2 * mm
    left_y = H - TM - 62 * mm
    left_w = 110 * mm
    left_h = 28 * mm
    c.rect(left_x, left_y, left_w, left_h)

    c.setFont("Helvetica-Bold", 8)
    c.drawString(left_x + 2 * mm, left_y + left_h - 6 * mm, "To,")

    c.drawString(left_x + 2 * mm, left_y + left_h - 11 * mm, safe_str(row["PartyName"]))

    c.setFont("Helvetica", 7.5)
    c.drawString(left_x + 2 * mm, left_y + left_h - 15 * mm, safe_str(row["PartyAddress"]))
    c.drawString(left_x + 2 * mm, left_y + left_h - 19 * mm,
                 f"{row['PartyCity']}, {row['PartyState']} {row['PartyPincode']}")

    c.setFont("Helvetica-Bold", 7.5)
    c.drawString(left_x + 2 * mm, left_y + left_h - 23 * mm,
                 f"GSTIN: {row['PartyGSTIN']}")

    c.setFont("Helvetica", 7.5)
    c.drawString(left_x + 2 * mm, left_y - 5 * mm,
                 f"From location: {safe_str(row.get('FromLocation'))}")

    # Right box
    rb_w = 85 * mm
    rb_h = 28 * mm
    rb_x = W - RM - rb_w - 2 * mm
    rb_y = left_y
    c.rect(rb_x, rb_y, rb_w, rb_h)

    c.setFont("Helvetica-Bold", 8)
    c.drawString(rb_x + 4 * mm, rb_y + rb_h - 8 * mm,
                 f"Freight Bill No: {safe_str(row.get('FreightBillNo'))}")
    c.drawString(rb_x + 4 * mm, rb_y + rb_h - 14 * mm,
                 f"Invoice Date: {safe_str(row.get('InvoiceDate'))}")
    c.drawString(rb_x + 4 * mm, rb_y + rb_h - 20 * mm,
                 f"Due Date: {safe_str(row.get('DueDate'))}")

    # ---------------- TABLE ----------------
    table_x = LM + 2 * mm
    table_top = left_y - 18 * mm
    table_w = (W - LM - RM) - 4 * mm
    header_h = 12 * mm
    row_h = 10 * mm

    cols = [
        ("S.\nno.", 10), ("Shipment\nDate", 20), ("LR No.", 14),
        ("Destination", 22), ("CN\nNumber", 18), ("Truck No", 18),
        ("Invoice No", 18), ("Pkgs", 10), ("Weight\n(kgs)", 14),
        ("Date of\nArrival", 16), ("Date of\nDelivery", 16),
        ("Truck\nType", 14), ("Freight\nAmt", 16),
        ("To Point\nCharges", 16), ("Unloading\nCharge", 16),
        ("Source\nDetention", 16), ("Destination\nDetention", 16),
        ("Total\nAmount", 18),
    ]

    scale = table_w / sum(w for _, w in cols)
    cols = [(n, w * scale) for n, w in cols]

    header_bottom = table_top - header_h
    c.rect(table_x, header_bottom, table_w, header_h)

    c.setFont("Helvetica-Bold", 6.5)
    x = table_x
    for name, w in cols:
        c.line(x, header_bottom, x, table_top)
        cx = x + w / 2
        yy = table_top - 4 * mm
        for p in name.split("\n"):
            c.drawCentredString(cx, yy, p)
            yy -= 3 * mm
        x += w
    c.line(table_x + table_w, header_bottom, table_x + table_w, table_top)

    data_bottom = header_bottom - row_h
    c.rect(table_x, data_bottom, table_w, row_h)

    total_amt = (
        safe_float(row.get("FreightAmt")) +
        safe_float(row.get("ToPointCharges")) +
        safe_float(row.get("UnloadingCharge")) +
        safe_float(row.get("SourceDetention")) +
        safe_float(row.get("DestinationDetention"))
    )

    data = [
        "1",
        safe_str(row.get("ShipmentDate")),
        safe_str(row.get("LRNo")),
        safe_str(row.get("Destination")),
        safe_str(row.get("CNNumber")),
        safe_str(row.get("TruckNo")),
        safe_str(row.get("InvoiceNo")),
        safe_str(row.get("Pkgs")),
        safe_str(row.get("WeightKgs")),
        safe_str(row.get("DateArrival")),
        safe_str(row.get("DateDelivery")),
        safe_str(row.get("TruckType")),
        money(row.get("FreightAmt")),
        money(row.get("ToPointCharges")),
        money(row.get("UnloadingCharge")),
        money(row.get("SourceDetention")),
        money(row.get("DestinationDetention")),
        money(total_amt),
    ]

    c.setFont("Helvetica", 7)
    x = table_x
    for (_, w), txt in zip(cols, data):
        c.line(x, data_bottom, x, header_bottom)
        c.drawCentredString(x + w / 2, data_bottom + 3.5 * mm, safe_str(txt))
        x += w
    c.line(table_x + table_w, data_bottom, table_x + table_w, header_bottom)

    # Words
    try:
        words = num2words(int(total_amt), lang="en").title() + " Rupees Only"
    except:
        words = ""

    c.setFont("Helvetica", 7)
    c.drawString(table_x + 4 * mm, data_bottom - 5 * mm,
                 f"Total in words (Rs.): {words}")

    c.showPage()
    c.save()

# ---------------- ROUTE ----------------
@app.route("/", methods=["GET", "POST"])
def index():
    if request.method == "POST":
        file = request.files.get("file")
        if not file:
            return "No file uploaded", 400

        path = os.path.join(UPLOAD_FOLDER, file.filename)
        file.save(path)

        df = pd.read_excel(path, dtype=str).fillna("")
        df.columns = [c.strip() for c in df.columns]

        missing = [h for h in REQUIRED_HEADERS if h not in df.columns]
        if missing:
            return f"Missing columns: {missing}", 400

        pdfs = []
        for _, r in df.iterrows():
            row = r.to_dict()
            name = f"{row.get('FreightBillNo','BILL')}_{datetime.now().strftime('%H%M%S')}.pdf"
            out = os.path.join(OUTPUT_FOLDER, name)
            generate_invoice_pdf(row, out)
            pdfs.append(out)

        zip_name = f"BILLS_{datetime.now().strftime('%Y%m%d_%H%M%S')}.zip"
        zip_path = os.path.join(OUTPUT_FOLDER, zip_name)

        with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as z:
            for p in pdfs:
                z.write(p, os.path.basename(p))

        return send_file(zip_path, as_attachment=True)

    return render_template("index.html")


if __name__ == "__main__":
    app.run(debug=True)
