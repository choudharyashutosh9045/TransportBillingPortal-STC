import os
import pandas as pd
from flask import Flask, render_template, request, send_file
from reportlab.lib.pagesizes import A4, landscape
from reportlab.pdfgen import canvas
from reportlab.lib.units import mm
from reportlab.lib.utils import ImageReader
from datetime import datetime
from num2words import num2words

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
UPLOAD_FOLDER = os.path.join(BASE_DIR, "uploads")
OUTPUT_FOLDER = os.path.join(BASE_DIR, "generated")

os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)

app = Flask(__name__)

# ---------------- FIXED DATA ----------------

FIXED_PARTY = {}
FIXED_STC_BANK = {
    "PANNo": "BSSPG9414K",
    "STCGSTIN": "05BSSPG9414K1ZA",
    "STCStateCode": "5",
    "AccountName": "South Transport Company",
    "AccountNo": "364205500142",
    "IFSCode": "ICIC0003642"
}

# ---------------- HELPERS ----------------

def safe_str(v):
    return "" if pd.isna(v) else str(v)

def format_date(v):
    if pd.isna(v) or v == "":
        return ""
    try:
        return pd.to_datetime(v).strftime("%d %b %Y")
    except:
        return str(v)

def money(v):
    try:
        return f"{float(v):.2f}"
    except:
        return "0.00"

def calc_total(r):
    return sum([
        float(r.get("FreightAmt", 0) or 0),
        float(r.get("ToPointCharges", 0) or 0),
        float(r.get("UnloadingCharge", 0) or 0),
        float(r.get("SourceDetention", 0) or 0),
        float(r.get("DestinationDetention", 0) or 0),
    ])

# ---------------- PDF GENERATOR (UNCHANGED FORMAT) ----------------

def generate_invoice_pdf(row: dict, pdf_path: str):
    row = {**FIXED_PARTY, **FIXED_STC_BANK, **row}

    W, H = landscape(A4)
    c = canvas.Canvas(pdf_path, pagesize=(W, H))

    LM = RM = TM = BM = 10 * mm

    c.setLineWidth(1)
    c.rect(LM, BM, W - LM - RM, H - TM - BM)

    c.setFont("Helvetica-Bold", 14)
    c.drawCentredString(W / 2, H - TM - 8 * mm, "SOUTH TRANSPORT COMPANY")
    c.setFont("Helvetica", 8)
    c.drawCentredString(W / 2, H - TM - 12 * mm, "Dehradun Road Near power Grid Bhagwanpur")
    c.drawCentredString(W / 2, H - TM - 15 * mm, "Roorkee,Haridwar, U.K. 247661, India")
    c.setFont("Helvetica-Bold", 10)
    c.drawCentredString(W / 2, H - TM - 22 * mm, "INVOICE")

    logo_path = os.path.join(BASE_DIR, "logo.png")
    if os.path.exists(logo_path):
        c.drawImage(ImageReader(logo_path),
                    LM + 6 * mm,
                    H - TM - 33 * mm,
                    width=58 * mm,
                    height=28 * mm,
                    mask="auto")

    left_y = H - TM - 62 * mm
    c.rect(LM + 2 * mm, left_y, 110 * mm, 28 * mm)

    c.setFont("Helvetica-Bold", 8)
    c.drawString(LM + 4 * mm, left_y + 20 * mm, "To,")
    c.drawString(LM + 4 * mm, left_y + 15 * mm, safe_str(row["PartyName"]))
    c.setFont("Helvetica", 7.5)
    c.drawString(LM + 4 * mm, left_y + 10 * mm, safe_str(row["PartyAddress"]))
    c.drawString(LM + 4 * mm, left_y + 6 * mm,
                 f"{safe_str(row['PartyCity'])}, {safe_str(row['PartyState'])} {safe_str(row['PartyPincode'])}")
    c.setFont("Helvetica-Bold", 7.5)
    c.drawString(LM + 4 * mm, left_y + 2 * mm, f"GSTIN: {safe_str(row['PartyGSTIN'])}")

    c.setFont("Helvetica", 7.5)
    c.drawString(LM + 4 * mm, left_y - 5 * mm, f"From location: {safe_str(row['FromLocation'])}")

    total = calc_total(row)
    words = num2words(int(total), lang="en").title() + " Rupees Only"

    c.drawString(LM + 4 * mm, BM + 35 * mm, "Total in words (Rs.) :")
    c.drawString(LM + 45 * mm, BM + 35 * mm, words)
    c.drawRightString(W - RM - 4 * mm, BM + 35 * mm, money(total))

    c.showPage()
    c.save()

# ---------------- ROUTES ----------------

@app.route("/", methods=["GET", "POST"])
def index():
    if request.method == "POST":
        f = request.files["file"]
        path = os.path.join(UPLOAD_FOLDER, f.filename)
        f.save(path)

        df = pd.read_excel(path)
        row = df.iloc[0].to_dict()

        pdf_name = f"{row['FreightBillNo']}_{row['LRNo']}.pdf"
        pdf_path = os.path.join(OUTPUT_FOLDER, pdf_name)

        generate_invoice_pdf(row, pdf_path)

        return send_file(pdf_path, as_attachment=True)

    return render_template("index.html")

if __name__ == "__main__":
    app.run(debug=True)
