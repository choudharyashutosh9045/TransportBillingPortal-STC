from flask import Flask, render_template, request, send_file, jsonify
import pandas as pd
import os, zipfile, uuid
from datetime import datetime
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas

app = Flask(__name__)

UPLOAD_DIR = "uploads"
OUTPUT_DIR = "output"
os.makedirs(UPLOAD_DIR, exist_ok=True)
os.makedirs(OUTPUT_DIR, exist_ok=True)

# ---------------- FIXED DATA ----------------
FIXED_PARTY = {
    "PartyName": "South Transport Company",
    "PartyAddress": "Roorkee, Uttarakhand"
}

FIXED_STC_BANK = {
    "BankName": "HDFC Bank",
    "AccountNo": "XXXXXXXXXX",
    "IFSC": "HDFC000XXXX"
}

# ---------------- CORE FIX ----------------
def sanitize_row(row: dict):
    clean = {}
    for k, v in row.items():

        if pd.isna(v):
            clean[k] = ""
            continue

        if isinstance(v, (pd.Timestamp, datetime)):
            clean[k] = v.strftime("%d-%m-%Y")
            continue

        if isinstance(v, (int, float)):
            if isinstance(v, float) and v.is_integer():
                clean[k] = str(int(v))
            else:
                clean[k] = str(v)
            continue

        clean[k] = str(v).strip()

    return clean

# ---------------- PDF GENERATION ----------------
def generate_invoice_pdf(row: dict, pdf_path: str):

    # 🔐 MOST IMPORTANT LINE
    row = sanitize_row(row)

    # merge fixed data
    row = {**FIXED_PARTY, **FIXED_STC_BANK, **row}

    c = canvas.Canvas(pdf_path, pagesize=A4)
    w, h = A4

    y = h - 40

    # -------- PDF FORMAT (UNCHANGED LOGIC) --------
    c.setFont("Helvetica-Bold", 14)
    c.drawString(40, y, "FREIGHT BILL")
    y -= 30

    c.setFont("Helvetica", 9)

    def line(label, value):
        nonlocal y
        c.drawString(40, y, f"{label}:")
        c.drawString(160, y, value)
        y -= 14

    line("Freight Bill No", row.get("FreightBillNo", ""))
    line("Invoice Date", row.get("InvoiceDate", ""))
    line("Due Date", row.get("DueDate", ""))
    line("LR No", row.get("LRNo", ""))
    line("Truck No", row.get("TruckNo", ""))
    line("Invoice No", row.get("InvoiceNo", ""))
    line("From", row.get("FromLocation", ""))
    line("Destination", row.get("Destination", ""))
    line("Truck Type", row.get("TruckType", ""))

    y -= 10
    line("Pkgs", row.get("Pkgs", ""))
    line("Weight (Kgs)", row.get("WeightKgs", ""))

    y -= 10
    line("Freight Amount", row.get("FreightAmt", ""))
    line("To Point Charges", row.get("ToPointCharges", ""))
    line("Unloading Charge", row.get("UnloadingCharge", ""))
    line("Source Detention", row.get("SourceDetention", ""))
    line("Destination Detention", row.get("DestinationDetention", ""))

    c.showPage()
    c.save()

# ---------------- ROUTES ----------------
@app.route("/", methods=["GET", "POST"])
def index():
    if request.method == "POST":
        file = request.files.get("file")
        if not file:
            return "No file", 400

        path = os.path.join(UPLOAD_DIR, file.filename)
        file.save(path)

        df = pd.read_excel(path)

        zip_name = f"bills_{uuid.uuid4().hex}.zip"
        zip_path = os.path.join(OUTPUT_DIR, zip_name)

        with zipfile.ZipFile(zip_path, "w") as z:
            for i, row in df.iterrows():
                row_dict = row.to_dict()
                pdf_name = f"{row_dict.get('FreightBillNo','bill')}_{i}.pdf"
                pdf_path = os.path.join(OUTPUT_DIR, pdf_name)
                generate_invoice_pdf(row_dict, pdf_path)
                z.write(pdf_path, pdf_name)
                os.remove(pdf_path)

        return send_file(zip_path, as_attachment=True)

    return render_template("invoice.html")

@app.route("/preview", methods=["POST"])
def preview():
    file = request.files.get("file")
    if not file:
        return jsonify(ok=False, error="No file")

    df = pd.read_excel(file)
    rows = []

    for _, r in df.head(10).iterrows():
        row = sanitize_row(r.to_dict())
        total = (
            float(r.get("FreightAmt", 0) or 0)
            + float(r.get("ToPointCharges", 0) or 0)
            + float(r.get("UnloadingCharge", 0) or 0)
        )
        row["TotalAmount"] = str(total)
        rows.append(row)

    return jsonify(ok=True, count=len(df), rows=rows)

@app.route("/api/history")
def history():
    return jsonify([])

if __name__ == "__main__":
    app.run(debug=True)
