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
TEMPLATE_FOLDER = os.path.join(BASE_DIR, "templates")
DATA_FOLDER = os.path.join(BASE_DIR, "data")

os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)
os.makedirs(DATA_FOLDER, exist_ok=True)

app.config["UPLOAD_FOLDER"] = UPLOAD_FOLDER


# ==========================
# ✅ FIXED DETAILS (SAME ALWAYS)
# ==========================
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

# Excel headers required
REQUIRED_HEADERS = [
    "FreightBillNo","InvoiceDate","DueDate","FromLocation",
    "ShipmentDate","LRNo","Destination","CNNumber","TruckNo","InvoiceNo",
    "Pkgs","WeightKgs","DateArrival","DateDelivery","TruckType",
    "FreightAmt","ToPointCharges","UnloadingCharge","SourceDetention","DestinationDetention"
]


# ---------------- HELPERS ----------------
def safe_str(v):
    if pd.isna(v):
        return ""
    return str(v).strip()

def safe_float(v):
    try:
        if pd.isna(v):
            return 0.0
        if str(v).strip() == "":
            return 0.0
        return float(v)
    except:
        return 0.0

def format_date(v):
    s = safe_str(v)
    if not s:
        return ""
    try:
        dt = pd.to_datetime(s, dayfirst=True, errors="coerce")
        if pd.isna(dt):
            return s
        return dt.strftime("%d %b %Y")
    except:
        return s

def money(v):
    return f"{safe_float(v):.2f}"

def calc_total(row):
    return (
        safe_float(row.get("FreightAmt")) +
        safe_float(row.get("ToPointCharges")) +
        safe_float(row.get("UnloadingCharge")) +
        safe_float(row.get("SourceDetention")) +
        safe_float(row.get("DestinationDetention"))
    )

def now_ts():
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


# ==========================
# ✅ DATABASE (PostgreSQL Render)
# ==========================
def get_db_conn():
    db_url = os.environ.get("DATABASE_URL")
    if not db_url:
        return None
    return psycopg2.connect(db_url, sslmode="require")

def init_db():
    conn = get_db_conn()
    if not conn:
        print("⚠️ DATABASE_URL not set. DB disabled.")
        return

    with conn.cursor() as cur:
        cur.execute("""
            CREATE TABLE IF NOT EXISTS bill_history (
                id SERIAL PRIMARY KEY,
                created_at TIMESTAMP DEFAULT NOW(),
                source_excel VARCHAR(255),
                bill_no VARCHAR(100),
                lr_no VARCHAR(100),
                invoice_date VARCHAR(50),
                due_date VARCHAR(50),
                destination VARCHAR(200),
                total_amount NUMERIC(12,2),
                zip_name VARCHAR(255)
            );
        """)
        conn.commit()

    conn.close()
    print("✅ DB ready: bill_history table created/checked.")


def add_history_entry_db(source_excel, df, zip_name):
    conn = get_db_conn()
    if not conn:
        return

    with conn.cursor() as cur:
        for _, r in df.iterrows():
            row = r.to_dict()
            bill_no = safe_str(row.get("FreightBillNo"))
            lr_no = safe_str(row.get("LRNo"))
            invoice_date = format_date(row.get("InvoiceDate"))
            due_date = format_date(row.get("DueDate"))
            destination = safe_str(row.get("Destination"))
            total_amount = float(calc_total(row))

            cur.execute("""
                INSERT INTO bill_history
                (source_excel, bill_no, lr_no, invoice_date, due_date, destination, total_amount, zip_name)
                VALUES (%s,%s,%s,%s,%s,%s,%s,%s)
            """, (source_excel, bill_no, lr_no, invoice_date, due_date, destination, total_amount, zip_name))

        conn.commit()

    conn.close()


def get_history_db(limit=10):
    conn = get_db_conn()
    if not conn:
        return []

    with conn.cursor(cursor_factory=RealDictCursor) as cur:
        cur.execute("""
            SELECT id, created_at, source_excel, bill_no, lr_no, invoice_date, due_date, destination, total_amount, zip_name
            FROM bill_history
            ORDER BY created_at DESC
            LIMIT %s
        """, (limit,))
        rows = cur.fetchall()

    conn.close()

    # convert datetime to string for JSON
    for r in rows:
        if r.get("created_at"):
            r["created_at"] = r["created_at"].strftime("%d %b %Y %I:%M %p")
    return rows


# ---------------- PDF GENERATOR ----------------
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

    # Logo (bigger + up)
    logo_path = os.path.join(BASE_DIR, "logo.png")
    if os.path.exists(logo_path):
        try:
            img = ImageReader(logo_path)
            logo_w = 58 * mm
            logo_h = 28 * mm
            logo_x = LM + 6 * mm
            logo_y = H - TM - 40 * mm
            c.drawImage(img, logo_x, logo_y, width=logo_w, height=logo_h, mask="auto", preserveAspectRatio=True)
        except Exception as e:
            print("LOGO ERROR:", e)

    # Left box (To)
    left_box_x = LM + 2 * mm
    left_box_y = H - TM - 62 * mm
    left_box_w = 110 * mm
    left_box_h = 28 * mm
    c.rect(left_box_x, left_box_y, left_box_w, left_box_h)

    c.setFont("Helvetica-Bold", 8)
    c.drawString(left_box_x + 2 * mm, left_box_y + left_box_h - 6 * mm, "To,")

    c.setFont("Helvetica-Bold", 8)
    c.drawString(left_box_x + 2 * mm, left_box_y + left_box_h - 11 * mm, safe_str(row.get("PartyName")))

    c.setFont("Helvetica", 7.5)
    c.drawString(left_box_x + 2 * mm, left_box_y + left_box_h - 15 * mm, safe_str(row.get("PartyAddress")))
    c.drawString(left_box_x + 2 * mm, left_box_y + left_box_h - 19 * mm,
                 f"{safe_str(row.get('PartyCity'))}, {safe_str(row.get('PartyState'))} {safe_str(row.get('PartyPincode'))}".strip(", "))

    c.setFont("Helvetica-Bold", 7.5)
    c.drawString(left_box_x + 2 * mm, left_box_y + left_box_h - 23 * mm, f"GSTIN: {safe_str(row.get('PartyGSTIN'))}")

    # From location
    c.setFont("Helvetica", 7.5)
    c.drawString(left_box_x + 2 * mm, left_box_y - 5 * mm, f"From location: {safe_str(row.get('FromLocation'))}")

    # Right box (bill details)
    right_box_w = 85 * mm
    right_box_h = 28 * mm
    right_box_x = W - RM - right_box_w - 2 * mm
    right_box_y = left_box_y
    c.rect(right_box_x, right_box_y, right_box_w, right_box_h)

    c.setFont("Helvetica-Bold", 8)
    c.drawString(right_box_x + 4 * mm, right_box_y + right_box_h - 8 * mm, f"Freight Bill No:  {safe_str(row.get('FreightBillNo'))}")
    c.drawString(right_box_x + 4 * mm, right_box_y + right_box_h - 14 * mm, f"Invoice Date:      {format_date(row.get('InvoiceDate'))}")
    c.drawString(right_box_x + 4 * mm, right_box_y + right_box_h - 20 * mm, f"Due Date:          {format_date(row.get('DueDate'))}")

    # Table
    table_x = LM + 2 * mm
    table_top = left_box_y - 18 * mm
    table_w = (W - LM - RM) - 4 * mm

    header_h = 12 * mm
    row_h = 10 * mm
    words_h = 7 * mm

    cols = [
        ("S.\nno.", 10),
        ("Shipment\nDate", 20),
        ("LR No.", 14),
        ("Destination", 22),
        ("CN\nNumber", 18),
        ("Truck No", 18),
        ("Invoice No", 18),
        ("Pkgs", 10),
        ("Weight\n(kgs)", 14),
        ("Date of\nArrival", 16),
        ("Date of\nDelivery", 16),
        ("Truck\nType", 14),
        ("Freight\nAmt (Rs.)", 16),
        ("To Point\nCharges(Rs.)", 16),
        ("Unloading\nCharge (Rs.)", 16),
        ("Source\nDetention\n(Rs.)", 16),
        ("Destination\nDetention\n(Rs.)", 16),
        ("Total\nAmount (Rs.)", 18),
    ]

    total_units = sum(w for _, w in cols)
    scale = table_w / total_units
    cols = [(n, w * scale) for n, w in cols]

    # Header box
    header_bottom = table_top - header_h
    c.rect(table_x, header_bottom, table_w, header_h)

    c.setFont("Helvetica-Bold", 6.5)
    x = table_x
    for name, wcol in cols:
        c.line(x, header_bottom, x, table_top)
        cx = x + wcol / 2
        parts = name.split("\n")
        yy = table_top - 4 * mm
        for p in parts:
            c.drawCentredString(cx, yy, p)
            yy -= 3 * mm
        x += wcol
    c.line(table_x + table_w, header_bottom, table_x + table_w, table_top)

    # Data row
    data_top = header_bottom
    data_bottom = data_top - row_h
    c.rect(table_x, data_bottom, table_w, row_h)

    total_amt = calc_total(row)

    data_list = [
        "1",
        format_date(row.get("ShipmentDate")),
        safe_str(row.get("LRNo")),
        safe_str(row.get("Destination")),
        safe_str(row.get("CNNumber")),
        safe_str(row.get("TruckNo")),
        safe_str(row.get("InvoiceNo")),
        safe_str(row.get("Pkgs")),
        safe_str(row.get("WeightKgs")),
        format_date(row.get("DateArrival")),
        format_date(row.get("DateDelivery")),
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
    for (name, wcol), txt in zip(cols, data_list):
        c.line(x, data_bottom, x, data_top)
        c.drawCentredString(x + wcol / 2, data_bottom + 3.5 * mm, safe_str(txt))
        x += wcol
    c.line(table_x + table_w, data_bottom, table_x + table_w, data_top)

    # Total in words row
    words_top = data_bottom
    words_bottom = words_top - words_h
    c.rect(table_x, words_bottom, table_w, words_h)

    words = num2words(int(round(total_amt)), lang="en").title() + " Rupees Only"
    c.setFont("Helvetica-Bold", 7)
    c.drawString(table_x + 4 * mm, words_bottom + 2.2 * mm, "Total in words (Rs.) :")
    c.setFont("Helvetica", 7)
    c.drawString(table_x + 32 * mm, words_bottom + 2.2 * mm, words)
    c.setFont("Helvetica-Bold", 7)
    c.drawRightString(table_x + table_w - 2 * mm, words_bottom + 2.2 * mm, money(total_amt))

    # Note
    note_y = words_bottom - 10 * mm
    c.setFont("Helvetica", 7)
    c.drawString(
        table_x,
        note_y,
        'Any changes or discrepancies should be highlighted within 5 working days else it would be considered final. '
        'Please send all remittance details to "southtptrk@gmail.com".'
    )

    # Bank table bottom left
    bank_x = table_x
    bank_y = BM + 8 * mm
    bank_w = 85 * mm
    bank_h = 32 * mm
    c.rect(bank_x, bank_y, bank_w, bank_h)

    bank_rows = [
        ("Our PAN No.", safe_str(row.get("PANNo"))),
        ("STC GSTIN", safe_str(row.get("STCGSTIN"))),
        ("STC State Code", safe_str(row.get("STCStateCode"))),
        ("Account name", safe_str(row.get("AccountName"))),
        ("Account no", safe_str(row.get("AccountNo"))),
        ("IFS Code", safe_str(row.get("IFSCode"))),
    ]

    r_h = bank_h / len(bank_rows)
    c.setFont("Helvetica-Bold", 7)
    for i, (k, v) in enumerate(bank_rows):
        y1 = bank_y + bank_h - (i + 1) * r_h
        c.line(bank_x, y1, bank_x + bank_w, y1)
        c.drawString(bank_x + 2 * mm, y1 + 2 * mm, k)
        c.drawString(bank_x + 32 * mm, y1 + 2 * mm, v)

    c.line(bank_x + 30 * mm, bank_y, bank_x + 30 * mm, bank_y + bank_h)

    # Signatory right
    sign_x = W - RM - 80 * mm
    sign_y = BM + 10 * mm
    c.setFont("Helvetica-Bold", 8)
    c.drawString(sign_x, sign_y + 18 * mm, "For SOUTH TRANSPORT COMPANY")
    c.setLineWidth(0.8)
    c.line(sign_x + 10 * mm, sign_y + 8 * mm, sign_x + 70 * mm, sign_y + 8 * mm)
    c.setFont("Helvetica", 7)
    c.drawRightString(sign_x + 75 * mm, sign_y + 2 * mm, "(Authorized Signatory)")

    c.showPage()
    c.save()


# ---------------- ROUTES ----------------
@app.route("/", methods=["GET", "POST"])
def index():
    if request.method == "POST":
        file = request.files.get("file")
        if not file:
            return "No file uploaded", 400

        filepath = os.path.join(app.config["UPLOAD_FOLDER"], file.filename)
        file.save(filepath)

        df = pd.read_excel(filepath)
        df.columns = [str(c).strip() for c in df.columns]

        # validate headers
        missing = [h for h in REQUIRED_HEADERS if h not in df.columns]
        if missing:
            return f"Missing columns in Excel: {missing}", 400

        generated = []
        for _, r in df.iterrows():
            row = r.to_dict()

            bill_no = safe_str(row.get("FreightBillNo", "BILL"))
            lr_no = safe_str(row.get("LRNo", "LR"))

            ts = datetime.now().strftime("%H%M%S")
            pdf_name = f"{bill_no}_LR{lr_no}_{ts}.pdf"
            pdf_path = os.path.join(OUTPUT_FOLDER, pdf_name)

            generate_invoice_pdf(row, pdf_path)
            generated.append(pdf_path)

        zip_name = f"Bills_{datetime.now().strftime('%Y%m%d_%H%M%S')}.zip"
        zip_path = os.path.join(OUTPUT_FOLDER, zip_name)

        with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as zf:
            for p in generated:
                zf.write(p, arcname=os.path.basename(p))

        # ✅ Save history to DB
        add_history_entry_db(file.filename, df, zip_name)

        return send_file(zip_path, as_attachment=True)

    return render_template("index.html")


@app.route("/api/history", methods=["GET"])
def api_history():
    return jsonify(get_history_db(10))


@app.route("/download-template", methods=["GET"])
def download_template():
    template_path = os.path.join(OUTPUT_FOLDER, "Excel_Template_STC.xlsx")
    df = pd.DataFrame(columns=REQUIRED_HEADERS)
    df.to_excel(template_path, index=False)
    return send_file(template_path, as_attachment=True)


@app.route("/preview", methods=["POST"])
def preview():
    file = request.files.get("file")
    if not file:
        return jsonify({"ok": False, "error": "No file uploaded"}), 400

    filepath = os.path.join(app.config["UPLOAD_FOLDER"], "preview_" + file.filename)
    file.save(filepath)

    df = pd.read_excel(filepath)
    df.columns = [str(c).strip() for c in df.columns]

    missing = [h for h in REQUIRED_HEADERS if h not in df.columns]
    if missing:
        return jsonify({"ok": False, "error": f"Missing columns: {missing}"}), 400

    df2 = df.head(10).copy()

    df2["TotalAmount"] = (
        df2["FreightAmt"].fillna(0).astype(float) +
        df2["ToPointCharges"].fillna(0).astype(float) +
        df2["UnloadingCharge"].fillna(0).astype(float) +
        df2["SourceDetention"].fillna(0).astype(float) +
        df2["DestinationDetention"].fillna(0).astype(float)
    )

    return jsonify({
        "ok": True,
        "rows": df2.fillna("").to_dict(orient="records"),
        "count": len(df)
    })


# ✅ init DB on start
init_db()

if __name__ == "__main__":
    print("RUNNING APP VERSION: PORTAL-UI + PREVIEW + HISTORY + TEMPLATE + DB")
    app.run(debug=True)
