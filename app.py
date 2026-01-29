from flask import Flask, render_template, request, send_file
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
from reportlab.lib.units import mm
from reportlab.lib.utils import ImageReader
from datetime import datetime
import os

app = Flask(__name__)

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
OUTPUT_DIR = os.path.join(BASE_DIR, "output")n
if not os.path.exists(OUTPUT_DIR):
    os.makedirs(OUTPUT_DIR)

@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        data = request.form.to_dict()
        filename = generate_invoice_pdf(data)
        return send_file(filename, as_attachment=True)
    return render_template('index.html')

def generate_invoice_pdf(row: dict, pdf_path: str):
    row = {**FIXED_PARTY, **FIXED_STC_BANK, **row}

    W, H = landscape(A4)
    c = canvas.Canvas(pdf_path, pagesize=(W, H))

    LM = 10 * mm
    RM = 10 * mm
    TM = 10 * mm
    BM = 10 * mm

    c.setLineWidth(1)
    c.rect(LM, BM, W - LM - RM, H - TM - BM)

    c.setFont("Helvetica-Bold", 14)
    c.drawCentredString(W / 2, H - TM - 8 * mm, "SOUTH TRANSPORT COMPANY")

    c.setFont("Helvetica", 8)
    c.drawCentredString(W / 2, H - TM - 12 * mm, "Dehradun Road Near power Grid Bhagwanpur")
    c.drawCentredString(W / 2, H - TM - 15 * mm, "Roorkee, Haridwar, U.K. 247661, India")

    c.setFont("Helvetica-Bold", 10)
    c.drawCentredString(W / 2, H - TM - 22 * mm, "INVOICE")

    logo_path = os.path.join(BASE_DIR, "logo.png")
    if os.path.exists(logo_path):
        img = ImageReader(logo_path)
        c.drawImage(img, LM + 6 * mm, H - TM - 36 * mm, 75 * mm, 38 * mm, mask="auto")

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
    c.drawString(
        left_x + 2 * mm,
        left_y + left_h - 19 * mm,
        f"{row['PartyCity']}, {row['PartyState']} {row['PartyPincode']}"
    )

    c.setFont("Helvetica-Bold", 7.5)
    c.drawString(left_x + 2 * mm, left_y + left_h - 23 * mm, f"GSTIN: {row['PartyGSTIN']}")

    c.setFont("Helvetica", 7.5)
    c.drawString(left_x + 2 * mm, left_y - 5 * mm, f"From location: {row.get('FromLocation','')}")

    right_w = 85 * mm
    right_h = 28 * mm
    right_x = W - RM - right_w - 2 * mm
    right_y = left_y

    c.rect(right_x, right_y, right_w, right_h)
    c.setFont("Helvetica-Bold", 8)
    c.drawString(right_x + 4 * mm, right_y + right_h - 8 * mm, f"Freight Bill No: {row['FreightBillNo']}")
    c.drawString(right_x + 4 * mm, right_y + right_h - 14 * mm, f"Invoice Date: {format_date(row['InvoiceDate'])}")
    c.drawString(right_x + 4 * mm, right_y + right_h - 20 * mm, f"Due Date: {format_date(row['DueDate'])}")

    table_x = LM + 2 * mm
    table_top = left_y - 18 * mm
    table_w = W - LM - RM - 4 * mm

    header_h = 12 * mm
    row_h = 10 * mm
    words_h = 7 * mm

    cols = [
        ("S.\nno.", 10), ("Shipment\nDate", 20), ("LR No.", 14), ("Destination", 22),
        ("CN\nNumber", 18), ("Truck No", 18), ("Invoice No", 18), ("Pkgs", 10),
        ("Weight\n(kgs)", 14), ("Date of\nArrival", 16), ("Date of\nDelivery", 16),
        ("Truck\nType", 14), ("Freight\nAmt (Rs.)", 16), ("To Point\nCharges", 16),
        ("Unloading\nCharge", 16), ("Source\nDetention", 16),
        ("Destination\nDetention", 16), ("Total\nAmount", 18)
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

    data_top = header_bottom
    data_bottom = data_top - row_h
    c.rect(table_x, data_bottom, table_w, row_h)

    total = calc_total(row)

    data = [
        "1", format_date(row["ShipmentDate"]), row["LRNo"], row["Destination"],
        row["CNNumber"], row["TruckNo"], row["InvoiceNo"], row["Pkgs"],
        row["WeightKgs"], format_date(row["DateArrival"]),
        format_date(row["DateDelivery"]), row["TruckType"],
        money(row["FreightAmt"]), money(row["ToPointCharges"]),
        money(row["UnloadingCharge"]), money(row["SourceDetention"]),
        money(row["DestinationDetention"]), money(total)
    ]

    c.setFont("Helvetica", 7)
    x = table_x
    for (_, w), txt in zip(cols, data):
        c.line(x, data_bottom, x, data_top)
        c.drawCentredString(x + w / 2, data_bottom + 3.5 * mm, safe_str(txt))
        x += w
    c.line(table_x + table_w, data_bottom, table_x + table_w, data_top)

    words_bottom = data_bottom - words_h
    c.rect(table_x, words_bottom, table_w, words_h)

    words = num2words(int(round(total)), lang="en").title() + " Rupees Only"
    c.setFont("Helvetica-Bold", 7)
    c.drawString(table_x + 4 * mm, words_bottom + 2.2 * mm, "Total in words (Rs.) :")
    c.setFont("Helvetica", 7)
    c.drawString(table_x + 32 * mm, words_bottom + 2.2 * mm, words)
    c.drawRightString(table_x + table_w - 2 * mm, words_bottom + 2.2 * mm, money(total))

    c.showPage()
    c.save()

    return file_path

if __name__ == '__main__':
    app.run(debug=True)
