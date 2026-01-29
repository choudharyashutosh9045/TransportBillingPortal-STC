from flask import Flask, render_template, request, send_file
import pandas as pd
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
from reportlab.lib.units import mm
from num2words import num2words
from datetime import datetime
import os

app = Flask(__name__)
UPLOAD_FOLDER = "uploads"
OUTPUT_FOLDER = "output"

os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)

@app.route("/", methods=["GET", "POST"])
def index():
    if request.method == "POST":
        file = request.files["excel"]
        filepath = os.path.join(UPLOAD_FOLDER, file.filename)
        file.save(filepath)

        df = pd.read_excel(filepath)
        pdf_path = generate_invoice_pdf(df)

        return send_file(pdf_path, as_attachment=True)

    return render_template("index.html")


def generate_invoice_pdf(df):
    invoice_no = df.iloc[0]["FreightBillNo"]
    pdf_file = f"{OUTPUT_FOLDER}/{invoice_no}.pdf"

    c = canvas.Canvas(pdf_file, pagesize=A4)
    width, height = A4

    # ===== HEADER =====
    c.setFont("Helvetica-Bold", 14)
    c.drawCentredString(width / 2, height - 30, "SOUTH TRANSPORT COMPANY")

    c.setFont("Helvetica", 9)
    c.drawCentredString(width / 2, height - 45,
        "Dehradun Road Near Power Grid Bhagwanpur, Roorkee, Haridwar, U.K. 247661, India"
    )

    c.setFont("Helvetica-Bold", 11)
    c.drawCentredString(width / 2, height - 65, "INVOICE")

    # ===== PARTY DETAILS =====
    c.rect(20, height - 170, 260, 90)
    c.setFont("Helvetica-Bold", 9)
    c.drawString(25, height - 85, "To,")
    c.setFont("Helvetica", 9)
    c.drawString(25, height - 100, "Grivaa Springs Private Ltd.")
    c.drawString(25, height - 115, "Khasra no 135, Tanshipur, Roorkee")
    c.drawString(25, height - 130, "Roorkee, Uttarakhand 247656")
    c.drawString(25, height - 145, "GSTIN: 05AAICG4793P1ZV")

    # ===== INVOICE BOX =====
    c.rect(width - 220, height - 170, 200, 90)
    c.setFont("Helvetica-Bold", 9)
    c.drawString(width - 210, height - 100, f"Freight Bill No : {invoice_no}")
    c.drawString(width - 210, height - 120,
        f"Invoice Date : {df.iloc[0]['InvoiceDate'].strftime('%d-%m-%Y')}"
    )
    c.drawString(width - 210, height - 140,
        f"Due Date : {df.iloc[0]['DueDate'].strftime('%d-%m-%Y')}"
    )

    c.setFont("Helvetica", 9)
    c.drawString(25, height - 190, f"From location: {df.iloc[0]['FromLocation']}")

    # ===== TABLE HEADER =====
    y = height - 230
    c.setFont("Helvetica-Bold", 7)

    headers = [
        "S.No", "Shipment Date", "LR No", "Destination", "CN Number",
        "Truck No", "Invoice No", "Pkgs", "Weight (kgs)",
        "Date Arrival", "Date Delivery", "Truck Type",
        "Freight Amt", "To Point", "Unloading",
        "Source Det.", "Dest. Det.", "Total"
    ]

    col_widths = [25, 55, 45, 55, 55, 55, 55, 30, 40, 55, 55, 40, 45, 40, 45, 45, 45, 50]

    x = 20
    for i, h in enumerate(headers):
        c.drawString(x + 2, y, h)
        x += col_widths[i]

    c.line(20, y - 2, width - 20, y - 2)

    # ===== TABLE DATA =====
    y -= 15
    total_amount = 0

    c.setFont("Helvetica", 7)

    for idx, row in df.iterrows():
        x = 20
        row_total = (
            row["FreightAmt"] +
            row["ToPointCharges"] +
            row["UnloadingCharge"] +
            row["SourceDetention"] +
            row["DestinationDetention"]
        )

        values = [
            idx + 1,
            row["ShipmentDate"].strftime("%d-%m-%Y"),
            row["LRNo"],
            row["Destination"],
            row["CNNumber"],
            row["TruckNo"],
            row["InvoiceNo"],
            row["Pkgs"],
            row["WeightKgs"],
            row["DateArrival"].strftime("%d-%m-%Y"),
            row["DateDelivery"].strftime("%d-%m-%Y"),
            row["TruckType"],
            f"{row['FreightAmt']:.2f}",
            f"{row['ToPointCharges']:.2f}",
            f"{row['UnloadingCharge']:.2f}",
            f"{row['SourceDetention']:.2f}",
            f"{row['DestinationDetention']:.2f}",
            f"{row_total:.2f}",
        ]

        for i, v in enumerate(values):
            c.drawString(x + 2, y, str(v))
            x += col_widths[i]

        total_amount += row_total
        y -= 12

        if y < 100:
            c.showPage()
            y = height - 50

    # ===== TOTAL =====
    c.setFont("Helvetica-Bold", 9)
    c.drawRightString(width - 25, y - 10, f"{total_amount:.2f}")

    amount_words = num2words(total_amount, to="currency", lang="en_IN").title()
    c.drawString(25, y - 10, f"Total in words (Rs.): {amount_words} Only")

    # ===== FOOTER =====
    c.setFont("Helvetica", 8)
    c.drawString(25, 80,
        'Any changes or discrepancies should be highlighted within 5 working days else it would be considered final.'
    )

    # ===== BANK DETAILS =====
    c.rect(20, 20, 260, 120)
    c.setFont("Helvetica", 8)
    bank_lines = [
        ("Our PAN No.", "BSSPG9414K"),
        ("STC GSTIN", "05BSSPG9414K1ZA"),
        ("STC State Code", "5"),
        ("Account name", "South Transport Company"),
        ("Account no", "36420500142"),
        ("IFS Code", "ICIC0003642"),
    ]

    yb = 120
    for k, v in bank_lines:
        c.drawString(25, yb, k)
        c.drawString(140, yb, v)
        yb -= 18

    c.drawRightString(width - 50, 60, "For SOUTH TRANSPORT COMPANY")
    c.drawRightString(width - 50, 40, "Authorized Signatory")

    c.save()
    return pdf_file


if __name__ == "__main__":
    app.run(debug=True)
