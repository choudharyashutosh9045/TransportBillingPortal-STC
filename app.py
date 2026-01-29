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

def draw_wrapped_text(c, text, x, y, max_width, leading=9):
    words = str(text).split(" ")
    line = ""
    for word in words:
        test_line = line + word + " "
        if c.stringWidth(test_line, "Helvetica", 7) <= max_width:
            line = test_line
        else:
            c.drawString(x, y, line)
            y -= leading
            line = word + " "
    if line:
        c.drawString(x, y, line)
    return y


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

    # ================= LOGO (SMALLER) =================
    logo_path = "static/logo.png"   # apna logo path
    if os.path.exists(logo_path):
        c.drawImage(
            logo_path,
            25,
            height - 95,
            width=80,
            height=45,
            preserveAspectRatio=True,
            mask='auto'
        )

    # ================= HEADER =================
    c.setFont("Helvetica-Bold", 13)
    c.drawCentredString(width / 2, height - 50, "SOUTH TRANSPORT COMPANY")

    c.setFont("Helvetica", 9)
    c.drawCentredString(
        width / 2,
        height - 64,
        "Dehradun Road Near Power Grid Bhagwanpur, Roorkee, Haridwar, U.K. 247661, India"
    )

    c.setFont("Helvetica-Bold", 11)
    c.drawCentredString(width / 2, height - 80, "INVOICE")

    # ================= LEFT BOX (SMALLER) =================
    c.rect(20, height - 180, 250, 85)
    c.setFont("Helvetica", 8)
    c.drawString(25, height - 105, "To,")
    c.drawString(25, height - 118, "Grivaa Springs Private Ltd.")
    c.drawString(25, height - 131, "Khasra no 135, Tanshipur, Roorkee")
    c.drawString(25, height - 144, "Roorkee, Uttarakhand 247656")
    c.drawString(25, height - 157, "GSTIN: 05AAICG4793P1ZV")

    # ================= RIGHT BOX (SMALLER) =================
    c.rect(width - 230, height - 180, 210, 85)
    c.setFont("Helvetica", 8)
    c.drawString(width - 220, height - 115, f"Freight Bill No : {invoice_no}")
    c.drawString(
        width - 220,
        height - 130,
        f"Invoice Date : {df.iloc[0]['InvoiceDate'].strftime('%d-%m-%Y')}"
    )
    c.drawString(
        width - 220,
        height - 145,
        f"Due Date : {df.iloc[0]['DueDate'].strftime('%d-%m-%Y')}"
    )

    c.drawString(25, height - 195, f"From location: {df.iloc[0]['FromLocation']}")

    # ================= TABLE HEADER =================
    y = height - 230
    c.setFont("Helvetica-Bold", 7)

    headers = [
        "S.No", "Shipment\nDate", "LR No", "Destination", "CN No",
        "Truck No", "Invoice No", "Pkgs", "Weight",
        "Arrival", "Delivery", "Truck Type",
        "Freight", "To Point", "Unload",
        "Src Det", "Dst Det", "Total"
    ]

    col_widths = [22, 45, 35, 48, 40, 45, 48, 26, 35, 40, 40, 40, 42, 38, 40, 38, 38, 45]

    x = 20
    for i, h in enumerate(headers):
        c.drawString(x + 2, y, h)
        x += col_widths[i]

    c.line(20, y - 2, width - 20, y - 2)
    y -= 14

    # ================= TABLE DATA =================
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
            row["FreightAmt"],
            row["ToPointCharges"],
            row["UnloadingCharge"],
            row["SourceDetention"],
            row["DestinationDetention"],
            row_total
        ]

        max_row_height = y
        for i, val in enumerate(values):
            new_y = draw_wrapped_text(
                c,
                val,
                x + 2,
                y,
                col_widths[i] - 4
            )
            max_row_height = min(max_row_height, new_y)
            x += col_widths[i]

        y = max_row_height - 8
        total_amount += row_total

    # ================= TOTAL =================
    c.setFont("Helvetica-Bold", 8)
    c.drawRightString(width - 25, y, f"{total_amount:.2f}")

    amount_words = num2words(total_amount, to="currency", lang="en_IN").title()
    c.drawString(22, y, f"Total in words (Rs.): {amount_words} Only")

    c.save()
    return pdf_file


if __name__ == "__main__":
    app.run(debug=True)
