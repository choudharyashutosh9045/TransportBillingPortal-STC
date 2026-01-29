from flask import Flask, render_template, request, send_file
import pandas as pd
import os
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
from reportlab.lib.units import mm
from reportlab.lib.utils import ImageReader
from datetime import datetime
import math

app = Flask(__name__)

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
UPLOAD_DIR = os.path.join(BASE_DIR, 'uploads')
OUTPUT_DIR = os.path.join(BASE_DIR, 'output')
LOGO_PATH = os.path.join(BASE_DIR, 'static', 'logo.png')

os.makedirs(UPLOAD_DIR, exist_ok=True)
os.makedirs(OUTPUT_DIR, exist_ok=True)


def safe(val, default=""):
    if val is None:
        return default
    if isinstance(val, float) and math.isnan(val):
        return default
    return str(val)


def generate_invoice_pdf(row, pdf_path):
    folder = os.path.dirname(pdf_path)
    os.makedirs(folder, exist_ok=True)

    c = canvas.Canvas(pdf_path, pagesize=A4)
    width, height = A4

    c.rect(15, 15, width - 30, height - 30)

    if os.path.exists(LOGO_PATH):
        c.drawImage(ImageReader(LOGO_PATH), 25, height - 120, 120, 80, mask='auto')

    c.setFont("Helvetica-Bold", 14)
    c.drawCentredString(width / 2, height - 40, "SOUTH TRANSPORT COMPANY")

    c.setFont("Helvetica", 9)
    c.drawCentredString(width / 2, height - 55, "Dehradun Road Near power Grid Bhagwanpur")
    c.drawCentredString(width / 2, height - 68, "Roorkee, Haridwar, U.K. 247661, India")

    c.setFont("Helvetica-Bold", 11)
    c.drawCentredString(width / 2, height - 90, "INVOICE")

    y = height - 120

    c.setFont("Helvetica", 9)
    c.drawString(25, y, f"Invoice No: {safe(row.get('Invoice No'))}")
    c.drawRightString(width - 25, y, f"Date: {safe(row.get('Invoice Date'))}")

    y -= 20

    c.setFont("Helvetica-Bold", 9)
    c.drawString(25, y, "Bill To:")

    y -= 12
    c.setFont("Helvetica", 9)
    c.drawString(25, y, safe(row.get('Customer Name')))
    y -= 12
    c.drawString(25, y, safe(row.get('Customer Address')))

    y -= 25

    c.setFont("Helvetica-Bold", 9)
    headers = ["LR No", "Vehicle No", "From", "To", "Material", "Weight", "Rate", "Amount"]
    x_positions = [25, 90, 160, 230, 300, 380, 440, 500]

    for h, x in zip(headers, x_positions):
        c.drawString(x, y, h)

    y -= 10
    c.line(25, y, width - 25, y)

    y -= 15

    c.setFont("Helvetica", 9)
    c.drawString(25, y, safe(row.get('LR No')))
    c.drawString(90, y, safe(row.get('Vehicle No')))
    c.drawString(160, y, safe(row.get('From')))
    c.drawString(230, y, safe(row.get('To')))
    c.drawString(300, y, safe(row.get('Material')))
    c.drawRightString(420, y, safe(row.get('Weight')))
    c.drawRightString(470, y, safe(row.get('Rate')))
    c.drawRightString(width - 30, y, safe(row.get('Amount')))

    y -= 30

    c.setFont("Helvetica-Bold", 9)
    c.drawRightString(470, y, "Total")
    c.drawRightString(width - 30, y, safe(row.get('Amount')))

    y -= 40

    c.setFont("Helvetica", 9)
    c.drawString(25, y, "Bank Name: SOUTH TRANSPORT COMPANY")
    y -= 12
    c.drawString(25, y, "A/C No: 1234567890")
    y -= 12
    c.drawString(25, y, "IFSC: SBIN000000")

    y -= 30

    c.drawRightString(width - 25, y, "For SOUTH TRANSPORT COMPANY")

    y -= 40

    c.drawRightString(width - 25, y, "Authorised Signatory")

    c.save()


@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        file = request.files['file']
        if not file:
            return "No file uploaded"

        filepath = os.path.join(UPLOAD_DIR, file.filename)
        file.save(filepath)

        df = pd.read_excel(filepath)

        generated = []

        for _, row in df.iterrows():
            fy = safe(row.get('FY', '2025-26'))
            invoice_no = safe(row.get('Invoice No'))
            pdf_name = f"{invoice_no}.pdf"
            pdf_path = os.path.join(OUTPUT_DIR, fy, pdf_name)

            generate_invoice_pdf(row, pdf_path)
            generated.append(pdf_path)

        return send_file(generated[0], as_attachment=True)

    return render_template('index.html')


if __name__ == '__main__':
    app.run(debug=True)