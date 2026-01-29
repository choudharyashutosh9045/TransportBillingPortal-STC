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

def generate_invoice_pdf(data):
    file_path = os.path.join(OUTPUT_DIR, f"{data['FreightBillNo']}.pdf")
    c = canvas.Canvas(file_path, pagesize=A4)
    width, height = A4

    c.rect(15*mm, 15*mm, width-30*mm, height-30*mm)

    logo_path = os.path.join(BASE_DIR, 'static', 'logo.png')
    if os.path.exists(logo_path):
        c.drawImage(ImageReader(logo_path), 25*mm, height-55*mm, width=40*mm, height=40*mm, mask='auto')

    c.setFont('Helvetica-Bold', 14)
    c.drawCentredString(width/2, height-30*mm, 'SOUTH TRANSPORT COMPANY')

    c.setFont('Helvetica', 9)
    c.drawCentredString(width/2, height-36*mm, 'Dehradun Road Near Power Grid Bhagwanpur')
    c.drawCentredString(width/2, height-40*mm, 'Roorkee, Haridwar, U.K. 247661, India')

    c.setFont('Helvetica-Bold', 11)
    c.drawCentredString(width/2, height-48*mm, 'INVOICE')

    c.rect(25*mm, height-95*mm, 80*mm, 35*mm)
    c.setFont('Helvetica-Bold', 9)
    c.drawString(27*mm, height-65*mm, 'To,')
    c.setFont('Helvetica', 9)
    c.drawString(27*mm, height-70*mm, data['PartyName'])
    c.drawString(27*mm, height-75*mm, data['PartyAddress'])
    c.drawString(27*mm, height-80*mm, data['PartyCity'])
    c.drawString(27*mm, height-85*mm, f"GSTIN: {data['PartyGST']}")

    c.rect(width-105*mm, height-95*mm, 80*mm, 35*mm)
    c.setFont('Helvetica-Bold', 9)
    c.drawString(width-103*mm, height-70*mm, f"Freight Bill No: {data['FreightBillNo']}")
    c.drawString(width-103*mm, height-78*mm, f"Invoice Date: {data['InvoiceDate']}")
    c.drawString(width-103*mm, height-86*mm, f"Due Date: {data['DueDate']}")

    c.setFont('Helvetica', 9)
    c.drawString(25*mm, height-105*mm, f"From location: {data['FromLocation']}")

    start_y = height-120*mm
    col_x = [20, 30, 50, 65, 85, 110, 135, 155, 170, 190]
    headers = ['S.', 'Ship Dt', 'LR No', 'Dest', 'CN No', 'Truck No', 'Invoice No', 'Pkgs', 'Wt', 'Amount']

    c.setFont('Helvetica-Bold', 8)
    for i, h in enumerate(headers):
        c.drawString(col_x[i]*mm, start_y, h)

    c.line(20*mm, start_y-2*mm, width-20*mm, start_y-2*mm)

    c.setFont('Helvetica', 8)
    y = start_y-8*mm
    c.drawString(col_x[0]*mm, y, '1')
    c.drawString(col_x[1]*mm, y, data['ShipmentDate'])
    c.drawString(col_x[2]*mm, y, data['LRNo'])
    c.drawString(col_x[3]*mm, y, data['Destination'])
    c.drawString(col_x[4]*mm, y, data['CNNumber'])
    c.drawString(col_x[5]*mm, y, data['TruckNo'])
    c.drawString(col_x[6]*mm, y, data['InvoiceNo'])
    c.drawString(col_x[7]*mm, y, data['Pkgs'])
    c.drawString(col_x[8]*mm, y, data['WeightKgs'])
    c.drawRightString(width-22*mm, y, data['TotalAmount'])

    c.line(20*mm, y-4*mm, width-20*mm, y-4*mm)

    c.setFont('Helvetica-Bold', 9)
    c.drawString(22*mm, y-10*mm, f"Total in words (Rs.): {data['AmountWords']}")
    c.drawRightString(width-22*mm, y-10*mm, data['TotalAmount'])

    c.rect(22*mm, 40*mm, 80*mm, 35*mm)
    c.setFont('Helvetica', 8)
    c.drawString(24*mm, 68*mm, f"PAN: {data['PAN']}")
    c.drawString(24*mm, 62*mm, f"GSTIN: {data['STCGST']}")
    c.drawString(24*mm, 56*mm, f"State Code: {data['StateCode']}")
    c.drawString(24*mm, 50*mm, f"Account Name: {data['AccountName']}")
    c.drawString(24*mm, 44*mm, f"Account No: {data['AccountNo']}")
    c.drawString(24*mm, 38*mm, f"IFSC: {data['IFSC']}")

    c.setFont('Helvetica', 9)
    c.drawRightString(width-22*mm, 55*mm, 'For SOUTH TRANSPORT COMPANY')
    c.line(width-80*mm, 40*mm, width-22*mm, 40*mm)
    c.drawRightString(width-22*mm, 34*mm, '(Authorized Signatory)')

    c.showPage()
    c.save()
    return file_path

if __name__ == '__main__':
    app.run(debug=True)
