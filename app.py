from flask import Flask, render_template, request, send_file
import pandas as pd
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.units import mm
from reportlab.pdfgen import canvas
from reportlab.lib import colors
from num2words import num2words
import os

app = Flask(__name__)

UPLOAD_FOLDER = "uploads"
OUTPUT_FOLDER = "output"
LOGO_PATH = "static/logo.png"

os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)


def wrap_text(c, text, x, y, max_width, font_name="Helvetica", font_size=7, line_height=8):
    """Wrap text to fit within max_width"""
    words = str(text).split()
    lines = []
    current_line = ""
    
    c.setFont(font_name, font_size)
    
    for word in words:
        test_line = current_line + word + " "
        if c.stringWidth(test_line, font_name, font_size) <= max_width:
            current_line = test_line
        else:
            if current_line:
                lines.append(current_line.strip())
            current_line = word + " "
    
    if current_line:
        lines.append(current_line.strip())
    
    # Draw lines centered
    start_y = y + (len(lines) - 1) * line_height / 2
    for i, line in enumerate(lines):
        c.drawCentredString(x, start_y - (i * line_height), line)
    
    return y - (len(lines) * line_height)


@app.route("/", methods=["GET", "POST"])
def index():
    if request.method == "POST":
        try:
            if "file" not in request.files:
                print("ERROR: No file in request.files")
                return "No file uploaded", 400

            file = request.files["file"]
            if file.filename == "":
                print("ERROR: Empty filename")
                return "No file selected", 400

            print(f"File received: {file.filename}")
            path = os.path.join(UPLOAD_FOLDER, file.filename)
            file.save(path)
            print(f"File saved to: {path}")

            df = pd.read_excel(path)
            print(f"Excel loaded. Rows: {len(df)}, Columns: {list(df.columns)}")
            
            # Convert dates to datetime if they're not already
            date_columns = ['InvoiceDate', 'DueDate', 'ShipmentDate', 'DateArrival', 'DateDelivery']
            for col in date_columns:
                if col in df.columns:
                    df[col] = pd.to_datetime(df[col], errors='coerce')
            
            pdf_path = generate_pdf(df)
            print(f"PDF generated: {pdf_path}")

            return send_file(pdf_path, as_attachment=True)
        
        except Exception as e:
            print(f"ERROR OCCURRED: {str(e)}")
            import traceback
            traceback.print_exc()
            return f"Error: {str(e)}", 500

    return render_template("index.html")


def generate_pdf(df):
    bill_no = str(df.iloc[0]["FreightBillNo"]).replace("/", "_")
    pdf_path = f"{OUTPUT_FOLDER}/{bill_no}.pdf"

    c = canvas.Canvas(pdf_path, pagesize=landscape(A4))
    width, height = landscape(A4)

    # ================= OUTER BORDER =================
    c.rect(15, 15, width - 30, height - 30, stroke=1, fill=0)

    # ================= LOGO =================
    if os.path.exists(LOGO_PATH):
        c.drawImage(LOGO_PATH, 55, height - 140, width=100, height=80, preserveAspectRatio=True, mask='auto')

    # ================= HEADER =================
    c.setFont("Helvetica-Bold", 16)
    c.drawCentredString(width / 2, height - 70, "SOUTH TRANSPORT COMPANY")
    
    c.setFont("Helvetica", 9)
    c.drawCentredString(width / 2, height - 85, "Dehradun Road Near power Grid Bhagwanpur")
    c.drawCentredString(width / 2, height - 98, "Roorkee, Haridwar, U.K. 247661, India")
    
    c.setFont("Helvetica-Bold", 12)
    c.drawCentredString(width / 2, height - 125, "INVOICE")

    # ================= LEFT BOX (To Details) =================
    box_top = height - 160
    c.rect(30, box_top - 110, 260, 110, stroke=1, fill=0)
    
    c.setFont("Helvetica-Bold", 10)
    c.drawString(40, box_top - 20, "To,")
    
    c.setFont("Helvetica", 9)
    c.drawString(40, box_top - 35, "Grivaa Springs Private Ltd.")
    c.drawString(40, box_top - 50, "Khasra no 135, Tansipur, Roorkee")
    c.drawString(40, box_top - 65, "Roorkee, Uttarakhand 247656")
    c.drawString(40, box_top - 85, "GSTIN: 05AAICG4793P1ZV")

    # ================= RIGHT BOX (Bill Details) =================
    c.rect(width - 290, box_top - 110, 260, 110, stroke=1, fill=0)
    
    c.setFont("Helvetica-Bold", 10)
    c.drawString(width - 280, box_top - 25, f"Freight Bill No: {df.iloc[0]['FreightBillNo']}")
    
    c.setFont("Helvetica", 9)
    c.drawString(width - 280, box_top - 45, f"Invoice Date:      {df.iloc[0]['InvoiceDate'].strftime('%d %b %Y')}")
    c.drawString(width - 280, box_top - 65, f"Due Date:          {df.iloc[0]['DueDate'].strftime('%d-%m-%y')}")

    # ================= FROM LOCATION =================
    c.setFont("Helvetica", 9)
    c.drawString(30, box_top - 125, f"From location: {df.iloc[0]['FromLocation']}")

    # ================= TABLE =================
    table_top = box_top - 155
    
    # Table headers
    headers = [
        "S.\nno.", "Shipment\nDate", "LR No.", "Destination", "CN\nNumber",
        "Truck No", "Invoice No", "Pkgs", "Weight\n(Kgs)", "Date of\nArrival",
        "Date of\nDelivery", "Truck\nType", "Freight\nAmt (Rs.)", "To Point\nCharges(Rs.)",
        "Unloading\nCharge (Rs.)", "Source\nDetention\n(Rs.)", "Destination\nDetention\n(Rs.)",
        "Total\nAmount (Rs.)"
    ]
    
    # Column widths - adjusted to fit within border (width - 60 = ~780 points)
    col_widths = [25, 42, 32, 48, 32, 45, 55, 28, 38, 42, 42, 38, 48, 48, 48, 48, 52, 52]
    
    # Draw header background
    c.setFillColor(colors.lightgrey)
    c.rect(30, table_top - 30, sum(col_widths), 30, stroke=1, fill=1)
    
    # Draw headers
    c.setFillColor(colors.black)
    c.setFont("Helvetica-Bold", 7)
    
    x = 30
    for i, header in enumerate(headers):
        # Multi-line header
        lines = header.split('\n')
        y_offset = table_top - 9
        for line in lines:
            c.drawCentredString(x + col_widths[i]/2, y_offset, line)
            y_offset -= 8
        x += col_widths[i]
    
    # Draw vertical lines for header
    x = 30
    for width_val in col_widths:
        c.line(x, table_top, x, table_top - 30)
        x += width_val
    c.line(x, table_top, x, table_top - 30)  # Last line
    
    # ================= TABLE DATA =================
    c.setFont("Helvetica", 7)
    y = table_top - 30
    total_amount = 0
    
    for idx, row in df.iterrows():
        row_height = 35  # Increased for text wrapping
        y -= row_height
        
        # Calculate row total
        row_total = (
            float(row["FreightAmt"]) + 
            float(row["ToPointCharges"]) + 
            float(row["UnloadingCharge"]) + 
            float(row["SourceDetention"]) + 
            float(row["DestinationDetention"])
        )
        total_amount += row_total
        
        # Data values with indices for wrapping
        values = [
            (str(idx + 1), False),  # 0: S.no
            (row["ShipmentDate"].strftime("%d %b %Y"), False),  # 1: Shipment Date
            (str(row["LRNo"]), False),  # 2: LR No
            (str(row["Destination"]), False),  # 3: Destination
            (str(row["CNNumber"]), False),  # 4: CN Number
            (str(row["TruckNo"]), False),  # 5: Truck No
            (str(row["InvoiceNo"]), True),  # 6: Invoice No - WRAP
            (str(int(row["Pkgs"])), False),  # 7: Pkgs
            (str(int(row["WeightKgs"])), False),  # 8: Weight
            (row["DateArrival"].strftime("%d %b %Y"), False),  # 9: Date Arrival
            (row["DateDelivery"].strftime("%d %b %Y"), False),  # 10: Date Delivery
            (str(row["TruckType"]), True),  # 11: Truck Type - WRAP
            (f"{float(row['FreightAmt']):.2f}", False),  # 12: Freight
            (f"{float(row['ToPointCharges']):.2f}", False),  # 13: To Point
            (f"{float(row['UnloadingCharge']):.2f}", False),  # 14: Unloading
            (f"{float(row['SourceDetention']):.2f}", False),  # 15: Source
            (f"{float(row['DestinationDetention']):.2f}", False),  # 16: Destination
            (f"{row_total:.2f}", False)  # 17: Total
        ]
        
        # Draw row with wrapping for specific columns
        x = 30
        for i, (val, should_wrap) in enumerate(values):
            if should_wrap:
                # Use wrap_text for Invoice No and Truck Type
                wrap_text(c, val, x + col_widths[i]/2, y + row_height/2, col_widths[i] - 4)
            else:
                c.drawCentredString(x + col_widths[i]/2, y + row_height/2, val)
            x += col_widths[i]
        
        # Draw horizontal line
        c.line(30, y, 30 + sum(col_widths), y)
    
    # Draw vertical lines for data rows
    x = 30
    for width_val in col_widths:
        c.line(x, table_top - 30, x, y)
        x += width_val
    c.line(x, table_top - 30, x, y)  # Last line
    
    # ================= TOTAL ROW (INSIDE BORDER) =================
    total_row_height = 25
    y -= total_row_height
    
    # Draw total row border
    c.rect(30, y, sum(col_widths), total_row_height, stroke=1, fill=0)
    
    # Total in words on left side
    c.setFont("Helvetica-Bold", 8)
    total_words = num2words(total_amount, to='currency', lang='en_IN').title()
    c.drawString(35, y + 10, f"Total in words (Rs.) :  {total_words} Only")
    
    # Total amount on right side (below Total Amount column)
    c.setFont("Helvetica-Bold", 9)
    # Calculate x position for last column (Total Amount)
    total_col_x = 30 + sum(col_widths[:-1])
    c.drawCentredString(total_col_x + col_widths[-1]/2, y + 10, f"{total_amount:.2f}")
    
    # Vertical line before total amount
    c.line(total_col_x, y, total_col_x, y + total_row_height)

    # ================= NOTE =================
    c.setFont("Helvetica", 7)
    note_y = y - 15
    c.drawString(30, note_y, 'Any changes or discrepancies should be highlighted within 5 working days else it would be considered final. Please send all remittance details to "southtprk@gmail.com".')

    # ================= BANK DETAILS TABLE (INSIDE BORDER) =================
    bank_y = note_y - 25
    
    # Bank details data
    bank_details = [
        ("Our PAN No.", "BSSPG9414K"),
        ("STC GSTIN", "05BSSPG9414K1ZA"),
        ("STC State Code", "5"),
        ("Account name", "South Transport Company"),
        ("Account no", "364205500142"),
        ("IFS Code", "ICIC0003642")
    ]
    
    # Draw bank table
    row_height = 13
    
    c.setFont("Helvetica", 7)
    for i, (label, value) in enumerate(bank_details):
        y_pos = bank_y - (i * row_height)
        
        # Draw cells
        c.rect(30, y_pos - row_height, 100, row_height, stroke=1, fill=0)
        c.rect(130, y_pos - row_height, 100, row_height, stroke=1, fill=0)
        
        # Draw text
        c.drawString(35, y_pos - 15, label)
        c.drawString(135, y_pos - 15, value)

    # ================= SIGNATURE (INSIDE BORDER) =================
    sig_y = bank_y - 15
    c.setFont("Helvetica-Bold", 9)
    c.drawRightString(width - 35, sig_y, "For SOUTH TRANSPORT COMPANY")
    
    c.setFont("Helvetica", 7)
    c.drawRightString(width - 35, sig_y - 50, "(Authorized Signatory)")
    c.line(width - 180, sig_y - 52, width - 35, sig_y - 52)

    c.save()
    return pdf_path


if __name__ == "__main__":
    app.run(debug=True)
