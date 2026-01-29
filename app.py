from flask import Flask, render_template, request, send_file
import pandas as pd
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.units import mm
from reportlab.pdfgen import canvas
from reportlab.lib import colors
from num2words import num2words
import os
from PIL import Image
import zipfile

app = Flask(__name__)

UPLOAD_FOLDER = "uploads"
OUTPUT_FOLDER = "output"
LOGO_PATH = "static/logo.png"

os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)


def wrap_text(c, text, x, y, max_width, font_name="Helvetica", font_size=7, line_height=8):
    """Wrap text to fit within max_width and return list of lines"""
    words = str(text).split()
    lines = []
    current_line = ""
    
    c.setFont(font_name, font_size)
    
    for word in words:
        test_line = current_line + (" " if current_line else "") + word
        if c.stringWidth(test_line, font_name, font_size) <= max_width:
            current_line = test_line
        else:
            if current_line:
                lines.append(current_line)
            current_line = word
    
    if current_line:
        lines.append(current_line)
    
    return lines


def draw_wrapped_text(c, text, x, y, max_width, font_name="Helvetica", font_size=7, line_height=8):
    """Draw wrapped text centered"""
    lines = wrap_text(c, text, x, y, max_width, font_name, font_size, line_height)
    
    # Calculate starting y position to center vertically
    total_height = len(lines) * line_height
    start_y = y + (total_height / 2) - line_height/2
    
    c.setFont(font_name, font_size)
    for i, line in enumerate(lines):
        c.drawCentredString(x, start_y - (i * line_height), line)


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
            
            # Generate PDFs for each unique FreightBillNo
            pdf_files = generate_multiple_pdfs(df)
            print(f"Generated {len(pdf_files)} PDF(s)")
            
            # If only one PDF, send it directly
            if len(pdf_files) == 1:
                return send_file(pdf_files[0], as_attachment=True)
            
            # If multiple PDFs, create a zip file
            zip_path = os.path.join(OUTPUT_FOLDER, "invoices.zip")
            with zipfile.ZipFile(zip_path, 'w') as zipf:
                for pdf_file in pdf_files:
                    zipf.write(pdf_file, os.path.basename(pdf_file))
            
            return send_file(zip_path, as_attachment=True)
        
        except Exception as e:
            print(f"ERROR OCCURRED: {str(e)}")
            import traceback
            traceback.print_exc()
            return f"Error: {str(e)}", 500

    return render_template("index.html")


def generate_multiple_pdfs(df):
    """Generate separate PDF for each unique FreightBillNo"""
    pdf_files = []
    
    # Group by FreightBillNo
    grouped = df.groupby('FreightBillNo')
    
    for bill_no, group_df in grouped:
        print(f"\n=== Generating PDF for Bill: {bill_no} ===")
        pdf_path = generate_pdf(group_df.reset_index(drop=True))
        pdf_files.append(pdf_path)
    
    return pdf_files


def generate_pdf(df):
    bill_no = str(df.iloc[0]["FreightBillNo"]).replace("/", "_")
    pdf_path = f"{OUTPUT_FOLDER}/{bill_no}.pdf"

    c = canvas.Canvas(pdf_path, pagesize=landscape(A4))
    width, height = landscape(A4)

    # ================= OUTER BORDER =================
    margin = 15
    c.rect(margin, margin, width - 2*margin, height - 2*margin, stroke=1, fill=0)

    # ================= LOGO WITH ERROR HANDLING =================
    try:
        if os.path.exists(LOGO_PATH):
            try:
                img = Image.open(LOGO_PATH)
                img.verify()
                print(f"✓ Logo found and verified: {LOGO_PATH}")
                
                c.drawImage(LOGO_PATH, 55, height - 140, width=100, height=80, preserveAspectRatio=True)
                print("✓ Logo successfully added to PDF")
                
            except Exception as img_error:
                print(f"⚠ Logo file is corrupted or invalid: {img_error}")
                c.setFont("Helvetica-Bold", 10)
                c.drawString(55, height - 80, "[LOGO]")
        else:
            print(f"⚠ Logo not found at: {LOGO_PATH}")
            c.setFont("Helvetica-Bold", 10)
            c.drawString(55, height - 80, "[LOGO]")
            
    except Exception as e:
        print(f"⚠ Error loading logo: {e}")

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

    # ================= TABLE (OPTIMIZED COLUMN WIDTHS) =================
    table_top = box_top - 155
    table_left = 30  # Left margin
    table_right = width - 30  # Right margin
    table_width = table_right - table_left  # Total available width = ~812 points
    
    # Table headers
    headers = [
        "S.\nno.", "Shipment\nDate", "LR\nNo.", "Destination", "CN\nNumber",
        "Truck No", "Invoice No", "Pkgs", "Weight\n(Kgs)", "Date of\nArrival",
        "Date of\nDelivery", "Truck\nType", "Freight\nAmt (Rs.)", "To Point\nCharges(Rs.)",
        "Unloading\nCharge (Rs.)", "Source\nDetention\n(Rs.)", "Destination\nDetention\n(Rs.)",
        "Total\nAmount (Rs.)"
    ]
    
    # OPTIMIZED Column widths - Better space distribution
    # Total should be ~812 to fit border to border with equal margins
    col_widths = [
        25,   # 0: S.no
        48,   # 1: Shipment Date (increased)
        28,   # 2: LR No
        52,   # 3: Destination (increased)
        32,   # 4: CN Number
        48,   # 5: Truck No (increased)
        85,   # 6: Invoice No (INCREASED significantly - was 60, now 85)
        28,   # 7: Pkgs
        40,   # 8: Weight (increased)
        48,   # 9: Date Arrival (increased)
        48,   # 10: Date Delivery (increased)
        48,   # 11: Truck Type (increased for wrapping)
        52,   # 12: Freight Amt (increased)
        52,   # 13: To Point Charges (increased)
        52,   # 14: Unloading (increased)
        52,   # 15: Source Detention (increased)
        52,   # 16: Destination Detention (increased)
        62    # 17: Total Amount (increased from end)
    ]
    
    # Verify total width
    total_col_width = sum(col_widths)
    print(f"Table width: {total_col_width}, Available: {table_width}")
    
    # Draw header background
    c.setFillColor(colors.lightgrey)
    c.rect(table_left, table_top - 30, total_col_width, 30, stroke=1, fill=1)
    
    # Draw headers
    c.setFillColor(colors.black)
    c.setFont("Helvetica-Bold", 7)
    
    x = table_left
    for i, header in enumerate(headers):
        lines = header.split('\n')
        y_offset = table_top - 9
        for line in lines:
            c.drawCentredString(x + col_widths[i]/2, y_offset, line)
            y_offset -= 8
        x += col_widths[i]
    
    # Draw vertical lines for header
    x = table_left
    for width_val in col_widths:
        c.line(x, table_top, x, table_top - 30)
        x += width_val
    c.line(x, table_top, x, table_top - 30)  # Last line
    
    # ================= TABLE DATA =================
    c.setFont("Helvetica", 7)
    y = table_top - 30
    total_amount = 0
    
    for idx, row in df.iterrows():
        row_height = 35
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
        
        # Data values
        values = [
            str(idx + 1),  # 0: S.no
            row["ShipmentDate"].strftime("%d %b %Y"),  # 1: Shipment Date
            str(row["LRNo"]),  # 2: LR No
            str(row["Destination"]),  # 3: Destination
            str(row["CNNumber"]),  # 4: CN Number
            str(row["TruckNo"]),  # 5: Truck No
            str(row["InvoiceNo"]),  # 6: Invoice No - NEEDS WRAPPING
            str(int(row["Pkgs"])),  # 7: Pkgs
            str(int(row["WeightKgs"])),  # 8: Weight
            row["DateArrival"].strftime("%d %b %Y"),  # 9: Date Arrival
            row["DateDelivery"].strftime("%d %b %Y"),  # 10: Date Delivery
            str(row["TruckType"]),  # 11: Truck Type - NEEDS WRAPPING
            f"{float(row['FreightAmt']):.2f}",  # 12: Freight
            f"{float(row['ToPointCharges']):.2f}",  # 13: To Point
            f"{float(row['UnloadingCharge']):.2f}",  # 14: Unloading
            f"{float(row['SourceDetention']):.2f}",  # 15: Source
            f"{float(row['DestinationDetention']):.2f}",  # 16: Destination
            f"{row_total:.2f}"  # 17: Total
        ]
        
        # Columns that need wrapping - Invoice No (6) and Truck Type (11)
        wrap_columns = [6, 11]
        
        # Draw row
        x = table_left
        for i, val in enumerate(values):
            if i in wrap_columns:
                # Use proper wrapping with more padding for Invoice No
                padding = 8 if i == 6 else 6  # More padding for Invoice No
                draw_wrapped_text(c, val, x + col_widths[i]/2, y + row_height/2, 
                                col_widths[i] - padding, "Helvetica", 7, 8)
            else:
                c.drawCentredString(x + col_widths[i]/2, y + row_height/2, val)
            x += col_widths[i]
        
        # Draw horizontal line
        c.line(table_left, y, table_left + total_col_width, y)
    
    # Draw vertical lines for data rows
    x = table_left
    for width_val in col_widths:
        c.line(x, table_top - 30, x, y)
        x += width_val
    c.line(x, table_top - 30, x, y)  # Last line
    
    # ================= TOTAL ROW =================
    total_row_height = 25
    y -= total_row_height
    
    c.rect(table_left, y, total_col_width, total_row_height, stroke=1, fill=0)
    
    c.setFont("Helvetica-Bold", 8)
    total_words = num2words(total_amount, to='currency', lang='en_IN').title()
    c.drawString(table_left + 5, y + 10, f"Total in words (Rs.) :  {total_words} Only")
    
    c.setFont("Helvetica-Bold", 9)
    total_col_x = table_left + sum(col_widths[:-1])
    c.drawCentredString(total_col_x + col_widths[-1]/2, y + 10, f"{total_amount:.2f}")
    
    c.line(total_col_x, y, total_col_x, y + total_row_height)

    # ================= NOTE =================
    c.setFont("Helvetica", 7)
    note_y = y - 15
    c.drawString(30, note_y, 'Any changes or discrepancies should be highlighted within 5 working days else it would be considered final. Please send all remittance details to "southtprk@gmail.com".')

    # ================= BANK DETAILS TABLE =================
    bank_y = note_y - 25
    
    bank_details = [
        ("Our PAN No.", "BSSPG9414K"),
        ("STC GSTIN", "05BSSPG9414K1ZA"),
        ("STC State Code", "5"),
        ("Account name", "South Transport Company"),
        ("Account no", "364205500142"),
        ("IFS Code", "ICIC0003642")
    ]
    
    row_height = 13
    
    c.setFont("Helvetica", 7)
    for i, (label, value) in enumerate(bank_details):
        y_pos = bank_y - (i * row_height)
        
        c.rect(30, y_pos - row_height, 100, row_height, stroke=1, fill=0)
        c.rect(130, y_pos - row_height, 100, row_height, stroke=1, fill=0)
        
        c.drawString(35, y_pos - 10, label)
        c.drawString(135, y_pos - 10, value)

    # ================= SIGNATURE =================
    sig_y = bank_y - 15
    c.setFont("Helvetica-Bold", 9)
    c.drawRightString(width - 35, sig_y, "For SOUTH TRANSPORT COMPANY")
    
    c.setFont("Helvetica", 7)
    c.drawRightString(width - 35, sig_y - 50, "(Authorized Signatory)")
    c.line(width - 180, sig_y - 52, width - 35, sig_y - 52)

    c.save()
    print(f"✓ PDF saved successfully: {pdf_path}")
    return pdf_path


if __name__ == "__main__":
    app.run(debug=True)