from flask import Flask, render_template, request, send_file, jsonify
import pandas as pd
from reportlab.lib.pagesizes import A4, landscape
from reportlab.pdfgen import canvas
from reportlab.lib import colors
from num2words import num2words
import os
from PIL import Image
import zipfile
from datetime import datetime
import json

app = Flask(__name__)

UPLOAD_FOLDER = "uploads"
OUTPUT_FOLDER = "output"
LOGO_PATH = "static/logo.png"
HISTORY_FILE = "history.json"
TEMPLATE_FILE = "excel_template.xlsx"

os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)
os.makedirs("static", exist_ok=True)


def load_history():
    """Load history from JSON file"""
    if os.path.exists(HISTORY_FILE):
        try:
            with open(HISTORY_FILE, 'r') as f:
                return json.load(f)
        except:
            return []
    return []


def save_history(entry):
    """Save history entry"""
    history = load_history()
    history.insert(0, entry)
    history = history[:10]  # Keep last 10
    with open(HISTORY_FILE, 'w') as f:
        json.dump(history, f, indent=2)


def create_excel_template():
    """Create Excel template file if it doesn't exist"""
    if not os.path.exists(TEMPLATE_FILE):
        columns = [
            'FreightBillNo', 'InvoiceDate', 'DueDate', 'FromLocation',
            'ShipmentDate', 'LRNo', 'Destination', 'CNNumber', 'TruckNo',
            'InvoiceNo', 'Pkgs', 'WeightKgs', 'DateArrival', 'DateDelivery',
            'TruckType', 'FreightAmt', 'ToPointCharges', 'UnloadingCharge',
            'SourceDetention', 'DestinationDetention'
        ]
        
        # Create sample data
        sample_data = {
            'FreightBillNo': ['FB/2025/001'],
            'InvoiceDate': ['2025-01-15'],
            'DueDate': ['2025-02-15'],
            'FromLocation': ['Roorkee'],
            'ShipmentDate': ['2025-01-10'],
            'LRNo': ['LR12345'],
            'Destination': ['Delhi'],
            'CNNumber': ['CN001'],
            'TruckNo': ['UK01AB1234'],
            'InvoiceNo': ['INV001'],
            'Pkgs': [10],
            'WeightKgs': [500],
            'DateArrival': ['2025-01-12'],
            'DateDelivery': ['2025-01-13'],
            'TruckType': ['Open Body'],
            'FreightAmt': [5000],
            'ToPointCharges': [500],
            'UnloadingCharge': [300],
            'SourceDetention': [0],
            'DestinationDetention': [0]
        }
        
        df = pd.DataFrame(sample_data)
        df.to_excel(TEMPLATE_FILE, index=False)
        print(f"✓ Template created: {TEMPLATE_FILE}")


def wrap_text_lines(c, text, max_width, font_name="Helvetica", font_size=7):
    """Wrap text to fit within max_width"""
    text = str(text)
    lines = []
    current_line = ""
    
    c.setFont(font_name, font_size)
    
    if '/' in text:
        parts = text.split('/')
        for i, part in enumerate(parts):
            test_line = current_line + part
            if i < len(parts) - 1:
                test_line += "/"
            
            if c.stringWidth(test_line, font_name, font_size) <= max_width:
                current_line = test_line
            else:
                if current_line:
                    lines.append(current_line)
                current_line = part
                if i < len(parts) - 1:
                    current_line += "/"
        
        if current_line:
            lines.append(current_line)
    else:
        words = text.split()
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
    
    return lines if lines else [text]


def draw_wrapped_text(c, text, x, y, max_width, font_name="Helvetica", font_size=7, line_height=7):
    """Draw wrapped text centered"""
    lines = wrap_text_lines(c, text, max_width, font_name, font_size)
    
    total_height = len(lines) * line_height
    start_y = y + (total_height / 2) - (line_height / 2)
    
    c.setFont(font_name, font_size)
    for i, line in enumerate(lines):
        c.drawCentredString(x, start_y - (i * line_height), line)


@app.route("/")
def index():
    """Serve the main page"""
    return render_template("index.html")


@app.route("/download-template")
def download_template():
    """Download Excel template"""
    create_excel_template()
    return send_file(TEMPLATE_FILE, as_attachment=True, download_name="STC_Template.xlsx")


@app.route("/", methods=["POST"])
def generate_bills():
    """Generate PDF bills from Excel"""
    try:
        if "file" not in request.files:
            return jsonify({"error": "No file uploaded"}), 400

        file = request.files["file"]
        if file.filename == "":
            return jsonify({"error": "No file selected"}), 400

        print(f"📄 File received: {file.filename}")
        path = os.path.join(UPLOAD_FOLDER, file.filename)
        file.save(path)

        # Read Excel
        df = pd.read_excel(path)
        print(f"✓ Excel loaded: {len(df)} rows")
        
        # Convert dates
        date_columns = ['InvoiceDate', 'DueDate', 'ShipmentDate', 'DateArrival', 'DateDelivery']
        for col in date_columns:
            if col in df.columns:
                df[col] = pd.to_datetime(df[col], errors='coerce')
        
        # Generate PDFs
        pdf_files = generate_multiple_pdfs(df)
        print(f"✓ Generated {len(pdf_files)} PDF(s)")
        
        # Save to history with bill numbers
        bill_numbers = df['FreightBillNo'].unique().tolist()
        history_entry = {
            "time": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "file": file.filename,
            "rows": len(df),
            "bills": [str(b) for b in bill_numbers[:5]],  # First 5 bill numbers
            "pdf_files": [os.path.basename(f) for f in pdf_files]
        }
        save_history(history_entry)
        
        # Create ZIP
        zip_path = os.path.join(OUTPUT_FOLDER, "Bills.zip")
        with zipfile.ZipFile(zip_path, 'w') as zipf:
            for pdf_file in pdf_files:
                zipf.write(pdf_file, os.path.basename(pdf_file))
        
        print(f"✓ ZIP created: {zip_path}")
        return send_file(zip_path, as_attachment=True, download_name="Bills.zip")
        
    except Exception as e:
        print(f"❌ ERROR: {str(e)}")
        import traceback
        traceback.print_exc()
        return jsonify({"error": str(e)}), 500


@app.route("/preview", methods=["POST"])
def preview():
    """Preview Excel data"""
    try:
        if "file" not in request.files:
            return jsonify({"ok": False, "error": "No file uploaded"}), 400

        file = request.files["file"]
        if file.filename == "":
            return jsonify({"ok": False, "error": "No file selected"}), 400

        path = os.path.join(UPLOAD_FOLDER, f"preview_{file.filename}")
        file.save(path)

        # Read Excel
        df = pd.read_excel(path)
        
        # Convert dates for display
        date_columns = ['InvoiceDate', 'DueDate', 'ShipmentDate', 'DateArrival', 'DateDelivery']
        for col in date_columns:
            if col in df.columns:
                df[col] = pd.to_datetime(df[col], errors='coerce')
        
        # Get ALL rows for search functionality
        rows = []
        for idx, row in df.iterrows():
            total = (
                float(row.get("FreightAmt", 0)) + 
                float(row.get("ToPointCharges", 0)) + 
                float(row.get("UnloadingCharge", 0)) + 
                float(row.get("SourceDetention", 0)) + 
                float(row.get("DestinationDetention", 0))
            )
            
            rows.append({
                "FreightBillNo": str(row.get("FreightBillNo", "")),
                "LRNo": str(row.get("LRNo", "")),
                "TruckNo": str(row.get("TruckNo", "")),
                "InvoiceNo": str(row.get("InvoiceNo", "")),
                "Destination": str(row.get("Destination", "")),
                "TotalAmount": f"₹{total:.2f}"
            })
        
        # Clean up preview file
        os.remove(path)
        
        return jsonify({
            "ok": True,
            "count": len(df),
            "rows": rows  # Return all rows
        })
        
    except Exception as e:
        print(f"Preview ERROR: {str(e)}")
        import traceback
        traceback.print_exc()
        return jsonify({"ok": False, "error": str(e)}), 500


@app.route("/api/history")
def get_history():
    """Get history"""
    try:
        history = load_history()
        return jsonify(history if history else [])
    except Exception as e:
        print(f"History error: {e}")
        return jsonify([])


@app.route("/api/bills/<filename>")
def get_bill(filename):
    """Download a specific bill PDF"""
    try:
        file_path = os.path.join(OUTPUT_FOLDER, filename)
        if os.path.exists(file_path):
            return send_file(file_path, as_attachment=True, download_name=filename)
        else:
            return jsonify({"error": "File not found"}), 404
    except Exception as e:
        return jsonify({"error": str(e)}), 500


def generate_multiple_pdfs(df):
    """Generate separate PDF for each FreightBillNo"""
    pdf_files = []
    grouped = df.groupby('FreightBillNo')
    
    for bill_no, group_df in grouped:
        print(f"  → Generating: {bill_no}")
        pdf_path = generate_pdf(group_df.reset_index(drop=True))
        pdf_files.append(pdf_path)
    
    return pdf_files


def generate_pdf(df):
    """Generate single PDF"""
    bill_no = str(df.iloc[0]["FreightBillNo"]).replace("/", "_")
    pdf_path = f"{OUTPUT_FOLDER}/{bill_no}.pdf"

    c = canvas.Canvas(pdf_path, pagesize=landscape(A4))
    width, height = landscape(A4)

    margin = 15
    c.rect(margin, margin, width - 2*margin, height - 2*margin, stroke=1, fill=0)

    # Logo
    try:
        if os.path.exists(LOGO_PATH):
            try:
                img = Image.open(LOGO_PATH)
                img.verify()
                c.drawImage(LOGO_PATH, 55, height - 140, width=100, height=80, preserveAspectRatio=True)
            except:
                c.setFont("Helvetica-Bold", 10)
                c.drawString(55, height - 80, "[LOGO]")
        else:
            c.setFont("Helvetica-Bold", 10)
            c.drawString(55, height - 80, "[LOGO]")
    except:
        pass

    # Header
    c.setFont("Helvetica-Bold", 16)
    c.drawCentredString(width / 2, height - 70, "SOUTH TRANSPORT COMPANY")
    
    c.setFont("Helvetica", 9)
    c.drawCentredString(width / 2, height - 85, "Dehradun Road Near power Grid Bhagwanpur")
    c.drawCentredString(width / 2, height - 98, "Roorkee, Haridwar, U.K. 247661, India")
    
    c.setFont("Helvetica-Bold", 12)
    c.drawCentredString(width / 2, height - 125, "INVOICE")

    # Left Box
    box_top = height - 160
    c.rect(30, box_top - 110, 260, 110, stroke=1, fill=0)
    
    c.setFont("Helvetica-Bold", 10)
    c.drawString(40, box_top - 20, "To,")
    
    c.setFont("Helvetica", 9)
    c.drawString(40, box_top - 35, "Grivaa Springs Private Ltd.")
    c.drawString(40, box_top - 50, "Khasra no 135, Tansipur, Roorkee")
    c.drawString(40, box_top - 65, "Roorkee, Uttarakhand 247656")
    c.drawString(40, box_top - 85, "GSTIN: 05AAICG4793P1ZV")

    # Right Box
    c.rect(width - 290, box_top - 110, 260, 110, stroke=1, fill=0)
    
    c.setFont("Helvetica-Bold", 10)
    c.drawString(width - 280, box_top - 25, f"Freight Bill No: {df.iloc[0]['FreightBillNo']}")
    
    c.setFont("Helvetica", 9)
    c.drawString(width - 280, box_top - 45, f"Invoice Date:      {df.iloc[0]['InvoiceDate'].strftime('%d %b %Y')}")
    c.drawString(width - 280, box_top - 65, f"Due Date:          {df.iloc[0]['DueDate'].strftime('%d-%m-%y')}")

    c.setFont("Helvetica", 9)
    c.drawString(30, box_top - 125, f"From location: {df.iloc[0]['FromLocation']}")

    # Table
    table_top = box_top - 155
    table_left = 30
    table_right = width - 30
    
    headers = [
        "S.\nno.", "Shipment\nDate", "LR\nNo.", "Destination", "CN\nNumber",
        "Truck No", "Invoice No", "Pkgs", "Weight\n(Kgs)", "Date of\nArrival",
        "Date of\nDelivery", "Truck\nType", "Freight\nAmt (Rs.)", "To Point\nCharges(Rs.)",
        "Unloading\nCharge (Rs.)", "Source\nDetention\n(Rs.)", "Destination\nDetention\n(Rs.)",
        "Total\nAmount (Rs.)"
    ]
    
    col_widths = [22, 45, 27, 48, 30, 44, 70, 26, 38, 45, 45, 45, 48, 48, 48, 48, 50, 55]
    total_col_width = sum(col_widths)
    
    # Header
    c.setFillColor(colors.lightgrey)
    c.rect(table_left, table_top - 30, total_col_width, 30, stroke=1, fill=1)
    
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
    
    x = table_left
    for width_val in col_widths:
        c.line(x, table_top, x, table_top - 30)
        x += width_val
    c.line(x, table_top, x, table_top - 30)
    
    # Data
    c.setFont("Helvetica", 7)
    y = table_top - 30
    total_amount = 0
    
    for idx, row in df.iterrows():
        row_height = 35
        y -= row_height
        
        row_total = (
            float(row["FreightAmt"]) + 
            float(row["ToPointCharges"]) + 
            float(row["UnloadingCharge"]) + 
            float(row["SourceDetention"]) + 
            float(row["DestinationDetention"])
        )
        total_amount += row_total
        
        values = [
            str(idx + 1),
            row["ShipmentDate"].strftime("%d %b %Y"),
            str(row["LRNo"]),
            str(row["Destination"]),
            str(row["CNNumber"]),
            str(row["TruckNo"]),
            str(row["InvoiceNo"]),
            str(int(row["Pkgs"])),
            str(int(row["WeightKgs"])),
            row["DateArrival"].strftime("%d %b %Y"),
            row["DateDelivery"].strftime("%d %b %Y"),
            str(row["TruckType"]),
            f"{float(row['FreightAmt']):.2f}",
            f"{float(row['ToPointCharges']):.2f}",
            f"{float(row['UnloadingCharge']):.2f}",
            f"{float(row['SourceDetention']):.2f}",
            f"{float(row['DestinationDetention']):.2f}",
            f"{row_total:.2f}"
        ]
        
        wrap_columns = {6, 11}
        
        x = table_left
        for i, val in enumerate(values):
            if i in wrap_columns:
                draw_wrapped_text(c, val, x + col_widths[i]/2, y + row_height/2, 
                                col_widths[i] - 6, "Helvetica", 6, 7)
            else:
                c.drawCentredString(x + col_widths[i]/2, y + row_height/2, val)
            x += col_widths[i]
        
        c.line(table_left, y, table_left + total_col_width, y)
    
    x = table_left
    for width_val in col_widths:
        c.line(x, table_top - 30, x, y)
        x += width_val
    c.line(x, table_top - 30, x, y)
    
    # Total Row
    total_row_height = 25
    y -= total_row_height
    
    c.rect(table_left, y, total_col_width, total_row_height, stroke=1, fill=0)
    
    c.setFont("Helvetica-Bold", 8)
    rupees = int(total_amount)
    paise = int((total_amount - rupees) * 100)
    
    if paise > 0:
        total_words = f"{num2words(rupees, lang='en_IN').title()} Rupees and {num2words(paise, lang='en_IN').title()} Paise"
    else:
        total_words = f"{num2words(rupees, lang='en_IN').title()} Rupees"
    
    c.drawString(table_left + 5, y + 10, f"Total in words (Rs.) :  {total_words} Only")
    
    c.setFont("Helvetica-Bold", 9)
    total_col_x = table_left + sum(col_widths[:-1])
    c.drawCentredString(total_col_x + col_widths[-1]/2, y + 10, f"{total_amount:.2f}")
    
    c.line(total_col_x, y, total_col_x, y + total_row_height)

    # Note
    c.setFont("Helvetica", 7)
    note_y = y - 15
    c.drawString(30, note_y, 'Any changes or discrepancies should be highlighted within 5 working days else it would be considered final. Please send all remittance details to "southtprk@gmail.com".')

    # Bank Details
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

    # Signature
    sig_y = bank_y - 15
    c.setFont("Helvetica-Bold", 9)
    c.drawRightString(width - 35, sig_y, "For SOUTH TRANSPORT COMPANY")
    
    c.setFont("Helvetica", 7)
    c.drawRightString(width - 35, sig_y - 50, "(Authorized Signatory)")
    c.line(width - 180, sig_y - 52, width - 35, sig_y - 52)

    c.save()
    return pdf_path


if __name__ == "__main__":
    create_excel_template()  # Create template on startup
    app.run(debug=True, port=5000)