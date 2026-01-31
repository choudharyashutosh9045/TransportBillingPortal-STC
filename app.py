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
HISTORY_FILE = "history.json"

os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)
os.makedirs("static/logos", exist_ok=True)

# Company configurations
COMPANIES = {
    "stc": {
        "name": "SOUTH TRANSPORT COMPANY",
        "address_line1": "Dehradun Road Near power Grid Bhagwanpur",
        "address_line2": "Roorkee, Haridwar, U.K. 247661, India",
        "logo": "static/logos/stc_logo.png",
        "type": "basic",  # Basic format
        "customer": {
            "name": "Grivaa Springs Private Ltd.",
            "address_line1": "Khasra no 135, Tansipur, Roorkee",
            "address_line2": "Roorkee, Uttarakhand 247656",
            "gstin": "05AAICG4793P1ZV"
        },
        "bank": {
            "pan": "BSSPG9414K",
            "gstin": "05BSSPG9414K1ZA",
            "state_code": "5",
            "account_name": "South Transport Company",
            "account_no": "364205500142",
            "ifsc": "ICIC0003642",
            "email": "southtprk@gmail.com"
        }
    },
    "ntc": {
        "name": "NORTH TRANSPORT COMPANY",
        "address_line1": "Industrial Area, Sector 5",
        "address_line2": "Delhi, 110001, India",
        "logo": "static/logos/ntc_logo.png",
        "type": "basic",
        "customer": {
            "name": "ABC Industries Ltd.",
            "address_line1": "Plot No 45, Industrial Estate",
            "address_line2": "Delhi, 110002",
            "gstin": "07AAICA1234P1ZV"
        },
        "bank": {
            "pan": "ABCDN1234K",
            "gstin": "07ABCDN1234K1ZA",
            "state_code": "7",
            "account_name": "North Transport Company",
            "account_no": "123456789012",
            "ifsc": "HDFC0001234",
            "email": "contact@ntc.com"
        }
    },
    "wtc": {
        "name": "WEST TRANSPORT COMPANY",
        "address_line1": "Ring Road, Navrangpura",
        "address_line2": "Ahmedabad, Gujarat 380009, India",
        "logo": "static/logos/wtc_logo.png",
        "type": "basic",
        "customer": {
            "name": "Gujarat Trading Co.",
            "address_line1": "CG Road, Ahmedabad",
            "address_line2": "Gujarat 380006",
            "gstin": "24AABCG5678P1ZV"
        },
        "bank": {
            "pan": "WESTG5678K",
            "gstin": "24WESTG5678K1ZA",
            "state_code": "24",
            "account_name": "West Transport Company",
            "account_no": "987654321098",
            "ifsc": "SBIN0001234",
            "email": "info@wtc.com"
        }
    },
    "transin": {
        "name": "TRANSIN LOGISTICS PRIVATE LIMITED",
        "address_line1": "Plot No. 17 & 18, Vishnu Avenue, Flat No.304, 3rd Floor, VIP Hills, Jaihind Enclave",
        "address_line2": "Madhapur, Hyderabad 500081, Telangana, India",
        "logo": "static/logos/transin_logo.png",
        "type": "transin",  # Special Transin format
        "customer": {
            "name": "Balaji Action Buildwell Pvt Ltd.",
            "address_line1": "Sitarganj-U.K",
            "address_line2": "Uttrakhand code: 262405",
            "gstin": "05AAKCB1853F1ZW"
        },
        "bank": {
            "pan": "AAFCT6966J",
            "gstin": "36AAFCT6966J1ZP",
            "state_code": "36",
            "sac_code": "996791",
            "account_name": "Transin Logistics Pvt Ltd",
            "account_no": "N/A",
            "ifsc": "N/A",
            "email": "receivables@onmove.in"
        },
        "gst_note": 'Freight Claimed is exclusive of GST which has to be submitted by you to the Government. "We hereby confirm that GST ITC for providing the taxable service has not been taken by us under the provisions mentioned in of the GST Rules, 2017". Tax is payable on reverse charge basis.',
        "digital_signature": "Anchit Maheshwari"
    }
}


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
    history = history[:20]
    with open(HISTORY_FILE, 'w') as f:
        json.dump(history, f, indent=2)


def create_excel_template(company_code="stc"):
    """Create Excel template file based on company type"""
    template_file = f"excel_template_{company_code}.xlsx"
    
    if company_code == "transin":
        # Transin format - NO DateArrival, DateDelivery, TruckType
        columns = [
            'FreightBillNo', 'InvoiceDate', 'DueDate', 'FromLocation',
            'ShipmentDate', 'LRNo', 'Destination', 'CNNumber', 'TruckNo',
            'InvoiceNo', 'Pkgs', 'WeightKgs', 'FreightAmt', 'ToPointCharges', 
            'UnloadingCharge', 'SourceDetention', 'DestinationDetention'
        ]
        
        sample_data = {
            'FreightBillNo': ['DBLT1-2526-228'],
            'InvoiceDate': ['2026-01-18'],
            'DueDate': ['2026-02-18'],
            'FromLocation': ['Kichha'],
            'ShipmentDate': ['2025-12-09'],
            'LRNo': ['11376'],
            'Destination': ['Ahmedabad'],
            'CNNumber': ['DT1225559770'],
            'TruckNo': ['UP21ET3805'],
            'InvoiceNo': ['F22511136438'],
            'Pkgs': [282],
            'WeightKgs': [15390],
            'FreightAmt': [38530],
            'ToPointCharges': [0],
            'UnloadingCharge': [400],
            'SourceDetention': [0],
            'DestinationDetention': [0]
        }
    else:
        # Basic format - with all fields
        columns = [
            'FreightBillNo', 'InvoiceDate', 'DueDate', 'FromLocation',
            'ShipmentDate', 'LRNo', 'Destination', 'CNNumber', 'TruckNo',
            'InvoiceNo', 'Pkgs', 'WeightKgs', 'DateArrival', 'DateDelivery',
            'TruckType', 'FreightAmt', 'ToPointCharges', 'UnloadingCharge',
            'SourceDetention', 'DestinationDetention'
        ]
        
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
    df.to_excel(template_file, index=False)
    print(f"✓ Template created: {template_file}")
    return template_file


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
    return render_template("index.html", companies=COMPANIES)


@app.route("/download-template")
def download_template():
    """Download Excel template"""
    company_code = request.args.get("company", "stc")
    template_file = create_excel_template(company_code)
    return send_file(template_file, as_attachment=True, download_name=f"{company_code.upper()}_Template.xlsx")


@app.route("/api/companies")
def get_companies():
    """Get list of all companies"""
    companies_list = []
    for code, data in COMPANIES.items():
        companies_list.append({
            "code": code,
            "name": data["name"],
            "address": data["address_line1"]
        })
    return jsonify(companies_list)


@app.route("/api/company/<company_code>")
def get_company(company_code):
    """Get specific company details"""
    if company_code in COMPANIES:
        return jsonify(COMPANIES[company_code])
    return jsonify({"error": "Company not found"}), 404


@app.route("/", methods=["POST"])
def generate_bills():
    """Generate PDF bills from Excel"""
    try:
        if "file" not in request.files:
            return jsonify({"error": "No file uploaded"}), 400

        file = request.files["file"]
        company_code = request.form.get("company", "stc")
        
        if file.filename == "":
            return jsonify({"error": "No file selected"}), 400

        if company_code not in COMPANIES:
            return jsonify({"error": "Invalid company selected"}), 400

        print(f"📄 File received: {file.filename}")
        print(f"🏢 Company: {COMPANIES[company_code]['name']}")
        
        path = os.path.join(UPLOAD_FOLDER, file.filename)
        file.save(path)

        # Read Excel
        df = pd.read_excel(path)
        print(f"✓ Excel loaded: {len(df)} rows")
        
        # Convert dates
        date_columns = ['InvoiceDate', 'DueDate', 'ShipmentDate']
        if company_code != "transin":
            date_columns.extend(['DateArrival', 'DateDelivery'])
            
        for col in date_columns:
            if col in df.columns:
                df[col] = pd.to_datetime(df[col], errors='coerce')
        
        # Generate PDFs
        pdf_files = generate_multiple_pdfs(df, company_code)
        print(f"✓ Generated {len(pdf_files)} PDF(s)")
        
        # Save to history
        bill_numbers = df['FreightBillNo'].unique().tolist()
        history_entry = {
            "time": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "file": file.filename,
            "company": COMPANIES[company_code]["name"],
            "company_code": company_code,
            "rows": len(df),
            "bills": [str(b) for b in bill_numbers[:5]],
            "pdf_files": [os.path.basename(f) for f in pdf_files]
        }
        save_history(history_entry)
        
        # Create ZIP
        zip_filename = f"{company_code.upper()}_Bills.zip"
        zip_path = os.path.join(OUTPUT_FOLDER, zip_filename)
        with zipfile.ZipFile(zip_path, 'w') as zipf:
            for pdf_file in pdf_files:
                zipf.write(pdf_file, os.path.basename(pdf_file))
        
        print(f"✓ ZIP created: {zip_path}")
        return send_file(zip_path, as_attachment=True, download_name=zip_filename)
        
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
        
        # Convert dates
        date_columns = ['InvoiceDate', 'DueDate', 'ShipmentDate', 'DateArrival', 'DateDelivery']
        for col in date_columns:
            if col in df.columns:
                df[col] = pd.to_datetime(df[col], errors='coerce')
        
        # Get ALL rows
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
        
        os.remove(path)
        
        return jsonify({
            "ok": True,
            "count": len(df),
            "rows": rows
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


def generate_multiple_pdfs(df, company_code):
    """Generate separate PDF for each FreightBillNo"""
    pdf_files = []
    grouped = df.groupby('FreightBillNo')
    
    for bill_no, group_df in grouped:
        print(f"  → Generating: {bill_no}")
        pdf_path = generate_pdf(group_df.reset_index(drop=True), company_code)
        pdf_files.append(pdf_path)
    
    return pdf_files


def generate_pdf(df, company_code):
    """Generate PDF based on company type"""
    company = COMPANIES[company_code]
    
    if company.get("type") == "transin":
        return generate_transin_pdf(df, company_code)
    else:
        return generate_basic_pdf(df, company_code)


def generate_transin_pdf(df, company_code):
    """Generate Transin Logistics format PDF"""
    company = COMPANIES[company_code]
    
    bill_no = str(df.iloc[0]["FreightBillNo"]).replace("/", "_")
    pdf_path = f"{OUTPUT_FOLDER}/{company_code}_{bill_no}.pdf"

    c = canvas.Canvas(pdf_path, pagesize=landscape(A4))
    width, height = landscape(A4)

    margin = 15
    c.rect(margin, margin, width - 2*margin, height - 2*margin, stroke=1, fill=0)

    # Header
    c.setFont("Helvetica-Bold", 14)
    c.drawCentredString(width / 2, height - 50, company["name"])
    
    c.setFont("Helvetica", 8)
    c.drawCentredString(width / 2, height - 65, company["address_line1"])
    c.drawCentredString(width / 2, height - 78, company["address_line2"])
    
    c.setFont("Helvetica-Bold", 12)
    c.drawCentredString(width / 2, height - 100, "INVOICE")
    
    c.setFont("Helvetica", 7)
    c.drawCentredString(width / 2, height - 113, "@ This is system generated invoice")

    # Left Box - Customer
    box_top = height - 140
    c.rect(30, box_top - 90, 260, 90, stroke=1, fill=0)
    
    c.setFont("Helvetica-Bold", 9)
    c.drawString(40, box_top - 18, "To,")
    
    c.setFont("Helvetica", 8)
    c.drawString(40, box_top - 32, company["customer"]["name"])
    c.drawString(40, box_top - 45, company["customer"]["address_line1"])
    c.drawString(40, box_top - 58, company["customer"]["address_line2"])
    c.drawString(40, box_top - 73, f"GSTIN: {company['customer']['gstin']}")

    # Right Box - Invoice Details
    c.rect(width - 280, box_top - 90, 250, 90, stroke=1, fill=0)
    
    c.setFont("Helvetica-Bold", 9)
    c.drawString(width - 270, box_top - 18, f"Freight Bill No: {df.iloc[0]['FreightBillNo']}")
    
    c.setFont("Helvetica", 8)
    c.drawString(width - 270, box_top - 35, f"Invoice Date: {df.iloc[0]['InvoiceDate'].strftime('%d %b %Y')}")
    c.drawString(width - 270, box_top - 50, f"Due Date: {df.iloc[0]['DueDate'].strftime('%d %b %Y')}")
    
    c.setFont("Helvetica-Bold", 8)
    c.drawString(width - 270, box_top - 68, f"Our PAN No. {company['bank']['pan']}")

    c.setFont("Helvetica", 7)
    c.drawString(30, box_top - 105, f"From location: {df.iloc[0]['FromLocation']}")

    # Table - Transin Format (NO DateArrival, DateDelivery, TruckType)
    table_top = box_top - 130
    table_left = 30
    
    headers = [
        "S.\nno.", "Shipment\nDate", "LR\nNo.", "Destination", "CN\nNumber",
        "Truck No", "Invoice No", "Pkgs", "Weight\n(kgs)", 
        "Freight\nAmt (Rs.)", "To Point\nCharges\n(Rs.)", "Unloading\nCharge\n(Rs.)", 
        "Source\nDetention\n(Rs.)", "Destination\nDetention\n(Rs.)", "Total\nAmount\n(Rs.)"
    ]
    
    col_widths = [25, 50, 35, 55, 50, 50, 85, 30, 45, 50, 50, 50, 50, 55, 55]
    total_col_width = sum(col_widths)
    
    # Header Row
    c.setFillColor(colors.lightgrey)
    c.rect(table_left, table_top - 30, total_col_width, 30, stroke=1, fill=1)
    
    c.setFillColor(colors.black)
    c.setFont("Helvetica-Bold", 7)
    
    x = table_left
    for i, header in enumerate(headers):
        lines = header.split('\n')
        y_offset = table_top - 8
        for line in lines:
            c.drawCentredString(x + col_widths[i]/2, y_offset, line)
            y_offset -= 8
        x += col_widths[i]
    
    # Vertical lines in header
    x = table_left
    for width_val in col_widths:
        c.line(x, table_top, x, table_top - 30)
        x += width_val
    c.line(x, table_top, x, table_top - 30)
    
    # Data Rows
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
            row["ShipmentDate"].strftime("%d-%m-%Y"),
            str(row["LRNo"]),
            str(row["Destination"]),
            str(row["CNNumber"]),
            str(row["TruckNo"]),
            str(row["InvoiceNo"]),
            str(int(row["Pkgs"])),
            str(int(row["WeightKgs"])),
            f"{float(row['FreightAmt']):.1f}",
            f"{float(row['ToPointCharges']):.1f}",
            f"{float(row['UnloadingCharge']):.1f}",
            f"{float(row['SourceDetention']):.1f}",
            f"{float(row['DestinationDetention']):.1f}",
            f"{row_total:.1f}"
        ]
        
        wrap_columns = {6}  # Invoice No
        
        x = table_left
        for i, val in enumerate(values):
            if i in wrap_columns:
                draw_wrapped_text(c, val, x + col_widths[i]/2, y + row_height/2, 
                                col_widths[i] - 6, "Helvetica", 6, 7)
            else:
                c.drawCentredString(x + col_widths[i]/2, y + row_height/2, val)
            x += col_widths[i]
        
        c.line(table_left, y, table_left + total_col_width, y)
    
    # Vertical lines
    x = table_left
    for width_val in col_widths:
        c.line(x, table_top - 30, x, y)
        x += width_val
    c.line(x, table_top - 30, x, y)
    
    # Total Row
    total_row_height = 20
    y -= total_row_height
    
    c.rect(table_left, y, total_col_width, total_row_height, stroke=1, fill=0)
    
    c.setFont("Helvetica-Bold", 8)
    rupees = int(total_amount)
    paise = int((total_amount - rupees) * 100)
    
    if paise > 0:
        total_words = f"{num2words(rupees, lang='en_IN').title()} Rupees and {num2words(paise, lang='en_IN').title()} Paise"
    else:
        total_words = f"{num2words(rupees, lang='en_IN').title()} Rupees"
    
    c.drawString(table_left + 5, y + 8, f"Total in words (Rs.) : {total_words} Only")
    
    c.setFont("Helvetica-Bold", 9)
    total_col_x = table_left + sum(col_widths[:-1])
    c.drawCentredString(total_col_x + col_widths[-1]/2, y + 8, f"{total_amount:.1f}")
    
    c.line(total_col_x, y, total_col_x, y + total_row_height)

    # Notes
    c.setFont("Helvetica", 6)
    note_y = y - 12
    c.drawString(30, note_y, 'Any changes or discrepancies should be highlighted within 5 working days else it would be considered final. Please send all remittance details to "receivables@onmove.in".')
    
    note_y -= 10
    c.drawString(30, note_y, company.get("gst_note", ""))
    
    note_y -= 10
    c.setFont("Helvetica", 7)
    c.drawString(30, note_y, f"Transin GSTIN {company['bank']['gstin']}")
    c.drawString(200, note_y, f"SAC code {company['bank']['sac_code']}")
    c.drawString(350, note_y, f"Transin State Code {company['bank']['state_code']}")
    
    note_y -= 10
    c.drawString(30, note_y, "Tax Details - 5% IGST or (2.5% SGST+2.5% CGST) as applicable")

    # Digital Signature
    sig_y = note_y - 30
    c.setFont("Helvetica-Bold", 9)
    c.drawRightString(width - 35, sig_y, f"For {company['name']}")
    
    c.setFont("Helvetica", 7)
    c.drawRightString(width - 35, sig_y - 40, "(Authorized Signatory)")
    
    if "digital_signature" in company:
        c.setFont("Helvetica-Oblique", 7)
        c.drawRightString(width - 35, sig_y - 55, f"Digitally signed by {company['digital_signature']}")
    
    c.line(width - 180, sig_y - 42, width - 35, sig_y - 42)

    c.save()
    return pdf_path


def generate_basic_pdf(df, company_code):
    """Generate basic format PDF (STC, NTC, WTC)"""
    company = COMPANIES[company_code]
    
    bill_no = str(df.iloc[0]["FreightBillNo"]).replace("/", "_")
    pdf_path = f"{OUTPUT_FOLDER}/{company_code}_{bill_no}.pdf"

    c = canvas.Canvas(pdf_path, pagesize=landscape(A4))
    width, height = landscape(A4)

    margin = 15
    c.rect(margin, margin, width - 2*margin, height - 2*margin, stroke=1, fill=0)

    # Logo
    try:
        if os.path.exists(company["logo"]):
            try:
                img = Image.open(company["logo"])
                img.verify()
                c.drawImage(company["logo"], 55, height - 140, width=100, height=80, preserveAspectRatio=True)
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
    c.drawCentredString(width / 2, height - 70, company["name"])
    
    c.setFont("Helvetica", 9)
    c.drawCentredString(width / 2, height - 85, company["address_line1"])
    c.drawCentredString(width / 2, height - 98, company["address_line2"])
    
    c.setFont("Helvetica-Bold", 12)
    c.drawCentredString(width / 2, height - 125, "INVOICE")

    # Left Box
    box_top = height - 160
    c.rect(30, box_top - 110, 260, 110, stroke=1, fill=0)
    
    c.setFont("Helvetica-Bold", 10)
    c.drawString(40, box_top - 20, "To,")
    
    c.setFont("Helvetica", 9)
    c.drawString(40, box_top - 35, company["customer"]["name"])
    c.drawString(40, box_top - 50, company["customer"]["address_line1"])
    c.drawString(40, box_top - 65, company["customer"]["address_line2"])
    c.drawString(40, box_top - 85, f"GSTIN: {company['customer']['gstin']}")

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
    c.drawString(30, note_y, f'Any changes or discrepancies should be highlighted within 5 working days else it would be considered final. Please send all remittance details to "{company["bank"]["email"]}".')

    # Bank Details
    bank_y = note_y - 25
    
    bank_details = [
        ("Our PAN No.", company["bank"]["pan"]),
        (f"{company_code.upper()} GSTIN", company["bank"]["gstin"]),
        (f"{company_code.upper()} State Code", company["bank"]["state_code"]),
        ("Account name", company["bank"]["account_name"]),
        ("Account no", company["bank"]["account_no"]),
        ("IFS Code", company["bank"]["ifsc"])
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
    c.drawRightString(width - 35, sig_y, f"For {company['name']}")
    
    c.setFont("Helvetica", 7)
    c.drawRightString(width - 35, sig_y - 50, "(Authorized Signatory)")
    c.line(width - 180, sig_y - 52, width - 35, sig_y - 52)

    c.save()
    return pdf_path


if __name__ == "__main__":
    app.run(debug=True, port=5000)