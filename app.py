from flask import Flask, render_template, request, send_file
import pandas as pd
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
from num2words import num2words
import os

app = Flask(__name__)

UPLOAD_FOLDER = "uploads"
OUTPUT_FOLDER = "output"
LOGO_PATH = "static/logo.png"

os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)


def wrap_text(c, text, x, y, width, leading=9):
    words = str(text).split()
    line = ""
    for word in words:
        if c.stringWidth(line + word + " ", "Helvetica", 7) <= width:
            line += word + " "
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
        try:
            # FIXED: Changed "excel" to "file" to match HTML form
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
    # Convert dates to datetime if they're not already
    date_columns = ['InvoiceDate', 'DueDate', 'ShipmentDate', 'DateArrival', 'DateDelivery']
    for col in date_columns:
        if col in df.columns:
            # Use coerce to handle any format automatically
            df[col] = pd.to_datetime(df[col], errors='coerce')
    
    bill_no = str(df.iloc[0]["FreightBillNo"]).replace("/", "_")
    pdf_path = f"{OUTPUT_FOLDER}/{bill_no}.pdf"

    c = canvas.Canvas(pdf_path, pagesize=A4)
    w, h = A4

    # ================= LOGO =================
    if os.path.exists(LOGO_PATH):
        c.drawImage(LOGO_PATH, 30, h - 100, width=90, height=50, mask="auto")

    # ================= HEADER =================
    c.setFont("Helvetica-Bold", 14)
    c.drawCentredString(w / 2, h - 55, "SOUTH TRANSPORT COMPANY")

    c.setFont("Helvetica", 9)
    c.drawCentredString(
        w / 2,
        h - 70,
        "Dehradun Road Near Power Grid Bhagwanpur, Roorkee, Haridwar, U.K. 247661, India"
    )

    c.setFont("Helvetica-Bold", 11)
    c.drawCentredString(w / 2, h - 90, "INVOICE")

    # ================= LEFT BOX =================
    c.rect(25, h - 195, 260, 90)
    c.setFont("Helvetica", 9)
    c.drawString(30, h - 120, "To,")
    c.drawString(30, h - 135, "Grivaa Springs Private Ltd.")
    c.drawString(30, h - 150, "Khasra no 135, Tanshipur, Roorkee")
    c.drawString(30, h - 165, "Roorkee, Uttarakhand 247656")
    c.drawString(30, h - 180, "GSTIN: 05AAICG4793P1ZV")

    # ================= RIGHT BOX =================
    c.rect(w - 255, h - 195, 230, 90)
    c.drawString(w - 245, h - 130, f"Freight Bill No : {bill_no}")
    c.drawString(
        w - 245, h - 150,
        f"Invoice Date : {df.iloc[0]['InvoiceDate'].strftime('%d-%m-%Y')}"
    )
    c.drawString(
        w - 245, h - 170,
        f"Due Date : {df.iloc[0]['DueDate'].strftime('%d-%m-%Y')}"
    )

    c.drawString(30, h - 210, f"From location: {df.iloc[0]['FromLocation']}")

    # ================= TABLE HEADER =================
    y = h - 250
    c.setFont("Helvetica-Bold", 7)

    headers = [
        "S.\nNo", "Shipment\nDate", "LR No", "Destination", "CN\nNumber",
        "Truck No", "Invoice No", "Pkgs", "Weight\n(kgs)",
        "Date of\nArrival", "Date of\nDelivery", "Truck\nType",
        "Freight\nAmt", "To Point\nCharges", "Unloading\nCharge",
        "Source\nDetention", "Destination\nDetention", "Total\nAmount"
    ]

    col_widths = [25, 45, 35, 50, 40, 45, 50, 30, 40, 45, 45, 40, 45, 45, 45, 45, 50, 50]

    x = 25
    for i, htxt in enumerate(headers):
        c.drawString(x + 2, y, htxt)
        x += col_widths[i]

    c.line(25, y - 2, w - 25, y - 2)
    y -= 18

    # ================= TABLE DATA =================
    c.setFont("Helvetica", 7)
    total = 0

    for idx, row in df.iterrows():
        x = 25

        row_total = (
            row["FreightAmt"]
            + row["ToPointCharges"]
            + row["UnloadingCharge"]
            + row["SourceDetention"]
            + row["DestinationDetention"]
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

        min_y = y
        for i, val in enumerate(values):
            ny = wrap_text(c, val, x + 2, y, col_widths[i] - 4)
            min_y = min(min_y, ny)
            x += col_widths[i]

        y = min_y - 10
        total += row_total

    # ================= TOTAL =================
    c.setFont("Helvetica-Bold", 8)
    c.drawRightString(w - 30, y, f"{total:.2f}")

    words = num2words(total, to="currency", lang="en_IN").title()
    c.drawString(30, y, f"Total in words (Rs.): {words} Only")

    c.save()
    return pdf_path


if __name__ == "__main__":
    app.run(debug=True)