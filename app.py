from flask import Flask, render_template, request, send_file
import pandas as pd
import os
import uuid

app = Flask(__name__)

UPLOAD_FOLDER = "uploads"
OUTPUT_FOLDER = "output"

os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)


# ---------------- PDF FUNCTION (EXACT NAME IMPORTANT) ----------------
def generate_invoice_pdf(row, pdf_path):
    """
    ⚠️ YAHAN TERA PURANA PDF CODE HOGA
    ⚠️ LOGO / FORMAT / ALIGNMENT KUCH CHANGE NAHI
    """
    from reportlab.lib.pagesizes import A4
    from reportlab.pdfgen import canvas

    c = canvas.Canvas(pdf_path, pagesize=A4)

    # Example (replace with your existing logic)
    c.drawString(50, 800, "Transport Bill")
    c.drawString(50, 780, f"LR No: {row.get('LR No','')}")

    c.save()


# ---------------- MAIN ROUTE ----------------
@app.route("/", methods=["GET", "POST"])
def index():
    if request.method == "POST":

        if "file" not in request.files:
            return "No file part"

        file = request.files["file"]

        if file.filename == "":
            return "No selected file"

        # Save Excel
        excel_path = os.path.join(
            UPLOAD_FOLDER, f"{uuid.uuid4()}_{file.filename}"
        )
        file.save(excel_path)

        # Read Excel
        df = pd.read_excel(excel_path)

        pdf_files = []

        for _, row in df.iterrows():
            pdf_name = f"{uuid.uuid4()}.pdf"
            pdf_path = os.path.join(OUTPUT_FOLDER, pdf_name)

            generate_invoice_pdf(row, pdf_path)
            pdf_files.append(pdf_path)

        # For now return first PDF (safe)
        return send_file(pdf_files[0], as_attachment=True)

    return render_template("index.html")


# ---------------- RUN ----------------
if __name__ == "__main__":
    app.run(debug=True)
