from flask import Flask, render_template, request, send_file
import os, zipfile
import pandas as pd
from datetime import datetime
from num2words import num2words
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.units import mm
from reportlab.lib.utils import ImageReader

app = Flask(__name__)

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
UPLOAD_FOLDER = os.path.join(BASE_DIR, "uploads")
OUTPUT_FOLDER = os.path.join(BASE_DIR, "output")

os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)

# ================= FIXED DETAILS =================
FIXED_PARTY = {
    "PartyName": "Grivaa Springs Private Ltd.",
    "PartyAddress": "Khasra no 135, Tansipur, Roorkee",
    "PartyCity": "Roorkee",
    "PartyState": "Uttarakhand",
    "PartyPincode": "247656",
    "PartyGSTIN": "05AAICG4793P1ZV",
}

FIXED_BANK = {
    "PAN": "BSSPG9414K",
    "GSTIN": "05BSSPG9414K1ZA",
    "STATE": "5",
    "ACC_NAME": "South Transport Company",
    "ACC_NO": "364205500142",
    "IFS": "ICIC0003642",
}

REQUIRED_HEADERS = [
    "FreightBillNo","InvoiceDate","DueDate","FromLocation","ShipmentDate",
    "LRNo","Destination","CNNumber","TruckNo","InvoiceNo","Pkgs","WeightKgs",
    "DateArrival","DateDelivery","TruckType","FreightAmt","ToPointCharges",
    "UnloadingCharge","SourceDetention","DestinationDetention"
]

# ================= HELPERS =================
def s(v): return "" if pd.isna(v) else str(v)
def f(v): return float(v) if str(v).strip() else 0.0
def d(v): return pd.to_datetime(v, dayfirst=True).strftime("%d %b %Y")
def money(v): return f"{v:.2f}"

def total(row):
    return (
        f(row["FreightAmt"]) +
        f(row["ToPointCharges"]) +
        f(row["UnloadingCharge"]) +
        f(row["SourceDetention"]) +
        f(row["DestinationDetention"])
    )

# ================= PDF =================
def generate_pdf(row, path):
    row = {**FIXED_PARTY, **FIXED_BANK, **row}
    W, H = landscape(A4)
    c = canvas.Canvas(path, pagesize=(W, H))

    LM, TM, BM = 10*mm, 10*mm, 10*mm
    c.rect(LM, BM, W-2*LM, H-TM-BM)

    # Header
    c.setFont("Helvetica-Bold", 14)
    c.drawCentredString(W/2, H-20, "SOUTH TRANSPORT COMPANY")
    c.setFont("Helvetica", 8)
    c.drawCentredString(W/2, H-34, "Dehradun Road Near Power Grid Bhagwanpur")
    c.drawCentredString(W/2, H-46, "Roorkee, Haridwar, U.K. 247661")
    c.setFont("Helvetica-Bold", 10)
    c.drawCentredString(W/2, H-65, "INVOICE")

    # Logo
    logo = os.path.join(BASE_DIR, "logo.png")
    if os.path.exists(logo):
        c.drawImage(logo, LM+10, H-120, 120, 60, mask="auto")

    # Left Box
    c.rect(LM+10, H-210, 350, 80)
    c.setFont("Helvetica-Bold", 8)
    c.drawString(LM+15, H-150, "To,")
    c.drawString(LM+15, H-165, row["PartyName"])
    c.setFont("Helvetica", 8)
    c.drawString(LM+15, H-180, row["PartyAddress"])
    c.drawString(LM+15, H-195, f'{row["PartyCity"]}, {row["PartyState"]} {row["PartyPincode"]}')
    c.setFont("Helvetica-Bold", 8)
    c.drawString(LM+15, H-210, f'GSTIN: {row["PartyGSTIN"]}')

    c.setFont("Helvetica", 8)
    c.drawString(LM+10, H-230, f'From location: {row["FromLocation"]}')

    # Right Box
    c.rect(W-260, H-210, 240, 80)
    c.setFont("Helvetica-Bold", 8)
    c.drawString(W-250, H-160, f'Freight Bill No: {row["FreightBillNo"]}')
    c.drawString(W-250, H-180, f'Invoice Date: {d(row["InvoiceDate"])}')
    c.drawString(W-250, H-200, f'Due Date: {d(row["DueDate"])}')

    # Table Header
    y = H-270
    headers = [
        "S.No","Shipment Date","LR No","Destination","CN No","Truck No",
        "Invoice No","Pkgs","Weight","Arrival","Delivery","Truck Type",
        "Freight","To Point","Unload","Src Det","Dst Det","Total"
    ]
    widths = [30,70,50,80,50,70,80,40,60,60,60,70,70,60,60,60,60,70]

    x = LM+10
    c.setFont("Helvetica-Bold", 7)
    for h,w in zip(headers,widths):
        c.rect(x,y,w,25)
        c.drawCentredString(x+w/2,y+8,h)
        x+=w

    # Data Row
    y -= 25
    x = LM+10
    amt = total(row)

    values = [
        "1", d(row["ShipmentDate"]), row["LRNo"], row["Destination"],
        row["CNNumber"], row["TruckNo"], row["InvoiceNo"], row["Pkgs"],
        row["WeightKgs"], d(row["DateArrival"]), d(row["DateDelivery"]),
        row["TruckType"], money(f(row["FreightAmt"])), money(f(row["ToPointCharges"])),
        money(f(row["UnloadingCharge"])), money(f(row["SourceDetention"])),
        money(f(row["DestinationDetention"])), money(amt)
    ]

    c.setFont("Helvetica", 7)
    for v,w in zip(values,widths):
        c.rect(x,y,w,22)
        c.drawCentredString(x+w/2,y+7,str(v))
        x+=w

    # Total in words
    y -= 22
    c.rect(LM+10,y,sum(widths),22)
    words = num2words(int(amt)).title() + " Rupees Only"
    c.drawString(LM+15,y+7,f"Total in words (Rs.): {words}")
    c.drawRightString(LM+10+sum(widths)-5,y+7,money(amt))

    # Bank Box
    bx, by = LM+10, BM+40
    c.rect(bx,by,350,100)
    bank = [
        ("Our PAN No.",row["PAN"]),
        ("STC GSTIN",row["GSTIN"]),
        ("STC State Code",row["STATE"]),
        ("Account Name",row["ACC_NAME"]),
        ("Account No",row["ACC_NO"]),
        ("IFS Code",row["IFS"]),
    ]
    yy = by+80
    for k,v in bank:
        c.drawString(bx+10,yy,k)
        c.drawString(bx+150,yy,v)
        yy-=15

    # Sign
    c.drawString(W-300,by+60,"For SOUTH TRANSPORT COMPANY")
    c.line(W-300,by+30,W-80,by+30)
    c.drawString(W-200,by+10,"(Authorized Signatory)")

    c.showPage()
    c.save()

# ================= ROUTE =================
@app.route("/", methods=["GET","POST"])
def index():
    if request.method == "POST":
        file = request.files["file"]
        path = os.path.join(UPLOAD_FOLDER,file.filename)
        file.save(path)

        df = pd.read_excel(path)
        df.columns = [c.strip() for c in df.columns]

        for h in REQUIRED_HEADERS:
            if h not in df.columns:
                return f"Missing column: {h}"

        files=[]
        for _,r in df.iterrows():
            pdf = f'{r["FreightBillNo"]}.pdf'
            pdf_path = os.path.join(OUTPUT_FOLDER,pdf)
            generate_pdf(r,pdf_path)
            files.append(pdf_path)

        zip_path = os.path.join(OUTPUT_FOLDER,"Bills.zip")
        with zipfile.ZipFile(zip_path,"w") as z:
            for f in files:
                z.write(f,os.path.basename(f))

        return send_file(zip_path,as_attachment=True)

    return '''
    <form method="post" enctype="multipart/form-data">
      <input type="file" name="file">
      <button>Upload</button>
    </form>
    '''

if __name__ == "__main__":
    app.run(debug=True)
