import pandas as pd
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
from num2words import num2words
import os

# Test your Excel file directly
excel_path = "test.xlsx"  # Replace with your Excel filename

print("Step 1: Reading Excel...")
try:
    df = pd.read_excel(excel_path)
    print(f"✓ Excel loaded successfully! Rows: {len(df)}")
    print(f"✓ Columns: {list(df.columns)}")
    print("\nFirst row data:")
    print(df.iloc[0])
except Exception as e:
    print(f"✗ Error reading Excel: {e}")
    import traceback
    traceback.print_exc()
    exit()

print("\n\nStep 2: Converting dates...")
try:
    date_columns = ['InvoiceDate', 'DueDate', 'ShipmentDate', 'DateArrival', 'DateDelivery']
    for col in date_columns:
        if col in df.columns:
            print(f"Converting {col}...")
            df[col] = pd.to_datetime(df[col], format='%d-%m-%Y', errors='coerce')
    print("✓ Dates converted successfully!")
except Exception as e:
    print(f"✗ Error converting dates: {e}")
    import traceback
    traceback.print_exc()
    exit()

print("\n\nStep 3: Creating PDF...")
try:
    bill_no = str(df.iloc[0]["FreightBillNo"]).replace("/", "_")
    print(f"Bill No: {bill_no}")
    
    pdf_path = f"test_{bill_no}.pdf"
    c = canvas.Canvas(pdf_path, pagesize=A4)
    w, h = A4
    
    # Test basic drawing
    c.setFont("Helvetica-Bold", 14)
    c.drawCentredString(w / 2, h - 55, "SOUTH TRANSPORT COMPANY")
    
    print("✓ PDF canvas created!")
    
    # Try formatting date
    test_date = df.iloc[0]['InvoiceDate']
    print(f"Test date: {test_date}")
    formatted = test_date.strftime('%d-%m-%Y')
    print(f"Formatted date: {formatted}")
    
    c.save()
    print(f"✓ PDF saved successfully as {pdf_path}")
    
except Exception as e:
    print(f"✗ Error creating PDF: {e}")
    import traceback
    traceback.print_exc()
    exit()

print("\n\n✅ All tests passed! Your data format is correct.")