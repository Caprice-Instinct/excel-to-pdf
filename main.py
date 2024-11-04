import pandas as pd
import glob
from fpdf import FPDF
from pathlib import Path

# Use glob to access the multiple filepaths
filepaths = glob.glob("invoices/*.xlsx")

# For loop to extract data from the different filepaths
for filepath in filepaths:
    # Create dataframes
    df = pd.read_excel(filepath, sheet_name="Sheet 1")

    # Extract the filename from filepath without the extension
    filename = Path(filepath).stem

    # Extract the invoice number and date
    invoice_nr, invoice_date = filename.split("-")

    # Pdf creation
    pdf = FPDF(orientation="P", unit="mm", format="A4")
    pdf.add_page()

    pdf.set_font(family="Times", style="B", size=16)
    pdf.cell(w=50, h=8, txt=f"Invoice nr. {invoice_nr}", ln=1)

    pdf.set_font(family="Times", style="B", size=16)
    pdf.cell(w=50, h=8, txt=f"Date: {invoice_date}", ln=1)

    # Produce pdf output
    pdf.output(f"PDFs/{filename}.pdf")