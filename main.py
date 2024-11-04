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

    # Extract the first number of the filename only
    invoice_nr = filename.split("-")[0]

    # Pdf creation
    pdf = FPDF(orientation="P", unit="mm", format="A4")
    pdf.add_page()
    pdf.set_font(family="Times", style="B", size=16)
    pdf.cell(w=50, h=8, txt=f"Invoice nr. {invoice_nr}")
    pdf.output(f"PDFs/{filename}.pdf")