import pandas as pd
import glob
from fpdf import FPDF
from pathlib import Path

# Use glob to access the multiple filepaths
filepaths = glob.glob("invoices/*.xlsx")

# For loop to extract data from the different filepaths
for filepath in filepaths:
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

    # Create dataframes
    df = pd.read_excel(filepath, sheet_name="Sheet 1")

    # Create a list for the Excel columns
    columns = list(df.columns)

    # Modify the column headers
    # columns = [item.replace("_", " ").title() for item in columns]
    # Remove underscore
    columns = [item.replace("_", " ") for item in columns]
    # Title the headers
    columns = [item.title() for item in columns]

    # Display the column headers
    pdf.set_font(family="Times", size=10, style="B")

    pdf.cell(w=30, h=10, txt=columns[0], border=1)
    pdf.cell(w=50, h=10, txt=columns[1], border=1)
    pdf.cell(w=30, h=10, txt=columns[2], border=1)
    pdf.cell(w=30, h=10, txt=columns[3], border=1)
    pdf.cell(w=30, h=10, txt=columns[4], border=1, ln=1)

    # Add rows to the table
    for index, row in df.iterrows():
        pdf.set_font(family="Times", size=10)
        pdf.set_text_color(80, 80, 80)
        pdf.cell(w=30, h=10, txt=str(row['product_id']), border=1)
        pdf.cell(w=50, h=10, txt=str(row['product_name']), border=1)
        pdf.cell(w=30, h=10, txt=str(row['amount_purchased']), border=1)
        pdf.cell(w=30, h=10, txt=str(row['price_per_unit']), border=1)
        pdf.cell(w=30, h=10, txt=str(row['total_price']), border=1, ln=1)

    # Add total price for invoice
    total_sum = df['total_price'].sum()

    pdf.set_font(family="Times", size=10)
    pdf.set_text_color(80, 80, 80)
    pdf.cell(w=30, h=10, txt="", border=1)
    pdf.cell(w=50, h=10, txt="", border=1)
    pdf.cell(w=30, h=10, txt="", border=1)
    pdf.cell(w=30, h=10, txt="", border=1)
    pdf.cell(w=30, h=10, txt=str(total_sum), border=1, ln=1)

    pdf.ln(10)
    # Total sum sentence
    pdf.set_font(family="Times", size=10, style="B")
    pdf.cell(w=30, h=10, txt=f"The total price is {total_sum}", ln=1)

    # Add company name and logo
    pdf.cell(w=15, h=10, txt=f"LinetCo")
    pdf.image("LinetCo.png", w=10)

    # Produce pdf output
    pdf.output(f"PDFs/{filename}.pdf")