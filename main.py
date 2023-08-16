import pandas as pd
import glob
from fpdf import FPDF
from pathlib import Path

# Glob allows us to access all the Excel files in the folder
filepaths = glob.glob("invoices/*.xlsx")

# Get the input data and load into python
for filepath in filepaths:

    pdf = FPDF(orientation="P", unit="mm", format="A4")
    pdf.add_page()

    # Extract invoice nr
    """ filename = filepath.strip("invoices/")
    new_filename = (filename[:5])
    print(new_filename)
    """
    filename = Path(filepath).stem
    invoice_no, date = filename.split("-")
    # invoice_no = filename.split("-")[0]
    # date = filename.split("-")[1]

    # Set the header
    pdf.set_font(family="Times", style="B", size=16)
    pdf.cell(w=50, h=8, txt=f"Invoice no. {invoice_no}", ln=1)

    # Set the date header
    pdf.set_font(family="Times", style="B", size=16)
    pdf.cell(w=50, h=8, txt=f"Date: {date}", ln=1)

    # Add column headers
    df = pd.read_excel(filepath, sheet_name="Sheet 1")
    columns = df.columns
    columns = [item.replace("_", " ").title() for item in columns]
    pdf.set_font(family="Times", style="B", size=10)
    pdf.set_text_color(80, 80, 80)
    pdf.cell(w=30, h=8, txt=columns[0], border=1)
    pdf.cell(w=70, h=8, txt=columns[1], border=1)
    pdf.cell(w=35, h=8, txt=columns[2], border=1)
    pdf.cell(w=30, h=8, txt=columns[3], border=1)
    pdf.cell(w=30, h=8, txt=columns[4], border=1, ln=1)

    # Add rows
    for index, row in df.iterrows():
        pdf.set_font(family="Times", size=10)
        pdf.set_text_color(80, 80, 80)
        pdf.cell(w=30, h=8, txt=str(row["product_id"]), border=1)
        pdf.cell(w=70, h=8, txt=str(row["product_name"]), border=1)
        pdf.cell(w=35, h=8, txt=str(row["amount_purchased"]), border=1)
        pdf.cell(w=30, h=8, txt=str(row["price_per_unit"]), border=1)
        pdf.cell(w=30, h=8, txt=str(row["total_price"]), border=1, ln=1)

    pdf.output(f"PDF/ {filename}.pdf")
