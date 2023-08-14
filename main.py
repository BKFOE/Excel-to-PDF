import pandas as pd
import glob
from fpdf import FPDF
from pathlib import Path

# Glob allows us to access all the Excel files in the folder
filepaths = glob.glob("invoices/*.xlsx")

# Get the input data and load into python
for filepath in filepaths:
    df = pd.read_excel(filepath, sheet_name="Sheet 1")
    pdf = FPDF(orientation="P", unit="mm", format="A4")
    pdf.add_page()

    # Extract invoice nr
    """ filename = filepath.strip("invoices/")
    new_filename = (filename[:5])
    print(new_filename)
    """
    filename = Path(filepath).stem
    invoice_nr = filename.split("-")[0]

    # Set the header
    pdf.set_font(family="Times", style="B", size=16)
    pdf.cell(w=50, h=8, txt=f"Invoice nr. {invoice_nr}")
    pdf.output(f"PDF/ {filename}.pdf")
