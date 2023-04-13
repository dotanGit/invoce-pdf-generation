import pandas as pd
import glob   # we use this library when we hava multiple file, and we save those path inside a list
from fpdf import FPDF
from pathlib import Path
from datetime import datetime



filepaths = glob.glob('invoices/*.xlsx')   # *.xlsx mean that we import every file in the folder that is xlsx type
print(filepaths)

for path in filepaths:
    df = pd.read_excel(path, sheet_name="Sheet 1")   # when we use multiple files we need to define the key sheet file
    pdf = FPDF(orientation='P', unit='mm', format='A4')
    pdf.add_page()

    filename = Path(path).stem
    invoice_nr = filename.split('-')[0]
    format_date = datetime.now().strftime('%d-%m-%y')

    pdf.set_font(family="Arial", size=16, style="B")
    pdf.cell(w=50, h=8, txt=f"Invoice number.{invoice_nr}", ln=1)

    pdf.set_font(family="Arial", size=16, style="B")
    pdf.cell(w=50, h=8, txt=f"Date: {format_date}")



    pdf.output(f'PDFs/{filename}.pdf')


