import pandas as pd
import glob   # we use this library when we hava multiple file, and we save those path inside a list
from fpdf import FPDF
from pathlib import Path
from datetime import datetime



filepaths = glob.glob('invoices/*.xlsx')   # *.xlsx mean that we import every file in the folder that is xlsx type
print(filepaths)

for path in filepaths:

    pdf = FPDF(orientation='P', unit='mm', format='A4')
    pdf.add_page()

    filename = Path(path).stem
    invoice_nr = filename.split('-')[0]
    format_date = datetime.now().strftime('%d-%m-%y')

    # The name of the file
    pdf.set_font(family="Arial", size=16, style="B")
    pdf.cell(w=50, h=8, txt=f"Invoice number.{invoice_nr}", ln=1)

    # The date the file was created
    pdf.set_font(family="Arial", size=16, style="B")
    pdf.cell(w=50, h=12, txt=f"Date: {format_date}", ln=20)
    pdf.cell(w=0, h=10, ln=1)


    df = pd.read_excel(path, sheet_name="Sheet 1")  # when we use multiple files we need to define the key sheet file

    # Add a header
    columns_a = df.columns
    columns_a = [item.replace("_", " ").title() for item in columns_a]  # this is exactly like for loop just in one line
    pdf.set_font(family="Arial", size=10, style='B')
    pdf.set_text_color(40, 40, 40)
    pdf.cell(w=30, h=8, txt=str(columns_a[0]), border=1, align='C')
    pdf.cell(w=70, h=8, txt=str(columns_a[1]), border=1, align='C')
    pdf.cell(w=35, h=8, txt=str(columns_a[2]), border=1, align='C')
    pdf.cell(w=30, h=8, txt=str(columns_a[3]), border=1, align='C')
    pdf.cell(w=30, h=8, txt=str(columns_a[4]), border=1, ln=1, align='C')

    # Add rows to the table
    for index, row in df.iterrows():
        pdf.set_font(family="Arial", size=12)
        pdf.set_text_color(80, 80, 80)
        pdf.cell(w=30, h=8, txt=str(row['product_id']), border=1, align='C')
        pdf.cell(w=70, h=8, txt=str(row['product_name']), border=1, align='C')
        pdf.cell(w=35, h=8, txt=str(row['amount_purchased']), border=1, align='C')
        pdf.cell(w=30, h=8, txt=str(row['price_per_unit']), border=1, align='C')
        pdf.cell(w=30, h=8, txt=str(row['total_price']), border=1, ln=1, align='C')


    pdf.output(f'PDFs/{filename}.pdf')


