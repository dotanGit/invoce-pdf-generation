import pandas as pd
import glob   # library for working with file paths
from fpdf import FPDF   # library for creating PDFs
from pathlib import Path   # library for working with file paths
from datetime import datetime   # library for working with dates and times

filepaths = glob.glob('invoices/*.xlsx')   # get all Excel files in the "invoices" folder

for path in filepaths:
    pdf = FPDF(orientation='P', unit='mm', format='A4')   # create new PDF with specified orientation, unit, and format
    pdf.add_page()

    filename = Path(path).stem   # get the name of the file without the file extension
    invoice_nr = filename.split('-')[0]   # extract the invoice number from the file name
    format_date = datetime.now().strftime('%d-%m-%y')   # get the current date in the specified format

    # Add the invoice number to the PDF
    pdf.set_font(family="Arial", size=16, style="B")
    pdf.cell(w=50, h=8, txt=f"Invoice number.{invoice_nr}", ln=1)

    # Add the date to the PDF
    pdf.set_font(family="Arial", size=16, style="B")
    pdf.cell(w=50, h=12, txt=f"Date: {format_date}", ln=20)
    pdf.cell(w=0, h=10, ln=1)

    # Read the Excel file into a pandas DataFrame
    df = pd.read_excel(path, sheet_name="Sheet 1")

    # Add a header row to the table
    columns_a = df.columns
    columns_a = [item.replace("_", " ").title() for item in columns_a]  # capitalize the column names and replace underscores with spaces
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

    # Add the sum total row to the table
    sum_total = str(df['total_price'].sum())

print("PDF were created successfully")
