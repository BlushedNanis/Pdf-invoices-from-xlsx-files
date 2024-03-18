import pandas as pd
import glob
from fpdf import FPDF


filepaths = glob.glob("Xlsx files/*.xlsx")

for filepath in filepaths:
    pdf = FPDF()
    df = pd.read_excel(filepath, "Sheet 1")
    invoice_number = filepath[11:-15]
    invoice_date = filepath[-14:-5]
    pdf.add_page()
    pdf.set_font('times', 'B', 18)
    pdf.cell(0, 5, f"Invoice nr. {invoice_number}", 0, 1, 'L')
    pdf.cell(0, 15, f"Date {invoice_date}", 0, 2, 'L')
    pdf.set_font('times', 'B', 8)
    columns = [item.replace("_", " ").capitalize() for item in df.columns]
    for column in columns:
        pdf.cell(37, 10, column, 1, 0, 'C')
    pdf.set_font('times', '', 8)
    for index, row in df.iterrows():
        pdf.ln(10)
        pdf.cell(37, 10, str(row['product_id']), 1, 0, 'L')
        pdf.cell(37, 10, str(row['product_name']), 1, 0, 'L')
        pdf.cell(37, 10, str(row['amount_purchased']), 1, 0, 'C')
        pdf.cell(37, 10, str(row['price_per_unit']), 1, 0, 'C')
        pdf.cell(37, 10, str(row['total_price']), 1, 0, 'C')
    pdf.ln(10)
    for i in range(len(columns)-1):
        pdf.cell(37, 10, '', 1, 0)
    pdf.set_font('Times', 'B', 8)
    pdf.cell(37, 10, str(df.sum()['total_price']), 1, 1, 'C')
    pdf.ln(10)
    pdf.set_font('times', 'B', 10)
    pdf.cell(0, 10, f"Total due amount is ${str(row['total_price'])}",
             0, 1)
    pdf.cell(30, 10, "Blushed Nanis Inc.", 0, 0)
    pdf.image("logo\gizmo.png", w=7)
    pdf.output(f"pdf invoices\{filepath[11:-5]}.pdf")