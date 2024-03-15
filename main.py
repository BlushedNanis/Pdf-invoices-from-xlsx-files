import pandas as pd
import glob
from fpdf import FPDF


filepaths = glob.glob("Xlsx files/*.xlsx")
pdf = FPDF()

for filepath in filepaths:
    df = pd.read_excel(filepath, "Sheet 1")
    invoice_number = filepath[11:-15]
    invoice_date = filepath[-14:-5]
    pdf.add_page()
    pdf.set_font('times', 'B', 20)
    pdf.cell(0, 10, f"Invoice nr. {invoice_number}", 0, 1, 'L')
    pdf.cell(0, 10, f"Date {invoice_date}", 0, 1, 'L')
    for index, row in df.iterrows():
        break
    pdf.output(f"pdf invoices\{filepath[11:-5]}.pdf")