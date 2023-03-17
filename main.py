import pandas as pd
from fpdf import FPDF
import glob
from pathlib import Path

filepaths = glob.glob('Invoices/*.xlsx')

for filepath in filepaths:
    pdf = FPDF(orientation='P', unit='mm', format='A4')
    pdf.add_page()

    filename = Path(filepath).stem
    inv_num, date = filename.split('-')

    pdf.set_font(family='Times', size=16, style='B')
    pdf.cell(w=50, h=8, txt=f'Invoices num {inv_num}', ln=1)

    pdf.set_font(family='Times', size=16, style='B')
    pdf.cell(w=50, h=8, txt=f'Date {date}', ln=2)

    df = pd.read_excel(filepath, sheet_name='Лист1')

    # Remove underscore and titled a headers name
    columns = list(map(lambda x: x.replace('_', ' ').title(), df.columns))

    # Add a header
    pdf.set_font(family='Times', size=9, style='B')
    pdf.set_text_color(80, 80, 80)
    pdf.cell(w=30, h=8, txt=columns[0], border=1)
    pdf.cell(w=70, h=8, txt=columns[1], border=1)
    pdf.cell(w=30, h=8, txt=columns[2], border=1)
    pdf.cell(w=30, h=8, txt=columns[3], border=1)
    pdf.cell(w=30, h=8, txt=columns[4], border=1, ln=1)

    for idx, row in df.iterrows():
        pdf.set_font(family='Times', size=8, style='B')
        pdf.set_text_color(80, 80, 80)
        pdf.cell(w=30, h=8, txt=str(row['Product_id']), border=1)
        pdf.cell(w=70, h=8, txt=row['Product_name'], border=1)
        pdf.cell(w=30, h=8, txt=str(row['Amount_purchased']), border=1)
        pdf.cell(w=30, h=8, txt=str(row['Price_per_unit']), border=1)
        pdf.cell(w=30, h=8, txt=str(row['Total_price']), border=1, ln=1)

    pdf.output(f'PDFs/{filename}.pdf')
