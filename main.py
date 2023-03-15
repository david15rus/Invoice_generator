import pandas as pd
from fpdf import FPDF
import glob
from pathlib import Path

filepaths = glob.glob('Invioces/*.xlsx')

for filepath in filepaths:
    df = pd.read_excel(filepath, sheet_name='Лист1')
    pdf = FPDF(orientation='P', unit='mm', format='A4')
    pdf.add_page()
    filename = Path(filepath).stem
    filename_list = filename.split('-')
    inv_num = filename_list[0]
    date = f'{filename_list[1]}.{filename_list[2]}.{filename_list[3]}'
    pdf.set_font(family='Times', size=16, style='B')
    pdf.cell(w=50, h=8, txt=f'Invioces num {inv_num}', ln=1)
    pdf.cell(w=50, h=8, txt=f'Date {date}', ln=2)
    pdf.output(f'PDFs/{filename}.pdf')
