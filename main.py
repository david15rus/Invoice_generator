import pandas as pd
from fpdf import FPDF
import glob

filepaths = glob.glob('Invioces/*.xlsx')

for filepath in filepaths:
    df = pd.read_excel(filepath, sheet_name='Лист1')
    print(df)
