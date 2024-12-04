from fpdf import FPDF
import pandas as pd

data=pd.read_csv('topics.csv')

pdf = FPDF('P', 'mm', 'A4')

for index, row in data.iterrows():

    for i in range(int(row['Pages'])):
        pdf.add_page()
        pdf.set_font('Arial','B',16)
        pdf.set_text_color(100,100,100)

        if i==0:
            pdf.cell(0,16,row['Topic'])

        for j in range(24,270,10):
            pdf.line(10,j,200,j)

        pdf.ln(256)
        pdf.set_font('Arial', 'I', 8)
        pdf.set_text_color(160,160,160)
        pdf.cell(0, 10, row['Topic'] , 0, 0, 'R')

pdf.output('output.pdf')