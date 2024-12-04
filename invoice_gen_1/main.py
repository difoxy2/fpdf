from fpdf import FPDF
import glob
from pathlib import Path
import pandas as pd

filepaths=glob.glob('resources/*.xlsx')

for filepath in filepaths:
    filename=Path(filepath).stem
    invoice_no, date = filename.split('-')
    df=pd.read_excel(filepath)
    

    #new page
    pdf=FPDF('P','mm','A4')
    pdf.add_page()

    #headdings
    pdf.set_font('Times','B',16)
    pdf.cell(0,8,txt='Invoice nr. '+invoice_no,ln=1)
    pdf.cell(0,8,txt='Date '+date,ln=1)
    pdf.cell(30,6,'',0,align='C',ln=1)
    
    #table headding
    pdf.set_font('Times','B',12)
    headding_list = [i.replace('_'," ").title() for i in list(df.columns)]
    pdf.cell(25,6,headding_list[0],1)
    pdf.cell(70,6,headding_list[1],1)
    pdf.cell(40,6,headding_list[2],1)
    pdf.cell(30,6,headding_list[3],1)
    pdf.cell(30,6,headding_list[4],1,ln=1)

    #table body
    pdf.set_font('Times','',12)
    for index,row in df.iterrows():
        pdf.cell(25,6,str(row['product_id']),1)
        pdf.cell(70,6,str(row['product_name']),1)
        pdf.cell(40,6,str(row['amount_purchased']),1,align='C')
        pdf.cell(30,6,str(row['price_per_unit']),1,align='C')
        pdf.cell(30,6,str(row['total_price']),1,align='C',ln=1)
    
    #table total amount
    pdf.set_font('Times','B',12)
    pdf.cell(25,6,'',1)
    pdf.cell(70,6,'',1)
    pdf.cell(40,6,'',1,align='C')
    pdf.cell(30,6,'',1,align='C')
    pdf.cell(30,6,str(df['total_price'].sum()),1,align='C',ln=1)
    pdf.cell(30,6,'',0,align='C',ln=1)

    #summery below table
    pdf.set_font('Times','B',12)
    pdf.cell(0,6,txt=f"TOTAL AMOUNT IS {df['total_price'].sum()} EUROS.",ln=1)
    pdf.cell(23,6,txt='PythonHow')
    pdf.image('resources/pythonhow.png', w=10)


    #output pdf
    pdf.output(invoice_no+'.pdf')