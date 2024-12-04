from fpdf import FPDF
import pandas as pd
import glob, pathlib

filepaths=glob.glob('resources/*.xlsx')
for filepath in filepaths:
    pdf=FPDF(orientation='P',unit='mm',format='A4')
    pdf.add_page()

    #read data
    df_invoice=pd.read_excel(filepath,usecols=range(0,5))
    print(df_invoice)
    df_receiver=pd.read_excel(filepath,usecols=range(5,9))
    invoice_nr,date_str=pathlib.Path(filepath).stem.split('-')
    date_y,date_m,date_d=date_str.split('.')
   
    #1st Line   
    pdf.set_font(family='Times',style='B',size=20)
    pdf.set_text_color(255,255,255)
    pdf.set_fill_color(0, 0, 0)
        #black padding
    pdf.cell(90,10,txt='',fill=True,ln=1)
        #Company name
    pdf.cell(90,14,txt='PYTHONHOW',fill=True,border=0,align='C')
        #Invoice Nr + black blank
    pdf.set_text_color(0,0,0)
    pdf.set_fill_color(255,255,255)
    pdf.cell(90,14,txt=f"Invoice #{invoice_nr}",border=0,align='R')
        #right black padding
    pdf.set_fill_color(0,0,0)
    pdf.cell(5,14,txt='',fill=True,border=0,align='C',ln=1)

    #2nd Line
        #Catch phase
    pdf.set_font(family='Times',style='B',size=10)
    pdf.set_text_color(255,255,255)
    pdf.set_fill_color(0, 0, 0)
    pdf.cell(90,6,txt='Your Trusworthly invoice company',fill=True,border=0,align='C')
        #Date
    pdf.set_font(family='Times',style='B',size=12)
    pdf.set_text_color(160,160,160)
    pdf.cell(90,6,txt=f"{date_d}/{date_m}/{date_y}",border=0,align='R')
        #right black padding
    pdf.set_fill_color(0,0,0)
    pdf.cell(5,6,txt='',fill=True,border=0,align='C',ln=1)
        #black padding
    pdf.cell(90,10,txt='',fill=True,ln=1) 

    pdf.ln(20)
    pdf.cell(0,1,txt='',fill=True,ln=1) #black line

    #Customer detail table
        #headding
    pdf.set_font('Times','B',10)
    pdf.set_text_color(0,0,0)
    headding_list = [i.replace('_'," ").upper() for i in list(df_receiver.columns)]
    pdf.cell(52,10,headding_list[0],align='L',border='B')
    pdf.cell(52,10,headding_list[1],align='L',border='B')
    pdf.cell(52,10,headding_list[2],align='L',border='B')
    pdf.cell(0,10,headding_list[3],align='L',border='B',ln=1)
        #data
    pdf.set_font('Times','',10)
    pdf.set_text_color(100,100,100)
    for jndex, row in df_receiver.iterrows():
        pdf.cell(52,6,txt='' if str(row['To']) == 'nan' else str(row['To']),align='L')
        pdf.cell(52,6,txt='' if str(row['customer_id']) == 'nan' else str(row['customer_id']),align='L')
        pdf.cell(52,6,txt='' if str(row['address']) == 'nan' else str(row['address']),align='L')
        pdf.cell(0,6,txt='' if str(row['phone']) == 'nan' else str(row['phone']),align='L',ln=1)


    pdf.ln(20)
    pdf.cell(0,1,txt='',fill=True,ln=1) #black line
    
    #Invoice table
        #headding
    pdf.set_font('Times','B',10)
    pdf.set_text_color(0,0,0)
    headding_list = [i.replace('_'," ").upper() for i in list(df_invoice.columns)]
    pdf.cell(30,10,headding_list[0],align='L',border='B')
    pdf.cell(55,10,headding_list[1],align='L',border='B')
    pdf.cell(25,10,'QTY',align='L',border='B')
    pdf.cell(40,10,headding_list[3],align='L',border='B')
    pdf.cell(0,10,headding_list[4],align='L',border='B',ln=1)
        #data
    pdf.set_font('Times','',10)
    pdf.set_text_color(100,100,100)
    for kndex, row in df_invoice.iterrows():
        pdf.cell(30,8,txt='' if str(row['product_id']) == 'nan' else str(row['product_id']),align='L')
        pdf.cell(55,8,txt='' if str(row['product_name']) == 'nan' else str(row['product_name']),align='L')
        pdf.cell(25,8,txt='' if str(row['amount_purchased']) == 'nan' else str(row['amount_purchased']),align='L')
        pdf.cell(40,8,txt='' if str(row['price_per_unit']) == 'nan' else str(row['price_per_unit']),align='L')
        pdf.cell(0,8,txt='' if str(row['total_price']) == 'nan' else str(row['total_price']),align='L',ln=1)
    #black line
    pdf.cell(0,1,txt='',border='B',ln=1) 
    #sub total
    pdf.cell(110,10,txt='') 
    pdf.set_font('Times','B',10)
    pdf.set_text_color(0,0,0)
    pdf.cell(40,10,'SUBTOTAL')
    pdf.set_text_color(90,90,90)
    pdf.cell(0,10,'$ '+str(df_invoice['total_price'].sum()),ln=1)
    #sales tax
    pdf.cell(110,10,txt='')   
    pdf.set_font('Times','B',10)
    pdf.set_text_color(0,0,0)
    pdf.cell(40,10,'SALES TAX')
    pdf.set_text_color(90,90,90)
    pdf.cell(0,10,'0.06',ln=1)
    #black line
    pdf.cell(110,1,txt='') 
    pdf.cell(0,1,txt='',fill=True,ln=1) 
    #total
    pdf.cell(110,10,txt='')   
    pdf.set_font('Times','B',10)
    pdf.set_text_color(0,0,0)
    pdf.cell(40,10,'TOTAL')
    pdf.set_text_color(90,90,90)
    total=str(round(df_invoice['total_price'].sum()*1.06,2))
    pdf.cell(0,10,'$ '+total,ln=1) 
    #black line
    pdf.cell(110,1,txt='') 
    pdf.cell(0,1,txt='',border='B',ln=1)    

    #bottom left padding
    pdf.set_y(240.0)
    pdf.cell(5,20,'',fill=True)

    #My company details
    pdf.set_font('Times','',10)
    pdf.set_text_color(100,100,100)
    pdf.set_xy(20.0,238.0)
    pdf.cell(60,8,txt='236-555-0126',ln=1)
    pdf.set_x(20.0)
    pdf.cell(60,8,txt='pythonhow@trustworthly.mail.com',ln=1)
    pdf.set_x(20.0)
    pdf.cell(60,8,txt='4321 AppleTree Lane | Naville 1452',ln=1)
    
    #Thank you
    pdf.set_xy(105.0,230.0)
    pdf.set_text_color(255,255,255)
    pdf.set_font('Times','B',20)
    pdf.cell(0,40,'THANK YOU',fill=True,align='C')
    


    pdf.output(invoice_nr+'.pdf')