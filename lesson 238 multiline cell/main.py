from fpdf import FPDF
import glob
from pathlib import Path

filepaths=glob.glob('Text+Files/*.txt')

pdf=FPDF("P","mm","A4")

for filepath in filepaths:

    pdf.add_page()

    filename = Path(filepath).stem.capitalize()
    pdf.set_font('Times','B',16)
    pdf.cell(0,8,txt=filename,ln=1)

    with open(Path(filepath),'r') as f:
        data=f.read()
        pdf.set_font('Times','',12)
        pdf.multi_cell(0,6,txt=data)
    

pdf.output('lesson238.pdf')



