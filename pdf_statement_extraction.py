from img2table.document import PDF
from natsort import natsorted, ns
import os

"""
this script extracts image based pdf in excel file in this case bank statment
img2table == 0.0.16
natsort == 8.3.1
"""

fileList = natsorted(os.listdir(os.getcwd()), key=lambda y: y.lower())#sort alphanumeric in ascending order

for pdf in fileList:
    if pdf.endswith(".pdf"):
        pdf_obj = PDF(os.path.join(os.getcwd(),pdf), dpi=200)	
        pdf_obj.to_xlsx(pdf + ".xlsx")

