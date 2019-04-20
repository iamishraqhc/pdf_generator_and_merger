import pandas as pd
import pdfkit
from PyPDF2 import PdfFileMerger
import os
import math

df = pd.read_excel('CALCOLATORE PREVENTIVI.xls', sheet_name='CALCOLATORE') # col = 1, row = 2

var1 = df.loc[3][7]
var2 = df.loc[5][3] # 7 4
var3 = df.loc[7][3] # 9 4
var4 = df.loc[11][2] # 13 3
var5_1 = df.loc[19][3] # 21 4
var5_2 = df.loc[19][5] # 21 6
var5_3 = df.loc[19][7] # 21 8
if var5_1 == "x" or var5_1 == "X":
    var5 = "16"
if var5_2 == "x" or var5_2 == "X":
    var5 = "32"
if var5_3 == "x" or var5_3 == "X":
    var5 = "48"
var7 = df.loc[31][3] # 33 4
var8 = df.loc[3][3] # 5 4
var9 = df.loc[31][6] # 33 7

pdf_1 = df.loc[35][4] # 37 5
pdf_2 = df.loc[37][4] # 39 5
pdf_3 = df.loc[39][4] # 41 5
pdf_4 = df.loc[41][4] # 43 5
pdf_5 = df.loc[43][4] # 45 5

# var5, var7, var9, var3, var8, var4, var2, var1_1-5

final_text_html = ''
with open('test.html','r') as file:
    text_html = file.read()
    text_html = text_html.replace('var5', var5)
    text_html = text_html.replace('var7', var7)
    text_html = text_html.replace('var9', var9)
    text_html = text_html.replace('var3', var3)
    text_html = text_html.replace('var8', var8)
    text_html = text_html.replace('var4', str(var4))
    text_html = text_html.replace('var2', var2)
    text_html = text_html.replace('var1_1', 'sdasdada') # special consideration
    final_text_html = text_html
    # print(text_html)

# Generate pdf from this text (1)
# Read pdf from template

# Output 1 + 2

with open('test2.html','w') as file2:
    file2.write(final_text_html)
config_path = pdfkit.configuration(wkhtmltopdf=bytes(r'C:\Program Files\wkhtmltopdf\bin\wkhtmltopdf.exe', 'utf8'))

pdfkit.from_file('test2.html','output1.pdf',configuration=config_path)
print("Output generated")

# path = r'C:\Users\Ishraq Haider\PycharmProjects\pdfGenerator'

pdfs = ['output1.pdf']
if pdf_1 == "x" or pdf_1 == "X":
    pdfs.append("CARTA D'IDENTITA'.pdf")
if pdf_2 == "x" or pdf_2 == "X":
    pdfs.append("DICHIARAZIONE PER MEPA.pdf")
if pdf_3 == "x" or pdf_3 == "X":
    pdfs.append("MANIFESTAZIONE D'INTERESSE.pdf")
if pdf_4 == "x" or pdf_4 == "X":
    pdfs.append("VISURA CAMERALE.pdf")
if pdf_5 == "x" or pdf_5 == "X":
    pdfs.append("AUTODICHIARAZIONE.pdf")

print(pdfs)

# pdf_files = ['output1.pdf','AUTODICHIARAZIONE.pdf']

merger = PdfFileMerger(strict=False)

for files in pdfs:
    merger.append(files)
if not os.path.exists('Final Output.pdf'):
    merger.write('Final Output.pdf')
merger.close()