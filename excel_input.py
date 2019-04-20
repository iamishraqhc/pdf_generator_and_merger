import pandas as pd
import pdfkit
from PyPDF2 import PdfFileMerger
import os

df = pd.read_excel('CALCOLATORE PREVENTIVI.xls', sheet_name='CALCOLATORE') # col = 1, row = 2

var1 = df.loc[3][7]
var2 = df.loc[5][3] # 7 4
var3 = df.loc[7][3] # 9 4
var4 = df.loc[11][2] # 13 3
var5_1 = df.loc[19][3] # 21 4
var5_2 = df.loc[19][5] # 21 6
var5_3 = df.loc[19][7] # 21 8
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
    text_html = text_html.replace('var5', 'asdasdasd')
    text_html = text_html.replace('var7', 'dasdasdasd')
    text_html = text_html.replace('var9', 'asdadasdasd')
    text_html = text_html.replace('var3', 'dasdasdasd')
    text_html = text_html.replace('var8', 'sdasdaseq')
    text_html = text_html.replace('var4', 'dasdqwe')
    text_html = text_html.replace('var2', 'sdadqw')
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

pdf_files = ['output1.pdf','AUTODICHIARAZIONE.pdf']

merger = PdfFileMerger()

for files in pdf_files:
    merger.append(files)
if not os.path.exists('Final Output.pdf'):
    merger.write('Final Output.pdf')
merger.close()