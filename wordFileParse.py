#specific to extracting information from word documents
import os
import zipfile

#other tools useful in extracting the information from our document
import re
#to pretty print our xml:
import xml.dom.minidom

# Writing to an excel  
# sheet using Python 
import xlwt 
from xlwt import Workbook 

document = zipfile.ZipFile('claire.docx')

print(document.namelist())

uglyXml = xml.dom.minidom.parseString(document.read('word/document.xml')).toprettyxml(indent='  ')

text_re = re.compile('>\n\s+([^<>\s].*?)\n\s+</', re.DOTALL)
prettyXml = text_re.sub('>\g<1></', uglyXml)

# regex = r"[A-Z][a-z]+"
regex = r"[0-9]+\scup"

link_list = re.findall(regex,prettyXml)[1:]
# link_list = [x[:-1] for x in link_list]
print(link_list)




  
# Workbook is created 
wb = Workbook() 
  
# add_sheet is used to create sheet. 
sheet1 = wb.add_sheet('Sheet 1') 

for idx, link in enumerate(link_list):
  sheet1.write(idx, 0, link)

wb.save('xlwt example.xls') 