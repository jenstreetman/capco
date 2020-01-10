#specific to extracting information from word documents
import os
import zipfile

#other tools useful in extracting the information from our document
import re
#to pretty print our xml:
import xml.dom.minidom

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