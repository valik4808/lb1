import sys
import zipfile
from lxml import etree
import codext

with zipfile.ZipFile(sys.argv[1]) as docx:
    docx.extractall('./tmp')
    doc = etree.parse("./tmp/word/document.xml")

root = doc.getroot()
raw_data = ''

for wr_roots in doc.xpath('//w:r', namespaces=root.nsmap):
    if wr_roots.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}rsidRPr') == '00F30162':
        raw_data += '1' * len(wr_roots.find('w:t', namespaces=root.nsmap).text)
    elif len(wr_roots.find('w:t', namespaces=root.nsmap).text) >= 8:
        break
    else:
        raw_data += '0' * len(wr_roots.find('w:t', namespaces=root.nsmap).text)

data_cp866 = ''
data_cp1251 = ''
data_koi8r = ''
data_baudot = codext.decode(raw_data[:-(len(raw_data) % 5)], "baudot-ita2")

i = 0
while i + 8 < len(raw_data):
    data_cp866 += bytes.fromhex(hex(int(raw_data[i:i+8], 2))[2:]).decode(encoding="cp866")
    data_cp1251 += bytes.fromhex(hex(int(raw_data[i:i+8], 2))[2:]).decode(encoding="cp1251")
    data_koi8r += bytes.fromhex(hex(int(raw_data[i:i+8], 2))[2:]).decode(encoding="koi8-r")
    i += 8

print(f'RAW: {raw_data}\nCP866: {data_cp866}\nWindows-1251: {data_cp1251}\nKOI8R: {data_koi8r}\nBAUDOT (MTK-2): ↴\n{data_baudot}')