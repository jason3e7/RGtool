import subprocess
import os
import re
from lxml import etree

outputPath = os.getcwd() + "/report/docxtemp"
filepath = "../RGtool/report/"
filename = "hello.docx"

## decompress from docx 
subprocess.call(['7z', 'x', filepath + filename, '-o' + outputPath, '-y'], stdout=open(os.devnull, 'wb'))

## read xml file
f = open(outputPath + '/word/document.xml', 'r+')
xml = f.read()
#print xml

## root is element
root = etree.fromstring(xml)
#print(etree.tostring(root, pretty_print=True))

## ns is namespace
m = re.match('\{.*\}', root.tag)
ns = m.group(0)

## remove namespace
nsEnd = root.tag.find('}')
root.tag = root.tag[nsEnd + 1:]
for elem in root.getiterator():
	nsEnd = elem.tag.find('}')
	elem.tag = elem.tag[nsEnd + 1:]

'''
print root
print root.tag
print root[0]
for elem in root.getiterator():
	print elem

#nsmap = {k:v for k,v in root.nsmap.iteritems() if k}
#print nsmap
'''

'''
document = root.xpath('/document/')
print document
'''

#body = root.iterfind('p')
#print body
#print body[0]

'''
body = body.find('p')
print body
'''

## add namespace
for elem in root.getiterator():
	elem.tag = ns + elem.tag

## write xml file
f.seek(0)
f.write('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n')
f.write(etree.tostring(root))
f.truncate()
f.close()

## compress to zip 
subprocess.call(['7z', 'a', filepath + 'test.docx', outputPath + "/*"], stdout=open(os.devnull, 'wb'))
