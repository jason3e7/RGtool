import subprocess
import os
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

root = etree.fromstring(xml)

#parser = etree.XMLParser(ns_clean = True)
#root = etree.parse(xml, parser)

#print root
#print root.find('t', namespaces='w')
#print root.find("a")


## write xml file
f.seek(0)
f.write('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n')
f.write(etree.tostring(root))
f.truncate()
f.close()

## compress to zip 
subprocess.call(['7z', 'a', filepath + 'test.docx', outputPath + "/*"], stdout=open(os.devnull, 'wb'))
