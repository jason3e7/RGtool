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
for elem in root.getiterator():
	nsEnd = elem.tag.find('}')
	elem.tag = elem.tag[nsEnd + 1:]
	for key in elem.keys():
		nsEnd = key.find('}')
		elem.set(key[nsEnd + 1:], elem.attrib[key])
		del elem.attrib[key]
	
'''
for elem in root.getiterator():
	print elem
	for item in elem.items():
		print item
'''	
#print(etree.tostring(root, pretty_print=True))

pgMar = root.xpath('/document/body/sectPr/pgMar')
pgBorders = etree.Element("pgBorders", offsetFrom="page")
pgMar[0].addnext(pgBorders)

locates = ['top', 'left', 'bottom', 'right']
for l in locates:
	etree.SubElement(pgBorders, l, val="thinThickSmallGap", sz="24", space="24", color="auto")

body = root.xpath('/document/body')
#p = etree.Element("p")
#print body[0]
#print p
etree.SubElement(body[0], 'p')
etree.SubElement(body[0], 'p')
etree.SubElement(body[0], 'p')

## company
p = etree.SubElement(body[0], 'p')
pPr = etree.SubElement(p, 'pPr')
etree.SubElement(pPr, 'jc', val='center')
rPr = etree.SubElement(pPr, 'rPr')
etree.SubElement(rPr, 'bdr', color='auto', space='0', sz='4', val='single')

r = etree.SubElement(p, 'r')
rPr = etree.SubElement(r, 'rPr')
etree.SubElement(rPr, 'rFonts', hint='eastAsia')
etree.SubElement(rPr, 'bdr', color='auto', space='0', sz='4', val='single')
t = etree.SubElement(r, 't')
t.text = 'company'

etree.SubElement(body[0], 'p')

## PT report
p = etree.SubElement(body[0], 'p')
pPr = etree.SubElement(p, 'pPr')
etree.SubElement(pPr, 'jc', val='center')
rPr = etree.SubElement(pPr, 'rPr')
etree.SubElement(rPr, 'bdr', color='auto', space='0', sz='4')

r = etree.SubElement(p, 'r')
rPr = etree.SubElement(r, 'rPr')
etree.SubElement(rPr, 'rFonts', hint='eastAsia')
etree.SubElement(rPr, 'bdr', color='auto', space='0', sz='4')
t = etree.SubElement(r, 't')
t.text = 'PT report'

etree.SubElement(body[0], 'p')


etree.SubElement(body[0], 'p')
etree.SubElement(body[0], 'p')
etree.SubElement(body[0], 'p')
#print(etree.tostring(root, pretty_print=True))

## add namespace
for elem in root.getiterator():
	elem.tag = ns + elem.tag
	for key in elem.keys():
		elem.set(ns + key, elem.attrib[key])
		del elem.attrib[key]

## write xml file
f.seek(0)
f.write('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n')
f.write(etree.tostring(root))
f.truncate()
f.close()

## compress to zip 
subprocess.call(['7z', 'a', filepath + 'test.docx', outputPath + "/*"], stdout=open(os.devnull, 'wb'))
