# -*- coding: utf8 -*-
#coding=utf-8 

import sys
reload(sys)
print sys.getdefaultencoding()
sys.setdefaultencoding('utf-8')
print sys.getdefaultencoding()

import win32com.client


wordFilePath = "../RGtool/report/hello.docx"
excelFilePath = "../RGtool/resource/test.xlsx"

excelapp = win32com.client.Dispatch("Excel.Application")
excelapp.Visible = 0
excelxls = excelapp.Workbooks.Open(excelFilePath)

ws = excelxls.Worksheets("Sheet1")
used = ws.UsedRange
nrows = used.Row + used.Rows.Count
ncols = used.Column + used.Columns.Count

#print nrows
#print ncols
'''
for i in range(1, nrows):
	for j in range(1, ncols):
		#print i
		#print j
		print ws.Cells(i, j)
'''
#print ws.Cells(2, 2)


#data = excelapp.Range("A1")
#print data.value


wordapp = win32com.client.Dispatch("Word.Application") # Create new Word Object
wordapp.Visible = 0 # Word Application should`t be visible
worddoc = wordapp.Documents.Add() # Create new Document Object
'''
worddoc.PageSetup.Orientation = 1 # Make some Setup to the Document:
worddoc.PageSetup.LeftMargin = 20
worddoc.PageSetup.TopMargin = 20
worddoc.PageSetup.BottomMargin = 20
worddoc.PageSetup.RightMargin = 20
worddoc.Content.Font.Size = 11
worddoc.Content.Paragraphs.TabStops.Add (100)
worddoc.Content.Text = "Hello world!"
worddoc.Content.MoveEnd
'''
rng = worddoc.Range(0,0)
for i in range(2, nrows):
	for j in range(2, ncols):
		rng.InsertAfter(ws.Cells(1, j))
		rng.InsertAfter(" : ")
		rng.InsertAfter(unicode(ws.Cells(i, j)))
		rng.InsertAfter("\r\n")
	rng.InsertAfter("\r\n\r\n")

worddoc.SaveAs(wordFilePath)
wordapp.Quit() # Close the Word Application


