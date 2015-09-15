# http://win32com.goermezer.de/content/view/173/192/
# https://mail.python.org/pipermail/tutor/2006-July/048131.html

import win32com.client
filepath = "C:\Users\Jason3e7\Documents\hello.docx"

wordapp = win32com.client.Dispatch("Word.Application") # Create new Word Object
wordapp.Visible = 0 # Word Application should`t be visible
worddoc = wordapp.Documents.Add() # Create new Document Object
worddoc.PageSetup.Orientation = 1 # Make some Setup to the Document:
worddoc.PageSetup.LeftMargin = 20
worddoc.PageSetup.TopMargin = 20
worddoc.PageSetup.BottomMargin = 20
worddoc.PageSetup.RightMargin = 20
worddoc.Content.Font.Size = 11
worddoc.Content.Paragraphs.TabStops.Add (100)
worddoc.Content.Text = "Hello world!"
worddoc.Content.MoveEnd
worddoc.SaveAs(filepath)
wordapp.Quit() # Close the Word Application
