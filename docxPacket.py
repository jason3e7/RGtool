import subprocess
import os

outputPath = os.getcwd() + "/report/docxtemp"
filepath = "../RGtool/report/"
filename = "hello.docx"

subprocess.call(['7z', 'x', filepath + filename, '-o' + outputPath, '-y'], stdout=open(os.devnull, 'wb'))

subprocess.call(['7z', 'a', filepath + 'test.docx', outputPath + "/*"], stdout=open(os.devnull, 'wb'))
