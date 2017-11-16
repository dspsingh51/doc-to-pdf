import sys
import os
import comtypes.client

wdFormatPDF = 17

in_file = os.path.abspath(sys.argv[1])
out_file = os.path.abspath(sys.argv[2])

word = comtypes.client.CreateObject('Word.Application')
doc = word.Documents.Open('path for input file')
doc.SaveAs('path for output file', FileFormat=wdFormatPDF)
doc.Close()
word.Quit()
