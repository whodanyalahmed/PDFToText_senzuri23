# importing required modules
import PyPDF2
import os
filename = "invoice.pdf"
# creating a pdf file object
pdfFileObj = open(filename, 'rb')

# creating a pdf reader object
pdfReader = PyPDF2.PdfFileReader(pdfFileObj)

# printing number of pages in pdf file
# print(pdfReader.numPages)

# creating a page object
pageObj = pdfReader.getPage(0)

txt = pageObj.extractText()
list_txt = txt.split('\n')
# split txt with new line
print(list_txt[-7])

# closing the pdf file object
pdfFileObj.close()
# check if folder exist
if not os.path.exists('done'):
    os.makedirs('done')
os.rename(filename, "done/"+filename)
