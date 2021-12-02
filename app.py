# importing required modules
import PyPDF2
import shutil
import os
from openpyxl import load_workbook


def NamefromPDF(filename, path):

    # creating a pdf file object
    pdfFileObj = open(path+filename, 'rb')

    # creating a pdf reader object
    pdfReader = PyPDF2.PdfFileReader(pdfFileObj)

    # printing number of pages in pdf file
    # print(pdfReader.numPages)

    # creating a page object
    pageObj = pdfReader.getPage(0)

    txt = pageObj.extractText()
    list_txt = txt.split('\n')
    # split txt with new line
    main_text = list_txt[-7]
    # closing the pdf file object
    pdfFileObj.close()
    return main_text


def moveFile(filename, path):
    if not os.path.exists(path+'done'):
        os.makedirs(path+'done')
        print('info : created done folder')
    # check if file exist in path
    counter = 1
    filename_new = filename.split('.')
    # check if file exist until it is not exist
    while True:
        counter += 1
        newName = filename_new[0]+'_'+str(counter)+'.'+filename_new[1]
        if os.path.exists(path+"done/"+newName):
            continue
        else:
            break
    # move file to done folder
    shutil.move(path+filename, path+'done/' +
                filename_new[0]+'_'+str(counter)+'.'+filename_new[1])

    print('success : successfully moved ' + newName + ' file to done folder')


def get_column_values(sheet, column):
    values = []
    for row in sheet.iter_rows(min_row=2, max_row=sheet.max_row, min_col=column, max_col=column):
        for cell in row:
            values.append(cell.value)
    return values
# set column value


if __name__ == '__main__':
    source_file = load_workbook('data.xlsx', data_only=True)
    sheet = source_file["Sheet2"]

    filename = "invoice.pdf"
    path = "C://Users//Daniyal\Downloads//"

    # loop on sheet
    txt = NamefromPDF(filename, path)
    for index, row in enumerate(sheet.iter_rows(min_row=2, max_row=sheet.max_row, min_col=3, max_col=3)):
        for cell in row:
            num = 'C'+str(index+2)
            print(index, sheet[num].value, sheet['D'+str(index+2)].value)
            # if index == 86:
            #     txt = NamefromPDF(filename, path)
            #     print(txt)
            #     sheet['D'+str(index+2)].value = txt
    # iterate on pair keys and values with index
    source_file.save('data.xlsx')
    print(txt)
    moveFile(filename, path)
