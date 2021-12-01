# importing required modules
import PyPDF2
import shutil
import os


def getNameFromPDF(filename, path):

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


if __name__ == '__main__':
    filename = "invoice.pdf"
    path = "C://Users//Daniyal\Downloads//"
    txt = getNameFromPDF(filename, path)
    print(txt)
    moveFile(filename, path)
