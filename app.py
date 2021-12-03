# importing required modules
import PyPDF2
import shutil
import os
from openpyxl import load_workbook
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
import time
import sys
import os
cur_path = sys.path[0]


def resource_path(relative_path):
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.dirname(__file__)
    return os.path.join(base_path, relative_path)


driverPath = resource_path('I://clients//chromedriver.exe')
driver = webdriver.Chrome(driverPath)

driver.maximize_window()

driver.implicitly_wait(10)

# Login


def login():
    try:
        driver.find_element_by_id(
            "ctl00_ctl00__dh_UserName").send_keys("Dlaing")
        time.sleep(1)

        driver.find_element_by_id(
            "ctl00_ctl00__dh_Password").send_keys("Hirosushi1")
        time.sleep(1)
        driver.find_element_by_xpath(
            "//html/body/form/div[4]/div[1]/span/div[1]/div[2]/div/span/fieldset/span[3]/input").click()
        time.sleep(1)
    except Exception as e:
        print("error: something went wrong while logging in..." + str(e))

        time.sleep(2)


def go_to_page(url):
    try:
        driver.get(url)
        time.sleep(2)
    except Exception as e:
        print("error: something went wrong while going to page..." + str(e))


def InputKey(key):
    try:
        inputTxt = driver.find_element_by_id(
            "ctl00_ctl00_c_c__txtInvoiceRef__t")
        # clear inputTxt
        inputTxt.clear()
        inputTxt.send_keys(key)
        inputTxt.send_keys(Keys.ENTER)

    except Exception as e:
        print("error: Cant find the element for input " + str(e))


def DownloadPDF():
    try:
        driver.find_element_by_id(
            "ctl00_ctl00_c_c__repeaterCustomers_ctl00__btnDownload").click()
    except Exception as e:
        print("error: cant downlaod the file something went wrong..." + str(e))


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
    websiteURL = "https://tas.driverhire.co.uk/tas/invoicing/"
    driver.get(websiteURL)
    login()
    go_to_page(websiteURL)

    # loop on sheet
    txt = NamefromPDF(filename, path)
    for index, row in enumerate(sheet.iter_rows(min_row=2, max_row=sheet.max_row, min_col=3, max_col=3)):
        for cell in row:

            num = 'C'+str(index+2)

            if (sheet['D'+str(index+2)].value == None or sheet['D'+str(index+2)].value == 0):
                print(index, sheet['D'+str(index+2)].value, sheet[num].value)
                InputKey(sheet[num].value)
                source_file.save('data.xlsx')
                txt = NamefromPDF(filename, path)
                sheet['D'+str(index+2)].value = txt
                source_file.save('data.xlsx')

            # if index == 86:
            #     print(txt)
    # iterate on pair keys and values with index
    print(txt)
    moveFile(filename, path)
