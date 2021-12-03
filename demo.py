from selenium import webdriver
from selenium.common.exceptions import TimeoutException, NoSuchElementException
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.options import Options
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


path = resource_path('chromedriver.exe')
driver = webdriver.Chrome(path)

driver.maximize_window()
websiteURL = "https://tas.driverhire.co.uk/tas/invoicing/"
driver.get(websiteURL)
driver.implicitly_wait(10)
# op = input("\n\n\nEnter credentials...press Y when done: ")
# while True:
#     if(op != "Y"):
#         op = input("\n\n\n error: wrong key...please Enter credentials...press Y when done: ")
#     else:
#         break
# Login
try:
    driver.find_element_by_id("ctl00_ctl00__dh_UserName").send_keys("Dlaing")
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
driver.implicitly_wait(5)
try:
    driver.get(websiteURL)
except Exception as e:
    print("error: Cant get to invoicing..." + str(e))
try:
    inputTxt = driver.find_element_by_id("ctl00_ctl00_c_c__txtInvoiceRef__t")
    inputTxt.send_keys("141865")
    inputTxt.send_keys(Keys.ENTER)

except Exception as e:
    print("error: Cant find the element for input " + str(e))

# id for view =ctl00_ctl00_c_c__repeaterCustomers_ctl00__btnDownload
try:
    driver.find_element_by_id(
        "ctl00_ctl00_c_c__repeaterCustomers_ctl00__btnDownload").click()
except Exception as e:
    print("error: cant downlaod the file something went wrong..." + str(e))
driver.close()
