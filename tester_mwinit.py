import openpyxl
from openpyxl import workbook
import getpass
import datetime
import traceback
from webdriver_manager.chrome import ChromeDriverManager
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from collections import OrderedDict
from selenium.webdriver.support.ui import Select
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
import time
import traceback
from datetime import date, datetime, timedelta
# import getpass

now = datetime.now()
print("Start_time:" + now.strftime("%Y-%m-%d %H:%M:%S"))


username = getpass.getuser()
# path = (r"C:\\Users\\" + username + "\\Desktop\\Work_Flow\\Procurement_input.xlsx")
# rows = getRowCount(path, 'Sheet1')
# columns = getColumnCount(path, 'Sheet1')
driver = webdriver.Chrome(executable_path=r"C:\\Users\\" + username + "\\Desktop\\Bin\\SIM\\chromedriver.exe")
driver.get('https://procurementportal-eu.corp.amazon.com/')
time.sleep(10)
asin = 'B095DNPH4R'
driver.find_element(By.XPATH, "/html/body/div/div/div/div/div[2]/div/div/div/div/div/div[1]/div/div/div[2]/div/div/div[2]/div/table/tbody/tr[1]/td[1]/kat-input").send_keys(asin)
time.sleep(2)
# drop = driver.find_element(By.XPATH,"""/html/body/div/div/div/div/div[2]/div/div/div/div/div/div[1]/div/div/div[2]/div/div/div[2]/div/table/tbody/tr[2]/td/kat-dropdown//div[1]/div[1]/div/div[1]/slot/text()""").click()
el = Select(driver.find_elements(By.TAG_NAME, "slot")
options = [x for x in el.find_elements(By.TAG_NAME, "kat-option")]
for element in options:
    print(element.get_attribute("DE"))

# dropdown_menu = Select(driver.find_element_by_tag_name("slot"))
# for option in dropdown_menu.options:
#     print(option.text, option.get_attribute('DE'))

# for r in range(2, rows - 1):
#     try:
#         writeData(path, "Sheet1", r, 399, " ")
#     except PermissionError:
#         driver.execute_script("alert('PLEASE CLOSE THE INPUT EXCEL AND RUN AGAIN');")
#     else:
#         pass
#     asin = readData(path, "Sheet1", r, 1)
#     driver.find_element(By.ID, "katal-id-46").send_keys(asin)
#     time.sleep(2)
#     select = Select(driver.find_element_by_id('katal-id-47'))
#     select.select_by_visible_text('DE')


