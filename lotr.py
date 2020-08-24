from selenium import webdriver
import os
import openpyxl
from openpyxl import Workbook

driver = webdriver.Chrome()
driver.get('http://www.ageofthering.com/atthemovies/scripts/fellowshipofthering1to4.php')

rows = len(driver.find_elements_by_xpath("//*[@id='AutoNumber1']/tbody/tr"))
columns = len(driver.find_elements_by_xpath("//*[@id='AutoNumber1']/tbody/tr[3]/td"))
first = "//*[@id='AutoNumber1']/tbody/tr["
second = "]/td["
third = "]"

for r in range(1, rows+1):
    for c in range(1, columns+1):
        try:
            final = first + str(r) + second + str(c) + third
            data = driver.find_element_by_xpath(final).text
            #print(data, end=" ")
            fname = 'script.xlsx'
            if(os.path.exists(fname)):
                workbook = openpyxl.load_workbook(fname)
                worksheet = workbook.get_sheet_by_name('Sheet')
            else:
                workbook = Workbook()
                worksheet = workbook.active
            worksheet.cell(row=r, column=c).value = data
            workbook.save(fname)
        except:
            continue
    #print("")
