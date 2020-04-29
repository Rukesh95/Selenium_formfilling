import os
from _collections import OrderedDict
from datetime import time
from telnetlib import EC

from selenium.webdriver.remote.webelement import WebElement
from selenium.webdriver.support.select import Select

from selenium import webdriver
from selenium.webdriver.common.by import By
from urllib3.util import wait
from xlutils.copy import copy
import xlrd

driver=webdriver.Chrome("C:\\Users\\Rukesh\\PycharmProjects\\workshop\\Drivers\\chromedriver.exe")
driver.maximize_window()
driver.get("https://forms.gle/dqBWNo9C4fQR5HC98")
driver.find_element_by_name('emailAddress').send_keys('wastebosy@gmail.com')
#driver.find_element_by_xpath('/html/body/div[1]/div[1]/div[2]/div/div[2]/div/div/div[2]/div/div[2]/div/div[1]/div/span/span').click();
#driver.implicitly_wait(10)
#driver.find_element_by_name('password').send_keys('$Waste1$')
#driver.find_element_by_xpath('//*[@id="passwordNext"]/span/span').click();

my_path = os.path.abspath(os.path.dirname(__file__))
excelpath = os.path.join(my_path,"C:\\Users\\Rukesh\\Desktop\\test.xls")
excelbook = xlrd.open_workbook(excelpath)
worksheet = excelbook.sheet_by_index(0)
name = worksheet.cell_value(1,0)
Gender = worksheet.cell_value(1,1)
Education = worksheet.cell_value(1,2)
Semester = worksheet.cell_value(1,3)

driver.implicitly_wait(10)

driver.find_element_by_name('entry.1520926271').send_keys(name);
driver.implicitly_wait(10)
if Gender == 'Male' or Gender == 'male':
    driver.find_element_by_xpath('//*[@id="mG61Hd"]/div/div/div[2]/div[3]/div/div[2]/div/span/div/div[1]/label/div/div[1]').click();
elif Gender == 'Female' or Gender == 'female':
    driver.find_element_by_xpath('//*[@id="mG61Hd"]/div/div/div[2]/div[3]/div/div[2]/div/span/div/div[2]/label/div/div[1]').click();
else :
    driver.find_element_by_xpath('//*[@id="mG61Hd"]/div/div/div[2]/div[3]/div/div[2]/div/span/div/div[3]/label/div/div[1]').click();
driver.implicitly_wait(10)
if Education == 'UG':
    driver.find_element_by_xpath('//*[@id="mG61Hd"]/div/div/div[2]/div[4]/div/div[2]/div[1]/div/label/div/div[1]').click();
elif Education == 'PG':
    driver.find_element_by_xpath('//*[@id="mG61Hd"]/div/div/div[2]/div[4]/div/div[2]/div[2]/div/label/div/div[1]').click();
else:
    driver.find_element_by_xpath('//*[@id="mG61Hd"]/div/div/div[2]/div[3]/div/div[2]/div[1]/div/label').click();
    driver.find_element_by_xpath('//*[@id="mG61Hd"]/div/div/div[2]/div[3]/div/div[2]/div[2]/div/label').click();
driver.implicitly_wait(10)
if Semester == 1:
    driver.find_element_by_xpath('//*[@id="mG61Hd"]/div/div/div[2]/div[5]/div/div[2]/div/span/div/div[1]/label/div/div[1]').click()
elif Semester == 2:
    driver.find_element_by_xpath('//*[@id="mG61Hd"]/div/div/div[2]/div[5]/div/div[2]/div/span/div/div[2]/label/div/div[1]').click()
elif Semester == 3:
    driver.find_element_by_xpath('//*[@id="mG61Hd"]/div/div/div[2]/div[5]/div/div[2]/div/span/div/div[3]/label/div/div[1]').click()
elif Semester == 4:
    driver.find_element_by_xpath('//*[@id="mG61Hd"]/div/div/div[2]/div[5]/div/div[2]/div/span/div/div[4]/label/div/div[1]').click()
elif Semester == 5:
    driver.find_element_by_xpath('//*[@id="mG61Hd"]/div/div/div[2]/div[5]/div/div[2]/div/span/div/div[5]/label/div/div[1]').click()
elif Semester == 6:
    driver.find_element_by_xpath('//*[@id="mG61Hd"]/div/div/div[2]/div[5]/div/div[2]/div/span/div/div[6]/label/div/div[1]').click()
driver.implicitly_wait(5)
#driver.find_element_by_xpath('//*[@id="mG61Hd"]/div/div/div[2]/div[5]/div/div[3]/span/span[2]').click();
#driver.implicitly_wait(10)
#driver.find_element_by_xpath('//*[@id="doclist"]/div/div[4]/div[2]/div/div[2]/div/div[1]/div[1]/div/div[2]/div[1]/div/div/div[3]').send_keys('guvfyg')
driver.find_element_by_xpath('//*[@id="mG61Hd"]/div/div/div[3]/div[1]/div/div').click()