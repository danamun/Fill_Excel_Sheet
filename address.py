""" ***Author: Dana Mun***
A simple program that uses Yelp.com to get random existing addresses based on the zipcode provided in fileName.xls (make sure it's in .xls extension).
This code is compatible with Python 2.7.3 and need to download modules:
	1. python pip -m install selenium 
	2. python pip -m install xlrd
	3. python pip -m install xlwt
	4. python pip -m install xlutils
To fill a certain excel sheet change field named fileName below. 

To run the file you do python27 address.py (make sure the python.exe in your python path is renamed to python27.exe)
"""

from selenium import webdriver
from selenium.webdriver.common.keys import Keys
import xlrd
from xlwt import Workbook
from xlutils.copy import copy

"""*Enter Name of File here*"""
fileName = 'zipcodes.xls'
website = "http://www.yelp.com"
driver = webdriver.PhantomJS()
#driver = webdriver.Firefox() 
workbook = xlrd.open_workbook(fileName)
worksheet = workbook.sheet_by_name('Sheet1')

wb = copy(workbook)
s = wb.get_sheet(0)

num_rows = worksheet.nrows - 1
curr_row = 0
correctAddr = ''
while curr_row < num_rows:
	curr_row += 1
	row = worksheet.row(curr_row)
	print ("Zip Code: ", int(row[0].value))
	driver.get(website)
	elem = driver.find_element_by_id("dropperText_Mast")
	driver.implicitly_wait(10)
	elem.clear()
	driver.implicitly_wait(10)
	elem.send_keys(str(int(row[0].value)))
	driver.implicitly_wait(10)
	elem.send_keys(Keys.RETURN)
	driver.implicitly_wait(10)
	address = driver.find_elements_by_tag_name('address')
	for i in range(0, len(address)):
		if str(int(row[0].value)) in address[i].text:
			correctAddr = address[i].text
			break
		i += 1
	s.write(curr_row, 1, correctAddr.replace('\n', ' '))
	wb.save(fileName)
	print ("Zip Code: ", row[0], "Address Found: ", correctAddr.replace('\n', ' '))
	correctAddr = ''

driver.close()