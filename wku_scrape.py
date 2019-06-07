#from bs4 import BeautifulSoup
#import requests
from selenium import webdriver
from selenium.webdriver.support.ui import Select
import xlsxwriter




workbook = xlsxwriter.Workbook('./data/data.xlsx')
bold = workbook.add_format({'bold':1})
url = 'https://acsapps.wku.edu/pls/prod/dirpkg.prompt'
driver = webdriver.Chrome()
driver.get(url)


departments = []
for i in driver.find_elements_by_xpath('//option'):
    departments.append(i.text)

driver.close()

for dept in range(len(departments)):
    try:
        worksheet_admin = workbook.add_worksheet(departments[dept][:28])
    except:
        worksheet_admin =  workbook.add_worksheet('data'+str(dept))
    print(departments[dept])
    worksheet_admin.write('A2', 'DATA EXTRACTED FROM')
    worksheet_admin.write('B3', url)
    worksheet_admin.write('A5', 'Name', bold)   
    worksheet_admin.write('B5', 'Phone', bold)
    worksheet_admin.write('C5', 'Designation', bold)
    worksheet_admin.write('D5', 'Email', bold)
    worksheet_admin.write('E5', 'Office', bold)
    
    url = 'https://acsapps.wku.edu/pls/prod/dirpkg.prompt'
    driver = webdriver.Chrome()
    driver.get(url)
    
    select = Select(driver.find_element_by_name('dept'))
    select.select_by_visible_text(departments[dept])
    driver.find_element_by_xpath('//input[@tabindex=10]').click()
    
    content = driver.find_element_by_xpath('//div[@class="one_column"]')
    td= content.find_elements_by_xpath('//td')
    list_ = []
    for i in td:
        list_.append(i.text)
    
    intro,b = [], 0
    
    z = int(len(list_)/6)
    for i in range(z):
        if i == 0:
            intro.append(list_[0+b:5+b])
            b +=5
        else:
            intro.append(list_[0+b+1:5+b+1])
            b +=6
    
    intro
    row = 5
    col = 0
    for i in range(len(intro)):
        worksheet_admin.write_string(row, col, intro[i][0])
        try:
            worksheet_admin.write_string(row, col+1, intro[i][1])
        except IndexError:
            worksheet_admin.write_string(row,col+1,'')
        try:
            worksheet_admin.write_string(row, col+2, intro[i][2])
        except IndexError:
            worksheet_admin.write_string(row,col+2,'')
        try:
            worksheet_admin.write_string(row, col+3, intro[i][3])
        except IndexError:
            worksheet_admin.write_string(row,col+3,'')
        try:
            worksheet_admin.write_string(row, col+4, intro[i][4])
        except IndexError:
            worksheet_admin.write_string(row, col+4, '')
        try:
            worksheet_admin.write_string(row, col+5, intro[i][5])
        except IndexError:
            worksheet_admin.write_string(row, col+5, '')
        
        row +=1

    driver.close() 


workbook.close()
    