import random
from time import sleep
import csv
from selenium.webdriver import Chrome
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
import re

search_string = '3д панели москва'
towns=['Москва','Санкт-Петербург','Новосибирск','Екатеринбург','Нижний Новгород','Казань','Челябинск','Омск','Самара',
       'Ростов','Уфа','Красноярск','Пермь','Воронеж','Волгоград']
all_sites=[]
browser = Chrome('C:/Users/xxx/Downloads/chromedriver_win32/chromedriver.exe')
browser.get('https://yandex.ru/')

search_field = browser.find_element_by_xpath('//*[@id="text"]')
search_field.send_keys(search_string)
search_button = browser.find_element_by_css_selector('button[class="button suggest2-form__button button_theme_websearch button_size_ws-head i-bem"]')
search_button.click()
i=0
while i<10:
    next = browser.find_element_by_link_text('дальше')
    next.click()

    result=browser.find_elements_by_xpath('/html/body/div[3]/div[1]/div[2]/div[1]/div[1]/ul/li[*]/div/div[*]/div[1]/a[1]/b')

    for r in result:
        if r.text not in all_sites:
            all_sites.append(r.text)
    i+=1

print(all_sites)

import xlsxwriter
workbook = xlsxwriter.Workbook('C:/Users/xxx/PycharmProjects/wordstat/myfile.xlsx')
worksheet = workbook.add_worksheet(name=search_string)
row = 0
col = 0

for key in all_sites:
    row += 1
    worksheet.write(row, col,     key)

    col = 0

workbook.close()