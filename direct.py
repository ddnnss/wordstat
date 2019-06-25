import random
from time import sleep
import csv
from selenium.webdriver import Chrome
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
import re
from collections import defaultdict

search_strings = ['3д панели москва','декоративный камень москва', 'самокаты оптом спб']

all_sites=defaultdict(list)

browser = Chrome('C:/Users/xxx/Downloads/chromedriver_win32/chromedriver.exe')
browser.get('https://yandex.ru/')
i=0


for search_string in search_strings:
    if i==0:
        search_field = browser.find_element_by_xpath('//*[@id="text"]')
        search_button = browser.find_element_by_css_selector(
            'button[class="button suggest2-form__button button_theme_websearch button_size_ws-head i-bem"]')
        search_field.clear()
    else:

        search_field = browser.find_element_by_css_selector('input[class="input__control"]')
        search_button = browser.find_element_by_css_selector('button[class="websearch-button suggest2-form__button"]')
        search_field.clear()
    search_field.send_keys(search_string)
    search_button.click()
    i+=1


    results=browser.find_elements_by_xpath('/html/body/div[3]/div[1]/div[2]/div[1]/div[1]/ul/li[*]')


    for result in results:
        if 'реклама' in result.text:
            # print(result.get_attribute('innerHTML'))
            print(result.find_element_by_css_selector('a[class="link link_theme_outer path__item i-bem"]').text)
            all_sites[search_string].append(result.find_element_by_css_selector('a[class="link link_theme_outer path__item i-bem"]').text)
print (all_sites)

import xlsxwriter
workbook = xlsxwriter.Workbook('C:/Users/xxx/PycharmProjects/wordstat/myfile.xlsx')
for sites in all_sites:
    print (sites)
    worksheet = workbook.add_worksheet(name=sites)
    row = 0
    col = 0
    for site in all_sites[sites]:
        print (site)
        row += 1
        worksheet.write(row, col,     site)

        col = 0


workbook.close()