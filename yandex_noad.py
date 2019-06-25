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



def search(search_list):
    pass




pages = 1

search_strings=[]
search_string_list = ['юридические услуги','интернет-магазин']
#filename = search_string
towns=['Москва','Санкт-Петербург','Новосибирск','Екатеринбург','Нижний Новгород','Казань','Челябинск','Омск','Самара',
      'Ростов','Уфа','Красноярск','Пермь','Воронеж','Волгоград']

for string in search_string_list:
    for town in towns:
        search_strings.append('{} {}'.format(string,town))

print(search_strings)

all_sites=defaultdict(list)

browser = Chrome('C:/Users/xxx/Downloads/chromedriver_win32/chromedriver.exe')
browser.get('https://yandex.ru/')
i=0


for search_string in search_strings:
    ii = 0
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
    sleep(random.randint(5, 10) / 100)
    search_button.click()
    i+=1
    sleep(random.randint(5, 10) / 100)
    while ii < pages:
        next = browser.find_element_by_link_text('дальше')
        next.click()


        results=browser.find_elements_by_xpath('/html/body/div[3]/div[1]/div[2]/div[1]/div[1]/ul/li[*]')


        for result in results:
            #if 'реклама' not in result.text:
                # print(result.get_attribute('innerHTML'))

            try:
                print(result.find_element_by_css_selector('a[class="link link_theme_outer path__item i-bem"]').text)
                all_sites[search_string].append(result.find_element_by_css_selector('a[class="link link_theme_outer path__item i-bem"]').text)
            except:
                pass
        ii+=1
print (all_sites)

# import xlsxwriter
# workbook = xlsxwriter.Workbook('C:/Users/xxx/PycharmProjects/wordstat/{}.xlsx'.format(filename))
# for sites in all_sites:
#     print (sites)
#     worksheet = workbook.add_worksheet(name=sites[len(filename):])
#     row = 0
#     col = 0
#     for site in all_sites[sites]:
#         print (site)
#         row += 1
#         worksheet.write(row, col,site)
#
#         col = 0


# workbook.close()

