import sqlite3
import xlsxwriter
from collections import defaultdict

result = defaultdict(list)
workbook = xlsxwriter.Workbook('C:/Users/xxx/PycharmProjects/wordstat/ya_xls.xlsx')
worksheet = workbook.add_worksheet()
row = 1
conn = sqlite3.connect('C:/Users/xxx/PycharmProjects/wordstat/data.sqlite', timeout=10)
cursor = conn.cursor()

cursor.execute("SELECT section_id FROM position_oil")
position_oil_section_id = cursor.fetchall()
print(position_oil_section_id)

for item_id in position_oil_section_id:
    cursor.execute("SELECT alias FROM catalog_section WHERE id=(?)", item_id)
    item_url = 'https://lm74.ru/' + cursor.fetchone()[0] + '.html'
    cursor.execute("SELECT title FROM catalog_section WHERE id=(?)", item_id)
    item_title = cursor.fetchone()[0]
    cursor.execute("SELECT parent_id FROM catalog_section WHERE id=(?)", item_id)
    item_parent_id = cursor.fetchone()[0]
    cursor.execute("SELECT title FROM catalog_section WHERE id=(?)", (item_parent_id,))
    item_category = cursor.fetchone()[0]
    cursor.execute("SELECT price FROM position_oil WHERE section_id=(?)", item_id)
    item_price = cursor.fetchone()[0]
    cursor.execute("SELECT article FROM position_oil WHERE section_id=(?)", item_id)
    item_article = cursor.fetchone()[0]
    cursor.execute("SELECT liters FROM position_oil WHERE section_id=(?)", item_id)
    item_liters = 'Объем: ' + cursor.fetchone()[0] + 'л.'
    cursor.execute("SELECT img FROM section_oil WHERE id=(?)", item_id)
    item_img = 'https://lm74.ru/' + cursor.fetchone()[0]
    cursor.execute("SELECT usages FROM section_oil WHERE id=(?)", item_id)
    item_desr = 'АРТИКУЛ: {} | '.format(item_article) + cursor.fetchone()[0]
    # id
    worksheet.write('A{}'.format(row), row)
    # status
    worksheet.write('B{}'.format(row), '')
    # Доставка
    worksheet.write('C{}'.format(row), '')
    #Стоимость доставки
    worksheet.write('D{}'.format(row), '')
    #Срок доставки
    worksheet.write('E{}'.format(row), '')
    #Самовывоз
    worksheet.write('F{}'.format(row), '')
    #Стоимость самовывоза
    worksheet.write('G{}'.format(row), '')
    #Срок самовывоза
    worksheet.write('H{}'.format(row), '')
    #Купить в магазине без заказа
    worksheet.write('I{}'.format(row), '')
    #Ссылка на товар на сайте магазина*
    worksheet.write('J{}'.format(row), item_url)
    #Производитель
    worksheet.write('K{}'.format(row), '')
    #Название*
    worksheet.write('L{}'.format(row), item_title)
    #Категория*
    worksheet.write('M{}'.format(row), item_category)
    #Цена*
    worksheet.write('N{}'.format(row), item_price)
    #Цена без скидки
    worksheet.write('O{}'.format(row), '')
    #Валюта*
    worksheet.write('P{}'.format(row), 'RUR')
    #Ссылка на картинку*
    worksheet.write('Q{}'.format(row), item_img)
    #Описание
    worksheet.write('R{}'.format(row), item_desr)
    #Характеристики товара
    worksheet.write('S{}'.format(row), item_liters)
    #Условия продажи
    worksheet.write('T{}'.format(row), '')
    #Гарантия производителя
    worksheet.write('U{}'.format(row), '')
    #Страна происхождения
    worksheet.write('V{}'.format(row), '')
    #Штрихкод
    worksheet.write('W{}'.format(row), '')
    #bid
    worksheet.write('X{}'.format(row), '')
    #Уцененный товар
    worksheet.write('Y{}'.format(row), '')
    #Причина уценки
    worksheet.write('Z{}'.format(row), '')


    row += 1
    print(item_img)

workbook.close()
