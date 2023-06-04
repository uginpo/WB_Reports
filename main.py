import os.path
import pandas as pd
import numpy as np
import openpyxl
import xlsxwriter
import sidetable
import locale
import sys

clients = ['ООО Исток', 'ИП Искандаров', 'ИП Бурдукова', 'ИП Тараканов']
years_list = ['2021', '2022', '2023']

client = clients[0]
year_c = years_list[0]

path_to_data = f'/Users/uginpo/OneDrive - gxog/Отчеты ВБ/{client}/Сведенные/{year_c}/'
print(path_to_data)

name__price = "Закупка.xlsx"  # файл закупочных цен для расчета прибыли
name__list_otchet = "Отчеты.xlsx"  # список отчетов с доп. информацией - хранение и пр.

name__analitic = 'Аналитика.xlsx'
report__pivot = 'Сводный отчет.xlsx'

name_price = path_to_data + name__price
name_list_otchet = path_to_data + name__list_otchet
name_analitic = path_to_data + name__analitic
report_pivot = path_to_data + report__pivot
