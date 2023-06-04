import os.path
import pandas as pd
import numpy as np
import openpyxl
import xlsxwriter

# Настройка форматов отображения данный

pd.set_option('display.float_format', '{:.2f}'.format)

# Определение переменных

clients = ['ООО Исток', 'ИП Искандаров', 'ИП Бурдукова', 'ИП Тараканов']
years_list = ['2021', '2022', '2023']

client = clients[1]
year_c = years_list[2]

path_to_data = f'/Users/uginpo/OneDrive - gxog/Отчеты ВБ/{client}/Сведенные/{year_c}/'

name__price = "Закупка.xlsx"  # файл закупочных цен для расчета прибыли
name__list_otchet = "Отчеты.xlsx"  # список отчетов с доп. информацией - хранение и пр.

name__analitic = 'Аналитика.xlsx'
report__pivot = 'Сводный отчет.xlsx'

name_price = path_to_data + name__price
name_list_otchet = path_to_data + name__list_otchet
name_analitic = path_to_data + name__analitic
report_pivot = path_to_data + report__pivot

# Открытие списка отчетов
# Открытие прайслиста
df_list_otchet = pd.read_excel(name_list_otchet, header=0)  # список отчетов
df_list_otchet['Дата начала']=df_list_otchet['Дата начала'].dt.strftime('%d/%m/%Y')

df_price = pd.read_excel(name_price, header=0)  # pricelist закупочных цен

if os.path.isfile(name_analitic):
    df_analitic = pd.read_excel(name_analitic, sheet_name=[0, 1], header=0)
else:
    df_analitic = False
# Необходимые колонки из отчета

my_headers = ['Артикул поставщика', 'Название', 'Кол-во', 'Вайлдберриз реализовал Товар (Пр)',
              'Возмещение за выдачу и возврат товаров на ПВЗ', 'Возмещение издержек по эквайрингу',
              'Вознаграждение Вайлдберриз (ВВ), без НДС', 'НДС с Вознаграждения Вайлдберриз',
              'К перечислению Продавцу за реализованный Товар', 'Количество доставок', 'Количество возврата',
              'Услуги по доставке товара покупателю', 'Общая сумма штрафов']

new_headers = ['Артикул поставщика', 'Название', 'Кол-во', 'Количество доставок', 'Количество возврата',
               'Вайлдберриз реализовал Товар (Пр)',
               'Вознаграждение ВБ', 'Услуги по доставке товара покупателю',
               'Очищенная выручка']

final_column_list = ['Артикул поставщика', 'Название', 'Кол-во', 'Кол-во доставок', 'Кол-во возврата',
                     'Реализация ВБ', 'Вознаграждение ВБ',
                     'Логистика',
                     'Очищенная выручка']

final_headers = ['Отчет', 'Дата начала', 'Артикул поставщика', 'Название', 'Кол-во', 'Кол-во доставок',
                 'Кол-во возврата', 'Реализация ВБ', 'Вознаграждение ВБ', 'Логистика',
                 'Очищенная выручка', 'Закупочная цена', 'Сумма закупки', 'Доход']

list_profit_lost = ['Отчет', 'Дата начала', 'Кол-во', 'Сумма закупки', 'Логистика',
                    'Хранение', 'Удержания', 'Очищенная выручка', 'Доход']

# Цикл по недельным отчетам

# In[272]:


for i in range(df_list_otchet.shape[0]):
    df_otchet = df_list_otchet.iloc[i, [0, 1, 3, 4]]
    name_otchet = path_to_data + str(df_otchet['Отчет']) + '.xlsx'

    df = pd.read_excel(name_otchet, header=0)  # недельный отчет ВБ
    if df.shape[0] == 0:
        continue

    df = df.loc[:, my_headers]

    df['Вознаграждение ВБ'] = df[my_headers[4:8]].sum(axis=1)
    df['Очищенная выручка'] = df['К перечислению Продавцу за реализованный Товар'] - df[
        'Услуги по доставке товара покупателю']

    df = df.loc[:, new_headers]
    df.columns = final_column_list

    df = df.groupby(['Артикул поставщика', 'Название'])[final_column_list[2:]].sum().reset_index()

    df_f = df.copy(deep=True)

    list_to_add = ['Отчет', 'Дата начала']
    df_f[list_to_add] = df_otchet.loc[list_to_add]

    list_to_add = ['Сумма закупки', 'Доход']
    df_f[list_to_add] = 0

    df_f = pd.merge(df_f, df_price)
    df_f = df_f.loc[:, final_headers]

    df_f['Сумма закупки'] = df_f['Кол-во'] * df_f['Закупочная цена']
    df_f['Доход'] = df_f['Очищенная выручка'] - df_f['Сумма закупки']

    # Отчет по продажам сформирован
    # Формирование отчета по прибылям и убыткам

    df_pr_lost = pd.DataFrame(np.zeros(9).reshape(1, 9), columns=list_profit_lost)

    #    Подсчет итогов (суммирование итогов недельных продаж

    list_itog = ['Кол-во', 'Сумма закупки', 'Логистика', 'Очищенная выручка', 'Доход']

    df_pr_lost[['Отчет', 'Дата начала']] = df_otchet[['Отчет', 'Дата начала']]
    df_pr_lost[list_itog] = df_f[list_itog].sum()
    df_pr_lost[['Хранение', 'Удержания']] = df_otchet[['Хранение', 'Удержания']]
    df_pr_lost['Доход'] = df_pr_lost['Доход'] - df_pr_lost['Хранение'] - df_pr_lost['Удержания']

    if not df_analitic:
        df_analitic = dict()
        df_analitic[0] = df_f.copy(deep=True)
        df_analitic[1] = df_pr_lost.copy(deep=True)
    else:
        df_analitic[0] = pd.concat([df_analitic[0], df_f], axis=0, ignore_index=True)
        df_analitic[1] = pd.concat([df_analitic[1], df_pr_lost], axis=0, ignore_index=True)

with pd.ExcelWriter(name_analitic, engine='xlsxwriter') as wb:
    df_analitic[0].to_excel(wb, sheet_name='sales', index=False, float_format="%0.2f")
    df_analitic[1].to_excel(wb, sheet_name='profits', index=False, float_format="%0.2f")
    sheet = wb.sheets['sales']
    # sheet.autofilter(0, 0, 0, 15)
    sheet.set_column('A:A', 15)
    sheet.set_column('B:B', 15)
    sheet.set_column('C:C', 45)
    sheet.set_column('D:P', 14)

    sheet0 = wb.sheets['profits']
    sheet0.set_column('A:I', 15)
