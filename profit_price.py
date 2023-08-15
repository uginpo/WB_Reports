import os.path
import pandas as pd
import numpy as np
import openpyxl
import xlsxwriter
import datetime
import sidetable
import locale
import sys


pd.set_option('display.float_format', '{:.2f}'.format)

path_to_database = f'/Users/uginpo/OneDrive - gxog/Отчеты ВБ/База данных/'

month_d = {'01': 'Январь', '02': 'Февраль', '03': 'Март', '04': 'Апрель', '05': 'Май', '06': 'Июнь',
           '07': 'Июль', '08': 'Август', '09': 'Сентябрь', '10': 'Октябрь', '11': 'Ноябрь', '12': 'Декабрь'}

# ### Блок настроек фильтров

clients = ['ООО Исток', 'ИП Искандаров', 'ИП Бурдукова', 'ИП Тараканов']
year_report = ['2020', '2021', '2022', '2023']

client = clients[0]
year_c = year_report[3]
period = ['02', '03', '04', '05', '06', '07']

df_profit_col = ['Месяц', 'Номер отчета', 'Кол-во', 'Сумма закупки', 'Сумма брака', 'Логистика', 'Хранение',
                 'Удержания', 'Очищенная выручка', 'Доход']

weekly_reports = pd.read_pickle(f'{path_to_database}weekly_reports_wb.pkl')
price_list = pd.read_pickle(f'{path_to_database}price_list_wb.pkl')
summary_report = pd.read_pickle(f'{path_to_database}summary_report_wb.pkl')

# удаление пустых строк
weekly_reports.dropna(subset=['Тип документа'], how='any', inplace=True)

ind = weekly_reports.loc[weekly_reports['Обоснование для оплаты'].isin(
    ['Возмещение издержек по перевозке'])].index
weekly_reports.drop(ind, axis=0, inplace=True)

# ### Накладываем фильтр по клиенту и году

filter_es = summary_report['Дата начала'].dt.year.astype(str) == year_c
summary_report = summary_report.loc[filter_es, :]
summary_report = summary_report.loc[summary_report['Клиент'] == client, :]
price_list = price_list.loc[price_list['Клиент'] == client, :]

weekly_reports = weekly_reports.loc[weekly_reports['Номер отчета'].isin(
    list(summary_report['Номер отчета'])), :]

# ### Меняем знак количества и сумм продаж для возвратов на отрицательный

col_float = ['Кол-во', 'Цена розничная', 'Вайлдберриз реализовал Товар (Пр)',
             'Согласованный продуктовый дисконт, %', 'Промокод %',
             'Итоговая согласованная скидка',
             'Цена розничная с учетом согласованной скидки',
             'Размер снижения кВВ из-за рейтинга, %',
             'Размер снижения кВВ из-за акции, %',
             'Скидка постоянного Покупателя (СПП)', 'Размер кВВ, %',
             'Размер  кВВ без НДС, % Базовый', 'Итоговый кВВ без НДС, %',
             'Вознаграждение с продаж до вычета услуг поверенного, без НДС',
             'Возмещение за выдачу и возврат товаров на ПВЗ',
             'Возмещение издержек по эквайрингу',
             'Вознаграждение Вайлдберриз (ВВ), без НДС',
             'НДС с Вознаграждения Вайлдберриз',
             'К перечислению Продавцу за реализованный Товар', 'Количество доставок',
             'Количество возврата', 'Услуги по доставке товара покупателю',
             'Общая сумма штрафов', 'Доплаты']

sum_col_float = ['Продажа', 'Согласованная скидка, %', 'К перечислению за товар',
                 'Стоимость логистики',
                 'Повышенная логистика согласно коэффициенту по обмерам',
                 'Другие виды штрафов', 'Общая сумма штрафов', 'Доплаты',
                 'Стоимость хранения', 'Стоимость платной приемки', 'Прочие удержания',
                 'Итого к оплате']

price_list['Закупочная цена'] = price_list['Закупочная цена'].astype(float)
price_list['Артикул поставщика'] = price_list['Артикул поставщика'].astype(str)
weekly_reports['Артикул поставщика'] = weekly_reports['Артикул поставщика'].astype(
    str)
weekly_reports[col_float] = weekly_reports[col_float].astype(float)
summary_report[sum_col_float] = summary_report[sum_col_float].astype(float)


# Меняем знак количества и сумм продаж для возвратов
group_col = ['Кол-во', 'Цена розничная с учетом согласованной скидки',
             'К перечислению Продавцу за реализованный Товар']
a = weekly_reports['Обоснование для оплаты'].isin(['Возврат', 'Сторно продаж'])
weekly_reports.loc[a, group_col] = weekly_reports.loc[a, group_col]*(-1)

# ### Меняем знак количества и сумм логистики при "сторно" на отрицательный

group_col = ['Услуги по доставке товара покупателю']
a = weekly_reports['Обоснование для оплаты'].isin(['Логистика сторно'])
weekly_reports.loc[a, group_col] = weekly_reports.loc[a, group_col]*(-1)

summary_report['Месяц'] = summary_report['Дата начала'].dt.month.astype(
    str).str.zfill(2)

# ### Формирование конечного отчета по артикулам

weekly_reports = pd.merge(weekly_reports, summary_report.loc[:, [
                          'Месяц', 'Номер отчета', 'Дата начала']], how='left')


weekly_reports = weekly_reports.sort_values(
    by='Дата начала').reset_index().drop('index', axis=1)

weekly_reports = weekly_reports.loc[weekly_reports['Месяц'].isin(period), :]

df_f = weekly_reports.copy(deep=True)

# ### Формируем таблицу отчетов по месяцам

df_f = df_f.merge(
    price_list.loc[:, ['Артикул поставщика', 'Закупочная цена']], how='left')

# ### Подсчет вознаграждения ВБ

df_f['Вознаграждение ВБ'] = df_f['Цена розничная с учетом согласованной скидки'] - \
    df_f['К перечислению Продавцу за реализованный Товар']

a = df_f['Услуги по доставке товара покупателю'] + \
    df_f['Общая сумма штрафов']+df_f['Доплаты']
df_f['Очищенная выручка'] = df_f['К перечислению Продавцу за реализованный Товар']-a
df_f['Сумма закупки'] = df_f['Закупочная цена'] * df_f['Кол-во']
df_f['Доход'] = df_f['Очищенная выручка']-df_f['Сумма закупки']

dict_col = {'Цена розничная с учетом согласованной скидки': 'Продажа со скидкой',
            'Услуги по доставке товара покупателю': 'Логистика',
            'Общая сумма штрафов': 'Штрафы'}
df_f.rename(columns=dict_col, inplace=True)


df_detal_col = ['Месяц', 'Предмет', 'Артикул поставщика', 'Название', 'Кол-во', 'Продажа со скидкой', 'Размер кВВ, %',
                'Вознаграждение ВБ', 'Логистика', 'Штрафы', 'Доплаты', 'Очищенная выручка', 'Закупочная цена',
                'Сумма закупки', 'Доход', 'Количество доставок', 'Количество возврата']
df_f = df_f.loc[:, df_detal_col]

index_list = ['Предмет', 'Артикул поставщика', 'Название']
col_list = ['Кол-во', 'Продажа со скидкой', 'Размер кВВ, %', 'Вознаграждение ВБ', 'Логистика', 'Штрафы', 'Доплаты',
            'Очищенная выручка', 'Сумма закупки', 'Доход', 'Количество доставок', 'Количество возврата', 'Закупочная цена']
func_list = ['sum', 'sum', 'max', 'sum', 'sum', 'sum',
             'sum', 'sum', 'sum', 'sum', 'sum', 'sum', 'max']
dict_group = dict(zip(col_list, func_list))
df_f = df_f.groupby(index_list).agg(dict_group).reset_index()

df_f = df_f.loc[df_f['Кол-во'] != 0, :]
df_f = df_f.loc[(df_f['Количество доставок'] +
                 df_f['Количество возврата']) != 0, :]
df_f = df_f.loc[df_f['Размер кВВ, %'] != 0, :]

df_f['Средняя цена продажи'] = df_f['Продажа со скидкой'] / df_f['Кол-во']

df_log = df_f['Логистика'] / \
    (df_f['Количество доставок'] + df_f['Количество возврата'])

df_fine = df_f['Штрафы'] / df_f['Кол-во']
df_extr = df_f['Доплаты'] / df_f['Кол-во']


# Рассчитываем цены исходя из 17% скидки для ВБ
discount = 0.17
disc = 1

df_f['Цена 0 (со скидкой)'] = (df_f['Закупочная цена']+df_log +
                               df_fine + df_extr)/(disc*(1-df_f['Размер кВВ, %']))

df_f['Цена 20% (со скидкой)'] = 1.2 * df_f['Цена 0 (со скидкой)'] / disc

df_f['Цена 40% (со скидкой)'] = 1.4 * df_f['Цена 0 (со скидкой)'] / disc


df_f['Ваше действие'] = (df_f['Цена 0 (со скидкой)']
                         * disc - df_f['Средняя цена продажи'])

df_f['Ваше действие'] = np.where(
    df_f['Ваше действие'] >= 0, 'Срочно увеличьте цену', '')

path_to_data = f'/Users/uginpo/OneDrive - gxog/Отчеты ВБ/{client}/'

report__pivot = f'/Аналитика/Точка_безубыточности_{client}_{year_c}.xlsx'
report_pivot = path_to_data + report__pivot

with pd.ExcelWriter(report_pivot, engine='xlsxwriter') as wb:
    df_f.to_excel(wb, sheet_name='price_0 ' + year_c, index=False,
                  float_format="%0.2f", freeze_panes=(1, 3))
    sheet = wb.sheets['price_0 ' + year_c]
    sheet.set_column('A:A', 20)
    sheet.set_column('B:B', 30)
    sheet.set_column('C:C', 35)
    sheet.set_column('D:T', 15)
    sheet.set_column('U:U', 20)

    cell_format = wb.book.add_format()
    cell_format.set_bold()
    cell_format.set_font_color('black')
    cell_format.set_num_format({'num_format': '# ##0'})
    # Установка стиля для строки 2 и высоты 40
    sheet.set_row(0, 30, cell_format)
    cell_format.set_font_size(14)
    cell_format.set_num_format({'num_format': '# ##0'})

# Add a header format.
    header_format = wb.book.add_format({
        'bold': True,
        'text_wrap': True,
        'valign': 'center',
        'fg_color': '#D7E4BC',
        'border': 1,
    })
    header_format.set_align('top')
    header_format.set_font_size(14)
    sheet.set_row(0, 40, header_format)
    # Write the column headers with the defined format.
    for col_num, value in enumerate(df_f.columns.values):
        sheet.write(0, col_num, value, header_format)

    cell_format = wb.book.add_format()
    cell_format.set_bold()
    cell_format.set_font_color('black')
    # cell_format.set_num_format(4)
    # for row_i in ind:
    #    sheet.set_row(row_i + 1, 25, cell_format)

    cell_format.set_bg_color('#A3E2B9')

    cell_format = wb.book.add_format()
    cell_format.set_bold()
    cell_format.set_font_color('black')
