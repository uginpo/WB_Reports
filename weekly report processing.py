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


clients = ['ООО Исток', 'ИП Искандаров', 'ИП Бурдукова', 'ИП Тараканов']
year_report = ['2020', '2021', '2022', '2023']

client = clients[2]
year_c = year_report[3]

path_to_database = f'/Users/uginpo/OneDrive - gxog/Отчеты ВБ/База данных/'

month_d = {'01': 'Январь', '02': 'Февраль', '03': 'Март', '04': 'Апрель', '05': 'Май', '06': 'Июнь',
           '07': 'Июль', '08': 'Август', '09': 'Сентябрь', '10': 'Октябрь', '11': 'Ноябрь', '12': 'Декабрь'}
month_list = ['Январь', 'Февраль', 'Март', 'Апрель', 'Май', 'Июнь',
              'Июль', 'Август', 'Сентябрь', 'Октябрь', 'Ноябрь', 'Декабрь']

# ## Исходные столбцы
# ### детализация за неделю и сводный за неделю

weekly_col = ['№', 'Номер поставки', 'Предмет', 'Код номенклатуры', 'Бренд',
              'Артикул поставщика', 'Название', 'Размер', 'Баркод', 'Тип документа',
              'Обоснование для оплаты', 'Дата заказа покупателем', 'Дата продажи',
              'Кол-во', 'Цена розничная', 'Вайлдберриз реализовал Товар (Пр)',
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
              'Общая сумма штрафов', 'Доплаты', 'Виды логистики, штрафов и доплат',
              'Стикер МП', 'Наименование банка-эквайера', 'Номер офиса',
              'Наименование офиса доставки', 'ИНН партнера', 'Партнер', 'Склад',
              'Страна', 'Тип коробов', 'Номер таможенной декларации',
              'Код маркировки', 'ШК', 'Rid', 'Srid',
              'Возмещение издержек по перевозке', 'Организатор перевозки',
              'Номер отчета', 'Клиент']

summary_col = ['Номер отчета', 'Дата начала', 'Дата конца', 'Дата формирования',
               'Продажа', 'Согласованная скидка, %', 'К перечислению за товар',
               'Стоимость логистики',
               'Повышенная логистика согласно коэффициенту по обмерам',
               'Другие виды штрафов', 'Общая сумма штрафов', 'Доплаты',
               'Стоимость хранения', 'Стоимость платной приемки', 'Прочие удержания',
               'Итого к оплате', 'Валюта', 'Статус отчета', 'Клиент']

pricelist_col = ['Предмет', 'Артикул поставщика', 'Закупочная цена', 'Клиент']


# ### Значения ключевых полей для работы над отчетом

# Перечень поля 'Тип документа'
list_tip_doc = ['Продажа', 'Возврат']

# Перечень поля 'Обоснование для оплаты'
list_obosn_oplaty = ['Продажа', 'Логистика', 'Штрафы', 'Сторно продаж', 'Частичная компенсация брака',
                     'Корректная продажа', 'Возврат', 'Возмещение издержек по перевозке', 'Доплаты',
                     'Авансовая оплата за товар без движения', 'Логистика сторно', 'Штрафы и доплаты']

# Перечень поля 'Виды логистики, штрафов и доплат'
list_type_logistic = ['К клиенту при продаже',
                      'Платное хранение возвратов на ПВЗ более 3 дней',
                      'От клиента при отмене', 'К клиенту при отмене',
                      'Возврат неопознанного товара (К продавцу)',
                      'Возврат неопознанного товара (От продавца при отмене)',
                      'Возврат своего товара (От продавца при отмене)',
                      'Возврат брака (К продавцу)', 'От клиента при возврате',
                      'Возврат брака (От продавца при отмене)',
                      'Возврат своего товара (К продавцу)',
                      'Доплата за логистику с палето-мест',
                      'Возврат по инициативе продавца (От продавца при отмене)',
                      'Штраф МП. Нарушение срока', 'Штраф МП. Невыполненный заказ',
                      'Сторно. Штраф МП. Невыполненный заказ',
                      'Выявленные расхождения в карточке товара после приемки на складе WB']

# ## Столбы итоговых отчетов
# ### детализация продаж по артикулам и отчет о прибылях и убытках

df_detal_col = ['Месяц', 'Номер отчета', 'Предмет', 'Артикул поставщика', 'Название', 'Кол-во', 'Продажа со скидкой',
                'Вознаграждение ВБ', 'Логистика', 'Штрафы', 'Доплаты', 'Очищенная выручка', 'Закупочная цена',
                'Сумма закупки', 'Доход']

df_profit_col = ['Месяц', 'Номер отчета', 'Кол-во', 'Сумма закупки', 'Сумма брака', 'Логистика', 'Хранение',
                 'Удержания', 'Очищенная выручка', 'Доход']

# ### Открытие списка отчетов и прайслиста

weekly_reports = pd.read_pickle(f'{path_to_database}weekly_reports_wb.pkl')
price_list = pd.read_pickle(f'{path_to_database}price_list_wb.pkl')
summary_report = pd.read_pickle(f'{path_to_database}summary_report_wb.pkl')

# ### Удаляем служебные строки WB
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

group_col = ['Услуги по доставке товара покупателю', 'Количество доставок']
a = weekly_reports['Обоснование для оплаты'].isin(['Логистика сторно'])
weekly_reports.loc[a, group_col] = weekly_reports.loc[a, group_col]*(-1)

summary_report['Месяц'] = summary_report['Дата начала'].dt.month.astype(
    str).str.zfill(2)

# ### Формирование конечного отчета по артикулам

weekly_reports = pd.merge(weekly_reports, summary_report.loc[:, [
                          'Месяц', 'Номер отчета', 'Дата начала']], how='left')

weekly_reports = weekly_reports.sort_values(
    by='Дата начала').reset_index().drop('index', axis=1)

# ### Фрейм weekly_reports содержит информацию о всех транзакциях за год, привязанные к недельным отчетам

df_f = weekly_reports.copy(deep=True)


# Формируем DataFrame логистика df_logistic

col_l = ['Месяц', 'Предмет', 'Артикул поставщика', 'Название', 'Количество доставок',
         'Количество возврата', 'Стоимость доставки', 'Стоимость возврата', 'Логистика всего']
index_list = ['Месяц', 'Предмет', 'Артикул поставщика', 'Название']
itog_str = f'- Итого за месяц'

df_logistic = df_f.loc[df_f['Обоснование для оплаты'].isin(
    ['Логистика', 'Логистика сторно']), :].reset_index()

df_logistic['Стоимость доставки'] = df_logistic['Количество доставок'] * \
    df_logistic['Услуги по доставке товара покупателю']

df_logistic['Стоимость возврата'] = df_logistic['Количество возврата'] * \
    df_logistic['Услуги по доставке товара покупателю']

df_logistic['Логистика всего'] = df_logistic['Стоимость доставки'] + \
    df_logistic['Стоимость возврата']

df_logistic = df_logistic.loc[:, col_l]

df_logistic = df_logistic.groupby(index_list).sum()
df_logistic = df_logistic.stb.subtotal(
    sub_level=[1], grand_label='Всего за период', sub_label='Итого за месяц').reset_index()

df_logistic = df_logistic.rename(
    columns={"level_0": "Месяц", "level_1": "Предмет"})
df_logistic = df_logistic.rename(
    columns={"level_2": "Артикул поставщика", "level_3": "Название"})

df_logistic = df_logistic.replace({'Месяц': month_d})
df_logistic['Предмет'].where(~(df_logistic['Предмет'].str.contains(
    itog_str)), other='Итого за месяц', inplace=True)


# Формируем DataFrame 'потеряшка' df_p

df_p = df_f.loc[df_f['Обоснование для оплаты'].isin(
    ['Частичная компенсация брака', 'Авансовая оплата за товар без движения']), :].reset_index()


# ### Формируем таблицу отчетов по месяцам

df_f = df_f.merge(
    price_list.loc[:, ['Артикул поставщика', 'Закупочная цена']], how='left')

# ### Подсчет вознаграждения ВБ

df_f['Вознаграждение ВБ'] = df_f['Цена розничная с учетом согласованной скидки'] - \
    df_f['К перечислению Продавцу за реализованный Товар']

a = df_f['Услуги по доставке товара покупателю'] + \
    df_f['Общая сумма штрафов']+df_f['Доплаты']
df_f['Очищенная выручка'] = df_f['К перечислению Продавцу за реализованный Товар']-a
df_f['Сумма закупки'] = df_f['Закупочная цена']*df_f['Кол-во']
df_f['Доход'] = df_f['Очищенная выручка']-df_f['Сумма закупки']

list_itog = ['Кол-во', 'Сумма закупки']
df_itog = df_f.groupby(['Номер отчета'])[list_itog].sum().reset_index()

dict_col = {'Цена розничная с учетом согласованной скидки': 'Продажа со скидкой',
            'Услуги по доставке товара покупателю': 'Логистика',
            'Общая сумма штрафов': 'Штрафы'}
df_f.rename(columns=dict_col, inplace=True)

df_detal_col = ['Месяц', 'Предмет', 'Артикул поставщика', 'Название', 'Кол-во', 'Продажа со скидкой',
                'Вознаграждение ВБ', 'Логистика', 'Штрафы', 'Доплаты', 'Очищенная выручка', 'Закупочная цена',
                'Сумма закупки', 'Доход']
df_f = df_f.loc[:, df_detal_col]

# ### Группировка значений и добавление промежуточных итогов

index_list = ['Месяц', 'Предмет', 'Артикул поставщика', 'Название']
col_list = ['Кол-во', 'Продажа со скидкой', 'Вознаграждение ВБ', 'Логистика', 'Штрафы', 'Доплаты',
            'Очищенная выручка', 'Закупочная цена', 'Сумма закупки', 'Доход']
func_list = ['sum', 'sum', 'sum', 'sum',
             'sum', 'sum', 'sum', 'max', 'sum', 'sum']
dict_group = dict(zip(col_list, func_list))

df_f = df_f.groupby(index_list).agg(dict_group)
df_f = df_f.stb.subtotal(sub_level=[
                         1], grand_label='Всего за период', sub_label='Итого за месяц').reset_index()
df_f = df_f.rename(columns={"level_0": "Месяц", "level_1": "Предмет"})
df_f = df_f.rename(
    columns={"level_2": "Артикул поставщика", "level_3": "Название"})
df_f = df_f.replace({'Месяц': month_d})
itog_str = f'- Итого за месяц'
df_f['Предмет'].where(~(df_f['Предмет'].str.contains(
    itog_str)), other='Итого за месяц', inplace=True)

path_to_data = f'/Users/uginpo/OneDrive - gxog/Отчеты ВБ/{client}/'

report__pivot = f'/Аналитика/Сводный отчет_{client}_{year_c}.xlsx'
report_pivot = path_to_data + report__pivot

# ## Формирование отчета по прибылям и убыткам

df_pr_f = summary_report.sort_values(
    by='Дата начала').reset_index().drop('index', axis=1)

# ### Подсчет итогов (суммирование итогов недельных продаж

df_pr_f = df_pr_f.merge(df_itog, how='left')

df_pr_f['Прибыль'] = df_pr_f['Итого к оплате']-df_pr_f['Сумма закупки']

# ### Сохранение исходного файла сводных недельных отчетов

df_pr_b = df_pr_f.copy(deep=True)

df_pr_b.columns = ['Номер отчета', 'Дата начала', 'Дата конца', 'Дата формирования',
                   'Продажа', 'Согласованная скидка, %', 'К перечислению за товар',
                   'Логистика', 'Повышенная логистика согласно коэффициенту по обмерам',
                   'Другие виды штрафов', 'Штрафы', 'Доплаты',
                   'Хранение', 'Платная приемка', 'Удержания',
                   'К оплате', 'Валюта', 'Статус отчета', 'Клиент', 'Месяц',
                   'Кол-во', 'Сумма закупки', 'Прибыль']

list_profit_lost = ['Месяц', 'Кол-во', 'Сумма закупки', 'Логистика', 'Хранение', 'Штрафы',
                    'Платная приемка', 'Удержания', 'К оплате', 'Прибыль']


df_pr_b = df_pr_b.loc[:, list_profit_lost]


sum_profit_lost = ['Кол-во', 'Сумма закупки', 'Логистика', 'Хранение', 'Штрафы',
                   'Платная приемка', 'Удержания', 'К оплате', 'Прибыль']

df_pr_b = df_pr_b.groupby('Месяц')[sum_profit_lost].sum()
df_pr_b = df_pr_b.stb.subtotal(
    sub_level=[1], grand_label='Всего за период').reset_index()
df_pr_b = df_pr_b.rename(columns={"index": "Месяц"})
df_pr_b = df_pr_b.replace({'Месяц': month_d})

if not os.path.isdir(path_to_data + "Аналитика"):
    os.mkdir(path_to_data + "Аналитика")

with pd.ExcelWriter(report_pivot, engine='xlsxwriter') as wb:
    df_f.to_excel(wb, sheet_name='sales ' + year_c, index=False,
                  float_format="%0.2f", freeze_panes=(1, 0))
    df_pr_b.to_excel(wb, sheet_name='profits ' + year_c,
                     index=False, float_format="%0.2f", freeze_panes=(1, 0))
    df_logistic.to_excel(wb, sheet_name='logistics ' + year_c,
                         index=False, float_format="%0.2f", freeze_panes=(1, 0))
    sheet = wb.sheets['sales ' + year_c]
    sheet.set_column('A:A', 20)
    sheet.set_column('B:B', 30)
    sheet.set_column('C:C', 35)
    sheet.set_column('D:N', 15)

    cell_format = wb.book.add_format()
    cell_format.set_bold()
    cell_format.set_font_color('black')
    # Установка стиля для строки 2 и высоты 40
    sheet.set_row(0, 30, cell_format)
    cell_format.set_font_size(14)
    cell_format.set_num_format(4)

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
    cell_format.set_num_format(4)
    for row_i in ind:
        # Установка стиля для строки 2 и высоты 40
        sheet.set_row(row_i + 1, 25, cell_format)

    cell_format.set_bg_color('#A3E2B9')

    cell_format = wb.book.add_format()
    cell_format.set_bold()
    cell_format.set_font_color('black')
    # Установка стиля для строки 2 и высоты 40
    sheet.set_row(len(df_f), 30, cell_format)
    cell_format.set_bg_color('#FDCFC5')
    cell_format.set_font_size(14)
    cell_format.set_num_format(4)
    # Форматирование второго и третьего листа
    sheet0 = wb.sheets['profits ' + year_c]
    sheet0.set_column('A:J', 20)

    sheet_log = wb.sheets['logistics ' + year_c]
    sheet_log.set_column('A:I', 20)

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
    sheet0.set_row(0, 40, header_format)

    sheet_log.set_row(0, 40, header_format)

    # Write the column headers with the defined format.
    for col_num, value in enumerate(df_pr_b.columns.values):
        sheet0.write(0, col_num, value, header_format)

    for col_num, value in enumerate(df_logistic.columns.values):
        sheet_log.write(0, col_num, value, header_format)

    cell_format = wb.book.add_format()
    cell_format.set_bold()
    cell_format.set_font_color('black')
    sheet0.set_row(len(df_pr_b), 30, cell_format)
