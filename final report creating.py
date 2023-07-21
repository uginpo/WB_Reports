import pandas as pd
import numpy as np
import openpyxl
import xlsxwriter
import os.path
import sidetable
import locale
import sys

clients = ['ООО Исток', 'ИП Искандаров', 'ИП Бурдукова', 'ИП Тараканов']
years_list = ['2021', '2022', '2023']

client = clients[3]
year_c = years_list[2]

path_to_data = f'/Users/uginpo/OneDrive - gxog/Отчеты ВБ/{client}/Сведенные/{year_c}/'

name__price = f'Закупка_{client}_{year_c}.xlsx'  # файл закупочных цен для расчета прибыли
name__list_otchet = f'Отчеты_{client}_{year_c}.xlsx'  # список отчетов с доп. информацией - хранение и пр.

name__analitic = f'/Аналитика/Аналитика_{client}_{year_c}.xlsx'
report__pivot = f'/Аналитика/Сводный отчет_{client}_{year_c}.xlsx'

name_price = path_to_data + name__price
name_list_otchet = path_to_data + name__list_otchet
name_analitic = path_to_data + name__analitic
report_pivot = path_to_data + report__pivot

if os.path.isfile(name_analitic):
    df_analitic = pd.read_excel(name_analitic, sheet_name=[0, 1], header=0)
else:
    df_analitic = False

# Настройка параметров отчета

month_r = 'all'
date_start = ''
date_final = ''
period_r = 'month'  # 'report'
goods_r = 'all'

month_d = {'01': 'Январь', '02': 'Февраль', '03': 'Март', '04': 'Апрель', '05': 'Май', '06': 'Июнь',
           '07': 'Июль', '08': 'Август', '09': 'Сентябрь', '10': 'Октябрь', '11': 'Ноябрь', '12': 'Декабрь'}

sum_list = ['Кол-во', 'Реализация ВБ', 'Вознаграждение ВБ', 'Логистика',
            'Очищенная выручка', 'Сумма закупки', 'Сумма брака', 'Доход']
sum_list_pr = ['Кол-во', 'Сумма закупки', 'Сумма брака', 'Логистика', 'Хранение', 'Удержания', 'Очищенная выручка',
               'Доход']

df_f = df_analitic[0].copy(deep=True)
df_pr_lost = df_analitic[1].copy(deep=True)

# Формирование отчета по продажам

list_df = [df_f, df_pr_lost]
for df_ind in list_df:
    df_ind['Дата начала'] = pd.to_datetime(df_ind['Дата начала'], format='%d/%m/%Y')
    df_ind["Месяц"] = df_ind['Дата начала'].dt.month.astype(str).str.zfill(2)

year_f = df_f['Дата начала'].dt.year.astype(str)[0]

df_f = df_f.groupby(['Месяц', 'Артикул поставщика', 'Название'])[sum_list].sum()
df_pr_lost = df_pr_lost.groupby(['Месяц'])[sum_list_pr].sum()

df_b = df_f.stb.subtotal(sub_level=[1], grand_label='Всего за период', sub_label='Итого за месяц').reset_index()
df_b = df_b.rename(columns={"level_0": "Месяц", "level_1": "Артикул поставщика", "level_2": "Название"})

df_pr_b = df_pr_lost.stb.subtotal(sub_level=[1], grand_label='Всего за период').reset_index()
df_pr_b = df_pr_b.rename(columns={"index": "Месяц"})

df_b = df_b.replace({'Месяц': month_d})
df_pr_b = df_pr_b.replace({'Месяц': month_d})

itog_str = f'- Итого за месяц'
df1 = df_b[df_b['Артикул поставщика'].str.contains(itog_str)]
ind = df1.index
df_b['Артикул поставщика'].where(~(df_b['Артикул поставщика'].str.contains(itog_str)),
                                 other='Итого за месяц', inplace=True)

with pd.ExcelWriter(report_pivot, engine='xlsxwriter') as wb:
    df_b.to_excel(wb, sheet_name='sales ' + year_f, index=False, float_format="%0.2f", freeze_panes=(1, 0))
    df_pr_b.to_excel(wb, sheet_name='profits ' + year_f, index=False, float_format="%0.2f", freeze_panes=(1, 0))
    sheet = wb.sheets['sales ' + year_f]
    sheet.set_column('A:A', 20)
    sheet.set_column('B:B', 30)
    sheet.set_column('C:C', 35)
    sheet.set_column('D:L', 15)

    cell_format = wb.book.add_format()
    cell_format.set_bold()
    cell_format.set_font_color('black')
    sheet.set_row(0, 30, cell_format)  # Установка стиля для строки 2 и высоты 40
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
    for col_num, value in enumerate(df_b.columns.values):
        sheet.write(0, col_num, value, header_format)

    cell_format = wb.book.add_format()
    cell_format.set_bold()
    cell_format.set_font_color('black')
    cell_format.set_num_format(4)
    for row_i in ind:
        sheet.set_row(row_i + 1, 25, cell_format)  # Установка стиля для строки 2 и высоты 40

    cell_format.set_bg_color('#A3E2B9')

    cell_format = wb.book.add_format()
    cell_format.set_bold()
    cell_format.set_font_color('black')
    sheet.set_row(len(df_b), 30, cell_format)  # Установка стиля для строки 2 и высоты 40
    cell_format.set_bg_color('#FDCFC5')
    cell_format.set_font_size(14)
    cell_format.set_num_format(4)

    # Форматирование второго листа
    sheet0 = wb.sheets['profits ' + year_f]
    sheet0.set_column('A:L', 20)

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

    # Write the column headers with the defined format.
    for col_num, value in enumerate(df_pr_b.columns.values):
        sheet0.write(0, col_num, value, header_format)

    cell_format = wb.book.add_format()
    cell_format.set_bold()
    cell_format.set_font_color('black')
    sheet0.set_row(len(df_pr_b), 30, cell_format)
    cell_format.set_bg_color('#FDCFC5')
    cell_format.set_font_size(14)
    cell_format.set_num_format(4)
