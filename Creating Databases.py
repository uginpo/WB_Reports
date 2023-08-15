import os.path
import pandas as pd


def open_other(rep, client_f, rep_dig=''):
    df_f = pd.read_excel(rep, header=0)
    if len(rep_dig) != 0:
        df_f['Номер отчета'] = rep_dig

    df_f['Клиент'] = client_f
    return df_f


clients = ['ООО Исток', 'ИП Искандаров', 'ИП Бурдукова', 'ИП Тараканов']

path_common_data = f'/Users/uginpo/OneDrive - gxog/Отчеты ВБ/'
path_to_database = f'/Users/uginpo/OneDrive - gxog/Отчеты ВБ/База данных/'

weekly_report = f'{path_to_database}weekly_reports_wb.pkl'
commission_wb = f'{path_to_database}commissions_wb.pkl'
price_list_wb = f'{path_to_database}price_list_wb.pkl'
summary_report = f'{path_to_database}summary_report_wb.pkl'
name_commission = f'{path_common_data}Комиссия.xlsx'

date_new_report = pd.to_datetime('13/09/2021',
                                 dayfirst=True).date()  # С этой даты изменились файлы еженедельных отчетов
#  Создание dataframe комиссии WB

pd.read_excel(name_commission, header=0).to_pickle(commission_wb)

#  Создание dataframe еженедельных отчетов WB

# Проверка существования БД еженедельных отчетов
weekly_reports = df_sum = df_price = list()

if os.path.isfile(weekly_report):
    weekly_reports = pd.read_pickle(weekly_report)
    flag_weekly_reports = True
else:
    flag_weekly_reports = False

# Функции открытия недельного отчета о продажах и добавление в него колонок

flag_price_summary = False

for client in clients:
    path_data = f'/Users/uginpo/OneDrive - gxog/Отчеты ВБ/{client}/'
    name_price = f'{path_data}Закупка_{client}.xlsx'
    name_summary = f'{path_data}Еженедельные отчёты_{client}.xlsx'

    df_price_client = open_other(name_price, client)
    df_sum_client = open_other(name_summary, client)

    if flag_price_summary:
        df_price = pd.concat([df_price, df_price_client],
                             axis=0, ignore_index=True)
        df_sum = pd.concat([df_sum, df_sum_client], axis=0, ignore_index=True)

    else:
        df_price = df_price_client.copy(deep=True)
        df_sum = df_sum_client.copy(deep=True)
        flag_price_summary = True

    path_to_data = f'/Users/uginpo/OneDrive - gxog/Отчеты ВБ/{client}/Сведенные/'

    df_sum['Дата начала'] = pd.to_datetime(df_sum['Дата начала'])
    old_report_list = df_sum.loc[df_sum['Дата начала'].dt.date <
                                 date_new_report, '№ отчета'].astype(str) + '.xlsx'

    num_reports = pd.Series(os.listdir(path_to_data))
    num_reports = num_reports[num_reports.str.match("^[0-9]")]
    new_report_list = pd.Series(list(set(num_reports) - set(old_report_list)))

    for report in new_report_list:
        name_list_report = f'{path_to_data}{report}'
        report_dig = report.split('.')[0]

        if flag_weekly_reports:
            if weekly_reports['Номер отчета'].isin([report_dig]).sum() == 0:
                df = open_other(name_list_report, client, report_dig)
                weekly_reports = pd.concat(
                    [weekly_reports, df], axis=0, ignore_index=True)
        else:
            df = open_other(name_list_report, client, report_dig)
            weekly_reports = df.copy(deep=True)
            flag_weekly_reports = True

df_sum.rename(columns={'№ отчета': 'Номер отчета'}, inplace=True)
df_sum['Номер отчета'] = df_sum['Номер отчета'].astype(str)
ds_sum_old = df_sum.loc[df_sum['Дата начала'].dt.date < date_new_report, :]
ds_sum = df_sum.loc[df_sum['Дата начала'].dt.date >= date_new_report, :]

df_price.to_pickle(price_list_wb)
df_sum.to_pickle(summary_report)
weekly_reports.to_pickle(weekly_report)
