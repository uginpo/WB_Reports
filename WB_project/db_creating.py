import os.path
import pandas as pd
from variables import *

weekly_report = f'{DATABASE}weekly_reports_wb.parquet'
commission_wb = f'{DATABASE}commissions_wb.parquet'
price_list_wb = f'{DATABASE}price_list_wb.parquet'
summary_report = f'{DATABASE}summary_report_wb.parquet'


init_commission = f'{DATABASE_REPORTS}Комиссия.xlsx'


def create_commision(initial, destination):
    pd.read_excel(initial, header=0).to_parquet(destination, engine="pyarrow")


def create_price_list(destination):
    flag = True
    for client in CLIENTS:
        name_price = f'{DATABASE_REPORTS}{client}/Закупка_{client}.xlsx'
        df_info = get_from_initial(name_price, client)
        if flag:
            df_price = df_info.copy(deep=True)
            flag = False
        else:
            df_price = pd.concat([df_price, df_info],
                                 axis=0, ignore_index=True)
    df_price['Артикул поставщика'] = df_price['Артикул поставщика'].astype(str)
    df_price['Закупочная цена'] = df_price['Закупочная цена'].astype(float)

    df_price.to_parquet(destination, engine="pyarrow")


def create_sum_report(destination):
    flag = True
    for client in CLIENTS:
        sum_report = f'{DATABASE_REPORTS}{client}/Еженедельные отчёты_{client}.xlsx'
        df_info = get_from_initial(sum_report, client)
        if flag:
            df_sum = df_info.copy(deep=True)
            flag = False
        else:
            df_sum = pd.concat([df_sum, df_info],
                               axis=0, ignore_index=True)

    date_new_report = pd.to_datetime(DATE_CHANGE_REPORT, dayfirst=True).date()
    df_sum['Дата начала'] = pd.to_datetime(df_sum['Дата начала'])
    df_sum.rename(columns={'№ отчета': 'Номер отчета'}, inplace=True)
    df_sum['Номер отчета'] = df_sum['Номер отчета'].astype(str)
    df_sum = df_sum.loc[df_sum['Дата начала'].dt.date >= date_new_report, :]
    # Не учитываются "Прочие удержания"
    df_sum['Итого к оплате'] = (df_sum['К перечислению за товар'] - df_sum['Стоимость логистики'] -
                                df_sum['Повышенная логистика согласно коэффициенту по обмерам'] -
                                df_sum['Общая сумма штрафов'] - df_sum['Доплаты'] -
                                df_sum['Стоимость хранения'] - df_sum['Стоимость платной приемки'])

    df_sum.to_parquet(destination, engine="pyarrow")
    return (df_sum)


def create_weekly_report(destination, df_files):
    flag = True
    for client in CLIENTS:
        directory = f'{DATABASE_REPORTS}{client}/Сведенные/'

        for root, dirs, files in os.walk(directory):
            if len(files) != len(df_files.loc[df_files['Клиент'] == client, :]):
                raise ('Количество отчетов не совпадает')

            for name in files:
                if name.count('.xlsx') == 0:
                    continue
                df_w = get_from_initial(os.path.join(root, name), client)
                df_w['Номер отчета'] = name.replace('.xlsx', '')
                df_w['Размер'] = df_w['Размер'].astype(str)
                df_w['Баркод'] = df_w['Баркод'].astype(str)
                df_w['Артикул поставщика'] = df_w['Артикул поставщика'].astype(
                    str)

                if flag:
                    weekly_reports = df_w.copy(deep=True)
                    flag = False

                else:
                    weekly_reports = pd.concat(
                        [weekly_reports, df_w], axis=0, ignore_index=True)

    weekly_reports.to_parquet(destination, engine="pyarrow")


def get_from_initial(init_file, client):
    df_f = pd.read_excel(init_file, header=0)
    df_f['Клиент'] = client
    return df_f


if __name__ == '__main__':
    # Creating DB (commissions WB)
    create_commision(init_commission, commission_wb)
    create_price_list(price_list_wb)
    df_files = create_sum_report(summary_report)
    create_weekly_report(weekly_report, df_files)
    print("Все прошло успешно")
