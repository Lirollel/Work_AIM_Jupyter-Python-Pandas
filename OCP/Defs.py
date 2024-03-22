

# import sys
# sys.path.append("C:\\Users\\KlimovaAnnaA\\Documents\\MyFiles\\Projects\\OCP")
# from Defs import merge_SalesUnits
# from Defs import merge_Mapping
# from Defs import Period
# from Defs import new_list
# from Defs import export_from_RISKCUSTOM

import pandas as pd
import numpy as np
import oracledb

# BS, Holding and country by id or sapid
def merge_SalesUnits(df, col, id_col: str ='id', merge_col: str = ['ocpSegment', 'holding', 'registryCountry']):

    data = pd.read_excel('C:\\Users\\KlimovaAnnaA\\Documents\\MyFiles\\Projects\\OCP\\salesUnits.xlsx', sheet_name='salesUnits')
    data = data[[id_col, merge_col]]
    data = data.dropna(subset=id_col).drop_duplicates(subset=id_col)

    df = df.reset_index(drop=True)

    merge_data = df.merge(data, how='left', left_on=col, right_on=id_col, validate='many_to_one').iloc[:, -1].fillna('External')

    return merge_data

# id by CompName
def merge_Mapping(df, col):

    data = pd.read_excel('C:\\Users\\KlimovaAnnaA\\Documents\\MyFiles\\Projects\\OCP\\Методология\\Mapping.xlsx', sheet_name='mapping')
    data = data.dropna(subset='CompName').drop_duplicates(subset='CompName')

    df = df.reset_index(drop=True)

    merge_data = df.merge(data, how='left', left_on=col, right_on='CompName', validate='many_to_one').iloc[:, -1].fillna('External')

    return merge_data

# Counting the Period by df, day for counting and column with date
def Period(df, day_for_count, col_with_date):

    day = pd.to_datetime(day_for_count)
    df = df.reset_index(drop=True)

    while True:
        if np.issubdtype(df[col_with_date].dtype, np.datetime64):

            df['Days'] = df[col_with_date] - day
            df['Days'] = df['Days'].dt.days
            df.loc[df[col_with_date].isna() ,'Days'] = 0

            df['Period'] = '2Y+'
            df.loc[pd.to_numeric(df['Days']) < 730, 'Period'] = '1Y-2Y'
            df.loc[pd.to_numeric(df['Days']) < 365, 'Period'] = '6M-1Y'
            df.loc[pd.to_numeric(df['Days']) < 182, 'Period'] = '3M-6M'
            df.loc[pd.to_numeric(df['Days']) < 91, 'Period'] = '1M-3M'
            df.loc[pd.to_numeric(df['Days']) < 30, 'Period'] = '<1M'
            break

        else:
            df[col_with_date] = pd.to_datetime(df[col_with_date], errors='coerce')
            continue
    
    return df

# Запись данных на новый лист существующего файла
def new_list(df, output_file : str, sheet_name : str, index : bool = False):

    with pd.ExcelWriter(output_file, engine='openpyxl', mode='a') as writer:
        df.to_excel(writer,sheet_name=sheet_name, index=index)

# Подключение к БД и выгрузка данных по запросу
def export_from_RISKCUSTOM(query):
    oracledb.init_oracle_client('C:\\Users\\KlimovaAnnaA\\Documents\\MyFiles\\Oracle\\instantclient_21_13')
    connection = oracledb.connect(user="RISKCUSTOM", password="xxRiscRccess174!", host="exatest2-scan.moscow.eurochem.ru", port=1521, service_name='riskdev.moscow.eurochem.ru', disable_oob= True)
    data_export = pd.read_sql(query, con=connection)
    return data_export

