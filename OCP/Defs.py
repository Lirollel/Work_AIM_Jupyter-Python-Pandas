

# import sys
# sys.path.append("C:\\Users\\KlimovaAnnaA\\Documents\\MyFiles\\Projects\\OCP")
# from Defs import merge_SalesUnits
# from Defs import merge_Mapping
# from Defs import Period
# from Defs import new_list
# from Defs import export_from_RISKCUSTOM
# from Defs import add_in_currency_column

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

    data = pd.read_csv('C:\\Users\\KlimovaAnnaA\\Documents\\MyFiles\\Projects\\OCP\\Методология\\Mapping.csv')
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
    connection.close()
    return data_export

# Создание столбца в нужной валюте
def add_in_currency_column(df, col_with_CCY, col_with_VAL, CCY_to, day_for_export : str = '29/02/24'):
    data_CCY_map = pd.read_csv('C:\\Users\\KlimovaAnnaA\\Documents\\MyFiles\\Projects\\OCP\\Методология\\CCY_mapping.csv')
    
    query = f"""
    select * from "RISKACCESS"."XXMR_MADAB_CONTENT"
    where
    "RISKACCESS"."XXMR_MADAB_CONTENT"."COMMODITY_ID" in (2354,322,311,312,314,315,318,2332,2334,2360,9321,9326,9331,9902,10014,7647,33051,9318,9886,2362,33447) and
    "RISKACCESS"."XXMR_MADAB_CONTENT"."PERIOD" = TO_DATE('{day_for_export}', 'DD/MM/YY')
    """
    # Смотри запрос

    data_export = export_from_RISKCUSTOM(query)[['COMMODITY_ID', 'VALUE1']]
    values_data = data_export.merge(data_CCY_map, how='left', left_on='COMMODITY_ID', right_on='id', validate='one_to_one')[['VALUE1','CCY_from', 'CCY_to']] 
    # Может возникнуть ошибка, если значений будет больше

    coef_dict = {}
    coef_dict[CCY_to] = 1
    for CCY_from in df[col_with_CCY].unique():
        if CCY_from != CCY_to:
            if CCY_from in values_data.CCY_from.tolist():
                coef_dict[CCY_from] = values_data.query('CCY_from == @CCY_from & CCY_to == @CCY_to').VALUE1.tolist()[0]
            if CCY_from in values_data.CCY_to.tolist():
                coef_dict[CCY_from] = 1/values_data.query('CCY_to == @CCY_from & CCY_from == @CCY_to').VALUE1.tolist()[0]
            else:
                continue

    df[f'Coef_to_{CCY_to}'] = df[col_with_CCY].replace(coef_dict)
    df[f'{col_with_VAL}_in_{CCY_to}'] = df[col_with_VAL] * df[f'Coef_to_{CCY_to}']

    return df

