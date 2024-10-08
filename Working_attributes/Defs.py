
import pandas as pd
import numpy as np
import oracledb
import sys
sys.path.append("C:\\Users\\KlimovaAnnaA\\Documents\\MyFiles\\Projects\\Working_attributes")
from BD_passwords import user_RISKCUSTOM, password_RISKCUSTOM, host_RISKCUSTOM
from BD_passwords import user_XXWH, password_XXWH, host_XXWH



# BS, Holding and country by id or sapid
def merge_SalesUnits(df, col, id_col: str ='id', merge_col: str = ['ocpSegment', 'holding', 'registryCountry', 'businessSegmentDetailed']):

    data = export_from_RISKCUSTOM(f"""select "{id_col}", "oebs12ShortCode", "sapId", "{merge_col}" from "RISKACCESS"."mdgSalesUnits" """)
    # data = data[[id_col,'oebs12ShortCode', merge_col]]
    id_cols_list = [f"{id_col}", "oebs12ShortCode", "sapId"]
    id_cols_dict = {}
    for id_colmn in id_cols_list:
        data_id_col = 0
        data_id_col = data.dropna(subset=id_colmn).drop_duplicates(subset=id_colmn)
        id_cols_dict[id_colmn] = data_id_col

    df = df.reset_index(drop=True)

    # merge_data = df.merge(data, how='left', left_on=col, right_on=id_col, validate='many_to_one').iloc[:, -1].fillna('External')
    merge_data = df
    for id_colmn in id_cols_list:
        merge_data[f'{id_colmn}_merge'] = df.merge(id_cols_dict[id_colmn], how='left', left_on=col, right_on=id_colmn, validate='many_to_one').iloc[:, -1].fillna('External')

    merge_data['last_merge'] = np.nan
    for id_colmn in id_cols_list:
        merge_data.loc[merge_data[f'{id_colmn}_merge'] != 'External', 'last_merge'] = merge_data.loc[merge_data[f'{id_colmn}_merge'] != 'External', f'{id_colmn}_merge']
    merge_data['last_merge'] = merge_data['last_merge'].fillna('External')
    merge_data = merge_data['last_merge']

    return merge_data

# id by CompName
def merge_Mapping(df, col):
    
    data = pd.read_csv('C:\\Users\\KlimovaAnnaA\\Documents\\MyFiles\\Projects\\Working_attributes\\Mapping.csv')
    data = data.dropna(subset='CompName').drop_duplicates(subset='CompName')

    df = df.reset_index(drop=True)

    merge_data = df.merge(data, how='left', left_on=col, right_on='CompName', validate='many_to_one').iloc[:, -1].fillna('External')

    return merge_data

# Counting the Period by df, day for counting and column with date
def Period(df, day_for_count, col_with_date, day_is_col=False):

    df = df.reset_index(drop=True)

    if day_is_col==False:
        day = pd.to_datetime(day_for_count)
        df[day_for_count] = day
    while True:
        if np.issubdtype(df[col_with_date].dtype, np.datetime64) and np.issubdtype(df[day_for_count].dtype, np.datetime64):

            df['Days'] = df[col_with_date] - df[day_for_count]
            df['Days'] = df['Days'].dt.days
            df['Days'] = df['Days'].fillna(0)

            df['Period'] = '2Y+'
            df.loc[pd.to_numeric(df['Days']) < 730, 'Period'] = '1Y-2Y'
            df.loc[pd.to_numeric(df['Days']) < 365, 'Period'] = '6M-1Y'
            df.loc[pd.to_numeric(df['Days']) < 182, 'Period'] = '3M-6M'
            df.loc[pd.to_numeric(df['Days']) < 91, 'Period'] = '1M-3M'
            df.loc[pd.to_numeric(df['Days']) < 30, 'Period'] = '<1M'
            break

        else:
            df[col_with_date] = pd.to_datetime(df[col_with_date], errors='coerce')
            df[day_for_count] = pd.to_datetime(df[day_for_count], errors='coerce')
            continue
    if day_is_col==False:
        df = df.drop(columns=day_for_count)
    return df

# Запись данных на новый лист существующего файла
def new_list(df, output_file : str, sheet_name : str, index : bool = False, startrow=0, startcol=0):

    with pd.ExcelWriter(output_file, engine='openpyxl', mode='a') as writer:
        df.to_excel(writer,sheet_name=sheet_name, index=index, startrow=startrow, startcol=startcol)

# Подключение к БД и выгрузка данных по запросу
def export_from_RISKCUSTOM(query):
    oracledb.init_oracle_client('C:\\Users\\KlimovaAnnaA\\Documents\\MyFiles\\Oracle\\instantclient_21_13')
    connection = oracledb.connect(user=user_RISKCUSTOM, password=password_RISKCUSTOM, host=host_RISKCUSTOM, port=1521, service_name='riskdev.moscow.eurochem.ru', disable_oob= True)
    data_export = pd.read_sql(query, con=connection)
    connection.close()
    return data_export

# Конкатенация столбцов
def concat_columns(df: pd.DataFrame, columns: list, sign: str = '_'):
    df['concat_columns'] = df[columns].astype(str).apply(lambda row: sign.join(row.values.astype(str)), axis=1)
    return df

# Создание столбца в нужной валюте
CCY_tech_dict = {'RUC':'CHF','RUE':'EUR','RUE2':'EUR','RUE3':'EUR','RUE5':'EUR','RUG':'GBP','RUJ':'JPY','RUK':'KZT','RUM':'MDL','RUU':'USD','RUU1':'USD','RUU1R':'USD','RUU2':'USD','RUU25':'USD','RUU26':'USD','RUU3':'USD','RUU5':'USD','RUY':'CNY','RUK5':'KZT','RUY3':'CNY','RUY5':'CNY', 'RUI':'INR'}

def add_in_currency_column(df: pd.DataFrame, CCY_to: str, col_with_CCY: str, date_is_column: bool, col_with_VAL: str, DATE: str = 'YYYY-MM-DD'):
    df_columns_list = df.columns.tolist()
    df['CCY_to'] = CCY_to
    df[col_with_CCY] = df[col_with_CCY].replace(CCY_tech_dict)
    
    if date_is_column == False:
        
        df['date'] = DATE
        df['date_CCY_from'] = df['date'].astype(str)+'_'+df[col_with_CCY].astype(str)
        df['date_CCY_to'] = df['date'].astype(str)+'_'+df.CCY_to.astype(str)

        Date_SQL_str = "TO_DATE('" + str(DATE) + "', 'YYYY-MM-DD')"

    if date_is_column == True:
        
        df[f'{DATE}_str'] = df[DATE].astype(str).str[:10]
        df['date_CCY_from'] = df[f'{DATE}_str'].astype(str)+'_'+df[col_with_CCY].astype(str)
        df['date_CCY_to'] = df[f'{DATE}_str'].astype(str)+'_'+df.CCY_to.astype(str)

        # Создание списка уникальных дат
        Date_unique_list = df[f'{DATE}_str'].unique().tolist()
        Date_SQL_list = ["TO_DATE('" + str(x) + "', 'YYYY-MM-DD')" for x in Date_unique_list]
        Date_SQL_str = str(Date_SQL_list).replace('"','')[1:-1]
    
    # Создание списка уникальных валют
    CCY_unique_list = df[col_with_CCY].unique().tolist()
    CCY_variations_list = [f"{CCY_to}/" + str(x) for x in CCY_unique_list] + [str(x) + f"/{CCY_to}" for x in CCY_unique_list]
    data_CCY_map = pd.read_csv('C:\\Users\\KlimovaAnnaA\\Documents\\MyFiles\\Projects\\Working_attributes\\CCY_mapping.csv')
    CCY_id_unique_list = data_CCY_map[data_CCY_map.CCY.isin(CCY_variations_list)].id.unique().tolist()
    CCY_id_unique_str = str(CCY_id_unique_list)[1:-1]

    # выгрузка из БД по списку уникальных дат и значений валют
    query = f"""select * from "RISKACCESS"."XXMR_MADAB_CONTENT"
    where "RISKACCESS"."XXMR_MADAB_CONTENT"."COMMODITY_ID" in ({CCY_id_unique_str})
    and "RISKACCESS"."XXMR_MADAB_CONTENT"."PERIOD" in ({Date_SQL_str})"""
    data_export = export_from_RISKCUSTOM(query)[['COMMODITY_ID', 'PERIOD', 'VALUE1']]
    values_data = data_export.merge(data_CCY_map, how='left', left_on='COMMODITY_ID', right_on='id')[['PERIOD', 'VALUE1','CCY_from', 'CCY_to']] 
    values_data['PERIOD_str'] = values_data.PERIOD.astype(str).str[:10]
    values_data = concat_columns(values_data, ['PERIOD_str', 'CCY_to']).rename(columns={'concat_columns': 'date_CCY_to'})
    values_data = concat_columns(values_data, ['PERIOD_str', 'CCY_from']).rename(columns={'concat_columns': 'date_CCY_from'})

    # Создание словаря значений валют
    coef_dict = {}
    for i in df.index.tolist():

        date_CCY_from = df.loc[i, 'date_CCY_from']
        date_CCY_to = df.loc[i, 'date_CCY_to']

        if date_CCY_from != date_CCY_to:

            if date_CCY_from in values_data.date_CCY_from.tolist():
                coef_dict[date_CCY_from] = values_data.query('date_CCY_from == @date_CCY_from & date_CCY_to == @date_CCY_to').VALUE1.tolist()[0]
            if date_CCY_from in values_data.date_CCY_to.tolist():
                coef_dict[date_CCY_from] = 1/values_data.query('date_CCY_to == @date_CCY_from & date_CCY_from == @date_CCY_to').VALUE1.tolist()[0]
        
        else:
            coef_dict[date_CCY_from] = 1

    df[f'Coef_to_{CCY_to}'] = df.date_CCY_from.replace(coef_dict).fillna(0)
    df[f'Coef_to_{CCY_to}'] = pd.to_numeric(df[f'Coef_to_{CCY_to}'], errors='coerce')

    df[f'{col_with_VAL}_in_{CCY_to}'] = df[col_with_VAL] * df[f'Coef_to_{CCY_to}']

    df_columns_list.append(f'Coef_to_{CCY_to}')
    df_columns_list.append(f'{col_with_VAL}_in_{CCY_to}')
    df = df[df_columns_list]

    return df

def export_from_WHWEEK(query):
    oracledb.init_oracle_client('C:\\Users\\KlimovaAnnaA\\Documents\\MyFiles\\Oracle\\instantclient_21_13')
    connection = oracledb.connect(user=user_XXWH, password=password_XXWH, host=host_XXWH, port=1521, service_name='whweek.moscow.eurochem.ru', disable_oob= True)
    data_export = pd.read_sql(query, con=connection)
    connection.close()
    return data_export

def is_approximately_equal_for_cols(df: pd.DataFrame, col_1:str ,col_2:str , tolerance=0.4, print_diff=False, name_col_with_diff=''):
    # tolerance - persentage of differences in max value

    name_diff_col = f'Diff_{name_col_with_diff}_more_{tolerance}'
    df[name_diff_col] = np.nan
    #compaire values
    max_value = df[[col_1, col_2]].abs().max(axis=1)
    diff = abs(df[col_1].abs()-df[col_2].abs())
    df[name_diff_col] = (diff/max_value)>=tolerance
    # if cols equal
    df.loc[df[col_1]==df[col_2], name_diff_col] = False
    # print diff
    if print_diff==True:
        df['Diff'] = diff
    
    return df


holding_list = ['EUROCHEM', 'SUEK']