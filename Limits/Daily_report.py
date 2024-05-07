#!/usr/bin/env python
# coding: utf-8

# In[1]:
print('The "Limits" just started running')

import pandas as pd
import numpy as np
from datetime import date, timedelta
import seaborn as sns
import matplotlib.pyplot as plt
import matplotlib.axes as ax
import openpyxl
from openpyxl.drawing.image import Image
import win32com.client as win32
import os

olApp = win32.Dispatch('Outlook.Application')
olNS = olApp.GetNameSpace('MAPI')

import sys
sys.path.append("C:\\Users\\KlimovaAnnaA\\Documents\\MyFiles\\Projects\\OCP")
from Defs import merge_SalesUnits
from Defs import merge_Mapping
from Defs import Period
from Defs import new_list
from Defs import export_from_RISKCUSTOM
from Defs import add_in_currency_column
from Defs import concat_columns


# In[2]:


# manual_sending = False # True/False Заполните это поле, если хотите отправить отчет даже после критичных уведомлений

# Print_qualuty_check = True # Вынести QC в отдельный excel-файл? True/False
# Display_QC_mail = True # Показать письмо QC для отправки? True/False
# Send_QC_mail = True # Создать и отправить письмо для QC? True/False

Print_to_excel = True # Создать excel-файл с расчетами? True/False
Display_mail = True # Показать письмо для отправки? True/False
Send_mail = True # Создать и отправить письмо с расчетами и графиком? True/False

mail_to = 'TarakanovMIu@aimmngt.com' # Получатель письма


# In[3]:


query = f"""
SELECT *
FROM "RISKACCESS"."bankAccountsBalanceDaily"
WHERE "reportDate" = (SELECT MAX("reportDate") FROM "RISKACCESS"."bankAccountsBalanceDaily")
"""
bankAccountsBalanceDaily_data = export_from_RISKCUSTOM(query) # export data
BABD_data_work = bankAccountsBalanceDaily_data # BABD_data_work,
BABD_data_work['balanceUsd_mln'] = BABD_data_work.balanceUsd/10**6 # balanceUsd_mln
# balanceUsd_activmoney_market
accountStatus_list = ['active', 'mmarket']
BABD_data_work['balanceUsd_activmoney_market'] = 0
BABD_data_work.loc[BABD_data_work.accountStatus.isin(accountStatus_list), 'balanceUsd_activmoney_market'] = BABD_data_work.loc[BABD_data_work.accountStatus.isin(accountStatus_list), 'balanceUsd_mln']

query = f"""
SELECT CONCAT(CONCAT("bankId", "holding"), "limitType") as "con", "activeFrom", "bankId", "limitType", "limit", "holding"
FROM "RISKACCESS"."xxmrBankLimits" 
WHERE "limitType" IN ('Transit', 'Deposit')
"""
Limit_export_data = export_from_RISKCUSTOM(query) # export data
LE_data_work = Limit_export_data.sort_values(by='activeFrom').drop_duplicates(subset='con', keep='last')  # Unique strings with lasr value for activeFrom
# pivot by bankid and mln of limit
LE_data_work['H_B_concat'] = LE_data_work.holding + '_' + LE_data_work.bankId
LE_data_group = pd.pivot_table(data=LE_data_work, index=['H_B_concat', 'bankId', 'holding'], values='limit', aggfunc='sum').reset_index()
LE_data_group.limit = LE_data_group.limit/10**6
# merge BABD_data_work and LE_data_group
BABD_data_work = BABD_data_work.reset_index(drop=True)
BABD_data_work['H_B_concat'] = BABD_data_work.holding + '_' + BABD_data_work.bankId
BABD_data_work['Limit'] = BABD_data_work.merge(LE_data_group, how='left', left_on='H_B_concat', right_on='H_B_concat', validate='many_to_one').iloc[:,-1]
# Usage and Usage_activmoney_market
BABD_data_work['Usage_activmoney_market'] = 0
BABD_data_work['Usage'] = 0
BABD_data_limit_not_na = BABD_data_work[~BABD_data_work.Limit.isna()]
BABD_data_limit_not_na['Usage_activmoney_market'] = (BABD_data_limit_not_na.balanceUsd_activmoney_market/BABD_data_limit_not_na.Limit)*100
BABD_data_limit_not_na['Usage'] = (BABD_data_limit_not_na.balanceUsd_mln/BABD_data_limit_not_na.Limit)*100
BABD_data_work[~BABD_data_work.Limit.isna()] = BABD_data_limit_not_na
BABD_data_work.loc[(BABD_data_work.Usage == np.inf) | (BABD_data_work.Usage.isna()), 'Usage'] = 0
BABD_data_work.loc[(BABD_data_work.Usage_activmoney_market == np.inf) | (BABD_data_work.Usage_activmoney_market.isna()), 'Usage_activmoney_market'] = 0
BABD_data_work['Segment'] = merge_SalesUnits(df=BABD_data_work, merge_col='businessSegmentDetailed', col='buCode') # merge Segment
# Bank_name
query = """SELECT "bankId", "name", "country"
FROM "RISKACCESS"."xxmrBankLimitsBanks"
"""
data_xxmrBankLimitsBanks = export_from_RISKCUSTOM(query)
data_BLB = data_xxmrBankLimitsBanks.drop_duplicates(subset='bankId').dropna()
BABD_data_work['bank_name'] = BABD_data_work.merge(data_BLB, how='left', left_on='bankId', right_on='bankId', validate='many_to_one').iloc[:, -2]

BABD_data_work = BABD_data_work.rename(columns={'balanceUsd_activmoney_market':'Active', 'balanceUsd_mln':'Total', 'Usage_activmoney_market':'%_active', 'Usage':'%'})
# Creating Excel Writer Object from Pandas 
report_date = str(BABD_data_work.reportDate.max())[:10]
Output_file = report_date + '_limits_report.xlsx'
if Print_to_excel == True:
    writer = pd.ExcelWriter(Output_file, engine='openpyxl')  
    workbook=writer.book
# by holding
holding_list = BABD_data_work.holding.unique().tolist()
for holding in holding_list:
    holding_data = BABD_data_work[BABD_data_work.holding == holding].reset_index(drop=True) # holding data
    # table 1
    tabel_bankName = pd.pivot_table(data=holding_data, 
                    index='bank_name', 
                    values=['Limit', 'Active', 'Total', '%_active', '%'], 
                    aggfunc={'Limit':'mean', 'Active':'sum', 'Total':'sum', '%_active':'sum', '%':'sum'},
                    fill_value=0)\
                    .reset_index()\
                    .sort_values('Active', ascending=False)
    # table 2
    table_bankCountryCode = pd.pivot_table(data=holding_data, 
                    index='bankCountryCode', 
                    values=['Active', 'Total'], 
                    aggfunc={'Active':'sum', 'Total':'sum'})\
                    .reset_index()\
                    .sort_values('Active', ascending=False)
    # table 3
    table_Segment = pd.pivot_table(data=holding_data, 
                    index='Segment', 
                    values=['Active', 'Total'], 
                    aggfunc={'Active':'sum', 'Total':'sum'})\
                    .reset_index()\
                    .sort_values('Active', ascending=False)
    # to excel
    if Print_to_excel == True:
            tabel_bankName.to_excel(writer, sheet_name=holding, index=False) # to excel
            table_bankCountryCode.to_excel(writer, sheet_name=holding, startcol=8, index=False) # to excel
            table_Segment.to_excel(writer, sheet_name=holding, startcol=13, index=False) 
if Print_to_excel == True:
    writer.close()


# In[5]:


# Отправка письма
mailItem = olApp.CreateItem(0)
mailItem.BodyFormat = 3

mailItem.Subject = f'Bank limits for {report_date}' # mail head
# mail body
html_body = f"""<html><body><p>Dear colleagues,<br><br>
Please read the attached daily report on bank limits for {report_date}:<br><br>
Best regards,<br>
Maksim Tarakanov<br><br>
Whatsapp: +7 915 161 29 12<br>
Financial risk management</p></body></html>"""
mailItem.To = mail_to # mail to
mail_from = 'KlimovaAnnaA@aimmngt.com' # mail from
# mail attachment
mail_attachment = Output_file 


mailItem._oleobj_.Invoke(*(64209, 0, 8, 0, olNS.Accounts.Item(mail_from)))
mailItem.Attachments.Add(os.path.join(os.getcwd(), mail_attachment))
mailItem.HTMLBody = html_body
mailItem.Sensitivity  = 2

# mailItem.Save()
if Display_mail == True: ### DISPLAY
    mailItem.Display()
if Send_mail == True: ### SEND
    mailItem.Send()


# In[ ]:


manual_map = BABD_data_work.loc[BABD_data_work['Segment'] == 'External', ['holding', 'buCode']]
manual_map
print('The "Limits" was finished')


