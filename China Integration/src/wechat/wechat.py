import pandas as pd
import datetime as dt
import glob
import re
from helper.helper import *

def process_wechat(df, df_internalID):
    
    # remove empty rows in barcode
    df.dropna(subset=['商家编码'], inplace=True, axis=0)

    df_upload = pd.DataFrame()

    # PO and Memo
    df_upload['PO'] = df["订单编号"].astype('str')
    df_upload['Memo'] = df["订单编号"].astype('str')

    # map item id to barcode
    df_upload['UPC Code'] = df['商家编码'].astype('str')
    df_upload = df_upload.merge(df_internalID, how='left', on='UPC Code')
    df_upload = df_upload.rename(columns={'Internal ID': 'Item ID'})

    # exclude this for now
    df_upload[df_upload['UPC Code'] != "1001-001"]
    
    df_upload.drop('UPC Code', axis=1, inplace=True)

    # quantity
    df_upload['Qty'] = df['商品数量']

    # gross amount
    df_upload['Gross Amount'] = round(df['商品单价']*df['商品数量'], 2)

    # rate
    df_upload['Rate'] = df_upload['Gross Amount'] / 1.13
    df_upload['Rate'] = round(df_upload['Rate'] / df_upload['Qty'], 2)

    # date
    df_upload['Date'] = df['下单时间']
    df_upload['Date'] = pd.to_datetime(df_upload['Date'])
    df_upload['Date'] = df_upload['Date'].dt.strftime('%d/%m/%Y')

    # get strings for external id
    df_upload['JY'] = "JY"
    df_upload['substr'] = df_upload['PO'].astype('str').apply(lambda x: x[3:8])
    df_upload['excelDate'] = pd.to_datetime(df_upload['Date'], format="%d/%m/%Y").apply(lambda x: convertDateToExcel(x.year, x.month, x.day))

    # external id
    df_upload['External ID'] = df_upload['JY'] + df_upload['substr'] + df_upload['excelDate']
    df_upload.drop(['JY','substr','excelDate'], axis=1, inplace=True)

    # constant values
    df_upload['Department'] = 22
    df_upload['Location'] = 76
    df_upload['Currency'] = "CNY"
    df_upload['Customer'] = 4191514
    df_upload['Price Level'] = 'Custom'
    df_upload['VAT'] = ""
    df_upload['Payment Method'] = ""
    df_upload['Account'] = 2085
    df_upload['Cus Checking'] = 15
    df_upload['New Cus Checking'] = "Wechat"

    df_upload = df_upload[["Item ID", "Qty", "Rate", "External ID", "Date", "Department", "Location", "Currency", "Customer", "Gross Amount", 
                            "Price Level", "VAT", "PO", "Payment Method", "Memo", "Account", "Cus Checking", "New Cus Checking"]]

    # sort based on external id
    df_upload.sort_values(by=['External ID'], inplace=True)
    df_upload.reset_index(drop=True, inplace=True)

    return df_upload

def generate_wechat():
    df_internalID = get_internalID()
    for file in glob.glob("Input/Wechat/ORDER_MANAGE_EXPORT*.csv"):
        
        # read file
        df = pd.read_csv(file)
        
        name = re.findall(r"ORDER_MANAGE_EXPORT[_\d]+", file)

        # process data
        df_upload = process_wechat(df, df_internalID)

        # write to csv
        df_upload.to_csv(f'Output/Wechat/CSV/{name[0]}.csv', index=False)

        # write to excel
        writer = pd.ExcelWriter(f'Output/Wechat/Excel/{name[0]}.xlsx', 
                                engine='xlsxwriter', 
                                engine_kwargs={'options': {'string_to_numbers': False}})
        df_upload.to_excel(writer, index=False, sheet_name='Wechat')
        writer.save()

        print(f"[GENERATED] Output/Wechat/CSV/{name[0]}.csv")
        print(f"[GENERATED] Output/Wechat/Excel/{name[0]}.xlsx")