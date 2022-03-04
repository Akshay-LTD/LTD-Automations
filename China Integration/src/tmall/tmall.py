import pandas as pd
import datetime as dt
import re
import glob
from helper.helper import *

def process_tmall(df, df_internalID):

    # remove empty rows in barcode
    df.dropna(subset=['外部系统编号'], inplace=True, axis=0)

    # exclude this for now
    df = df[df['外部系统编号'] != "1001-001"]
    df.reset_index(drop=True, inplace=True)

    df_upload = pd.DataFrame()

    # PO and Memo
    df_upload['PO'] = df["主订单编号"].apply(lambda x: re.sub(r"=\"(\d+)\"", r"\g<1>", x)).astype('str')
    df_upload['Memo'] = df["主订单编号"].apply(lambda x: re.sub(r"=\"(\d+)\"", r"\g<1>", x)).astype('str')

    # map item id to barcode
    df_upload['UPC Code'] = df['外部系统编号']
    df_upload['UPC Code'] = df_upload['UPC Code'].astype('str')
    df_upload = df_upload.merge(df_internalID, how='left', on='UPC Code')
    df_upload = df_upload.rename(columns={'Internal ID': 'Item ID'})
    df_upload.drop('UPC Code', axis=1, inplace=True)

    # # remove rows not mapped
    # df_upload.dropna(subset=['Item ID'], inplace=True, axis=0)
    

    # quantity
    df_upload['Qty'] = df['购买数量']

    # gross amount
    df_upload['Gross Amount'] = round(df['买家应付货款'], 2)

    # rate
    df_upload['Rate'] = df_upload['Gross Amount'] / df_upload['Qty']
    df_upload['Rate'] = round(df_upload['Rate'] / 1.13, 2)

    # date
    df_upload['Date'] = df['订单创建时间']
    df_upload['Date'] = pd.to_datetime(df_upload['Date'])
    df_upload['Date'] = df_upload['Date'].dt.strftime('%d/%m/%Y')

    # get strings for external id
    df_upload['JY'] = "JY"
    df_upload['substr'] = df_upload['PO'].astype('str').apply(lambda x: x[4:11])
    df_upload['excelDate'] = pd.to_datetime(df_upload['Date'], format="%d/%m/%Y").apply(lambda x: convertDateToExcel(x.year, x.month, x.day))

    # external id
    df_upload['External ID'] = df_upload['JY'] + df_upload['substr'] + df_upload['excelDate']
    df_upload.drop(['JY','substr','excelDate'], axis=1, inplace=True)

    # constant values
    df_upload['Department'] = 22
    df_upload['Location'] = 76
    df_upload['Currency'] = "CNY"
    df_upload['Customer'] = 3760095
    df_upload['Price Level'] = 'Custom'
    df_upload['VAT'] = ""
    df_upload['Payment Method'] = ""
    df_upload['Account'] = 2069
    df_upload['Cus Checking'] = 19
    df_upload['New Cus Checking'] = "Tmall"

    # rearrange
    df_upload = df_upload[["Item ID", "Qty", "Rate", "External ID", "Date", "Department", "Location", "Currency", "Customer", "Gross Amount", 
                            "Price Level", "VAT", "PO", "Payment Method", "Memo", "Account", "Cus Checking", "New Cus Checking"]]

    # sort based on external id
    df_upload.sort_values(by=['External ID'], inplace=True)
    df_upload.reset_index(drop=True, inplace=True)

    return df_upload

def generate_tmall():
    df_internalID = get_internalID()
    for file in glob.glob("Input/Tmall/ExportOrderDetailList*.csv"):
        
        # read file
        df = pd.read_csv(file, encoding='gbk')
        name = re.findall(r"ExportOrderDetailList[\d]+", file)

        # process data
        df_upload = process_tmall(df, df_internalID)

        # get suffix for file name
        suffix = get_date(df_upload)

        # write to csv
        df_upload.to_csv(f'Output/Tmall/CSV/Tmall SO Upload {suffix}.csv', index=False)
        
        # write to excel
        writer = pd.ExcelWriter(f'Output/Tmall/Excel/Tmall SO Upload {suffix}.xlsx', 
                                engine='xlsxwriter', 
                                engine_kwargs={'options': {'string_to_numbers': False}})
        df_upload.to_excel(writer, index=False, sheet_name='Tmall')
        writer.save()

        print(f"[GENERATED] Output/Tmall/CSV/Tmall SO Upload {suffix}.csv")
        print(f"[GENERATED] Output/Tmall/Excel/Tmall SO Upload {suffix}.xlsx")