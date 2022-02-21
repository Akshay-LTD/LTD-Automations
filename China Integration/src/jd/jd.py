import pandas as pd
import datetime as dt
import re
import glob
from helper.helper import *

def process_jd(df, df_internalID):

    # remove empty rows in barcode
    df.dropna(subset=['商家SKUID'], inplace=True, axis=0)

    df_upload = pd.DataFrame()

    # PO and Memo
    df_upload['PO'] = df["订单号"].astype('str')
    df_upload['Memo'] = df["订单号"].astype('str')

    # map item id to barcode
    df_upload['UPC Code'] = df['商家SKUID']
    df_upload['UPC Code'] = df_upload['UPC Code'].astype('str')
    df_upload = df_upload.merge(df_internalID, how='left', on='UPC Code')
    df_upload = df_upload.rename(columns={'Internal ID': 'Item ID'})

    # exclude this for now
    df_upload[df_upload['UPC Code'] != "1001-001"]
    
    df_upload.drop('UPC Code', axis=1, inplace=True)

    # quantity
    df_upload['Qty'] = df['订购数量']

    # gross amount
    df_upload['Gross Amount'] = round(df['结算金额'], 2)

    # rate
    df_upload['Rate'] = df_upload['Gross Amount'] / df_upload['Qty']
    df_upload['Rate'] = round(df_upload['Rate'] / 1.13, 2)

    # date
    df_upload['Date'] = df['下单时间'].apply(lambda x: re.sub(r"\t", "", x))
    df_upload['Date'] = pd.to_datetime(df_upload['Date'])
    df_upload['Date'] = df_upload['Date'].dt.strftime('%d/%m/%Y')

    # get strings for external id
    df_upload['JY'] = "JY"
    df_upload['substr'] = df_upload['PO'].astype('str').apply(lambda x: x[-7:])
    df_upload['excelDate'] = pd.to_datetime(df_upload['Date'], format="%d/%m/%Y").apply(lambda x: convertDateToExcel(x.year, x.month, x.day))

    # external id
    df_upload['External ID'] = df_upload['JY'] + df_upload['substr'] + df_upload['excelDate']
    df_upload.drop(['JY','substr','excelDate'], axis=1, inplace=True)

    # constant values
    df_upload['Department'] = 22
    df_upload['Location'] = 76
    df_upload['Currency'] = "CNY"
    df_upload['Customer'] = 3760094
    df_upload['Price Level'] = 'Custom'
    df_upload['VAT'] = ""
    df_upload['Payment Method'] = ""
    df_upload['Account'] = 2070
    df_upload['Cus Checking'] = 12
    df_upload['New Cus Checking'] = "JD"

    # ONLY FOR JD
    # divide rate and gross amount by number of same orders
    externalID_count = df_upload['External ID'].value_counts().reset_index()
    externalID_count.columns = ['External ID', 'count']
    df_upload = df_upload.merge(externalID_count, on="External ID", how="left")

    df_upload['Rate'] = round(df_upload['Rate'] / df_upload['count'], 2)
    df_upload['Gross Amount'] = round(df_upload['Gross Amount'] / df_upload['count'], 2)

    df_upload.drop(['count'], axis=1, inplace=True)

    # rearrange
    df_upload = df_upload[["Item ID", "Qty", "Rate", "External ID", "Date", "Department", "Location", "Currency", "Customer", "Gross Amount",
                            "Price Level", "VAT", "PO", "Payment Method", "Memo", "Account", "Cus Checking", "New Cus Checking"]]

    # sort based on external id
    df_upload.sort_values(by=['External ID'], inplace=True)
    df_upload.reset_index(drop=True, inplace=True)

    return df_upload

def generate_jd():
    df_internalID = get_internalID()
    for file in glob.glob("Input/JD Flagship/*.csv"):
        df = pd.read_csv(file, encoding='gbk')

        name = re.findall(r"[\d\-TO]+", file)

        # process data
        df_upload = process_jd(df, df_internalID)

        # write to csv
        df_upload.to_csv(f'Output/JD Flagship/CSV/{name[0]}.csv', index=False)

        # write to excel
        writer = pd.ExcelWriter(f'Output/JD Flagship/Excel/{name[0]}.xlsx', 
                                engine='xlsxwriter', 
                                engine_kwargs={'options': {'string_to_numbers': False}})
        df_upload.to_excel(writer, index=False, sheet_name='Wechat')
        writer.save()

        print(f"[GENERATED] Output/JD Flagship/CSV/{name[0]}.csv")
        print(f"[GENERATED] Output/JD Flagship/Excel/{name[0]}.xlsx")
