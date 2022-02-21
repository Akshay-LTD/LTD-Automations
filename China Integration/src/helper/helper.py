import datetime as dt
import shutil
import pandas as pd
import glob
import os
import re
import zipfile

def convertDateToExcel(year, month, day) :
    offset = dt.date(1899,12,30).toordinal()
    itime = dt.date(year,month,day).toordinal()
    
    return str(itime - offset)

def get_internalID():
    
    # read data
    df = pd.read_csv('Item Registry Link Barcode.csv')
    df['UPC Code'] = df['UPC Code'].astype('str')

    return df

def clear_dir():
    platforms = ["JD Flagship", "Tmall", "Wechat"]

    for p in platforms:
        for file in glob.glob(f"Input/{p}/*.csv"):
            os.remove(file)
            print(f"[DELETED] {file}")

    for p in platforms:
        for file in glob.glob(f"Output/{p}/CSV/*.csv"):
            os.remove(file)
            print(f"[DELETED] {file}")

        for file in glob.glob(f"Output/{p}/Excel/*.xlsx"):
            os.remove(file)
            print(f"[DELETED] {file}")

def move_files():

    # Tmall
    for path in glob.glob(f"Input/Raw Data/ExportOrderDetailList*.csv"):
        name = re.sub(r".*(ExportOrderDetailList.*\.csv).*", r"\g<1>", path)
        shutil.move(path, f"Input/Tmall/{name}")
        print(f"[MOVED] {name}")

    # Wechat
    for path in glob.glob(f"Input/Raw Data/订单_订单管理_普通订单*.zip"):
        name = re.sub(r".*(订单_订单管理_普通订单.*\.zip).*", r"\g<1>", path)

        # extract zip to Input/Wechat/
        with zipfile.ZipFile(path, "r") as zf:
            for elem in zf.namelist():
                member = zf.open(elem)
                with open(f"Input/Wechat/{elem}", 'wb') as outfile:
                    shutil.copyfileobj(member, outfile)

        print(f"[EXTRACTED] {name}")

    # JD
    for path in glob.glob(f"Input/Raw Data/[0-9]*.zip"):
        name = re.sub(r".*\\(\d+.*\.zip).*", r"\g<1>", path)
        
        # check for password
        with zipfile.ZipFile(path, "r") as zf:
            for zinfo in zf.infolist():
                is_encrypted = zinfo.flag_bits & 0x1
                if is_encrypted:
                    for elem in zf.namelist():
                        # in case wrong password
                        while True:
                            try:
                                # grab password from user
                                pwd = str(input(f"{name} is password protected. Please enter the password: "))
                                member = zf.open(elem, pwd=bytes(pwd, 'utf-8'))
                                break
                            except RuntimeError as e:
                                if 'Bad password' in str(e):
                                    print("Wrong password. Please try again or press Ctrl + C to exit the program.")
                                    

                        with open(f"Input/JD Flagship/{elem}", 'wb') as outfile:
                            shutil.copyfileobj(member, outfile)
                else:
                    # no password, extract normally
                    for elem in zf.namelist():
                        member = zf.open(elem)
                        with open(f"Input/JD Flagship/{elem}", 'wb') as outfile:
                            shutil.copyfileobj(member, outfile)

        print(f"[EXTRACTED] {name}")

def del_zip():
    for path in glob.glob(f"Input/Raw Data/订单_订单管理_普通订单*.zip"):
        os.remove(path)

    for path in glob.glob(f"Input/Raw Data/[0-9]*.zip"):
        os.remove(path)

if __name__ == "__main__":
    move_files()
    del_zip()