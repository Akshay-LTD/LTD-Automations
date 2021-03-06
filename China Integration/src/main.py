from jd.jd import *
from tmall.tmall import *
from wechat.wechat import *
from helper.helper import *
from helper.upload import *
from helper.drive import *

if __name__ == "__main__":

    # clear everything in Input and Output
    clear_dir()

    # extract raw data to Input
    move_files()
    del_zip()

    # generate output files for JD Flagship
    generate_jd()

    # generate output files for Tmall
    generate_tmall()

    # generare output files for Wechat
    generate_wechat()

    # upload files to google drive
    upload_drive()

    # upload files in output to NetSuite
    # nc = connect_ns()
    # upload_files(nc)