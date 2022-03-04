from pydrive.auth import GoogleAuth
from pydrive.drive import GoogleDrive
import glob
import re

def upload_drive():
    gauth = GoogleAuth()           
    # gauth.LocalWebserverAuth()
    drive = GoogleDrive(gauth)

    for file in glob.glob("Output/JD Flagship/CSV/*"):
        name = re.sub(r".*(JD SO Upload)", "\g<1>", file)
        if "SO Upload" in name:
            gfile = drive.CreateFile({
                'parents': [{'id': '1kUkb3nNZHXgz_vBtq5MQkpnZqVL9QmJ3'}],
                'title': name
            })
            gfile.SetContentFile(file)
            gfile.Upload()
            print(f"[UPLOADED TO DRIVE] {name}")

    for file in glob.glob("Output/Tmall/CSV/*"):
        name = re.sub(r".*(Tmall SO Upload)", "\g<1>", file)
        if "SO Upload" in name:
            gfile = drive.CreateFile({
                'parents': [{'id': '1kV6mJu8eK1D_KPel6xsWbv1YX5lCvImd'}],
                'title': name
            })
            gfile.SetContentFile(file)
            gfile.Upload()
            print(f"[UPLOADED TO DRIVE] {name}")

    for file in glob.glob("Output/Wechat/CSV/*"):
        name = re.sub(r".*(Wechat SO Upload)", "\g<1>", file)
        if "SO Upload" in name:
            gfile = drive.CreateFile({
                'parents': [{'id': '1Ij_mZ14Y8f1Lty12od2r8anal2wz1oUV'}],
                'title': name
            })
            gfile.SetContentFile(file)
            gfile.Upload()
            print(f"[UPLOADED TO DRIVE] {name}")

if __name__ == "__main__":
    upload_drive()