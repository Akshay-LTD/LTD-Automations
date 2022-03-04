import glob
from netsuitesdk import NetSuiteConnection
import re

# connect to netsuite
def connect_ns():
    # sandbox credentials
    NS_ACCOUNT = "4136761_SB1"
    NS_CONSUMER_KEY = "3b95cb1872e0a081efb641aed3e0be55a239bb5ea2e777b97ccf9be7f0940e28"
    NS_CONSUMER_SECRET = "d9977394ef821247bc182d9697e30f670da03594e82a5267f1241730968b7223"
    NS_TOKEN_KEY = "3343acef393fa4a06c42485423096e41b6da4965e3689c8ef984feb466c9a1b2"
    NS_TOKEN_SECRET = "3c540383d29559ac5b20a22f4c9f1b0bf85858321b2dede6945892632cd8df40"
    ns = NetSuiteConnection(
        account=NS_ACCOUNT,
        consumer_key=NS_CONSUMER_KEY,
        consumer_secret=NS_CONSUMER_SECRET,
        token_key=NS_TOKEN_KEY,
        token_secret=NS_TOKEN_SECRET
    )
    return ns

# create a folder in NetSuite file cabinet
def create_folder(nc, id, name):
    created_folder = nc.folders.post(
        {
            "externalId": id,
            "name": name
        }
    )
    return created_folder

# upload all files to netsuite file cabinet
def upload_files(nc):
    platforms = ["JD Flagship", "Tmall", "Wechat"]
    for p in platforms:
        for path in glob.glob(f"Output/{p}/CSV/*.csv"):
            # open file
            file = open(path, 'rb').read()
            # extract name from path
            regex = re.search('CSV\\\(.*\.csv)', path)
            name = regex.group(1)
            # upload file
            uploaded_file = nc.files.post({
                "externalId": name[:-4],
                "name": name,
                'content': file,
                'fileType': '_CSV',
                "folder": {
                            "name": 'Test Folder Shaun',
                            "internalId": 705428,
                            "externalId": 'test-shaun',
                            "type": "folder"
                        }
                }
            )
            print(f"[UPLOAD] {name}")
            print(uploaded_file)

    return uploaded_file

if __name__ == "__main__":
    ns = connect_ns()
    upload_files(ns)