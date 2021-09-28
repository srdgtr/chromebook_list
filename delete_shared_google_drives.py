import pickle
import os.path
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
import json
import pandas as pd
import numpy as np
from datetime import datetime
import pandas.io.formats.excel


SCOPES = [
    "https://www.googleapis.com/auth/drive",
]

"""
Get all info on drives.
"""
creds = None
# The file token.pickle stores the user's access and refresh tokens, and is
# created automatically when the authorization flow completes for the first
# time.
if os.path.exists("token2.pickle"):
    with open("token2.pickle", "rb") as token:
        creds = pickle.load(token)
# If there are no (valid) credentials available, let the user log in.
if not creds or not creds.valid:
    if creds and creds.expired and creds.refresh_token:
        creds.refresh(Request())
    else:
        flow = InstalledAppFlow.from_client_secrets_file("credentials.json", SCOPES)
        creds = flow.run_local_server(port=0)
    # Save the credentials for the next run
    with open("token2.pickle", "wb") as token:
        pickle.dump(creds, token)

service = build("drive", "v3", credentials=creds)
dfcols = [
    "kind",
    "drives",
    "name",
]

device_list = pd.DataFrame(columns=dfcols)

aNextPageToken = "one"
aPageToken = None

while aNextPageToken:
    get_chromebooks_list = service.drives().list(
        pageSize=100,
        pageToken=aPageToken,
        useDomainAdminAccess = True,
        fields="nextPageToken,drives(kind, id, name)",
    )
    chromebooks_list = get_chromebooks_list.execute()
    aNextPageToken = None

    if chromebooks_list:
        chromebooks_list_dict = json.loads(str(chromebooks_list["drives"]).replace("'", '"').replace("\\", ""))
        for aRow in chromebooks_list_dict:
            device_list = device_list.append(chromebooks_list_dict, ignore_index=True)
    if "nextPageToken" in chromebooks_list:
        aPageToken = chromebooks_list["nextPageToken"]
        aNextPageToken = chromebooks_list["nextPageToken"]
    else:
        break

device_list.to_csv("drives.csv")