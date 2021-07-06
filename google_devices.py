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

date_today = datetime.now().strftime("%c").replace(":", "-")


def process_get_detaild_chromebook_info(detaild_device_info):
    chromebook_info = pd.Series(
        [
            detaild_device_info.get("deviceId"),
            detaild_device_info.get("serialNumber"),
            detaild_device_info.get("status"),
            detaild_device_info.get("lastSync"),
            detaild_device_info.get("supportEndDate"),
            detaild_device_info.get("annotatedUser"),
            detaild_device_info.get("annotatedLocation"),
            detaild_device_info.get("annotatedAssetId"),
            detaild_device_info.get("notes"),
            detaild_device_info.get("model"),
            detaild_device_info.get("osVersion"),
            detaild_device_info.get("platformVersion"),
            detaild_device_info.get("firmwareVersion"),
            detaild_device_info.get("macAddress"),
            detaild_device_info.get("bootMode"),
            detaild_device_info.get("lastEnrollmentTime"),
            detaild_device_info.get("orgUnitPath"),
            detaild_device_info.get("recentUsers"),
            detaild_device_info.get("activeTimeRanges"),
            detaild_device_info.get("tpmVersionInfo"),
            detaild_device_info.get("cpuStatusReports"),
            detaild_device_info.get("manufactureDate"),
            detaild_device_info.get("autoUpdateExpiration"),
            detaild_device_info.get("lastKnownNetwork"),
        ],
        index=dfcols,
    )
    return chromebook_info

    # If modifying these scopes, delete the file token.pickle.


SCOPES = [
    "https://www.googleapis.com/auth/admin.directory.device.chromeos",
]

"""
Get all info on chromebooks.
"""
creds = None
# The file token.pickle stores the user's access and refresh tokens, and is
# created automatically when the authorization flow completes for the first
# time.
if os.path.exists("token.pickle"):
    with open("token.pickle", "rb") as token:
        creds = pickle.load(token)
# If there are no (valid) credentials available, let the user log in.
if not creds or not creds.valid:
    if creds and creds.expired and creds.refresh_token:
        creds.refresh(Request())
    else:
        flow = InstalledAppFlow.from_client_secrets_file("credentials.json", SCOPES)
        creds = flow.run_local_server(port=0)
    # Save the credentials for the next run
    with open("token.pickle", "wb") as token:
        pickle.dump(creds, token)

service = build("admin", "directory_v1", credentials=creds)

dfcols = [
    "deviceId",
    "serialNumber",
    "status",
    "lastSync",
    "supportEndDate",
    "annotatedUser",
    "annotatedLocation",
    "annotatedAssetId",
    "notes",
    "model",
    "osVersion",
    "platformVersion",
    "firmwareVersion",
    "macAddress",
    "bootMode",
    "lastEnrollmentTime",
    "orgUnitPath",
    "recentUsers",
    "activeTimeRanges",
    "tpmVersionInfo",
    "cpuStatusReports",
    "manufactureDate",
    "autoUpdateExpiration",
    "lastKnownNetwork",
]
device_list = pd.DataFrame(columns=dfcols)

aNextPageToken = "one"
aPageToken = None

while aNextPageToken:
    get_chromebooks_list = service.chromeosdevices().list(
        customerId="my_customer",
        orderBy="serialNumber",
        projection="FULL",
        pageToken=aPageToken,
        maxResults=200,
        sortOrder=None,
        query=None,
        fields="nextPageToken,chromeosdevices(deviceId, serialNumber, status, lastSync, supportEndDate,annotatedUser, annotatedLocation, annotatedAssetId,notes,model,meid,orderNumber,willAutoRenew,osVersion,platformVersion,firmwareVersion,macAddress,bootMode,lastEnrollmentTime, orgUnitPath, recentUsers, ethernetMacAddress, activeTimeRanges, tpmVersionInfo, cpuStatusReports, systemRamTotal, systemRamFreeReports, diskVolumeReports, manufactureDate, autoUpdateExpiration,lastKnownNetwork)",
    )
    chromebooks_list = get_chromebooks_list.execute()
    aNextPageToken = None

    if chromebooks_list:
        chromebooks_list_dict = json.loads(str(chromebooks_list["chromeosdevices"]).replace("'", '"').replace("\\", ""))
        for aRow in chromebooks_list_dict:
            if aRow["status"] == "ACTIVE":
                info_chromebook = process_get_detaild_chromebook_info(aRow)
                device_list = device_list.append(info_chromebook, ignore_index=True)
    if "nextPageToken" in chromebooks_list:
        aPageToken = chromebooks_list["nextPageToken"]
        aNextPageToken = chromebooks_list["nextPageToken"]
    else:
        break


def total_usage(time_range):
    if not isinstance(time_range, list):
        return "0"
    else:
        total = []
        for time in time_range:
            total.append(time["activeTime"])
        return int(sum(total) / 6000)


device_list["usage_minuten"] = device_list["activeTimeRanges"].apply(total_usage)

device_list["lastKnownNetwork"] = device_list["lastKnownNetwork"].apply(
    lambda x: x if not isinstance(x, list) else x[0] if len(x) else ""
)
device_list["ipaddress"] = device_list["lastKnownNetwork"][pd.notna(device_list["lastKnownNetwork"])].apply(
    lambda x: x["ipAddress"] if x is not np.nan else x
)
device_list["wanIpAddress"] = device_list["lastKnownNetwork"][pd.notna(device_list["lastKnownNetwork"])].apply(
    lambda x: x["wanIpAddress"] if x is not np.nan else x
)
device_list["recentUser"] = device_list["recentUsers"].apply(
    lambda x: x if not isinstance(x, list) else x[0] if len(x) else ""
)
device_list["lastuser"] = device_list["recentUser"][pd.notna(device_list["recentUser"])].apply(
    lambda x: x.get("email", "") if x is not np.nan else x
)
device_list["lastKnownNetwork"].fillna(value="onbekend", inplace=True)
os_versions = device_list["osVersion"].value_counts().reset_index()
os_versions.columns = ["os_versions", "aantal"]
chromebook_models = device_list["model"].value_counts().reset_index()
chromebook_models.columns = ["chromebook_models", "aantal"]
chromebook_location = device_list["annotatedLocation"].value_counts().reset_index()
chromebook_location.columns = ["chromebook_locatie", "aantal"]

nodig_voor_controlen = device_list[
    [
        "serialNumber",
        "lastuser",
        "annotatedAssetId",
        "annotatedLocation",
        "notes",
        "osVersion",
        "model",
        "lastKnownNetwork",
    ]
]

num_rows = len(device_list)

writer = pd.ExcelWriter("active_chrome_devices" + "_" + date_today + ".xlsx", engine="xlsxwriter")
pandas.io.formats.excel.ExcelFormatter.header_style = None
device_list.to_excel(writer, sheet_name="chromebooks", index=False, float_format="%.2f")
nodig_voor_controlen.to_excel(writer, sheet_name="controlelijst", index=False, float_format="%.2f")
os_versions.to_excel(writer, sheet_name="os_versions", index=False)
chromebook_models.to_excel(writer, sheet_name="chromebook_models", index=False)
writer.sheets["os_versions"].hide()
writer.sheets["chromebook_models"].hide()
workbook = writer.book
rotate_items = workbook.add_format({"rotation": "30"})
ean_format = workbook.add_format({"num_format": "000000000000000"})
noip = workbook.add_format({"bg_color": "#57a639"})
zoekgeraakt = workbook.add_format({"bg_color": "#a52019"})
# for sheet in writer.sheets:
worksheet = writer.sheets["chromebooks"]
worksheet.freeze_panes(1, 0)
worksheet.set_row(0, 40, rotate_items)
worksheet.set_column("C:C", 20)
worksheet.set_column("F:F", 20)
worksheet.set_column("G:G", 20)
worksheet.conditional_format(
    "$A$2:$B$%d" % (num_rows), {"type": "formula", "criteria": '=INDIRECT("X"&ROW())="onbekend"', "format": noip}
)
worksheet.conditional_format(
    "$C$2:$C$%d" % (num_rows), {"type": "formula", "criteria": '=INDIRECT("BY"&ROW())<>""', "format": zoekgeraakt}
)

# toevoegen van chart met os aantallen
worksheet = workbook.add_chartsheet("os_version")
chart = workbook.add_chart({"type": "column"})
chart.set_title({"name": "aantallen van elke os versie actieve chromebooks"})
chart.add_series(
    {
        "name": "aantallen",
        "values": "=os_versions!$B$2:$B$20",
        "categories": "=os_versions!$A$2:$A$20",
        "name_font": {"size": 14, "bold": True},
    }
)
chart.set_x_axis(
    {
        "name": "os versions",
        "name_font": {"size": 14, "bold": True},
    }
)
chart.set_y_axis({"major_unit": 10, "name": "aantal"})
chart.set_legend({"none": True})
worksheet.set_chart(chart)

# toevoegen van chart met os aantallen
worksheet = workbook.add_chartsheet("aantallen chromebook")
chart = workbook.add_chart({"type": "column"})
chart.set_title({"name": "aantallen van elke model actieve chromebooks"})
chart.add_series(
    {
        "name": "aantallen",
        "values": "=chromebook_models!$B$2:$B$10",
        "categories": "=chromebook_models!$A$2:$A$10",
        "name_font": {"size": 14, "bold": True},
    }
)
chart.set_x_axis(
    {
        "name": "chromebook model",
        "name_font": {"size": 14, "bold": True},
    }
)
chart.set_y_axis({"major_unit": 10, "name": "aantal"})
chart.set_legend({"none": True})
worksheet.set_chart(chart)

writer.save()
