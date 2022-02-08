import json
import os.path
import pickle
from datetime import datetime

import numpy as np
import pandas as pd
import pandas.io.formats.excel
from google.auth.transport.requests import Request
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build

date_today: str = datetime.now().strftime("%c").replace(":", "-")


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


SCOPES: list[str] = [
    "https://www.googleapis.com/auth/admin.directory.device.chromeos",
]

"""
Get all info on chromebooks.
"""
CREDS = None
# The file token.pickle stores the user's access and refresh tokens, and is
# created automatically when the authorization flow completes for the first
# time.
if os.path.exists("token.pickle"):
    with open("token.pickle", "rb") as token:
        CREDS = pickle.load(token)
# If there are no (valid) credentials available, let the user log in.
if not CREDS or not CREDS.valid:
    if CREDS and CREDS.expired and CREDS.refresh_token:
        CREDS.refresh(Request())
    else:
        flow = InstalledAppFlow.from_client_secrets_file("credentials.json", SCOPES)
        CREDS = flow.run_local_server(port=0)
    # Save the credentials for the next run
    with open("token.pickle", "wb") as token:
        pickle.dump(CREDS, token)

service = build("admin", "directory_v1", credentials=CREDS)

dfcols: list[str] = [
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

NEXT_PLACE_TOKEN: str | None = "one"
PAGE_TOKEN: str | None = None

while NEXT_PLACE_TOKEN:
    get_chromebooks_list = service.chromeosdevices().list(
        customerId="my_customer",
        orderBy="serialNumber",
        projection="FULL",
        pageToken=PAGE_TOKEN,
        maxResults=200,
        sortOrder=None,
        query=None,
        fields="nextPageToken,chromeosdevices(deviceId, serialNumber, status, lastSync,\
        supportEndDate,annotatedUser, annotatedLocation, annotatedAssetId,notes,model,\
        meid,orderNumber,willAutoRenew,osVersion,platformVersion,firmwareVersion,macAddress,\
        bootMode,lastEnrollmentTime, orgUnitPath, recentUsers, ethernetMacAddress, activeTimeRanges,\
        tpmVersionInfo, cpuStatusReports, systemRamTotal, systemRamFreeReports, diskVolumeReports,\
        manufactureDate, autoUpdateExpiration,lastKnownNetwork)",
    )
    chromebooks_list = get_chromebooks_list.execute()
    NEXT_PLACE_TOKEN = None

    if chromebooks_list:
        chromebooks_list_dict = json.loads(
            str(chromebooks_list["chromeosdevices"]).replace("'", '"').replace("\\", "")
        )
        for aRow in chromebooks_list_dict:
            if aRow["status"] == "ACTIVE":
                info_chromebook = process_get_detaild_chromebook_info(aRow)
                device_list = device_list.append(info_chromebook, ignore_index=True)
    if "nextPageToken" in chromebooks_list:
        PAGE_TOKEN = chromebooks_list["nextPageToken"]
        NEXT_PLACE_TOKEN = chromebooks_list["nextPageToken"]
    else:
        break


def total_usage(time_range: str | list[dict[str, str]]) -> int:
    if not isinstance(time_range, list):
        return 0
    total: list = []
    for time in time_range:
        total.append(time["activeTime"])
    return int(sum(total) / 6000)


device_list["usage_minuten"] = device_list["activeTimeRanges"].apply(total_usage)

device_list["lastKnownNetwork"] = device_list["lastKnownNetwork"].apply(
    lambda x: x if not isinstance(x, list) else x[0] if len(x) else ""
)
device_list["ipaddress"] = device_list["lastKnownNetwork"][
    pd.notna(device_list["lastKnownNetwork"])
].apply(lambda x: x["ipAddress"] if x is not np.nan else x)
device_list["wanIpAddress"] = device_list["lastKnownNetwork"][
    pd.notna(device_list["lastKnownNetwork"])
].apply(lambda x: x["wanIpAddress"] if x is not np.nan else x)
device_list["recentUser"] = device_list["recentUsers"].apply(
    lambda x: x if not isinstance(x, list) else x[0] if len(x) else ""
)
device_list["lastuser"] = device_list["recentUser"][
    pd.notna(device_list["recentUser"])
].apply(lambda x: x.get("email", "") if x is not np.nan else x)
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

num_rows: int = len(device_list)

writer = pd.ExcelWriter(
    "active_chrome_devices" + "_" + date_today + ".xlsx", engine="xlsxwriter"
)
pandas.io.formats.excel.ExcelFormatter.header_style = None
device_list.to_excel(writer, sheet_name="chromebooks", index=False, float_format="%.2f")
nodig_voor_controlen.to_excel(
    writer, sheet_name="controlelijst", index=False, float_format="%.2f"
)
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
    "$A$2:$B$%d" % (num_rows),
    {"type": "formula", "criteria": '=INDIRECT("X"&ROW())="onbekend"', "format": noip},
)
worksheet.conditional_format(
    "$C$2:$C$%d" % (num_rows),
    {"type": "formula", "criteria": '=INDIRECT("BY"&ROW())<>""', "format": zoekgeraakt},
)

# toevoegen van chart met os aantallen
worksheet = workbook.add_chartsheet("os_version")
chart = workbook.add_chart({"type": "column"})
chart.set_title({"name": "aantallen van elke os versie actieve chromebooks"})
chart.set_style(3)
chart.set_plotarea({"gradient": {"colors": ["#33ccff", "#80ffff", "#339966"]}})
chart.set_chartarea({"border": {"none": True}, "fill": {"color": "#bfbfbf"}})
chart.add_series(
    {
        "name": "aantallen",
        "values": "=os_versions!$B$2:$B$20",
        "categories": "=os_versions!$A$2:$A$20",
        "gap": 25,
        "name_font": {"size": 14, "bold": True},
        "data_labels": {
            "value": True,
            "position": "inside_end",
            "font": {"name": "Calibri", "color": "white", "rotation": 345},
        },
    }
)
chart.set_x_axis(
    {
        "name": "os versions",
        "name_font": {"size": 14, "bold": True},
    }
)
chart.set_y_axis(
    {
        "major_unit": 20,
        "name": "aantal",
        "major_gridlines": {"visible": False},
    }
)
chart.set_legend({"none": True})
worksheet.set_chart(chart)

# toevoegen van chart met os aantallen
worksheet = workbook.add_chartsheet("aantallen chromebook")
chart = workbook.add_chart({"type": "column"})
chart.set_title({"name": "aantallen van elke model actieve chromebooks"})
chart.set_style(3)
chart.set_plotarea({"gradient": {"colors": ["#33ccff", "#80ffff", "#339966"]}})
chart.set_chartarea({"border": {"none": True}, "fill": {"color": "#bfbfbf"}})
chart.add_series(
    {
        "name": "aantallen",
        "values": "=chromebook_models!$B$2:$B$10",
        "categories": "=chromebook_models!$A$2:$A$10",
        "gap": 25,
        "name_font": {"size": 14, "bold": True},
        "data_labels": {
            "value": True,
            "position": "inside_end",
            "font": {"name": "Calibri", "color": "white", "rotation": 345},
        },
    }
)
chart.set_x_axis(
    {
        "name": "chromebook model",
        "name_font": {"size": 14, "bold": True},
    }
)
chart.set_y_axis(
    {
        "major_unit": 20,
        "name": "aantal",
        "major_gridlines": {"visible": False},
    }
)
chart.set_legend({"none": True})
worksheet.set_chart(chart)

writer.save()
