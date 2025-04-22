#!/usr/bin/python3

import json
from datetime import datetime, timedelta
from sys import exit, argv
from tempfile import NamedTemporaryFile
from urllib.parse import quote as urlencode
from xml.etree import ElementTree

import requests
from docx import Document
from requests import post

with open("/opt/bots/config.json", "r") as configFile:
    config = json.load(configFile)


houseResidents = config["houseResidents"]
# Make sure the folder passed is correct
basePath = "/opt/bots/HouseDutyReminder/"
dutiesPath = config["weeklyDutiesPath"] + "/" + config["currentSemester"]["year"] + "_" + config["currentSemester"]["season"]
ncURL = config["nextcloudURL"]
password = config["botPassword"]
webhookUrl = config["houseDutyReminderURL"]

# The NextCloud API requires the filepaths to be URL encoded.
dutiesPath = urlencode(dutiesPath)

# See https://docs.nextcloud.com/server/19/developer_manual/client_apis/WebDAV/basic.html for info on the NC API.
# Request a list of all the files in the duty sheet folder with all of their file IDs.
dutiesResponse = requests.request(
    method="PROPFIND",
    url=ncURL + "remote.php/dav/files/bot" + dutiesPath,
    auth=("bot", password),
    data="""<?xml version="1.0" encoding="UTF-8"?>
  <d:propfind xmlns:d="DAV:">
    <d:prop xmlns:oc="http://owncloud.org/ns" xmlns:nc="http://nextcloud.org/ns">
      <oc:fileid/>
      <d:getlastmodified/>
    </d:prop>
  </d:propfind>"""
)
response = dutiesResponse.text
responseXML = ElementTree.fromstring(response)
# Pull out the file IDs from the response and associate them with their matching file name.
fileIDs = {}
files = responseXML.findall("{DAV:}response")
for file in files:
    fileID = None
    lastModified = None
    filePath = file.find("{DAV:}href").text
    for child in file.find("{DAV:}propstat").iter():
        if child.tag.endswith("fileid"):
            fileID = child.text
        elif child.tag.endswith("lastmodified"):
            lastModified = child.text
        if fileID is not None and lastModified is not None:
            break
    if fileID is not None:
        fileIDs[fileID] = (filePath[25:], datetime.strptime(lastModified, "%a, %d %b %Y %H:%M:%S %Z"))

fileID = max(fileIDs.keys())
newestDutySheetPath, lastModifiedTime = fileIDs[fileID]
newestDutySheet = requests.get(
    url=ncURL + "remote.php/dav/files/bot" + newestDutySheetPath,
    auth=("bot", password)
).content

test = len(argv) > 1 and argv[1] == "test"

now = datetime.now()
days = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"]
today = days[now.weekday()]

# Check if it's the same sheet as last week, if so, do nothing.
# This means the bot doesn't need to be manually disabled between semesters.

lastWednesday = now - timedelta(days=now.weekday() + 4)
delta = lastModifiedTime < lastWednesday
if delta >= timedelta(days = 7):
    post(webhookUrl, data=json.dumps({"content": "Either the semester is over, or <@&626093603286155264> forgot to make the duty sheet 👀"}), headers={"Content-Type": "application/json"})
    exit(0)
elif delta >= timedelta(days = 13):
    print("Duty sheet not for this week. Exiting.")
    exit(0)

# These will store the duties later.
kitchenCleanup = {}
weeklyDuties = {}


def substringBefore(s: str, delim: str):
    i = s.find(delim)
    return s[:i] if i != -1 else s

# Pull all the tables out of the documents
docxFile = NamedTemporaryFile()
docxFile.write(newestDutySheet)
doc = Document(docxFile)
tables = doc.tables
for table in tables:
    rows = table.rows
    if len(rows) > 1:
        day = substringBefore(rows[1].cells[0].text, " ")
        if day[-3:].lower() == "day":  # If the left cell on row 2 ends with day
            # Kitchen cleaning
            duties = rows[1:]
            for duty in duties:
                cells = duty.cells
                day = substringBefore(cells[0].text, " ")
                # Sort the duties by roster number
                responsible = sorted([cell.text or cell for cell in cells[1:3]])
                kitchenCleanup[day] = [houseResidents.get(cell) or cell for cell in responsible]
        else:
            # Weekly duties
            duties = rows[1:]
            for duty in duties:
                cells = duty.cells
                duty = cells[0].text
                # Sort the duties by roster number
                responsible = sorted([cell.text for cell in cells[1:3]])
                if responsible[0] == responsible[1]:
                    responsible.pop()
                weeklyDuties[duty] = [houseResidents.get(cell) or cell for cell in responsible]
docxFile.close()

msg = ""
if today == "Tuesday":
    msg = """Weekly Duties:\n{}\n\nDaily:\n""".format(
        "\n".join("{}: {}".format(" and ".join(responsible), duty) for duty, responsible in weeklyDuties.items()
                  if "HOUSE DAY" not in responsible),
    )
msg += "{}, you have daily kitchen cleanup today!".format(" and ".join(kitchenCleanup[today]))


print(msg)
if not test:
    post(webhookUrl, data=json.dumps({"content": msg}), headers={"Content-Type": "application/json"})
