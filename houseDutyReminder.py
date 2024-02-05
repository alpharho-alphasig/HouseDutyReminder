#!/usr/bin/python3

from os import listdir, remove
from pathlib import Path
from sys import exit, argv
from docx import Document
from requests import post
from datetime import datetime
import json

# If you're updating this script for the new semester, change these two variables!
folderPath = Path(
    "/zpool/docker/volumes/nextcloud_aio_nextcloud_data/_data/admin/files/"
    "NewDrive/ALPHA SIG GENERAL/02_COMMITTEES/07_HOUSING/WEEKLY DUTIES/2023_FALL")
houseResidents = {
    "741": "<@!331935257647513611>",  # Jon
    "760": "<@!172823556143579136>",  # Nick
    "762": "<@!402626429692543006>",  # Colton
    "763": "<@!110120636445114368>",  # Vinuth
    "766": "<@!267153432211750912>",  # Sergei
    "768": "<@!171062832870326273>",  # Zach
    "774": "<@!81999376616009728>",  # Peter
    "775": "<@!128535042111569921>",  # David
    "776": "<@!212611379058835456>",  # Adrian
    "777": "<@!974419797989285888>",  # Gabe
    "778": "<@!113728654176944136>",  # Tim
    "779": "<@!171470084101898240>",  # Evan
    "780": "<@!302822566614138880>",  # Martin
    "781": "<@!544310490357170177>",  # George
    "782": "<@!296825744544366592>",  # Monkey
    "783": "<@!162350127372173312>",  # Monti
    "787": "<@!694206889646620683>",  # Justin
    "788": "<@!143380850556403712>",  # Adam
    "789": "<@!174917188367417344>",  # Shane
    "790": "<@!199253257091022849>",  # Luis
    "791": "<@!690007038696489191>",  # John
    "792": "<@!276489852244197376>",  # Will
    "793": "<@!430179868060418058>",  # Malcolm
    "794": "<@!387421780815642625>",  # Kendall
    "795": "<@!445356465088364554>",  # Sai
    "796": "<@!111249399836860416>",  # Yezen
    "797": "<@!110168137667747840>",  # Akhil
    "798": "<@!392842218211377162>",  # Joey
    "799": "<@!230285108110688257>",  # Leo
    "800": "<@!263490770558779394>",  # Ian
    "801": "<@!131197852188803073>",  # Rasin
    "802": "<@!251559872389316608>",  # Tyler
    "803": "<@!267311410990546945>",  # Eddie
    "804": "<@!176172161281687552>",  #Ronny
    "805": "<@!641486102296920067>",  #Marco
    "806": "<@!319292710349570060>",  #Mike 
    "807": "<@!762479320718508092>",  #Jake
    "808": "<@!209100786980880385>",  #Kyle
    "809": "<@!1140720304893743165>", #Jacob
    "810": "<@!229788520569503746>", #Chris
    "811": "<@!257200789141979146>", #Jayden
    "812": "<@!787030554276659251>", #Bobby
    "813": "<@!307256572852305922>", #Blade
    "814": "<@!1108077393261891734>", #Coby
    "815": "<@!318863344046178304>", #Victor
}

# Make sure the folder passed is correct
basePath = "/opt/bots/HouseDutyReminder/"
if not folderPath.exists():
    print("No folder found at \"{}\".".format(folderPath))
    exit(1)


def readAndClose(filePath):
    with open(filePath, "r") as f:
        s = f.read().strip()
    return s


def lastSubstringAfter(s, delimiter):
    lastIndex = s.rfind(delimiter)
    if lastIndex == -1:
        return s
    return s[lastIndex + 1:]


# Find the newest house duty file
files = []
for file in listdir(folderPath):
    currentFilePath = folderPath.joinpath(str(file))
    if currentFilePath.suffix == ".docx":
        files.append((int(lastSubstringAfter(str(currentFilePath.stem), "_")), currentFilePath))
newestFile = max(files)[1]

test = len(argv) > 1 and argv[1] == "test"

now = datetime.now()
days = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"]
today = days[now.weekday()]

# Check if it's the same sheet as last week, if so, do nothing.
# This means the bot doesn't need to be manually disabled between semesters.

# If it's Tuesday, check if the path's the same as last week and create/remove the samePathAsLastWeek file accordingly.
if Path(basePath + "lastPath").exists() and today == "Tuesday":
    with open(basePath + "lastPath", "r") as lastPathFile:
        lastPath = lastPathFile.read().strip()
        if lastPath == str(newestFile):
            open(basePath + "samePathAsLastWeek", "w+").close()  # Create the samePathAsLastWeek file.
        else:
            remove(basePath + "samePathAsLastWeek")

if Path(basePath + "samePathAsLastWeek").exists():
    print("Same file as last week. Exiting.")
    exit(0)

# These will store the duties later.
kitchenCleanup = {}
weeklyDuties = {}

# Pull all the tables out of the documents
doc = Document(folderPath.joinpath(newestFile))
tables = doc.tables
for table in tables:
    rows = table.rows
    if len(rows) > 1:
        if rows[1].cells[0].text[-3:].lower() == "day":  # If the left cell on row 2 ends with day
            # Kitchen cleaning
            duties = rows[1:]
            for duty in duties:
                cells = duty.cells
                day = cells[0].text
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

lastPathFile = open(basePath + "lastPath", "w+")

with lastPathFile:
    print(newestFile, file=lastPathFile)

msg = ""
if today == "Tuesday":
    msg = """Weekly Duties:\n{}\n\nDaily:\n""".format(
        "\n".join("{}: {}".format(" and ".join(responsible), duty) for duty, responsible in weeklyDuties.items()
                  if "HOUSE DAY" not in responsible),
    )
msg += "{}, you have daily kitchen cleanup today!".format(" and ".join(kitchenCleanup[today]))

webhookUrl = readAndClose(basePath + "URLs/discordWebhook.txt")

print(msg)
if not test:
    post(webhookUrl, data=json.dumps({"content": msg}), headers={"Content-Type": "application/json"})
