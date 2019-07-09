import datetime
import os
import shutil
import configparser
import sys


def getFile(day):
    return root + day.strftime("%Y.%m.%d.xlsm")


def findPrev(day):
    for x in range(30):
        day = day - datetime.timedelta(1)
        file = getFile(day)
        print(":", file)
        if os.path.exists(file):
            return file


ini = "Hesap.ini"

if not os.path.exists(ini):
    print(ini, "not found")
    exit()

config = configparser.ConfigParser()
config.read(ini)

root = config["DEFAULT"]["root"]
report = config["DEFAULT"]["report"]

if root == "" or report == "":
    print("config key not found")
    exit()

doLast = len(sys.argv) > 1 and sys.argv[1] == "last"
if doLast:
    print("Open last.")
else:
    print("Copy new.")

now = datetime.datetime.now()
file = getFile(now)

if not os.path.exists(file):
    source = findPrev(now)
    if doLast:
        file = source
    else:
        shutil.copyfile(source, file)

if os.path.exists(file):
    if (not doLast and os.path.exists(report)):
        os.startfile(report)
    print(">", file)
    os.startfile(file)

else:
    print("oops.")
