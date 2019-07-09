import datetime
import os
import shutil
import configparser


def getFile(day):
    return root + day.strftime("%Y.%m.%d.xlsm")


def findPrev(day):
    for x in range(30):
        day = day - datetime.timedelta(1)
        file = getFile(day)
        print(file)
        if os.path.exists(file):
            return file


config = configparser.ConfigParser()
config.read('Hesap.ini')

root = config["DEFAULT"]["root"]
report = config["DEFAULT"]["report"]

if root == "":
    print("config not found")
    exit()

now = datetime.datetime.now()
file = getFile(now)
print(file)

if not os.path.exists(file):
    source = findPrev(now)
    shutil.copyfile(source, file)

if os.path.exists(file):
    if (os.path.exists(report)):
        os.startfile(report)
    os.startfile(file)

else:
    print("oops.")
