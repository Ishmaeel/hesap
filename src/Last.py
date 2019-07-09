import datetime
import os
import shutil


def getFile(day):
    return root + day.strftime("%Y.%m.%d.xlsm")


def findPrev(day):
    for x in range(30):
        day = day - datetime.timedelta(1)
        file = getFile(day)
        print(file)
        if os.path.exists(file):
            return file


root = r"D:\Dropbox\_Finance\\"

now = datetime.datetime.now()
file = getFile(now)

if not os.path.exists(file):
    file = findPrev(now)

if os.path.exists(file):
    print(file)
    os.startfile(file)

else:
    print("oops.")
