import glob
import shutil
import os

def GetSheet(index):    #Get data about spreadsheet
    # get the doc from the scripting context which is made available to all scripts
    desktop = XSCRIPTCONTEXT.getDesktop()
    model = desktop.getCurrentComponent()
    # get the XText interface
    sheet = model.Sheets.getByIndex(index)
    return sheet

def GetPath(*args):    # modify path to directory
    sheet = GetSheet(0)

    if  sheet.getCellByPosition(5, 1).String == "":
        sheet.getCellByPosition(5, 1).String=("Podaj katalog!")
        return
    else:
        path = sheet.getCellByPosition(5, 1).String
        path = path.replace("\\","\\\\") + "\\\\"
    return path

def ClearTable(*args):  #Clear results
    sheet = GetSheet(0)
    max = int(sheet.getCellByPosition(1, 1).Value)
    for l in range(max):
        sheet.getCellByPosition(0, l + 3).String =""
        sheet.getCellByPosition(1, l + 3).String =""
        sheet.getCellByPosition(2, l + 3).String =""
        sheet.getCellByPosition(3, l + 3).String =""
        sheet.getCellByPosition(4, l + 3).String =""
        sheet.getCellByPosition(5, l + 3).String =""
        sheet.getCellByPosition(9, l + 3).String =""

def GetIDs(sheetindex): # Getting Players IDs from
    i=0
    IDs = []
    sheet = GetSheet(sheetindex)
    while sheet.getCellByPosition(0, i).String != "":
        IDs.append([sheet.getCellByPosition(0, i).String,0])
        i +=1
    return IDs


def Copy(*args):
    # Path to tree
    path = GetPath()
    sheet = GetSheet(0)
    Dane=GetIDs(1)
    dircounter = 0
    i=0
    directories = [d for d in glob.glob(path + "*\\")]
    for d in directories:
        dircounter += 1
        files = [f for f in glob.glob(d + "**/*.txt", recursive=True)]
        for f in files:
            with open(f) as plik:
                hand = plik.read()
            for j in range(len(Dane)):
                if hand.find(Dane[j][0] + ": posts") > 0:
                    Output = path + "\\" + sheet.getCellByPosition(9, 1).String + "_" + os.path.split(path.rstrip("\\"))[1] + "\\" + Dane[j][0] + "\\" + os.path.split(os.path.dirname(f))[1]
                    if not os.path.isdir(Output):
                        os.makedirs(Output)
                    shutil.copy(f, Output + "\\" + os.path.basename(f))
                    sheet.getCellByPosition(0, i + 3).String = Dane[j][0]
                    Dane[j][1] += 1
                    sheet.getCellByPosition(5, i+3).String = f
                    sheet.getCellByPosition(9, i + 3).String = Output + "\\" + os.path.basename(f)
                    i += 1
    sheet.getCellByPosition(2, 1).Value = i
    sheet.getCellByPosition(4, 1).Value = dircounter

    for j in range(len(Dane)):
        sheet.getCellByPosition(1, j + 3).String = Dane[j][0]
        sheet.getCellByPosition(2, j + 3).Value = Dane[j][1]

def Move(*args):
    # Path to tree
    path = GetPath()
    sheet = GetSheet(0)
    Dane=GetIDs(1)
    dircounter = 0
    i=0
    directories = [d for d in glob.glob(path + "*\\")]
    for d in directories:
        dircounter += 1
        files = [f for f in glob.glob(d + "**/*.txt", recursive=True)]
        for f in files:
            with open(f) as plik:
                hand = plik.read()
            for j in range(len(Dane)):
                if hand.find(Dane[j][0] + ": posts") > 0:
                    Output = path + "\\" + sheet.getCellByPosition(9, 1).String + "_" + os.path.split(path.rstrip("\\"))[1] + "\\" + Dane[j][0] + "\\" + os.path.split(os.path.dirname(f))[1]
                    if not os.path.isdir(Output):
                        os.makedirs(Output)
                    shutil.move(f, Output + "\\" + os.path.basename(f))
                    sheet.getCellByPosition(0, i + 3).String = Dane[j][0]
                    Dane[j][1] += 1
                    sheet.getCellByPosition(5, i+3).String = f
                    sheet.getCellByPosition(9, i + 3).String = Output + "\\" + os.path.basename(f)
                    i += 1
                    break
    sheet.getCellByPosition(2, 1).Value = i
    sheet.getCellByPosition(4, 1).Value = dircounter

    for j in range(len(Dane)):
        sheet.getCellByPosition(1, j + 3).String = Dane[j][0]
        sheet.getCellByPosition(2, j + 3).Value = Dane[j][1]





