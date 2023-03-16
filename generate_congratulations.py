import random
import xlrd

from import_xls import Config


def abilityChecker(Config, sheet: xlrd.sheet.Sheet, congrats1, congrats2, congrats3):
    read = 0
    for i in range(sheet.nrows):
        ConfigSpec = sheet.cell(rowx=i, colx=0).value
        if ConfigSpec != '':
            read += 1
    if ( (len(congrats1) * len(congrats2) * len(congrats3)) < read ):
        return False
    return True

def generateTriads(congrats1: list, congrats2: list, congrats3: list, Config, sheet: xlrd.sheet.Sheet):
    if (not abilityChecker(Config, sheet, congrats1, congrats2, congrats3)):
        raise Exception("Not enough to generate all triads!!!")

    listOfTriads = list()

    for i in range(len(congrats1)):
        for j in range(len(congrats2)):
            for k in range(len(congrats3)):
                triad = congrats1[i] + ', ' + congrats2[j] + ', ' + congrats3[k]
                listOfTriads.append(triad)

    return listOfTriads