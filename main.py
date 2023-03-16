import os
import random
import datetime
import win32com.client as office

from exportToDocx import exportToDocx
from import_xls import importFromXls
from generate_congratulations import generateTriads, abilityChecker

xlsPath = 'data.xls'
configSheetName = 'config'
addressee = 'addressates'

congrats1, congrats2, congrats3, config, addresseeSheet, addresseeList = importFromXls(xlsPath, configSheetName, addressee)

print(addresseeList)
triads = generateTriads(congrats1, congrats2, congrats3, config, addresseeSheet)

exportToDocx(config, addresseeSheet,addresseeList,triads)