import xlrd
import random


# CONFIG INCLUDES ALL OF THE SPECS THAT SPECIFY CONGRATS GENERATOR SETTINGS
class Config:
    def __init__(self, f='Times New Roman', t='template.docx', o='out', a='addressates', hol='holidays', cc='3',
                 c='congrats', x='50', y='50', w='100',
                 h='100'):
        self.font = f
        self.template = t
        self.out = o
        self.addressates = a
        self.holidays = hol
        self.ccount = cc
        self.congrats = c
        self.text_box_pos_x = x
        self.text_box_pos_y = y
        self.text_box_width = w
        self.text_box_height = h

    def __setitem__(self, key, value):
        match key:
            case 'font':
                self.font = value
            case 'template':
                self.template = value
            case 'out':
                self.out = value
            case 'addressates':
                self.addressates = value
            case 'holidays':
                self.holidays = value
            case 'ccount':
                self.ccount = value
            case 'congrats':
                self.congrats = value
            case 'text_box_pos_x':
                self.text_box_pos_x = value
            case 'text_box_pos_y':
                self.text_box_pos_y = value
            case 'text_box_width':
                self.text_box_width = value
            case 'text_box_height':
                self.text_box_height = value


# OPENS XLS AS SHEETS IN DICTIONARY
def OpenXslxAsSheets(xlsPath):
    xls = xlrd.open_workbook(xlsPath, formatting_info=True)
    sheets = dict()
    for sheet in xls.sheets():
        sheets[sheet.name] = sheet
    return sheets


# READS AND INITIALIZES CONFIG FROM CONFIG SHEET
def InitializeConfig(sheet: xlrd.sheet.Sheet):
    config = Config()
    for i in range(0, sheet.nrows):
        ConfigSpec = sheet.cell(rowx=i, colx=0).value
        if ConfigSpec != '':
            configValue = sheet.cell(rowx=i, colx=1).value
            config[ConfigSpec] = configValue
    return config


def ImportSheet(sheet: xlrd.sheet.Sheet):
    values = []
    for i in range(sheet.nrows):
        for j in range(sheet.ncols):
            value = sheet.cell(rowx=i, colx=j).value
            if value != '':
                values.append(value)
    return values


# GENERATES 3 DIFFERENT INDEXES IN RANGE [indexLowest, indexHighest]
def RandomDifferentIndexes(indexLowest, indexHighest):
    return random.sample(range(indexLowest, indexHighest), 3)


def stringToList(config):
    congratList = config.congrats.split(",")
    return congratList

def getAddresseeList(Config, sheet: xlrd.sheet.Sheet):
    addresseeList = list()
    for i in range(sheet.nrows):
        read = sheet.cell(rowx=i, colx=0).value
        if read != '':
            addresseeList.append(read)
    return addresseeList

def importFromXls(xlsPath, configSheetName, addressates):
    # OPENING EXCEL FILE
    sheets = OpenXslxAsSheets(xlsPath)
    configSheet = sheets[configSheetName]
    config = InitializeConfig(configSheet)

    congratsSheetCounter = config.ccount

    congratList = stringToList(config)

    index1, index2, index3 = RandomDifferentIndexes(0, int(congratsSheetCounter))
    congrats1 = ImportSheet(sheets[congratList[index1]])
    congrats2 = ImportSheet(sheets[congratList[index2]])
    congrats3 = ImportSheet(sheets[congratList[index3]])

    addresseeList = getAddresseeList(config,sheets[addressates])

    print('Selected themes:', congratList[index1],',', congratList[index2],',',congratList[index3])

    return congrats1, congrats2, congrats3, config, sheets[addressates], addresseeList
