import os
import datetime
import win32com.client as office
import xlrd.sheet
from docxcompose.composer import Composer
from docx import Document as Document

from import_xls import Config

def ConcatenateWords(files, path):
    result = Document(files[0])
    result.add_page_break()
    composer = Composer(result)

    for i in range(1, len(files)):
        doc = Document(files[i])

        if i != len(files) - 1:
            doc.add_page_break()

        composer.append(doc)

    composer.save(path)

def Delete(files):
    for filename in files:
        if os.path.exists(filename) and filename != 'out.docx':
            os.remove(filename)

def exportToDocx(Config, addresseeSheet: xlrd.sheet.Sheet, addresseeList, triadList):
    if not os.path.exists(Config.out):
        os.mkdir(Config.out)

    word = office.gencache.EnsureDispatch('Word.Application')
    i = 0
    for addressee in addresseeList:
        triad = triadList[i]
        congrat = 'Дорогой(ая) ' + addressee + '! поздравляю тебя с днем Бравл Старса! Желаю тебе ' + triad + '!'
        doc = word.Documents.Open(f'{os.getcwd()}\\{Config.template}')
        time_words_folder = f'{os.getcwd()}\\{Config.out}'

        try:
            textbox = doc.Shapes.AddTextbox(1, Config.text_box_pos_x, Config.text_box_pos_y, Config.text_box_width,
                                            Config.text_box_height)
            textbox.TextFrame.TextRange.Text = congrat
            textbox.TextFrame.MarginTop = 0
            textbox.TextFrame.MarginLeft = 0
            textbox.Fill.Visible = 0
            textbox.Line.Visible = 0

            doc.SaveAs2(f'{time_words_folder}\\{i}.docx')
            doc.Close()

        except BaseException as exception:
            doc.Close()
            raise Exception(exception)
        i += 1
    word.Application.Quit()

    Temp = [f'{time_words_folder}\\{i}.docx' for i in range(len(addresseeList))]
    ConcatenateWords(Temp, f'{os.getcwd()}\\{Config.out}\\out.docx')
    Delete(Temp)

