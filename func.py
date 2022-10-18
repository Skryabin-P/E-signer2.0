import pandas as pd
import tempfile
import pdfkit
from win32com import client
def clear_input(window,sg):
    for key, element in window.key_dict.items():
        if isinstance(element, sg.Input):
            element.update(value='')


def check_input(values,sg):
    keys_missed = []
    key_dict = {'FILE': 'Не выбран файл для подписания',
                'PFX': 'Не выбран сертификат для подписания',
                'FOLDER': 'Не выбрана папка для сохранения подписанного документа'}
    for key in key_dict:
        if len(values[key]) < 1:
            keys_missed.append(key)

    if len(keys_missed) > 0:
        popup_string = 'Не удалось подписать потому что: \n'
        i = 0
        for key in keys_missed:
            i += 1
            popup_string += f'\n {i}. {key_dict[key]}'
        sg.PopupError(popup_string, title='Ахтунг!', keep_on_top=True)
        return False
    return True


def excel_to_pdf(values, filename: str):
    excel = client.Dispatch("Excel.Application")
    # excel.visible=0
    sheets = excel.Workbooks.Open(values['FILE'])
    sheets.application.displayalerts = False
    work_sheets = sheets.Worksheets[1]
    # Convert into PDF File
    work_sheets.ExportAsFixedFormat(0, f'{values["FOLDER"]}/2.pdf')
    sheets.Close()