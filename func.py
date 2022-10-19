import pandas as pd
import tempfile
import pdfkit
from win32com import client
from openpyxl import load_workbook

def initial_layout(sg):
    choose_sheet = [[sg.Text('Выберете лист для подписи'),sg.Combo(key='sheet_list',size=(25,3),values=[])]]
    common_layout = [
        [sg.Text('Выберете файл для подписания'), sg.Input(key='FILE', enable_events=True, readonly=False),
         sg.FileBrowse('выбрать',key='browse_files', file_types=[('Файлы для подписания', '*.pdf')])],
        [sg.Column(choose_sheet,key='sheets',visible=False)],
        [sg.Text('Выберете сертификат для подписания'), sg.Input(key='PFX', readonly=True, ),
         sg.FileBrowse('выбрать', file_types=[('Сертификат', "*.pfx")])],
        [sg.Text('Введите пароль от сертификата'), sg.InputText(key='PASSWORD', password_char='*')],
        [sg.Text('Выберете папку куда сохранить подписанный документ'), sg.Input(key='FOLDER', readonly=True, ),
         sg.FolderBrowse('выбрать')],
        # [sg.ProgressBar(max_value=100)],
        [sg.Button('Подписать', key='sign'), sg.Button('Сбросить', key='clear'), sg.Exit('Выйти', key='exit')]]
    return common_layout

def clear_input(window,sg):
    window['sheets'].update(visible=False)
    for key, element in window.key_dict.items():
        if isinstance(element, sg.Input):
            element.update(value='')


def check_input(values,sg):
    keys_missed = []
    key_dict = {'FILE': 'Не выбран файл для подписания',
                'PFX': 'Не выбран сертификат для подписания',
                'PASSWORD': 'Не введён пароль для сертификата',
                'FOLDER': 'Не выбрана папка для сохранения подписанного документа',
                 }
    if values['type_file'] == 'EXCEL':
        key_dict['sheet_list'] = 'Не выбран лист для подписи'
    print(values)
    for key in key_dict:
        print(f'{key}  {values[key]}')
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

def get_sheet_names(filepath):
    wb = load_workbook(filepath, read_only=True, keep_links=False)
    sheets = wb.sheetnames
    wb.close()
    return sheets
def excel_to_pdf(values, filename: str):
    excel = client.Dispatch("Excel.Application")
    # excel.visible=0
    wb = excel.Workbooks.Open(values['FILE'])
    wb.application.displayalerts = False

    print(get_sheet_names(wb))
    #
    # work_sheets = wb.Worksheets[1]
    # # Convert into PDF File
    # work_sheets.ExportAsFixedFormat(0, f'{values["FOLDER"]}/2.pdf')
    wb.Close()