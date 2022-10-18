
import PySimpleGUI as sg
import os
from win32com import client
from func import clear_input,check_input,excel_to_pdf
sg.theme('DarkTeal')
layout = [  [sg.Text('Добро пожаловать')],
            [sg.Text('Выберете файл для подписания'),sg.Input(key='FILE'), sg.FileBrowse('выбрать',file_types=[('Файлы для подписания','*.xlsx;*.xls;*.pdf')])],
            [sg.Text('Выберете сертификат для подписания'),sg.Input(key='PFX'),sg.FileBrowse('выбрать',file_types=[('Сертификат',"*.pfx")])],
            [sg.Text('Выберете папку куда сохранить подписанный документ'),sg.Input(key='FOLDER'),sg.FolderBrowse('выбрать')],
            # [sg.ProgressBar(max_value=100)],
            [sg.Button('Подписать',key='sign'), sg.Button('Сбросить',key='clear'), sg.Exit('Выйти',key='exit')] ]
window = sg.Window('Подписание электронной подписью Excel и PDF', layout, size=(800,400))

while True:
    event, values = window.read()
    if event == sg.WIN_CLOSED or event=='exit':
        break
    elif event == 'clear':
        clear_input(window,sg)

    elif event == 'sign':
        if check_input(values,sg):
            filename, file_extension = os.path.splitext(values['FILE'])
            excel_to_pdf(values,filename)

'''
TODO:
1. choosing sheet number by user ( or may be list of sheet name in workbook)
2. Get password from user
3. Add digital signature in PDF


'''

