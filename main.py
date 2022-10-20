import threading
import time

import PySimpleGUI as sg
import os
from win32com import client
from sign import sign
from func import clear_input,check_input,excel_to_pdf,initial_layout,get_sheet_names
import subprocess
sg.theme('DarkGreen7')


def make_window(type='PDF'):
    # if type == 'PDF':
    #     pdf=True
    #     excel = False
    # else:
    #     excel=True
    #     pdf=False
    # sg.Input
    layout = [  [sg.Text('Подписание НЭП',border_width=1,background_color='Blue',text_color='orange',size=(50,1),justification='center',auto_size_text=True,font=('Times new roman',20,'bold'))],
                [sg.T("Тип файла для подписи: "),
                 sg.Combo(['PDF', 'EXCEL'], size=(10, 3), key='type_file', enable_events=True, default_value=type,
                          readonly=True)],
                [sg.Column(initial_layout(sg), key='PDF')],
                ]

    return sg.Window('Подписание электронной подписью Excel и PDF', layout, size=(800,400),icon='plane.ico')

window = make_window()
# window.extend_layout(,)
# window.extend_layout(window['EXCEL'],)
while True:

    event, values = window.read()
    print(event)
    # print(values)
    if event:
        window['success'].update(visible=False)
    if event == sg.WIN_CLOSED or event=='exit':
        break
    elif event == 'clear':
        clear_input(window,sg)

    elif event == 'sign':

        if check_input(values,sg):
            # window['waiting'].update(visible=True)
            pfx = values['PFX']
            password = values['PASSWORD']
            folder = values['FOLDER']
            filename, file_extension = os.path.splitext(values['FILE'])
            if values['type_file'] == 'EXCEL':
                pdf_to_sign = excel_to_pdf(values)
                output_file = f'{folder}/{values["sheet_list"]}.pdf'
            else:
                pdf_to_sign = values['FILE']
                output_file = f"{folder}/{values['FILE'].split('/')[-1]}"
            if pdf_to_sign != False:
                res = sign(pfx,password,pdf_to_sign,output_file)
            else:
                res = False

            # time.sleep(0.2)
            if res:
                if values['type_file'] == 'EXCEL':
                    os.remove(pdf_to_sign)
                window['success'].update(visible=True)
                if values['open']:
                    subprocess.Popen([output_file],shell=True)
            else:
                sg.PopupError('Не удалось подписать документ, что-то пошло не так')
            # window['waiting'].update(visible=False)
    elif event == 'type_file':
        if values['type_file'] == 'PDF':
            window['sheets'].update(visible=False)
            window['FILE'].update('')
            window['caution'].update('Перед подписанием закройте PDF файл подлежащий подписанию!')
            window['browse_files'].FileTypes = [('Файлы для подписания', '*.pdf')]
        else:
            window['FILE'].update('')
            window['caution'].update('Сохраните и закройте все Excel файлы перед подписанием')
            window['browse_files'].FileTypes= [('Файлы для подписания', '*.xlsx;xls')]

    elif event == 'FILE':
        if values['type_file'] == 'EXCEL':
            sheets = get_sheet_names(values['FILE'])

            window['sheets'].update(visible=True)
            window['sheet_list'].update(values=sheets,readonly=True)
            # window['sheet_list'].DefaultText = sheets[0]
        # window.Refresh()
        # window.extend_layout(window['main'], common_layout)


'''
TODO:
1. 


'''

