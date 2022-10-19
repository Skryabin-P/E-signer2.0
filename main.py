
import PySimpleGUI as sg
import os
from win32com import client
from func import clear_input,check_input,excel_to_pdf,initial_layout,get_sheet_names
sg.theme('DarkTeal')


def make_window(type='PDF'):
    # if type == 'PDF':
    #     pdf=True
    #     excel = False
    # else:
    #     excel=True
    #     pdf=False

    layout = [  [sg.Text('Добро пожаловать')],
                [sg.T("Тип файла для подписи: "),
                 sg.Combo(['PDF', 'EXCEL'], size=(10, 3), key='type_file', enable_events=True, default_value=type,
                          readonly=True)],
                [sg.Column(initial_layout(sg), key='PDF')],
                ]

    return sg.Window('Подписание электронной подписью Excel и PDF', layout, size=(800,400))

window = make_window()
# window.extend_layout(,)
# window.extend_layout(window['EXCEL'],)
while True:

    event, values = window.read()
    print(event)
    # print(values)
    if event == sg.WIN_CLOSED or event=='exit':
        break
    elif event == 'clear':
        clear_input(window,sg)

    elif event == 'sign':
        if check_input(values,sg):
            filename, file_extension = os.path.splitext(values['FILE'])
            excel_to_pdf(values,filename)
    elif event == 'type_file':
        if values['type_file'] == 'PDF':
            window['sheets'].update(visible=False)
            window['FILE'].update('')
            window['browse_files'].FileTypes = [('Файлы для подписания', '*.pdf')]
        else:
            window['FILE'].update('')

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

