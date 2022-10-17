
import PySimpleGUI as sg

def clear_input(window):
    for key, element in window.key_dict.items():
        if isinstance(element, sg.Input):
            element.update(value='')

def check_input(values):
    keys_missed = []
    key_dict = {'FILE':'Не выбран файл для подписания',
                'PFX':'Не выбран сертификат для подписания',
                'FOLDER':'Не выбрана папка для сохранения подписанного документа'}
    for value in values.items():
        print(values)
        if len(values[0]) < 1:
            keys_missed.append(value[1])
    popup_string = 'Не удалось подписать потому что: \n'
    i = 0
    for key in keys_missed:
        i+=1
        popup_string+=f'\n {i}. {key_dict[key]}'
    sg.PopupError(popup_string,title='Ахтунг!', keep_on_top=True)
    return False



sg.theme('DarkTeal')
layout = [  [sg.Text('Добро пожаловать')],
            [sg.Text('Выберете файл для подписания'),sg.Input(), sg.FileBrowse('выбрать',file_types=[('Файлы для подписания','*.xlsx;*.xls;*.pdf')],key='FILE')],
            [sg.Text('Выберете сертификат для подписания'),sg.Input(),sg.FileBrowse('выбрать',file_types=[('Сертификат',"*.pfx")],key='PFX')],
            [sg.Text('Выберете папку куда сохранить подписанный документ'),sg.Input(),sg.FolderBrowse('выбрать',key='FOLDER')],
            [sg.ProgressBar(max_value=100)],
            [sg.Button('Подписать',key='sign'), sg.Button('Сбросить',key='clear')] ]
window = sg.Window('My File Browser', layout, size=(800,400))

while True:
    event, values = window.read()
    if event == sg.WIN_CLOSED:
        break
    elif event == 'clear':
        clear_input(window)

    elif event == 'sign':
        check_input(values)


    print('FILE: ', values['FILE'])
    print('PFX: ', values['PFX'])
    print('FOLDER: ', values['FOLDER'])

