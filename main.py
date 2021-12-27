# -*- coding: utf-8 -*-
#
import os
import io
import sys
import PySimpleGUIQt as sg
from pytz import timezone
from PIL import Image
from datetime import datetime
from threading import Thread
from api import DpcNetAPI

__version__ = 'beta-001'

try:
    f = open('file.txt', 'r')
except:
    with open('file.txt', 'w') as f:
        f.write(',,,')
        f.close()
    for line in open('file.txt', 'r').readlines():
        try:
            params = line.split(',')
            user_email = params[0]
            user_password = params[1]
            user_remember = params[2].replace('\n', '')
        except:
            pass


def resource_path(relative_path):
    base_path = getattr(sys, '_MEIPASS', os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base_path, relative_path)


def get_current_date():
    return datetime.now().astimezone(timezone("America/Sao_Paulo")).strftime("%d/%m/%Y %H:%M:%S")


def long_operation_thread(fun):
    global _result
    global auto_download
    fun.log = _result
    fun.export_to_excel(download_files=auto_download)


def remember():
    for line in open('file.txt', 'r').readlines():
        if line.split(',')[2] == 'False':
            params = ['', '', False]
            return params
        try:
            params = line.split(',')
            user_email = params[0]
            user_password = params[1]
            user_remember = params[2].replace('\n', '')
            return params
        except:
            params = ['', '', False]
            return params


def checked(object):
    if not object or object == '' or object == 'False':
        return False
    return True


def get_img_data(filename, maxsize=(150, 100), file_type=None):
    if not file_type:
        file_type = 'PNG'
    img = Image.open(resource_path(filename))
    img.thumbnail(maxsize)
    bio = io.BytesIO()
    img.save(bio, format=file_type)
    del img
    return bio.getvalue()


def login_layout():
    layout1 = [
        [sg.Text(' \n' * 4)],
        [sg.Text(' ' * 2), sg.Image(data=get_img_data(filename="src/logoDPC.png"), size=(15, 15)),
         sg.Text(' ' * 2)],
        [sg.Text(' \n' * 4)],
        [sg.Text('Usuário: ')], [sg.Input(remember()[0],
                                          font=("Verdana", 13), key='usuario', focus=True, size=(30, 1.1))],
        [sg.Text('Senha: ')],
        [sg.Input(remember()[1], font=("Verdana", 13), key='senha', password_char='*', size=(30, 1.1))],
        [sg.Text('Lembre-me')], [sg.Checkbox('', default=checked(remember()[2]), key='remember')],
        [sg.Text(''), sg.Image(key='loading'), sg.Text('')],
        [sg.Text(' \n' * 5)],
        [sg.Button('Entrar', border_width=True, size=(30, 1.1), key='Entrar')],
        [sg.Text(f'Horário: {str(get_current_date())}', auto_size_text=True, font='Courier 8', key='clock',
                 justification='right', visible=False)]
    ]
    return layout1


def extractor_layout():
    layout2 = [[sg.Text('')],
               [sg.Text('Planilha'), sg.Input(key='-sourcefile-', enable_events=True),
                sg.FileBrowse(file_types=(("Excel files", "*.xls*"),))],
               [sg.Text('\n')],
               [sg.Text('Buscar '), sg.Input(key='code', enable_events=True),
                sg.Button('Testar', bind_return_key=True, key='Test')],
               [sg.Text('')],
               [sg.Text('Baixar imagens?', key="checkbox_label")],
               [sg.Checkbox('', default=False, key='auto_download', font=("verdana", 11))],
               [sg.Column([[sg.Frame('', font='Any 15', layout=[
                   # [sg.Multiline(size=(50, 20), font='Consolas 10', key='Output')]],
                   # element_justification='center')]])],
                   [sg.Output(font='Courier 10', key='Output')]], element_justification='center')]])],
               [sg.Button('Extrair', bind_return_key=True, key='Extract', disabled=True, enable_events=True),
                sg.Button('Sair', button_color=('white', 'firebrick3'), key='Quit')],
               [sg.Text('')],
               [sg.Text('Desenvolvido por: (cleiton.leonel@gmail.com)', auto_size_text=True, font='Courier 8',
                        justification='right'),
                sg.Text(f'Horário: {str(get_current_date())}', auto_size_text=True, font='Courier 8', key='clock',
                        justification='right')]]
    return layout2


def create_window(layout, title):
    return sg.Window(f'DPC Extractor | {title}',
                     auto_size_buttons=False,
                     default_element_size=(20, 1),
                     text_justification='right',
                     size=[300, 800], finalize=True).Layout(layout)


if __name__ == '__main__':
    layout = login_layout()
    window = create_window(layout, title='Login')
    dpc = None
    auto_download = None
    _result = None
    while True:
        button, values = window.Read(timeout=0)
        window['clock'].update(get_current_date())
        if button == sg.WIN_CLOSED:
            sys.exit(0)
        if button in ('Exit', 'Quit', None):
            break
        if button == 'Entrar':
            dpc_user = values['usuario']
            dpc_pass = values['senha']
            if values['remember']:
                remember()
                if dpc_user != '' and dpc_pass != '':
                    with open('file.txt', 'w') as file:
                        file.write(dpc_user + ',' + dpc_pass + ',' + str(values['remember']) + ',\n')
                        file.close()
            else:
                with open('file.txt', 'w') as file:
                    file.write('' + ',' + '' + ',' + str(values['remember']) + ',\n')
                    file.close()
            dpc = DpcNetAPI(username=dpc_user, password=dpc_pass)
            dpc.auth()
            if dpc.current_token:
                window.close()
                layout = extractor_layout()
                window = create_window(layout, title='Home')
        if values.get('-sourcefile-'):
            if values['-sourcefile-'] != '':
                window['Extract'].update(disabled=False)
                window.Refresh()
        if dpc and button == 'Test':
            window['Output'].update('')
            code = values['code']
            print(dpc.get_product_info(code))
        if values.get('auto_download'):
            auto_download = values.get('auto_download')
        if button == 'Extract':
            source_file = values['-sourcefile-']
            dpc.source_file = source_file
            task_dpc = Thread(target=long_operation_thread, args=(dpc,), daemon=True)
            task_dpc.start()
            window['auto_download'].update(visible=False)
            window['checkbox_label'].update(visible=False)
            window.Refresh()
        if _result:
            print(_result)
    window.close()
quit()
