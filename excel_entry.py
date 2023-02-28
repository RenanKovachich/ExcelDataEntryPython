import PySimpleGUI as sg
import pandas as pd

sg.theme('Black')

EXCEL_FILE = 'dados.xlsx'
df = pd.read_excel(EXCEL_FILE)

layout = [
    [sg.Text("Digite os dados requisitados pelo formulário: ")],
    [sg.Text('Nome Completo:', size=(15, 1)), sg.InputText(key='Nome')],
    [sg.Text('Cidade que mora:', size=(15, 1)), sg.InputText(key='Cidade')],
    [sg.Text('Cor Favorita:', size=(15, 1)), sg.Combo(['Azul', 'Vermelho', 'Verde', 'Roxo', 'Amarelo', 'Laranja',
                                                       'Preto', 'Branco', 'Cinza'], key='Cor Favorita')],
    [sg.Text('Idiomas que domino', size=(15, 1)),
     sg.Checkbox('Portugues', key='Portugues'),
     sg.Checkbox('Espanhol', key='Espanhol'),
     sg.Checkbox('Italiano', key='Italiano'),
     sg.Checkbox('Inglês', key='Ingles'),
     sg.Checkbox('Francês', key='Frances'),
     sg.Checkbox('Outro(s)', key='Outro(s)')],
    [sg.Text('Idade: ', size=(15, 1)), sg.Spin([i for i in range(0, 100)],
                                               initial_value=0, key='Idade')],
    [sg.Submit(), sg.Button('Clear'), sg.Exit()]
]

window = sg.Window('Formulario de entrada de dados', layout)


def clear_input():
    for key in values:
        window[key]('')
    return None


while True:
    event, values = window.read()
    if event == sg.WIN_CLOSED or event == 'Exit':
        break
    if event == 'Clear':
        clear_input()
    if event == 'Submit':
        df = df.append(values, ignore_index=True)
        df.to_excel(EXCEL_FILE, index=False)
        sg.popup("Informações Salvas!")
        clear_input()

window.close()
