import PySimpleGUI as sg
import os.path
import os
import re
import pandas as pd


def create_window():
    file_list_column = [
        [
            sg.Text("Selecione o Diretório:         "),
            sg.In(size=(25, 1), enable_events=True, key="-FOLDER-"),
            sg.FolderBrowse(),
        ],
        [
            sg.Text("Selecione o arquivo matriz: "),
            sg.In(size=(25, 1), enable_events=True, key="-FILE-"),
            sg.FileBrowse(),
        ],
        [
            sg.Listbox(
                values=[], enable_events=True, size=(57, 20), key="-FILE LIST-"
            )
        ],
        [
            sg.Button("Rename"),

        ],
        [
            sg.Button("Cancel")
        ],
    ]

    layout = [
        [
            sg.Column(file_list_column)
        ]
    ]

    window = sg.Window("Rename Files", layout)

    # Run the Event Loop
    while True:
        event, values = window.read()
        if event == "Cancel" or event == sg.WIN_CLOSED:
            break
        # Folder name was filled in, make a list of files in the folder
        if event == "-FOLDER-":
            folder = values["-FOLDER-"]

            try:
                # Get list of files in folder
                file_list = os.listdir(folder)

            except:
                file_list = []

            fnames = [
                f
                for f in file_list
                if os.path.isfile(os.path.join(folder, f))
                   and f.lower().endswith((".pdf", ".xls", ".xlsx", ".doc", ".docx", ".txt", ".png", ".jpg"))
            ]
            window["-FILE LIST-"].update(fnames)
        elif event == "-FILE LIST-":  # A file was chosen from the listbox
            try:
                filename = os.path.join(
                    values["-FOLDER-"], values["-FILE LIST-"][0]
                )
                window["-TOUT-"].update(filename)

            except:
                pass

        if event == "-FILE-":
            file = values["-FILE-"]

        if event == "Rename":
            rename(file, folder)

    window.close()


def rename(file, folder):
    print("entrou na função")
    print("file :" + file)
    print("folder: " + folder)
    file_content = pd.read_excel(file)
    files = os.listdir(folder)
    print(file_content)
    print(files)
    for file in files:
        file_id = file_content['ID'].tolist()
        old_name = file_content['nome_atual'].tolist()
        new_name = file_content['novo_nome'].tolist()


        # if files[i] == old_name:
        #     os.rename(old_name, new_name)
        # else:
        #     print("Arquivo" + old_name + "nao encontrado")




    # try:
    #     planilha = file
    #     newPlanilha = (planilha[planilha['NOME ATUAL'] == int(numero)].head())
    #     nome = newPlanilha.values.tolist()
    #     print(nome[0][1])
    #     novo_nome = nome[0][1] + '.pdf'
    #     return novo_nome
    # except:
    #     return False
    #
    # # Cria o diretório.(Caso não exista.)
    # if not os.path.exists('./Renomeados'):
    #     os.mkdir('./Renomeados')
    #
    # # Navega pelo diretorio dos arquivos PDF's.
    # # Deixa o arquivo com numeros apenas.
    # for _, _, documento in os.walk(folder):
    #     for documento2 in documento:
    #         documento3 = re.sub('-', '', str(documento2))
    #         number = lerplanilha(re.sub('[^0-9]', ' ', documento3))
    #         print(f'Nome Antigo:' + str(documento2))
    #         print(f'Novo Nome:' + str(number))
    #         if number:
    #             os.rename(r'Y:/GIOVANNA MORETI/ACORDOS 27.01/' + str(documento2),
    #                       r'Y:/Daniella/renomear/Renomeados/' + str(number))
    #         else:
    #             print('Não foi possivel renomear o arquivo: ' + documento2)


if __name__ == '__main__':
    create_window()
