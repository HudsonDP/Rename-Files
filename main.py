from pathlib import PurePosixPath
from datetime import datetime
import PySimpleGUI as sg
import os.path
import os
import pandas as pd
import shutil


def create_window():
    file_list_column = [
        [
            sg.Text("Selecione o Diretório de Origem: "),
            sg.In(size=(25, 1), enable_events=True, key="-FOLDER-"),
            sg.FolderBrowse(),
        ],
        [
            sg.Listbox(
                values=[], enable_events=True, size=(57, 20), key="-FILE LIST-"
            )
        ],
        [
            sg.Button("Baixar Planilha Matriz")

        ],
        [
            sg.Text("Selecione o Arquivo Matriz:        "),
            sg.In(size=(25, 1), enable_events=True, key="-FILE-"),
            sg.FileBrowse(),
        ],
        [
            sg.Button("Rename")

        ],
        [
            sg.Button(" Cancel ")
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
        if event == " Cancel " or event == sg.WIN_CLOSED:
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
                   and f.lower().endswith(
                    (".rar", ".pdf", ".xls", ".xlsx", ".doc", ".docx", ".txt", ".png", ".jpg", ".bmp"))
            ]
            window["-FILE LIST-"].update(fnames)
        elif event == "-FILE LIST-":  # A file was chosen from the listbox
            try:
                filename = os.path.join(
                    values["-FOLDER-"], values["-FILE LIST-"][0]
                )
            except:
                pass

        if event == "-FILE-":
            file = values["-FILE-"]

        if event == "Baixar Planilha Matriz":
            create_matriz(folder)
            sg.PopupOK("Arquivo Matriz Criado.")

        if event == "Rename":
            try:
                backup(folder)
                try:
                    rename(file, folder)
                    sg.PopupOK('Processo concluido: Verifique a pasta indicada!')
                except ValueError as err:
                    error_rename = "Erro ao Renomear: " + str(err)
                    sg.PopupError(error_rename)
            except ValueError as err:
                error_backup = "Erro ao criar backup: " + str(err)
                sg.PopupError(error_backup)

    window.close()


def rename(file, folder):
    file_content = pd.read_excel(file)
    files = os.listdir(folder)

    i = 0
    for file in files:

        file_format = PurePosixPath(file).suffix
        file_name = file.split(str(file_format))[0]

        old_name = file_content['nome_atual'].tolist()[i]
        new_name = file_content['novo_nome'].tolist()[i] + str(file_format)

        if not os.path.exists(folder + '/Renomeados'):
            os.mkdir(folder + '/Renomeados')
            destino = folder + '/Renomeados'

        if str(file) == str(old_name):
            old_name = os.path.join(folder, old_name)
            new_name = os.path.join(destino, new_name)
            os.rename(old_name, new_name)
            i += 1
        else:
            sg.PopupError("A lista de arquivos da planilha nao coincide com a pasta indicada!")
            break


def backup(folder):

    date_time = datetime.now()
    date_string = date_time.strftime('%d-%m-%Y %H-%M-%s')

    try:
        destino_backup = ('./Backup' + '-' + str(date_string))
        if os.path.exists(destino_backup):
            sg.Popup(
                "Já existe uma pasta backup. Favor verificar o conteúdo antes de remove-la, em sequida tente novamente!")
        else:
            shutil.copytree(str(folder), str(destino_backup))
            sg.PopupOK("Backup Criado")

    except ValueError as err:
        error_backup_mkdir = "Erro ao criar pasta BACKUP: " + str(err)
        sg.PopupError(error_backup_mkdir)


def create_matriz(folder):
    try:
        linhas_planilha = []
        files = os.listdir(folder)
        for file in files:
            linha = {}
            linha['nome_atual'] = str(file)
            linha['novo_nome'] = '[INSIRA O NOME NOVO]'

            linhas_planilha.append(linha)
            try:
                data = pd.DataFrame(linhas_planilha)
                data.to_excel('Matriz.xlsx', index=False)
            except ValueError as err:
                error_matriz = "Erro ao salvar Matriz: " + str(err)
                sg.PopupError(error_matriz)
    except ValueError as err:
        error_matriz = "Erro ao Criar Matriz: " + str(err)
        sg.PopupError(error_matriz)


if __name__ == '__main__':
    create_window()
