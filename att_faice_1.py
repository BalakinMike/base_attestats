# подключаем библиотеки
import PySimpleGUI as sg
import os
import pandas as pd

def flick_through(folder, t_task, task, output_directory):
    for filename in os.listdir(folder):
            if filename.endswith(".xlsx"):
                result_filename.append(read_and_filter_excel(os.path.join(folder, filename),filename,t_task, task, output_directory))
            else:
                print(f"Skipping non-xlsx file: {filename}")

def read_and_filter_excel(full_filename, filename, t_task, task, output_directory):
    print(full_filename, t_task, task)
    try:
        df = pd.read_excel(full_filename)
        df_filtered = df.loc[df[t_task] == task]
        if df_filtered.size != 0:
            print(f'Searching data is in file {full_filename}')
            name = filename.split('.')[0]
            df_filtered.to_excel(rf'{output_directory}\{name}_'+f'{task}.xlsx')

            return f'{name}_'+f'{task}.xlsx'
    except Exception as e:
        print(f"Error processing file: {filename}")
        print(e)
        
file_list_column = [
    [
        sg.Text("Папка с исходными файлами"),
        sg.In(size=(25, 1), enable_events=True, key="-FOLDER-"),
        sg.FolderBrowse('Обзор'),
        
    ],
    [sg.Text("Ввод поисковых параметров"),],
    [sg.Text("Номер документа   "), sg.InputText(key="-DATA_N-")],
    [sg.Text("Фамилия выпускника"), sg.InputText(key="-DATA_LN-")],
    [sg.Text("Город рождения    "), sg.InputText(key="-DATA_Ct-")],
    [sg.Button('Поиск', enable_events=True, key="-SEARCH-")],
    [
        sg.Listbox(
            values=[], enable_events=True, size=(40, 20), key="-FILE LIST-"
        )
    ],
]

# For now will only show the name of the file that was chosen
output_viewer_column = [
    [sg.Text("Файлы с найденными сведениями:")],
    [
        sg.Listbox(
            values=[], enable_events=True, size=(40, 20), key="-DATA_OUTPUT-"
        )
    ],
]

# ----- Full layout -----
layout = [
    [
        sg.Column(file_list_column),
        sg.VSeperator(),
        sg.Column(output_viewer_column),
    ]
]

window = sg.Window("Поиск в базе аттестатов", layout)

while True:
    # получаем события, произошедшие в окне
    event, values = window.read()
    if event == "Exit" or event == sg.WIN_CLOSED:
        break
    # если нажали на кнопку
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
            and f.lower().endswith((".xlsx"))
        ]
        output_directory = (os.path.join(folder, "output"))
        
        try:
            os.makedirs(output_directory)
            print(f"Successfully created the directory: {output_directory}")
        except FileExistsError:
            print(f"The directory: {output_directory} already exists.")
        
        window["-FILE LIST-"].update(fnames)

    elif event == "-FILE LIST-":  # A file was chosen from the listbox
        try:
            filename = os.path.join(
                values["-FOLDER-"], values["-FILE LIST-"][0]
            )
            

        except:
            pass
    
    result_filename = []
    
    if event == "-SEARCH-":
        if values["-DATA_LN-"] != '':
            t_task = 'Фамилия получателя'
            task = values["-DATA_LN-"]
            flick_through(folder, t_task, task, output_directory)
        elif values["-DATA_N-"] != '':
            t_task = 'Номер документа'
            task = int(values["-DATA_N-"])
            flick_through(folder, t_task, task, output_directory)
        elif values["-DATA_Ct-"] !='':
            t_task = 'Место рождения'
            task = 'г.' + values["-DATA_Ct-"].title()
            flick_through(folder, t_task, task, output_directory)

        window["-DATA_OUTPUT-"].update(result_filename)
    
# закрываем окно и освобождаем используемые ресурсы
window.close()