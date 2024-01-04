import os
import pandas as pd
import numpy as np
from pandas import Series, DataFrame
import re

def search_parameters():
    print("You can choose only one parameter:\
                    \n1. Document number;\
                    \n2.Recipient's last name;\
                    \n3.Recipient's city of birth.\nEnter value or 0")
    task = int(input('Document number: '))
    if task != 0:
        t_task = 'Номер документа'
        print(t_task, task)
        return t_task, task
    task = input("Recipient's last name: ")
    if task != '0':
        t_task = 'Фамилия получателя'
        print(t_task, task)
        return t_task, task
    task = input("Recipient's city of birth: ").title()
    if task != '0':
        t_task = 'Место рождения'
        task = 'г. '+ task
        print(t_task, task)
        return t_task, task
    
def read_and_filter_excel(filename, t_task, task, output_directory):
    print(filename, t_task, task)
    try:
        df = pd.read_excel(filename)
        df_filtered = df.loc[df[t_task] == task]
        if df_filtered.size != 0:
            print(f'Searching data is in file {filename}')
            df_filtered.to_excel(rf'{output_directory}\{task}.xlsx')
            return f'{task}.xlsx'
    except Exception as e:
        print(f"Error processing file: {filename}")
        print(e)
        # return f'Error processing file: {filename}: {e}'

if __name__ == "__main__":

    output_directory = r"D:\ProjectPython\PROJECTS\attestats\output"

    try:
        os.makedirs(output_directory)
        print(f"Successfully created the directory: {output_directory}")
    except FileExistsError:
        print(f"The directory: {output_directory} already exists.")
    try:
        os.makedirs(output_directory)
        print(f"Successfully created the directory: {output_directory}")
    except FileExistsError:
        print(f"The directory: {output_directory} already exists.")

    directory = r"D:\ProjectPython\PROJECTS\attestats\input"

    t_task, task = search_parameters()
    t_task, task = search_parameters()

    if not os.path.exists(directory):
        raise FileNotFoundError(f"Directory not found: {directory}")
    if not os.path.exists(directory):
        raise FileNotFoundError(f"Directory not found: {directory}")

    for filename in os.listdir(directory):
        if filename.endswith(".xlsx"):
            read_and_filter_excel(os.path.join(directory, filename),t_task, task)
        else:
            print(f"Skipping non-xlsx file: {filename}")


    