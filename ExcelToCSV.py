import os
import pandas
import tkinter as tk
from tkinter import filedialog


def select_excel_folder():
    global excel_folder_path
    excel_folder_path = filedialog.askdirectory(initialdir=f'{os.getcwd()}', title='Excel 폴더 선택')
    button_excel_folder.configure(text=excel_folder_path)

def select_export_folder():
    global export_folder_path
    export_folder_path = filedialog.askdirectory(initialdir=f'{os.getcwd()}', title='CSV 생성 폴더 선택')
    button_export_folder.configure(text=export_folder_path)

def change_excel_to_csv():
    if excel_folder_path == '' or excel_folder_path == '':
        return

    excel_file_list = []

    for file in os.listdir(excel_folder_path):
        if file.endswith('.xlsx'):
            excel_file_list.append(file)

    for file in excel_file_list:
        excel_file = pandas.read_excel(f'{excel_folder_path}\{file}')
        name = file.replace('.xlsx', '.csv')
        excel_file.to_csv(f'{export_folder_path}\{name}', index=False)

    button_execute.configure(text='Complete')


window = tk.Tk()
window.title("Excel To CSV")

window.geometry("480x260")
window.resizable(False, False)

button_excel_folder = tk.Button(window, width=50, height=3, text="엑셀 폴더 선택", command=select_excel_folder)
button_excel_folder.pack(pady=10)

button_export_folder = tk.Button(window, width=50, height=3, text="CSV 생성 폴더 위치 선택", command=select_export_folder)
button_export_folder.pack(pady=10)

button_execute = tk.Button(window, width=50, height=3, text="실행하기", bg='yellow', command=change_excel_to_csv)
button_execute.pack(pady=10)

window.mainloop()
