from tkinter import *
from tkinter import Tk, ttk, filedialog
from glob import glob
import main
import os
# root.withdraw()  # скрываем экземпляр


def transform_in_docx(paths):
    for path_one in paths:
        result.set(path_one)
        main.save_as_docx(path_one)


def replace_txt(paths):
    for path_one in paths:
        if os.path.split(path_one)[1].split('.')[1] == 'docx':
            result.set(path_one)
            main.search_str(path_one, 'Лекунович М.К.', '****')


def open_directory():
    pat = filedialog.askdirectory()  # выбираем папку для преобразования
    paths_doc = glob(pat+'/*.doc', recursive=True)  # рекурсивно обходим папку с файлами *doc
    transform_in_docx(paths_doc)
    paths_doc = glob(pat + '/*.docx', recursive=True)
    replace_txt(paths_doc)


if __name__ == '__main__':

    root = Tk()
    root.title('Преобразовать, удалить, и заменить')
    result = StringVar()
    result.set('path')
    but = ttk.Button(root, text='open dir and transform', command=open_directory).grid(row=0, column=0)

    label = ttk.Label(root, text=result.get()).grid(row=1, column=0)



# frame = ttk.Frame(rootpadding="3 3 12 12")
# mainframe.grid(column=0, row=0, sticky=(N, W, E, S))
# root.columnconfigure(0, weight=1)
# root.rowconfigure(0, weight=1)
# feet = StringVar()
# feet_entry = ttk.Entry(mainframe, width=7, textvariable=feet)
# feet_entry.grid(column=2, row=1, sticky=(W, E))
#
# meters = StringVar()
# ttk.Label(mainframe, textvariable=meters).grid(column=2, row=2, sticky=(W, E))
#
# ttk.Button(mainframe, text="Calculate").grid(column=3, row=3, sticky=W)
#
# ttk.Label(mainframe, text="feet").grid(column=3, row=1, sticky=W)
# ttk.Label(mainframe, text="is equivalent to").grid(column=1, row=2, sticky=E)
# ttk.Label(mainframe, text="meters").grid(column=3, row=2, sticky=W)
    root.mainloop()
# import os
# import subprocess
# FILEBROWSER_PATH = os.path.join(os.getenv('WINDIR'), 'explorer.exe')
#
# def explore(path):
#     # explorer would choke on forward slashes
#     path = os.path.normpath(path)
#
#     if os.path.isdir(path):
#         subprocess.run([FILEBROWSER_PATH, path])
#     elif os.path.isfile(path):
#         subprocess.run([FILEBROWSER_PATH, '/select,', os.path.normpath(path)])

