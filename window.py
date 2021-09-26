from tkinter import *
from tkinter import Tk, ttk, filedialog
from glob import glob
import main
import os
# root.withdraw()  # скрываем экземпляр
# написать проверки на отсутствие папок и ошибки, переписать в классе


def transform_in_docx(paths):
    """конвертируем файлы doc в docx"""
    for path_one in paths:
        result.set(path_one)
        main.save_as_docx(path_one)


def replace_txt(paths):
    """изменяем информацию в файле docx"""
    for path_one in paths:
        if os.path.split(path_one)[1].split('.')[1] == 'docx':
            result.set(path_one)
            main.search_str(path_one, 'Лекунович М.К.', '****')


def open_directory():
    """при нажатии на кнопку открываем проводник, выбираем папку, собираем пути"""
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
    root.mainloop()
