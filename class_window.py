from tkinter import Tk, ttk, filedialog, Frame, StringVar
from glob import glob
import library
import os


class MyWind(Frame):
    def __init__(self, parent, *args, **kwargs):
        Frame.__init__(self, parent, *args, **kwargs)
        self.parent = parent

        self.pat = ''  # path
        self.var = StringVar()
        self.var.set('')
        self.hit = 0

        self.label1 = ttk.Label(parent, textvariable=self.var, text=self.var.get())
        self.label1.grid(row=0, column=0, padx=5, pady=5)
        self.label2 = ttk.Label(parent, text='Enter the text you want to change')
        self.label2.grid(row=1, column=0, padx=5, pady=5)

        self.entry_search = ttk.Entry(parent)  # entry search
        self.entry_search.grid(row=2, column=0, padx=5, pady=5)
        self.entry_search.insert(0, 'entry search')
        self.entry_replace = ttk.Entry(parent)  # entry replace
        self.entry_replace.grid(row=3, column=0, padx=5, pady=5)
        self.entry_replace.insert(0, 'entry replace')

        self.button = ttk.Button(root, text='open dir and transform', command=self.open_directory)
        self.button.grid(row=4, column=0)

    def transform_in_docx(self):
        """конвертируем файлы doc в docx"""
        for path_one in glob(self.pat + '/*.doc', recursive=True):
            self.var.set(path_one)
            library.save_as_docx(path_one)

    def replace_txt(self):
        """изменяем информацию в файле docx"""
        for path_one in glob(self.pat + '/*.docx', recursive=True):
            if os.path.split(path_one)[1].split('.')[1] == 'docx':
                self.var.set(path_one)
                if self.entry_search.get():
                    self.hit += library.search_str(path_one, self.entry_search.get(), self.entry_replace.get())
                else:
                    self.var.set('Repeat the entry of the text')
                    return 1
        self.var.set(f'END hit={self.hit}')
        self.hit = 0

    def open_directory(self):
        """при нажатии на кнопку открываем проводник, выбираем папку"""
        self.pat = filedialog.askdirectory()  # выбираем папку для преобразования
        self.transform_in_docx()
        self.replace_txt()


if __name__ == '__main__':
    root = Tk()
    root.geometry('200x180')
    root.title('Transform doc in docx and replace text in documents')
    MyWind(root).grid()
    root.mainloop()
