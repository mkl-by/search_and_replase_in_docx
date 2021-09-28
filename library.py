import re
import os
import win32com.client as win32
from win32com.client import constants
from docx import Document

# сделать exe!!!!!!!!!!!!!!!!!!!!!!!


def search_str(
        path_one: str,  # путь к изменяемому файлу
        search_str: str,  # строка которую необходимо найти
        what_to_replace: str  # на что заменить строку
                    ) -> int:
    """
    функция получает путь к файлу docx и текст который необходимо изменить и
    на что заменить, возвращает колличество совпадений.
    """
    hit = 0
    document = Document(path_one)
    for paragraph in document.paragraphs:
        if search_str in paragraph.text:
            hit += 1
            paragraph.text = re.sub(search_str, what_to_replace, paragraph.text).strip()
        document.save(path_one)
    return hit


# noinspection PyBroadException
def save_as_docx(path_one: str) -> None:
    """функция конвертации файла doc в docx"""
    # Opening MS Word
    word = win32.gencache.EnsureDispatch('Word.Application')
    try:
        doc = word.Documents.Open(path_one)
        doc.Activate()
        # Rename path with .docx
        new_file_abs = os.path.abspath(path_one)
        new_file_abs = re.sub(r'\.\w+$', '.docx', new_file_abs)
        # Save and Close
        word.ActiveDocument.SaveAs(new_file_abs, FileFormat=constants.wdFormatXMLDocument)
        doc.Close(False)
        os.remove(path_one)  # delete file 'path_one.doc'
    except:
        return

