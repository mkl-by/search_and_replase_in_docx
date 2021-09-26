import re
import os
import win32com.client as win32
from win32com.client import constants
from docx import Document

# дописать удаление doc файлов, сделать exe!!!!!!!!!!!!!!!!!!!!!!!


def search_str(
        path_one: str,  # путь к изменяемому файлу
        search_str: str,  # строка которую необходимо найти
        what_to_replace: str  # на что заменить строку
                ) -> None:
    """
    функция получает путь к файлу docx и текст который необходимо изменить и
    на что заменить
    """
    document = Document(path_one)
    for paragraph in document.paragraphs:
        if search_str in paragraph.text:
            print(paragraph.text, 'найдено')
            paragraph.text = re.sub(search_str, what_to_replace, paragraph.text).strip()
        document.save(path_one)


def save_as_docx(path_one: str) -> None:
    # Opening MS Word
    word = win32.gencache.EnsureDispatch('Word.Application')
    doc = word.Documents.Open(path_one)
    doc.Activate()
    # Rename path with .docx
    new_file_abs = os.path.abspath(path_one)
    new_file_abs = re.sub(r'\.\w+$', '.docx', new_file_abs)
    # Save and Close
    word.ActiveDocument.SaveAs(
        new_file_abs, FileFormat=constants.wdFormatXMLDocument
    )
    doc.Close(False)

