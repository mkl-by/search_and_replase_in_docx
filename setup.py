"""конвертируем с помощью в ехе файлы"""

import sys
from cx_Freeze import setup, Executable
build_exe_options = {"packages": ["os", "tkinter", "glob", "re", "docx", "library"],
                     "excludes": ["asyncio", "email", "html", "http", "logging",
                                  "multiprocessing", "unittest", "urllib"]
                     }

# base="Win32GUI" should be used only for Windows GUI app
base = None
if sys.platform == "win32":
    base = "Win32GUI"

setup(name="doc_search",
    version="0.1",
    description="My GUI application!",
    options={"build_exe": build_exe_options},
    executables=[Executable("class_window.py", base=base)]
      )
