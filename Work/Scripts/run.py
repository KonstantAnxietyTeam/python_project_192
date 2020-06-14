"""
Модуль запуска приложения
"""

import tkinter as tk
from GUI_main import MainWindow


def start_gui():
    """
    Создание окна и запуск приложения
    """
    root = tk.Tk()
    top = MainWindow(root)
    root.mainloop()


if __name__ == '__main__':
    start_gui()
