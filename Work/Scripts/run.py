"""
Модуль запуска приложения
"""

from GUI_main import MainWindow
import tkinter as tk


def start_gui():
    """
    Создание окна и запуск приложения
    """
    root = tk.Tk()
    top = MainWindow(root)
    root.mainloop()


if __name__ == '__main__':
    start_gui()
