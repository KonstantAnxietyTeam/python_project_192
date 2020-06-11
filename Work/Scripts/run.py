from GUI_main import MainWindow
import tkinter as tk


def start_gui():
    """Starting point when module is the main routine."""
    root = tk.Tk()
    top = MainWindow(root)
    root.mainloop()


if __name__ == '__main__':
    start_gui()
