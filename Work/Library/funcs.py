import sys
import tkinter as tk
import tkinter.ttk as ttk
import pandas as pd
import pickle as pk
import numpy as np
from tkinter import filedialog
from tkinter import messagebox as mb


def refreshFromExcel(filename):
    xls = pd.ExcelFile(filename)  #  your repository
    p = []
    for sheet in xls.sheet_names:
        p.append(pd.read_excel(xls, sheet))
    saveToPickle("../Data/db.pickle", p)


def saveToPickle(filename, obj):
    if (filename):
        db = open(filename, "wb")
        pk.dump(obj, db)
        db.close()
        
class ChangeDialog(tk.Toplevel):
    def __init__(self, parent, prompt):
        tk.Toplevel.__init__(self, parent)
        self.geometry("200x90+550+230")
        self.resizable(0, 0)
        self.title("")

        self.var = tk.StringVar()

        self.label = tk.Label(self, text=prompt)
        self.entry = tk.Entry(self, textvariable=self.var)
        self.ok_button = tk.Button(self, text="OK", command=self.on_ok)
        self.ok_button.place(relx=0.338, rely=0.55, relheight=0.35,
                             relwidth=0.3, bordermode='ignore')

        self.label.pack(side="top", fill="x")
        self.entry.pack(side="top", fill="x")

        self.entry.bind("<Return>", self.on_ok)

    def on_ok(self, event=None):
        self.destroy()

    # def addPar(self):
    #     newPar = self.entry.get()
        # select = list(top.Filter_List2.curselection())
        # top.Filter_List2.insert(1, newPar)

    def show(self):
        self.wm_deiconify()
        self.entry.focus_force()
        self.wait_window()
        return self.var.get()

class message(tk.Toplevel):
    def __init__(self, parent, prompt="Сообщение"):
        self.opacity = 3.0
        tk.Toplevel.__init__(self, parent)
        self.label = tk.Label(self, text=prompt, background='mistyrose')
        self.label.pack(side="top", fill="x")
        geom = "200x60+" + str(parent.winfo_screenwidth()-260) + "+" + str(parent.winfo_screenheight()-120)
        self.geometry(geom)
        self.resizable(0, 0)
        self.configure(background='lightcoral')
        self.overrideredirect(True)
        self.title("Сообщение")

    def fade(self):
        self.opacity -= 0.01
        if self.opacity <= 0.05:
            self.destroy()
            return
        self.wm_attributes('-alpha', self.opacity)
        self.after(10, self.fade)


class askValuesDialog(tk.Toplevel):
    def __init__(self, parent, labelTexts, currValues=None):
        tk.Toplevel.__init__(self, parent)
        self.geometry("300x400+500+300")
        self.resizable(0, 0)
        self.grab_set() # make modal
        self.focus()
        
        self.Labels = [None] * len(labelTexts)
        self.Edits = [None] * len(labelTexts)
        self.retDict = dict()
        for i in range(len(labelTexts)):
            self.retDict[labelTexts[i]] = tk.StringVar()
            editHeight = .8*400/len(labelTexts)
            self.Labels[i] = tk.Label(self, text=labelTexts[i]+":", anchor='e')
            self.Labels[i].place(relx=.1, y=40+i*editHeight, width=100)
        
            self.Edits[i] = tk.Entry(self, textvariable=self.retDict[labelTexts[i]])
            self.Edits[i].place(relx=.5, y=40+i*editHeight, width=100)
            if labelTexts[i] == 'Код':
                self.Edits[i].configure(state='disabled')
            if currValues:
                self.Edits[i].insert(0, currValues[i])
        
        self.ok_button = tk.Button(self, text="OK", command=self.on_ok)

        self.ok_button.place(relx=.5, rely=.9, relwidth=.4, height=30, anchor="c")

        self.bind("<Return>", self.on_ok)
        self.protocol("WM_DELETE_WINDOW", self.exit)

    def exit(self):
        self.retDict.clear()
        self.on_ok()
        
    def on_ok(self, event=None):
        self.destroy()

    def show(self):
        self.wm_deiconify()
        self.wait_window()
        return self.retDict  
