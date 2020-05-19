#!/usr/bin/env python3
import sys
import tkinter as tk
import tkinter.ttk as ttk
import pandas as pd
import pickle as pk
from tkinter import filedialog
from tkinter import messagebox as mb


def start_gui():
    """Starting point when module is the main routine."""
    global val, w, root
    root = tk.Tk()
    top = MainWindow(root)
    root.mainloop()


def refreshFromExcel(filename):
    xls = pd.ExcelFile(filename)  #  your repository
    p = []
    for sheet in xls.sheet_names:
        p.append(pd.read_excel(xls, sheet))
    print(p[0])
    saveToPickle("../Data/db.pickle", p)


def saveToPickle(filename, obj):
    if (filename):
        db = open(filename, "wb")
        pk.dump(obj, db)
        db.close()


def dataSort():
    refreshFromExcel()
    db = open("../db.pickle","rb")
    p = pk.load(db)
    db.close()
    db = open("../db.pickle", "wb")
    p.sort_values("Код работника")
    
    
def openFromFile(filename):
    if not filename:
        return
    if (filename[-6::] == "pickle"):
        try:
            dbf = open(filename, "rb")
        except FileNotFoundError:
            mb.showerror(title="Файл не найден!", message="По указанному пути не удалось открыть файл. Будет создана пустая база данных.")
            createEmptyDatabase()
        else:
            MainWindow.currentFile = filename
            MainWindow.db = pk.load(dbf)
            dbf.close()
            MainWindow.modified = False
    else:
        try:
            xls = pd.ExcelFile(filename)  #  your repository
        except FileNotFoundError:
            mb.showerror(title="Файл не найден!", message="По указанному пути не удалось открыть файл. Будет создана пустая база данных.")
            createEmptyDatabase()
        else:
            MainWindow.db = pd.read_excel(xls, list(range(5)))
            MainWindow.currentFile = ''
            MainWindow.modified = True


def createEmptyDatabase():
    MainWindow.db = [pd.DataFrame(columns=['Код', 'Тип выплаты', 'Дата выплаты', 'Сумма', 'Код работника']),
                     pd.DataFrame(columns=['Код', 'Код должности', 'Отделение']),
                     pd.DataFrame(columns=['Код', 'Название', 'Норма (ч)', 'Ставка (ч)']),
                     pd.DataFrame(columns=['Код', 'ФИО', 'Номер договора', 'Телефон', 'Образование', 'Адрес']),
                     pd.DataFrame(columns=['Код', 'Название', 'Телефон'])]
    MainWindow.modified = False
    MainWindow.currentFile = ''
    
    
def idToDec(strHexID):
    return (int(strHexID[1::], 16)-1)
    
    
class MainWindow:
    db = None
    currentFile = ''
    modified = False
    def __init__(self, top=None):
        """This class configures and populates the toplevel window.
           top is the toplevel containing window."""
        # refreshFromExcel("../Data/db.xlsx")  # use once for db.pickle
        openFromFile("../Data/db.pickle")
        
        top.geometry("1000x600+150+30")
        top.resizable(0, 0)
        top.title("База Данных")
        
#        self.Table_Frame = tk.LabelFrame(top)
#        self.Table_Frame.place(relx=0.023, rely=0.017, relheight=0.373,
#                               relwidth=0.207)
#        self.Table_Frame.configure(text='''Таблица''')
#        self.Table_Frame.configure(cursor="arrow")
#
#        self.Add_Button = tk.Button(self.Table_Frame)
#        self.Add_Button.place(relx=0.048, rely=0.625, height=32, width=88,
#                              bordermode='ignore')
#        self.Add_Button.configure(cursor="hand2")
#        self.Add_Button.configure(text='''Добавить''')
#
#        self.ExpTab_Button = tk.Button(self.Table_Frame)
#        self.ExpTab_Button.place(relx=0.531, rely=0.799, height=32, width=88,
#                                 bordermode='ignore')
#        self.ExpTab_Button.configure(cursor="hand2")
#        self.ExpTab_Button.configure(text='''Экспорт''')
#
#        self.Delete_Button = tk.Button(self.Table_Frame)
#        self.Delete_Button.place(relx=0.048, rely=0.799, height=32, width=88,
#                                 bordermode='ignore')
#        self.Delete_Button.configure(cursor="hand2")
#        self.Delete_Button.configure(text='''Удалить''')
#
#        self.Edit_Button = tk.Button(self.Table_Frame)
#        self.Edit_Button.place(relx=0.531, rely=0.625, height=32, width=88,
#                               bordermode='ignore')
#        self.Edit_Button.configure(cursor="hand2")
#        self.Edit_Button.configure(text='''Правка''')
#
#        self.Choice_Label = tk.Label(self.Table_Frame)
#        self.Choice_Label.place(relx=0.386, rely=0.089, height=25, width=65,
#                                bordermode='ignore')
#        self.Choice_Label.configure(text='''Выбор''')
        
        self.Analysis_Frame = tk.LabelFrame(top, text="Анализ")
        self.Analysis_Frame.place(relx=0.24, rely=0.017, relheight=0.373,
                                  relwidth=0.201)

        self.Method_Label = tk.Label(self.Analysis_Frame, text="Метод Анализа")
        self.Method_Label.place(relx=0.199, rely=0.134, height=17, width=127,
                                bordermode='ignore')
        
        self.Analysis_Button = tk.Button(self.Analysis_Frame, text="Анализ")
        self.Analysis_Button.place(relx=0.05, rely=0.795, height=32, width=88,
                                   bordermode='ignore')
        self.Analysis_Button.configure(cursor="hand2")

        self.ExpAn_Button = tk.Button(self.Analysis_Frame, text="Экспорт")
        self.ExpAn_Button.place(relx=0.547, rely=0.795, height=32, width=78,
                                bordermode='ignore')
        self.ExpAn_Button.configure(cursor="hand2")

        self.Analysis_List = tk.Listbox(self.Analysis_Frame)
        self.Analysis_List.place(relx=0.05, rely=0.268, relheight=0.46,
                                 relwidth=0.871, bordermode='ignore')

#        self.Table_List = tk.Listbox(top)
#        self.Table_List.place(relx=0.033, rely=0.097, relheight=0.128,
#                              relwidth=0.188)

        self.Filter_Frame = tk.LabelFrame(top, text="Фильтры")
        self.Filter_Frame.place(relx=0.45, rely=0.017, relheight=0.373,
                                relwidth=0.532)

        self.Filter_List1 = tk.Listbox(self.Filter_Frame)
        self.Filter_List1.place(relx=0.019, rely=0.268, relheight=0.46,
                                relwidth=0.301, bordermode='ignore')

        self.Filter_List2 = tk.Listbox(self.Filter_Frame)
        self.Filter_List2.place(relx=0.338, rely=0.268, relheight=0.46,
                                relwidth=0.301, bordermode='ignore')

        self.Change_Button = tk.Button(self.Filter_Frame, text="Изменить значения")
        self.Change_Button.place(relx=0.357, rely=0.804, height=32, width=148,
                                 bordermode='ignore')
        self.Change_Button.configure(cursor="hand2")

        self.Reset_Button = tk.Button(self.Filter_Frame, text="Сбросить выбор")
        self.Reset_Button.place(relx=0.677, rely=0.800, height=32, width=148,
                                bordermode='ignore')
        self.Reset_Button.configure(cursor="hand2")

        self.Param_Label = tk.Label(self.Filter_Frame, text="Параметры")
        self.Param_Label.place(relx=0.075, rely=0.134, height=25, width=97,
                               bordermode='ignore')
        
        self.Values_Label = tk.Label(self.Filter_Frame, text="Значения")
        self.Values_Label.place(relx=0.414, rely=0.152, height=15, width=83,
                                bordermode='ignore')

        self.Columns_Label = tk.Label(self.Filter_Frame, text="Столбцы")
        self.Columns_Label.place(relx=0.752, rely=0.134, height=24, width=86,
                                 bordermode='ignore')

        self.Filter_List3 = tk.Listbox(self.Filter_Frame)
        self.Filter_List3.place(relx=0.658, rely=0.268, relheight=0.46,
                                relwidth=0.32, bordermode='ignore')
        self.Filter_List3.configure(foreground="#000000")

        self.Filter_Button = tk.Button(self.Filter_Frame, text="Фильтровать")
        self.Filter_Button.place(relx=0.038, rely=0.804, height=32, width=148,
                                 bordermode='ignore')
        self.Filter_Button.configure(cursor="hand2")

        self.Data = ttk.Notebook(top)
        self.Data.place(relx=0.023, rely=0.417, relheight=0.528, relwidth=0.96)
        #  self.Data.configure(takefocus="")

        self.Data_t1 = tk.Frame(self.Data)
        self.Data.add(self.Data_t1, padding=3)
        self.Data.tab(0, text="Учёт")

        self.Data_t2 = tk.Frame(self.Data)
        self.Data.add(self.Data_t2, padding=3)
        self.Data.tab(1, text="Работники")

        self.Data_t3 = tk.Frame(self.Data)
        self.Data.add(self.Data_t3, padding=3)
        self.Data.tab(2, text="Должности")

        self.Data_t4 = tk.Frame(self.Data)
        self.Data.add(self.Data_t4, padding=3)
        self.Data.tab(3, text="Информация")

        self.Data_t5 = tk.Frame(self.Data)
        self.Data.add(self.Data_t5, padding=3)
        self.Data.tab(4, text="Отдел")

        # configure tables
        tabs = [self.Data_t1, self.Data_t2, self.Data_t3,
                self.Data_t4, self.Data_t5]

        self.tables = [0, 1, 2, 3, 4]

        for i in range(len(tabs)):
            self.tables[i] = TreeViewWithPopup(tabs[i])
            self.tables[i].place(relwidth=1.0, relheight=1.0)
            self.tables[i]["columns"] = list(MainWindow.db[i].columns)
            self.tables[i]['show'] = 'headings'
            columns = list(MainWindow.db[i].columns)
            self.tables[i].column("#0", minwidth=5, width=5, stretch=tk.NO)

            self.tables[i].heading("#0", text="")

            for j in range(len(columns)):
                if self.treeCheckForDigit(self.db[i], columns[j]):
                    self.tables[i].heading(columns[j], text=columns[j]+'       ▼▲',\
                            command= lambda _treeview = self.tables[i], _col=columns[j]:self.treeSort(_treeview, _col, False))
                else:
                    self.tables[i].heading(columns[j], text=columns[j]) 
                self.Data.update()
                width = int((self.Data.winfo_width()-30)/(len(columns)-1))
                self.tables[i].column(columns[j], width=width, stretch=tk.NO)

            self.tables[i].column(columns[0], width=30, stretch=tk.NO)

            for j in self.db[i].index:
                items = []
                for title in MainWindow.db[i].columns:
                    items.append(MainWindow.db[i][title][j])
                self.tables[i].insert("", "end", values=items)

        # configure scrolls
        self.scrolls = [0, 1, 2, 3, 4]
        for i in range(len(tabs)):
            self.scrolls[i] = ttk.Scrollbar(self.tables[i], orient="vertical",
                                            command=self.tables[i].yview)
            self.tables[i].config(yscrollcommand=self.scrolls[i].set)
            self.scrolls[i].pack(fill="y", side='right')
            self.scrolls[i] = ttk.Scrollbar(self.tables[i], orient="horizontal",
                                            command=self.tables[i].xview)
            self.tables[i].config(xscrollcommand=self.scrolls[i].set)
            self.scrolls[i].pack(fill="x", side='bottom')

        # menu
        menubar = tk.Menu(top)
        filemenu = tk.Menu(menubar, tearoff=0)
        filemenu.add_command(label="Новый", command=self.newDatabase)
        filemenu.add_command(label="Открыть", command=self.open)
        filemenu.add_command(label="Сохранить", command=self.save)
        filemenu.add_command(label="Сохранить как...", command=self.saveas)
        filemenu.add_separator()
        filemenu.add_command(label="Выход", command=self.exit)
        root.protocol("WM_DELETE_WINDOW", self.exit)
        menubar.add_cascade(label="Файл", menu=filemenu)
        
        helpmenu = tk.Menu(menubar, tearoff=0)
        helpmenu.add_command(label="Добавить", command=self.addRecord)
        helpmenu.add_command(label="Удалить", command=self.deleteRecords)
        helpmenu.add_command(label="Изменить", command=self.modRecord)
        menubar.add_cascade(label="Правка", menu=helpmenu)

        top.config(menu=menubar)

        # status bar
        self.statusbar = tk.Label(top, text="Oh hi. I didn't see you there...", bd=1,
                             relief=tk.SUNKEN, anchor=tk.W)
        self.statusbar.pack(side=tk.BOTTOM, fill=tk.X)
        self.statusUpdate()
        
        root.bind("<Control-a>", self.selectAll)        
        
    def selectAll(self, event=None):
        self.tables[self.Data.index("current")].selectAll()
        
    def modRecord(self, event=None):
        self.tables[self.Data.index("current")].modRecord()
        
    def addRecord(self, event=None):
        self.tables[self.Data.index("current")].addRecord()
        
    def deleteRecords(self, event=None):
        self.tables[self.Data.index("current")].deleteRecords()
    
    def newDatabase(self):
        createEmptyDatabase()
        self.loadTables()
        
    def loadTables(self):
        for tree in self.tables:
            for item in tree.get_children():
                tree.delete(item)
        for i in range(len(self.tables)):
            for j in MainWindow.db[i].index:
                items = []
                for title in MainWindow.db[i].columns:
                    items.append(MainWindow.db[i][title][j])
                self.tables[i].insert("", "end", values=items)
        
    def exit(self):
        if MainWindow.modified:
            ans = tk.messagebox.askyesnocancel("Несохраненные изменения", "Хотите сохранить изменения перед закрытием?")
            if ans:
                self.save()
            elif ans == None:
                return
        root.destroy()     
        
    def open(self):
        if MainWindow.modified:
            ans = tk.messagebox.askyesnocancel("Несохраненные изменения", "Хотите сохранить изменения перед закрытием?")
            if ans:
                self.save()
            elif ans == None:
                return
        file = filedialog.askopenfilename(filetypes = [("pickle files", "*.pickle"), ("Excel files", "*.xls *.xlsx")])
        openFromFile(file)
        self.loadTables()
        
    def save(self):
        if (MainWindow.currentFile != ''):
            saveToPickle(MainWindow.currentFile, MainWindow.db)
        else:
            self.saveas()
        
    def saveas(self):
        filename = filedialog.asksaveasfilename(filetypes = [], defaultextension=".pickle")
        MainWindow.currentFile = filename
        MainWindow.modified = False
        saveToPickle(filename, self.db)
   
    def statusUpdate(self, event=None):
        curTable = self.tables[self.Data.index(self.Data.select())]
        status = "Elements: "
        selected= len(curTable.selection())
        if selected == 0:
            status += str(len(curTable.get_children()))
        else:
            status += ("%d out of %d" % (selected, len(curTable.get_children())))
        self.statusbar.config(text=status)
        self.statusbar.update_idletasks()
        root.after(1, self.statusUpdate)
        
    def treeSort(self, treeview, col, reverse):
        l = [(float(treeview.set(k, col)), k) for k in treeview.get_children('')] 
        l.sort(reverse=reverse)
        for index, (val, k) in enumerate(l):
            treeview.move(k, '', index)
        
        if reverse:
            char = '        ▼'
        else:
            char = '        ▲'
        
        treeview.heading(col,text = col+char, command=lambda: self.treeSort(treeview, col, not reverse))

    def treeCheckForDigit(self, data, col):
        str = data[col][0]
        if type(str) == type(''):
            return False
        else:
            return True
        

class message(tk.Toplevel):
    def __init__(self, parent, prompt="Сообщение"):
        self.opacity = 3.0
        tk.Toplevel.__init__(self, parent)
        self.label = tk.Label(self, text=prompt, background='mistyrose')
        self.label.pack(side="top", fill="x")
        geom = "200x60+" + str(root.winfo_screenwidth()-260) + "+" + str(root.winfo_screenheight()-120)
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
    
    
class CustomDialog(tk.Toplevel):
    #print(CustomDialog(root, "Enter something:").show()) to show
    def __init__(self, parent, prompt):
        tk.Toplevel.__init__(self, parent)

        self.var = tk.StringVar()

        self.label = tk.Label(self, text=prompt)
        self.entry = tk.Entry(self, textvariable=self.var)
        self.ok_button = tk.Button(self, text="OK", command=self.on_ok)

        self.label.pack(side="top", fill="x")
        self.entry.pack(side="top", fill="x")
        self.ok_button.pack(side="right")

        self.entry.bind("<Return>", self.on_ok)

    def on_ok(self, event=None):
        self.destroy()

    def show(self):
        self.wm_deiconify()
        self.entry.focus_force()
        self.wait_window()
        return self.var.get()
        

class TreeViewWithPopup(ttk.Treeview):
    def __init__(self, parent, *args, **kwargs):
        ttk.Treeview.__init__(self, parent, *args, **kwargs)
        self.popup_menu = tk.Menu(self, tearoff=0)
        self.popup_menu.add_command(label="Удалить",
                                    command=self.deleteRecords)
        self.popup_menu.add_command(label="Выбрать все",
                                    command=self.selectAll)
        self.popup_menu.add_command(label="Изменить",
                                    command=self.modRecord)
        self.popup_menu.add_command(label="Добавить",
                                    command=self.addRecord)
        self.bind("<Delete>", self.deleteRecords)
        self.bind("<Button-3>", self.popup)
        
    def popup(self, event):
        try:
            self.popup_menu.tk_popup(event.x_root, event.y_root, 0)
        finally:
            self.popup_menu.grab_release()
            
    def selectAll(self, event=None):
        self.selection_set(tuple(self.get_children()))  
        
    def addRecord(self):
        nb = self.master.master
        nb = nb.index(nb.select())
        dic = askValuesDialog(root, MainWindow.db[nb].columns).show()
        values = list(dic.values())
        keys = list(dic.keys())
        if (len(values) and values[0].get() != ''): # TODO correct input validation
            MainWindow.modified = True
            MainWindow.db[nb] = MainWindow.db[nb].append(
                    pd.DataFrame([[item.get() for item in values]], 
                                   columns=keys), 
                                   ignore_index=True)
            self.insert("", "end", values=[item.get() for item in values])
            
    def deleteRecords(self, event=None):
        nb = self.master.master
        nb = nb.index(nb.select())
        selected = [idToDec(i) for i in self.selection()]
        if not len(selected):
            message(root, "Не выбран элемент").fade()
        else:
            MainWindow.modified = True
            MainWindow.db[nb] = MainWindow.db[nb].drop(selected)
            for item in self.selection():
                self.delete(item)
            
    def modRecord(self):
        nb = self.master.master
        nb = nb.index(nb.select())
        selected = self.selection()
        if not selected:
            message(root, "Не выбран элемент").fade()
        else:
            selected = selected[0]
            dic = askValuesDialog(root, MainWindow.db[nb].columns, currValues=MainWindow.db[nb].iloc[idToDec(selected)].tolist()).show()
            keys = list(dic.keys())
            values = list(dic.values())
            if (len(values) and values[0].get() != ''): # TODO correct input validation
                MainWindow.modified = True
                for i in range(len(keys)):
                    self.item(selected, values=[item.get() for item in values])
                    MainWindow.db[nb].loc[idToDec(selected), keys[i]] = values[i].get()

    def menuFunc(self):
        pass

if __name__ == '__main__':
    start_gui()
