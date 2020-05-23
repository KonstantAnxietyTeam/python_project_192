import sys
import tkinter as tk
import tkinter.ttk as ttk
import pandas as pd
import pickle as pk
import numpy as np
from tkinter import filedialog
from tkinter import messagebox as mb
sys.path.insert(1, '../Library')
from funcs import *


newPar = ""
select = []
selected_tab = 0


def start_gui():
    """Starting point when module is the main routine."""
    global val, w, root
    root = tk.Tk()
    top = MainWindow(root)
    root.mainloop()


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
            xls = pd.ExcelFile(filename)  # your repository
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


def saveAsExcel(tree, nb):
    file = filedialog.asksaveasfilename(title="Select file", initialdir='../Data/db1.xlsx', defaultextension=".xlsx", filetypes=[("Excel file", "*.xlsx")])
    if file:
        ids=tree.get_children()
        #dic = dict([tree.column(i)['id'] for i in tree["displaycolumns"]]) # TODO get displayed columns only
        dic = dict.fromkeys(MainWindow.db[nb].columns, [])
        keys = list(dic.keys())
        for i in range(len(keys)):
            dic[keys[i]] = []
        for iid in ids:
            for i in range(len(keys)):
                dic[keys[i]].append(tree.item(iid)["values"][i])

        dic = pd.DataFrame.from_dict(dic)
        try:
           dic.to_excel(file, engine='xlsxwriter',index= False)
        except:
           print("Close the file than retry")
    else:
        print("You did not save the file")


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
        top.minsize(width=1000, height=600)
        top.title("База Данных")

#        self.Table_Frame = tk.LabelFrame(top)
#        self.Table_Frame.place(relx=0.023, rely=0.017, relheight=0.373,
#                               relwidth=0.207)
#        self.Table_Frame.configure(text="Таблица")
#        self.Table_Frame.configure(cursor="arrow")
#
#        self.Add_Button = tk.Button(self.Table_Frame)
#        self.Add_Button.place(relx=0.048, rely=0.625, height=32, width=88,
#                              bordermode='ignore')
#        self.Add_Button.configure(cursor="hand2")
#        self.Add_Button.configure(text="Добавить")
#
#        self.ExpTab_Button = tk.Button(self.Table_Frame)
#        self.ExpTab_Button.place(relx=0.531, rely=0.799, height=32, width=88,
#                                 bordermode='ignore')
#        self.ExpTab_Button.configure(cursor="hand2")
#        self.ExpTab_Button.configure(text="Экспорт")
#
#        self.Delete_Button = tk.Button(self.Table_Frame)
#        self.Delete_Button.place(relx=0.048, rely=0.799, height=32, width=88,
#                                 bordermode='ignore')
#        self.Delete_Button.configure(cursor="hand2")
#        self.Delete_Button.configure(text="Удалить")
#
#        self.Edit_Button = tk.Button(self.Table_Frame)
#        self.Edit_Button.place(relx=0.531, rely=0.625, height=32, width=88,
#                               bordermode='ignore')
#        self.Edit_Button.configure(cursor="hand2")
#        self.Edit_Button.configure(text="Правка")
#
#        self.Choice_Label = tk.Label(self.Table_Frame)
#        self.Choice_Label.place(relx=0.386, rely=0.089, height=25, width=65,
#                                bordermode='ignore')
#        self.Choice_Label.configure(text="Выбор")

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

        self.Filter_Frame = tk.LabelFrame(top, text="Фильтры")
        self.Filter_Frame.place(relx=0.45, rely=0.017, relheight=0.373,
                                relwidth=0.532)

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
        for i in range(len(MainWindow.db)):
            self.tables[i] = TreeViewWithPopup(tabs[i])
            self.tables[i].place(relwidth=1.0, relheight=1.0)
            self.tables[i]["columns"] = list(MainWindow.db[i].columns)
            self.tables[i]['show'] = 'headings'
            columns = list(MainWindow.db[i].columns)
            self.tables[i].column("#0", minwidth=5, width=5, stretch=tk.NO)

            self.tables[i].heading("#0", text="")

            for j in range(len(columns)):
                self.tables[i].heading(columns[j], text=columns[j]+'       ▼▲',\
                           command= lambda _treeview = self.tables[i], _col=columns[j]:self.treeSort(_treeview, _col, False))
                self.Data.update()
                width = int((self.Data.winfo_width()-30)/(len(columns)-1))
                self.tables[i].column(columns[j], width=width, stretch=tk.NO)

            self.tables[i].column(columns[0], width=30, stretch=tk.NO)

            for j in MainWindow.db[i].index:
                items = []
                for title in MainWindow.db[i].columns:
                    items.append(MainWindow.db[i][title][j])
                self.tables[i].add("", values=items)

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

        # filters
        self.Data.bind("<<NotebookTabChanged>>", self.tabChoice)

        #  configure filter lists
        self.Filter_List1 = tk.Listbox(self.Filter_Frame, exportselection=0)
        self.Filter_List1.place(relx=0.019, rely=0.268, relheight=0.46,
                                relwidth=0.301, bordermode='ignore')

        self.Filter_List2 = tk.Listbox(self.Filter_Frame, exportselection=0)
        self.Filter_List2.place(relx=0.338, rely=0.268, relheight=0.46,
                                relwidth=0.301, bordermode='ignore')

        self.Filter_List1.insert('end', "Тип выплаты")
        self.Filter_List1.insert('end', "Дата выплаты")
        self.Filter_List1.insert('end', "Сумма")
        self.Filter_List1.insert('end', "Код работника")
        for i in range(4):
            self.Filter_List2.insert('end', "")

        self.Filter_scroll = tk.Scrollbar(self.Filter_List1)
        self.Filter_List1.config(yscrollcommand=self.Filter_scroll.set)
        self.Filter_List1.bind("<MouseWheel>", self.scrollList2)
        self.Filter_List2.config(yscrollcommand=self.Filter_scroll.set)
        self.Filter_List2.bind("<MouseWheel>", self.scrollList1)

        self.Change_Button = tk.Button(self.Filter_Frame)
        self.Change_Button.place(relx=0.357, rely=0.804, height=32, width=148,
                                 bordermode='ignore')
        self.Change_Button.configure(cursor="hand2")
        self.Change_Button.configure(text="Изменить значения", command=self.open_dialog)
        self.Filter_List1.bind("<<ListboxSelect>>", self.moveSelection2)
        self.Filter_List2.bind("<<ListboxSelect>>", self.moveSelection1)

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

        # Checkboxes
        self.Boxes_Frame = tk.LabelFrame(self.Filter_Frame, text="Столбцы")
        self.Boxes_Frame.place(relx=0.658, rely=0.130, relheight=0.65,
                               relwidth=0.32, bordermode='ignore')
        self.Cvars = []
        for i in range(16):
            self.Cvars.append(tk.BooleanVar())
            self.Cvars[i].set(1)

        self.Cbox0 = tk.Checkbutton(self.Boxes_Frame)
        self.Cbox0.grid(row=0, column=0, sticky='W')
        self.Cbox0.configure(justify='left')
        self.Cbox0.configure(text='''Тип выплаты''', variable=self.Cvars[0])

        self.Cbox1 = tk.Checkbutton(self.Boxes_Frame)
        self.Cbox1.grid(row=1, column=0, sticky='W')
        self.Cbox1.configure(justify='left')
        self.Cbox1.configure(text='''Дата выплаты''', variable=self.Cvars[1])

        self.Cbox2 = tk.Checkbutton(self.Boxes_Frame)
        self.Cbox2.grid(row=2, column=0, sticky='W')
        self.Cbox2.configure(justify='left')
        self.Cbox2.configure(text='''Сумма''', variable=self.Cvars[2])

        self.Cbox3 = tk.Checkbutton(self.Boxes_Frame)
        self.Cbox3.grid(row=3, column=0, sticky='W')
        self.Cbox3.configure(justify='left')
        self.Cbox3.configure(text='''Код работника''', variable=self.Cvars[3])

        self.Cbox4 = tk.Checkbutton(self.Boxes_Frame)
        self.Cbox4.grid(row=0, column=0, sticky='W')
        self.Cbox4.configure(justify='left')
        self.Cbox4.configure(text='''Код должности''', variable=self.Cvars[4])
        self.Cbox4.grid_forget()

        self.Cbox5 = tk.Checkbutton(self.Boxes_Frame)
        self.Cbox5.grid(row=1, column=0, sticky='W')
        self.Cbox5.configure(justify='left')
        self.Cbox5.configure(text='''Отделение''', variable=self.Cvars[5])
        self.Cbox5.grid_forget()

        self.Cbox6 = tk.Checkbutton(self.Boxes_Frame)
        self.Cbox6.grid(row=0, column=0, sticky='W')
        self.Cbox6.configure(justify='left')
        self.Cbox6.configure(text='''Название''', variable=self.Cvars[6])
        self.Cbox6.grid_forget()

        self.Cbox7 = tk.Checkbutton(self.Boxes_Frame)
        self.Cbox7.grid(row=1, column=0, sticky='W')
        self.Cbox7.configure(justify='left')
        self.Cbox7.configure(text='''Норма (ч)''', variable=self.Cvars[7])
        self.Cbox7.grid_forget()

        self.Cbox8 = tk.Checkbutton(self.Boxes_Frame)
        self.Cbox8.grid(row=2, column=0, sticky='W')
        self.Cbox8.configure(justify='left')
        self.Cbox8.configure(text='''Ставка (ч)''', variable=self.Cvars[8])
        self.Cbox8.grid_forget()

        self.Cbox9 = tk.Checkbutton(self.Boxes_Frame)
        self.Cbox9.grid(row=0, column=0, sticky='W')
        self.Cbox9.configure(justify='left')
        self.Cbox9.configure(text='''ФИО''', variable=self.Cvars[9])
        self.Cbox9.grid_forget()

        self.Cbox10 = tk.Checkbutton(self.Boxes_Frame)
        self.Cbox10.grid(row=1, column=0, sticky='W')
        self.Cbox10.configure(justify='left')
        self.Cbox10.configure(text='''Номер договора''', variable=self.Cvars[10])
        self.Cbox10.grid_forget()

        self.Cbox11 = tk.Checkbutton(self.Boxes_Frame)
        self.Cbox11.grid(row=2, column=0, sticky='W')
        self.Cbox11.configure(justify='left')
        self.Cbox11.configure(text='''Телефон''', variable=self.Cvars[11])
        self.Cbox11.grid_forget()

        self.Cbox12 = tk.Checkbutton(self.Boxes_Frame)
        self.Cbox12.grid(row=3, column=0, sticky='W')
        self.Cbox12.configure(justify='left')
        self.Cbox12.configure(text='''Образование''', variable=self.Cvars[12])
        self.Cbox12.grid_forget()

        self.Cbox13 = tk.Checkbutton(self.Boxes_Frame)
        self.Cbox13.grid(row=4, column=0, sticky='W')
        self.Cbox13.configure(justify='left')
        self.Cbox13.configure(text='''Адрес''', variable=self.Cvars[13])
        self.Cbox13.grid_forget()

        self.Cbox14 = tk.Checkbutton(self.Boxes_Frame)
        self.Cbox14.grid(row=0, column=0, sticky='W')
        self.Cbox14.configure(justify='left')
        self.Cbox14.configure(text='''Название''', variable=self.Cvars[14])
        self.Cbox14.grid_forget()

        self.Cbox15 = tk.Checkbutton(self.Boxes_Frame)
        self.Cbox15.grid(row=1, column=0, sticky='W')
        self.Cbox15.configure(justify='left')
        self.Cbox15.configure(text='''Телефон''', variable=self.Cvars[15])
        self.Cbox15.grid_forget()

        # menu
        menubar = tk.Menu(top)
        filemenu = tk.Menu(menubar, tearoff=0)
        filemenu.add_command(label="Новый", command=self.newDatabase)
        filemenu.add_command(label="Открыть", command=self.open)
        filemenu.add_command(label="Сохранить", command=self.save)
        filemenu.add_command(label="Сохранить как...", command=self.saveas)
        filemenu.add_command(label="Экспорт", command=self.saveAsExcel)
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

    def saveAsExcel(self):
        saveAsExcel(self.tables[self.Data.index("current")], self.Data.index("current"))

    def moveSelection1(self, event):
        global select
        select = list(self.Filter_List2.curselection())
        self.Filter_List1.select_clear(0, 'end')
        self.Filter_List1.selection_set(select[0])
        self.Filter_List1.select_anchor(select[0])

    def moveSelection2(self, event):
        global select
        select = list(self.Filter_List1.curselection())
        self.Filter_List2.select_clear(0, 'end')
        self.Filter_List2.selection_set(select[0])
        self.Filter_List2.select_anchor(select[0])

    def scrollList1(self, event):
        self.Filter_List1.yview_scroll(int(-4*(event.delta/120)), "units")

    def scrollList2(self, event):
        self.Filter_List2.yview_scroll(int(-4*(event.delta/120)), "units")

    def tabChoice(self, event):
        global selected_tab
        selected_tab = event.widget.select()
        if event.widget.index(selected_tab) == 0:
            self.parInsert(0)
            self.insertCheckBoxes0()
        elif event.widget.index(selected_tab) == 1:
            self.parInsert(1)
            self.insertCheckBoxes1()
        elif event.widget.index(selected_tab) == 2:
            self.parInsert(2)
            self.insertCheckBoxes2()
        elif event.widget.index(selected_tab) == 3:
            self.parInsert(3)
            self.insertCheckBoxes3()
        else:
            self.parInsert(4)
            self.insertCheckBoxes4()

    def hideCheckBox(self, CheckBox):
        CheckBox.grid_forget()

    def hideAll(self):
        self.hideCheckBox(self.Cbox0)
        self.hideCheckBox(self.Cbox1)
        self.hideCheckBox(self.Cbox2)
        self.hideCheckBox(self.Cbox3)
        self.hideCheckBox(self.Cbox4)
        self.hideCheckBox(self.Cbox5)
        self.hideCheckBox(self.Cbox6)
        self.hideCheckBox(self.Cbox7)
        self.hideCheckBox(self.Cbox8)
        self.hideCheckBox(self.Cbox9)
        self.hideCheckBox(self.Cbox10)
        self.hideCheckBox(self.Cbox11)
        self.hideCheckBox(self.Cbox12)
        self.hideCheckBox(self.Cbox13)
        self.hideCheckBox(self.Cbox14)
        self.hideCheckBox(self.Cbox15)

    def insertCheckBoxes0(self):
        self.hideAll()
        self.Cbox0.grid(row=0, column=0, sticky='W')
        self.Cbox1.grid(row=1, column=0, sticky='W')
        self.Cbox2.grid(row=2, column=0, sticky='W')
        self.Cbox3.grid(row=3, column=0, sticky='W')

    def insertCheckBoxes1(self):
        self.hideAll()
        self.Cbox4.grid(row=0, column=0, sticky='W')
        self.Cbox5.grid(row=1, column=0, sticky='W')

    def insertCheckBoxes2(self):
        self.hideAll()
        self.Cbox6.grid(row=0, column=0, sticky='W')
        self.Cbox7.grid(row=1, column=0, sticky='W')
        self.Cbox8.grid(row=2, column=0, sticky='W')

    def insertCheckBoxes3(self):
        self.hideAll()
        self.Cbox9.grid(row=0, column=0, sticky='W')
        self.Cbox10.grid(row=1, column=0, sticky='W')
        self.Cbox11.grid(row=2, column=0, sticky='W')
        self.Cbox12.grid(row=3, column=0, sticky='W')
        self.Cbox13.grid(row=4, column=0, sticky='W')

    def insertCheckBoxes4(self):
        self.hideAll()
        self.Cbox14.grid(row=0, column=0, sticky='W')
        self.Cbox15.grid(row=1, column=0, sticky='W')

    def parInsert(self, tab):
        self.Filter_List1.delete(0, 'end')
        self.Filter_List2.delete(0, 'end')
        cols = list(self.db[tab].columns)
        for i in range(len(cols)-1):
            self.Filter_List1.insert('end', cols[i+1])
            self.Filter_List2.insert('end', "")

    def open_dialog(self):
        global newPar, select
        if len(select) != 0:
            newPar = ChangeDialog(root, "Введите новое значение:").show()
            self.Filter_List2.delete(select[0])
            self.Filter_List2.insert(select[0], newPar)
            self.Filter_List2.selection_set(select[0])
            self.Filter_List2.select_anchor(select[0])
        else:
            message(root, "Не выбран элемент").fade()

    def filterTable(self):
        global selcted_tab
        filters = []
        tab = self.Data.index(selected_tab)
        cols = list(self.db[tab].columns)
        cols = cols[1:]
        df = self.db[tab]
        df.index = np.arange(len(df))
        check = True
        for i in range(len(cols)):
            filters.append(self.Filter_List2.get(i))
        for fil in filters:
            if fil != "":
                check = False
        if check:
            for i in self.tables[tab].get_children():
                self.tables[tab].delete(i)
            for j in self.db[tab].index:
                items = []
                for title in self.db[tab].columns:
                    items.append(self.db[tab][title][j])
                self.tables[tab].add("", values=items)
        else:
            for i in range(len(filters)):
                if filters[i] != "":
                    name = df.columns[i+1]
                    if (filters[i].isdigit()):
                        df = df.drop(np.where(df[name] != int(filters[i]))[0])
                        df.index = np.arange(len(df))
                    else:
                        df = df.drop(np.where(df[name] != filters[i])[0])
            for i in self.tables[tab].get_children():
                self.tables[tab].delete(i)
            for j in df.index:
                items = []
                for title in df.columns:
                    items.append(df[title][j])
                self.tables[tab].add("", values=items)

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
                self.tables[i].add("", values=items)

    def exit(self):
        if MainWindow.modified:
            ans = tk.messagebox.askyesnocancel("Несохраненные изменения", "Хотите сохранить изменения перед закрытием?")
            if ans:
                self.save()
            elif ans is None:
                return
        root.destroy()

    def open(self):
        if MainWindow.modified:
            ans = tk.messagebox.askyesnocancel("Несохраненные изменения", "Хотите сохранить изменения перед закрытием?")
            if ans:
                self.save()
            elif ans is None:
                return
        file = filedialog.askopenfilename(filetypes=[("pickle files", "*.pickle"), ("Excel files", "*.xls *.xlsx")])
        openFromFile(file)
        self.loadTables()

    def save(self):
        if (MainWindow.currentFile != ''):
            saveToPickle(MainWindow.currentFile, MainWindow.db)
        else:
            self.saveas()

    def saveas(self):
        filename = filedialog.asksaveasfilename(filetypes=[], defaultextension=".pickle")
        MainWindow.currentFile = filename
        MainWindow.modified = False
        saveToPickle(filename, MainWindow.db)

    def statusUpdate(self, event=None):
        curTable = self.tables[self.Data.index(self.Data.select())]
        status = "Elements: "
        selected = len(curTable.selection())
        if selected == 0:
            status += str(len(curTable.get_children()))
        else:
            status += ("%d out of %d" % (selected, len(curTable.get_children())))
        self.statusbar.config(text=status)
        self.statusbar.update_idletasks()
        root.after(1, self.statusUpdate)

    def treeSort(self, treeview, col, reverse):
        firstElement = treeview.set(treeview.get_children('')[0], col)
        if self.treeCheckForDigit(firstElement):
            l = [(float(treeview.set(k, col)), k) for k in treeview.get_children('')]
        else:
            l = [(str(treeview.set(k, col)), k) for k in treeview.get_children('')]
        l.sort(reverse=reverse)
        for index, (val, k) in enumerate(l):
            treeview.move(k, '', index)

        if reverse:
            char = '        ▼'
        else:
            char = '        ▲'

        treeview.heading(col, text=col+char, command=lambda: self.treeSort(treeview, col, not reverse))

    def treeCheckForDigit(self, string):
        # print(string, type(string))
        if string.isdigit():
            return True
        else:
            try:
                float(string)
                return True
            except ValueError:
                return False


class CustomDialog(tk.Toplevel):
    # print(CustomDialog(root, "Enter something:").show()) to show
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
        self.globalCounter = 0

    def add(self, parent, values):
        self.insert("", "end", iid=self.globalCounter, values=values)
        self.globalCounter += 1

    def popup(self, event):
        try:
            self.popup_menu.tk_popup(event.x_root, event.y_root)
        finally:
            self.popup_menu.grab_release()

    def selectAll(self, event=None):
        self.selection_set(tuple(self.get_children()))

    def addRecord(self):
        # print(MainWindow.db[0])
        nb = self.master.master
        nb = nb.index(nb.select())
        dic = askValuesDialog(root, MainWindow.db[nb].columns).show()
        values = list(dic.values())
        keys = list(dic.keys())
        if (len(values) and values[0].get() != ''):  # TODO correct input validation
            MainWindow.modified = True
            MainWindow.db[nb] = MainWindow.db[nb].append(
                    pd.DataFrame([[np.int64(item.get()) if item.get().isdigit() else item.get() for item in values]],
                                     columns=keys),
                                   ignore_index=True)
            self.add("", values=[item.get() for item in values])
        # for i in self.get_children():
        # print(i)

    def deleteRecords(self, event=None):
        nb = self.master.master
        nb = nb.index(nb.select())
        selected = [int(i) for i in self.selection()]
        if not len(selected):
            message(root, "Не выбран элемент").fade()
        else:
            MainWindow.modified = True
            for item in selected:
                itemId = int(self.item(item)['values'][0])
                MainWindow.db[nb] = MainWindow.db[nb].drop(MainWindow.db[nb].index[MainWindow.db[nb]['Код'] == itemId])
                self.delete(self.selection()[0])
            # print(MainWindow.db[nb]['Код'])

    def modRecord(self):
        nb = self.master.master
        nb = nb.index(nb.select())
        selected = self.selection()
        if not selected:
            message(root, "Не выбран элемент").fade()
        else:
            selected = int(selected[0])
            itemId = np.int64(self.item(selected)['values'][0])
            print(itemId)
            itemValues = MainWindow.db[nb][MainWindow.db[nb]['Код'] == itemId].values[0].tolist()
            dic = askValuesDialog(root, MainWindow.db[nb].columns, currValues=itemValues).show()
            keys = list(dic.keys())
            values = list(dic.values())
            if (len(values) and values[0].get() != ''):  # TODO correct input validation
                MainWindow.modified = True
                for i in range(len(keys)):
                    self.item(selected, values=[item.get() for item in values])
                    MainWindow.db[nb].loc[itemId-1, keys[i]] = values[i].get()


if __name__ == '__main__':
    start_gui()
