import sys
import tkinter as tk
import tkinter.ttk as ttk
import pandas as pd
import pickle as pk
import numpy as np
import matplotlib.pyplot as plt
from tkinter import filedialog
from tkinter import messagebox as mb
sys.path.insert(1, '../Library')
from funcs import *


newPar = ""
select = []
selected_tab = 0


quantParams = [{"Код", "Сумма", "Код работника", "Дата выплаты"},
               {"Код", "Код должности", "Отделение"},
               {"Код", "Норма (ч)", "Ставка (ч)"},
               {"Код", "Номер договора"},
               {"Код"}]


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


def saveAsExcel(tree):
    file = filedialog.asksaveasfilename(title="Select file", initialdir='../Data/db1.xlsx', defaultextension=".xlsx", filetypes=[("Excel file", "*.xlsx")])
    if file:
        ids=tree.get_children()
        #dic = dict([tree.column(i)['id'] for i in tree["displaycolumns"]]) # TODO need to get displayed columns only
        dic = dict.fromkeys(tree["columns"], [])
        keys = list(dic.keys())
        for i in range(len(keys)):
            dic[keys[i]] = []
        for iid in ids:
            for i in range(len(keys)):
                dic[keys[i]].append(tree.item(iid)["values"][i])

        dic = pd.DataFrame.from_dict(dic)
        try:
           dic.to_excel(file, engine='xlsxwriter',index=False)
           message(root, "Таблица сохранена", msgtype="success").fade()
        except:
           message(root, "Не удалось сохранить файл!\nВозможно, он открыт\nв другой программе", msgtype="error").fade()
    else:
        pass # pressed cancel


def configureGUI(scr, top):
    configureWidgets(scr, top)

    # configure tables
    scr.tabs = [scr.Data_t1, scr.Data_t2, scr.Data_t3,
            scr.Data_t4, scr.Data_t5]

    scr.tables = [None] * 5
    for i in range(len(MainWindow.db)):
        scr.tables[i] = TreeViewWithPopup(scr.tabs[i])
        scr.tables[i].place(relwidth=1.0, relheight=1.0)
        scr.tables[i]["columns"] = list(MainWindow.db[i].columns)
        scr.tables[i]['show'] = 'headings'
        columns = list(MainWindow.db[i].columns)
        scr.tables[i].column("#0", minwidth=5, width=5, stretch=tk.NO)
        scr.tables[i].heading("#0", text="")

        for j in range(len(columns)):
            scr.tables[i].heading(columns[j], text=columns[j]+'       ▼▲',\
                       command= lambda _treeview = scr.tables[i], _col=columns[j]:scr.treeSort(_treeview, _col, False))
            scr.Data.update()
            width = int((scr.Data.winfo_width()-30)/(len(columns)-1))
            scr.tables[i].column(columns[j], width=width, stretch=tk.NO)

        scr.tables[i].column(columns[0], width=30, stretch=tk.NO)

        for j in MainWindow.db[i].index:
            items = []
            for title in MainWindow.db[i].columns:
                items.append(MainWindow.db[i][title][j])
            scr.tables[i].add("", values=items)

    # configure scrolls
    scr.scrolls = [None] * 5
    for i in range(len(scr.tabs)):
        scr.scrolls[i] = ttk.Scrollbar(scr.tables[i], orient="vertical",
                                        command=scr.tables[i].yview)
        scr.tables[i].config(yscrollcommand=scr.scrolls[i].set)
        scr.scrolls[i].pack(fill="y", side='right')
        scr.scrolls[i] = ttk.Scrollbar(scr.tables[i], orient="horizontal",
                                        command=scr.tables[i].xview)
        scr.tables[i].config(xscrollcommand=scr.scrolls[i].set)
        scr.scrolls[i].pack(fill="x", side='bottom')

    # binds
    scr.Data.bind("<<NotebookTabChanged>>", scr.tabChoice)

    scr.Filter_List1.bind("<<ListboxSelect>>", scr.moveSelection2)
    scr.Filter_List2.bind("<<ListboxSelect>>", scr.moveSelection1)

    root.protocol("WM_DELETE_WINDOW", scr.exit)
    root.bind("<Control-a>", scr.selectAll)
    root.bind("<Control-n>", scr.addRecord)
    root.bind("<Delete>", scr.deleteRecords)
    root.bind("<Button-1>", scr.statusUpdate)

    scr.ComboAnalysis.bind("<<ComboboxSelected>>", scr.configAnalysisCombos)

    # start status bar
    scr.statusUpdate()

    scr.ComboAnalysis.current(2)
    scr.configAnalysisCombos()
    scr.updateCombos()


class MainWindow:
    db = None
    currentFile = ''
    modified = False
    col = []

    def __init__(self, top=None):
        """This class configures and populates the toplevel window.
           top is the toplevel containing window."""
        # refreshFromExcel("../Data/db.xlsx")  # use once for db.pickle
        openFromFile("../Data/db.pickle")

        top.geometry("1000x600+150+30")
        top.minsize(width=1000, height=600)
        top.title("База Данных")

        configureGUI(self, top)

    def configAnalysisCombos(self, event=None):
        anId = self.ComboAnalysis.current()
        if anId == 0:
            self.ComboQuant.configure(state="disabled")
            self.ComboQual.configure(state="normal")
        elif anId == 1:
            self.ComboQuant.configure(state="normal")
            self.ComboQual.configure(state="disabled")
        else:
            self.ComboQuant.configure(state="normal")
            self.ComboQual.configure(state="normal")
        if anId == 5:
            self.LabelQual.configure(text="Количественный")
        else:
            self.LabelQual.configure(text="Качественный")

    def paramsValid(self):
        return (self.ComboAnalysis.current() == -1 or \
            self.ComboQual.current() == -1 or \
            self.ComboQuant.current() == -1)

    def showReport(self):
        if self.paramsValid():
            message(root, "Не выбран элемент", msgtype="warning").fade()
            return
        nb = self.Data.index(self.Data.select())
        df = MainWindow.db[nb]
        if self.ComboAnalysis.current() == 2:
            plot, file = getBar(self, df)
        elif self.ComboAnalysis.current() == 3: # add analysis here
            plot, file = getHist(root, self, MainWindow.db)
        elif self.ComboAnalysis.current() == 4:
            plot, file = getBoxWhisker(root, self, MainWindow.db)
        if file and plot:
            plot.show()
        else:
            message(root, "Не удалось построить диаграмму,\nпопробуйте выбрать\n"+
                    "другие данные", msgtype="error").fade()

    def exportReport(self):
        if self.paramsValid():
            message(root, "Не выбран элемент", msgtype="warning").fade()
            return
        nb = self.Data.index(self.Data.select())
        df = MainWindow.db[nb]
        pltType = 'plot'
        if self.ComboAnalysis.current() == 2:
            plot, file = getBar(self, df)
        elif self.ComboAnalysis.current() == 3: # add analysis here
            plot, file = getHist(self, df)
        elif self.ComboAnalysis.current() == 4:
            plot, file = getBoxWhisker(root, self, fdf)
            if (plot is None and file is None):
                return
        plot.savefig(file)

    def saveAsExcel(self):
        saveAsExcel(self.tables[self.Data.index("current")])

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

    def updateCombos(self):
        pass
        #self.ComboQuant.set('')
        #self.ComboQual.set('')
        #nb = self.Data.index(self.Data.select())
        #self.ComboQuant.configure(values = [h for h in MainWindow.db[nb].columns if h in quantParams[nb]])
        #self.ComboQual.configure(values = [h for h in MainWindow.db[nb].columns if not h in quantParams[nb]])
        #if len(self.ComboQuant["values"]) == 0:
        #    self.ComboQuant.configure(state="disabled")
        #else:
        #    self.ComboQuant.configure(state="normal")
        #if len(self.ComboQual["values"]) == 0:
        #    self.ComboQual.configure(state="disabled")
        #else:
        #    self.ComboQual.configure(state="normal")

    def tabChoice(self, event):
        global selected_tab
        selected_tab = event.widget.select()
        if event.widget.index(selected_tab) == 0:
            self.parInsert(0)
            self.insertCheckBoxes(0)
        elif event.widget.index(selected_tab) == 1:
            self.parInsert(1)
            self.insertCheckBoxes(1)
        elif event.widget.index(selected_tab) == 2:
            self.parInsert(2)
            self.insertCheckBoxes(2)
        elif event.widget.index(selected_tab) == 3:
            self.parInsert(3)
            self.insertCheckBoxes(3)
        else:
            self.parInsert(4)
            self.insertCheckBoxes(4)

    def hideAll(self):
        for i in self.Cboxes:
            for j in i:
                j.grid_forget()

    def insertCheckBoxes(self, tab):
        self.hideAll()
        for i in range(len(self.Cboxes[tab])):
            self.Cboxes[tab][i].grid(row=i, column=0, sticky='W')

    def removeColumns(self):
        global selected_tab
        tab = self.Data.index(selected_tab)
        indTab = 0
        for i in range(tab):
            indTab += len(self.Cvars[i])
        exclude = []
        for i in range(len(self.Cvars[tab])):
            ind = indTab + i
            if self.Cvars[tab][i].get() is False:
                exclude.append(self.names[ind])
        display = []
        for col in self.tables[tab]["columns"]:
            if col not in exclude:
                display.append(col)
        self.tables[tab]["displaycolumns"] = (display)

    def reset(self):
        global selected_tab
        tab = self.Data.index(selected_tab)
        for box in self.Cboxes[tab]:
            box.select()
        exclude = []
        display = []
        for col in self.tables[tab]["columns"]:
            if col not in exclude:
                display.append(col)
        self.tables[tab]["displaycolumns"] = (display)

        self.Filter_List2.delete(0, 'end')
        for i in range(4):
            self.Filter_List2.insert('end', "")
        for i in self.tables[tab].get_children():
            self.tables[tab].delete(i)
        for j in MainWindow.db[tab].index:
            items = []
            for title in MainWindow.db[tab].columns:
                items.append(MainWindow.db[tab][title][j])
            self.tables[tab].add("", values=items)

        self.Filter_List2.selection_set(select[0])
        self.Filter_List2.select_anchor(select[0])

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
            self.filterTable()
        else:
            message(root, "Не выбран элемент", msgtype="warning").fade()

    def filterTable(self):
        global selcted_tab
        filters = []
        tab = self.Data.index(selected_tab)
        cols = list(MainWindow.db[tab].columns)
        cols = cols[1:]
        df = MainWindow.db[tab]
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
            for j in MainWindow.db[tab].index:
                items = []
                for title in MainWindow.db[tab].columns:
                    items.append(MainWindow.db[tab][title][j])
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
        exit()

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
        #root.after(100, self.statusUpdate)

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

    def genUID(self):
        uid = 1
        ids = self.get_children()
        ids = set([int(self.item(item)["values"][0]) for item in ids])
        while uid in ids:
            uid += 1
        return uid


    def addRecord(self):
        nb = self.master.master
        nb = nb.index(nb.select())
        dic = askValuesDialog(root, MainWindow.db[nb].columns).show()
        values = list(dic.values())
        keys = list(dic.keys())
        if (len(values)):
            values = [item.get() for item in values]
            values[0] = str(self.genUID())
            MainWindow.modified = True

            MainWindow.db[nb] = MainWindow.db[nb].append(
                    pd.DataFrame([[np.int64(item) if item.isdigit() else item for item in values]],
                                     columns=keys),
                                   ignore_index=True)
            self.add("", values=values)

    def deleteRecords(self, event=None):
        nb = self.master.master
        nb = nb.index(nb.select())
        selected = [int(i) for i in self.selection()]
        if not len(selected):
            message(root, "Не выбран элемент", msgtype="warning").fade()
        else:
            MainWindow.modified = True
            for item in selected:
                itemId = int(self.item(item)['values'][0])
                MainWindow.db[nb] = MainWindow.db[nb].drop(MainWindow.db[nb].index[MainWindow.db[nb]['Код'] == itemId])
                self.delete(self.selection()[0])

    def modRecord(self):
        nb = self.master.master
        nb = nb.index(nb.select())
        selected = self.selection()
        if not selected:
            message(root, "Не выбран элемент", msgtype="warning").fade()
        else:
            selected = int(selected[0])
            itemId = np.int64(self.item(selected)['values'][0])
            itemValues = MainWindow.db[nb][MainWindow.db[nb]['Код'] == itemId].values[0].tolist()
            dic = askValuesDialog(root, MainWindow.db[nb].columns, currValues=itemValues).show()
            keys = list(dic.keys())
            values = list(dic.values())
            if (len(values)):
                values = [item.get() for item in values]
                values[0] = itemId
                MainWindow.modified = True
                for i in range(len(keys)):
                    self.item(selected, values=values)
                    MainWindow.db[nb].loc[itemId-1, keys[i]] = values[i]


if __name__ == '__main__':
    start_gui()
