"""
В этом модуле описаны функции, специфичные для проекта
"""

import sys
import tkinter as tk
import tkinter.ttk as ttk
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from tkinter import filedialog
sys.path.insert(1, '../Library')
import funcs


newPar = ""
select = []
selected_tab = 0


quantParams = [{"Код", "Сумма", "Код работника", "Дата выплаты"},
               {"Код", "Код должности", "Отделение"},
               {"Код", "Норма (ч/мес)", "Ставка (ч)"},
               {"Код", "Номер договора"},
               {"Код"}]


class DB:
    """
    Данные об используемой базе данных

    :Автор(ы): Константинов
    """
    db = None
    currentFile = ''
    modified = False


def createEmptyDatabase():
    """
    Создание пустой базы данных

    :return: Объект базы данных
    :rtype: :class:`pandas.DataFrame`
    :return: Состояние базы данных (изменена)
    :rtype: :class:`boolean`
    :return: Путь к текущему файлу для сохранения
    :rtype: :class:`string`

    :Автор(ы): Константинов
    """
    db = [pd.DataFrame(columns=['Код', 'Отработано (ч)', 'Дата выплаты',
                                'Сумма', 'Код работника']),
          pd.DataFrame(columns=['Код', 'Код должности',
                                'Отделение']),
          pd.DataFrame(columns=['Код', 'Название', 'Норма (ч/мес)',
                                'Ставка (ч)']),
          pd.DataFrame(columns=['Код', 'ФИО', 'Номер договора',
                                'Телефон', 'Образование', 'Адрес']),
          pd.DataFrame(columns=['Код', 'Название',
                                'Телефон'])]
    modified = False
    currentFile = ''
    return db, modified, currentFile


def configureGUI(scr, top):
    """
    Создание уникального для директории имени файла в формате `spec_spec_spec_UID.ext`

    :param scr: Объект окна
    :type scr: MainWindow
    :param top: Корневой объект
    :type top: tk.Tk
    :Автор(ы): Константинов, Сидоров
    """
    winW = scr.root.winfo_screenwidth()
    scr.config["def_window_width"] = int(scr.config["def_window_width"])
    scr.config["def_window_height"] = int(scr.config["def_window_height"])
    dispX = str(int(winW / 2 - scr.config["def_window_width"] / 2)) # center window
    dispY = '30'
    geom = str(scr.config["def_window_width"])+'x'+str(scr.config["def_window_height"])+'+'+dispX+'+'+dispY
    scr.root.geometry(geom)
    scr.root.minsize(width=1000, height=600)
    if int(scr.config["maximize"]):
        scr.root.state('zoomed')
    scr.root.attributes('-fullscreen', scr.config["fullscreen"])
    scr.root.title("База Данных")
    scr.root.configure(background=scr.config["def_bg_color"])
    funcs.configureWidgets(scr, top)

    # configure tables
    scr.tabs = [scr.Data_t1, scr.Data_t2, scr.Data_t3,
                scr.Data_t4, scr.Data_t5]

    scr.tables = [None] * 5
    for i in range(len(DB.db)):
        scr.tables[i] = TreeViewWithPopup(scr.tabs[i], scr.config)
        scr.tables[i].place(relwidth=1.0, relheight=1.0)
        scr.tables[i]["columns"] = list(DB.db[i].columns)
        scr.tables[i]['show'] = 'headings'
        columns = list(DB.db[i].columns)
        scr.tables[i].column("#0", minwidth=5, width=5, stretch=tk.NO)
        scr.tables[i].heading("#0", text="")

        for j in range(len(columns)):
            scr.tables[i].heading(columns[j], text=columns[j]+'       ▼▲',
                                  command=lambda _treeview=scr.tables[i],
                                  _col=columns[j]: scr.treeSort(_treeview, _col,
                                                               False))
            scr.Data.update()
            width = int((scr.Data.winfo_width()-30)/(len(columns)-1))
            scr.tables[i].column(columns[j], width=width, stretch=tk.NO)

        scr.tables[i].column(columns[0], width=30, stretch=tk.NO)

        for j in DB.db[i].index:
            items = []
            for title in DB.db[i].columns:
                items.append(DB.db[i][title][j])
            scr.tables[i].add("", values=items)

    # configure scrolls
    scr.scrolls = [None] * 5
    for i in range(len(scr.tabs)):
        scr.scrolls[i] = ttk.Scrollbar(scr.tables[i], orient="vertical",
                                       command=scr.tables[i].yview)
        scr.tables[i].configure(yscrollcommand=scr.scrolls[i].set)
        scr.scrolls[i].pack(fill="y", side='right')
        scr.scrolls[i] = ttk.Scrollbar(scr.tables[i], orient="horizontal",
                                       command=scr.tables[i].xview)
        scr.tables[i].configure(xscrollcommand=scr.scrolls[i].set)
        scr.scrolls[i].pack(fill="x", side='bottom')

    # binds
    scr.Data.bind("<<NotebookTabChanged>>", scr.tabChoice)

    scr.Filter_List1.bind("<<ListboxSelect>>", scr.moveSelection2)
    scr.Filter_List2.bind("<<ListboxSelect>>", scr.moveSelection1)

    top.protocol("WM_DELETE_WINDOW", scr.exit)
    top.bind("<Control-Shift-N>", scr.addRecord)
    top.bind("<Control-n>", scr.newDatabase)
    top.bind("<Control-o>", scr.open)
    top.bind("<Control-s>", scr.save)
    top.bind("<Control-Shift-S>", scr.saveas)
    top.bind("<Control-e>", scr.saveAsExcel)
    top.bind("<Control-a>", scr.selectAll)
    top.bind("<Control-Shift-A>", scr.addRecord)
    top.bind("<Control-o>", scr.open)
    top.bind("<Control-q>", scr.exit)
    top.bind("<Control-r>", scr.modRecord)
    top.bind("<Delete>", scr.deleteRecords)
    top.bind("<Control-p>", scr.customizeGUI)
    top.bind("<Button-1>", scr.statusUpdate)

    scr.ComboAnalysis.bind("<<ComboboxSelected>>", scr.configAnalysisCombos)

    # start status bar
    scr.statusUpdate()

    scr.ComboAnalysis.current(2)
    scr.configAnalysisCombos()


class MainWindow:
    """
    Класс главного окна приложения

    :Автор(ы): Константинов, Сидоров, Березуцкий
    """
    def __init__(self, root=None):
        """
        Инициализация

        :param root: Корневой объект
        :type root: tk.Tk
        :Автор(ы): Константинов
        """
        self.root = root
        self.root.focus_force()
        DB.db, DB.modified, DB.currentFile = funcs.openFromFile("../Data/db.pickle",
                                                          DB.db, DB.modified,
                                                          DB.currentFile,
                                                          createEmptyDatabase)
        self.config = funcs.getConfig()
        configureGUI(self, self.root)
        funcs.message(self.root, "Документацию и руководство\nпользователя можно найти\nв каталоге Notes", msgtype="info").fade()
        
        self.updateTitle()

    def updateTitle(self):
        """
        Обновление заголовка окна (состояние текущего файла)

        :Автор(ы): Константинов
        """
        title = "База данных"
        if DB.currentFile != '' or DB.modified:
            title += ' - '
            if DB.currentFile != '':
                filename = DB.currentFile[DB.currentFile.rfind('/')+1:]
                title += filename
            if DB.modified:
                title += '*'
        self.root.title(title)

    def configAnalysisCombos(self, event=None):
        """
        Настройка меню выбора атрибутов в зависимости от выбранного вида отчета
        
        :param event: объект события
        :Автор(ы): Константинов
        """
        anId = self.ComboAnalysis.current()
        self.ComboQuant.set('')
        self.ComboQual.set('')
        if anId == 0:
            self.ComboQuant.configure(state="disabled")
            self.ComboQual.configure(state="normal")
        elif anId == 1:
            self.ComboQuant.configure(state="disabled")
            self.ComboQual.configure(state="disabled")
        else:
            self.ComboQuant.configure(state="normal")
            self.ComboQual.configure(state="normal")
        if anId == 5:
            self.ComboQuant2.configure(state="normal")
            self.ComboQuant.configure(values=funcs.quantComboValues)
        elif anId == 2:
            self.ComboQuant.configure(values=funcs.qualComboValues)
            self.LabelQuant.configure(text="Качественный")
        else:
            self.ComboQuant2.configure(state="disabled")
            self.LabelQual.configure(text="Качественный")
            self.LabelQuant.configure(text="Количественный")
            self.ComboQuant["values"] = (funcs.quantComboValues)

    def paramsValid(self):
        """
        Проверка факта выбора в трех выпадающих меню одновременно: выбор вида
        отчета и двух параметров

        :return: во всех трех меню выбран один из вариантов
        :rtype: bool
        :Автор(ы): Константинов, Березуцкий
        """
        return (self.ComboAnalysis.current() == -1) or \
                (self.ComboQuant.current() == -1 and self.ComboQuant['state'].string == "normal") \
                or (self.ComboQual.current() == -1 and self.ComboQual['state'].string == "normal")

    def showReport(self):
        """
        Отображение выбранного отчета

        :Автор(ы): Константинов, Сидоров, Березуцкий
        """
        if self.paramsValid():
            funcs.message(self.root, "Не выбран элемент", msgtype="warning").fade()
            return
        if self.ComboAnalysis.current() == 0:
            plot, file = funcs.getQualityStatistics(self.root, self, DB.db,
                                self.config["def_graph_dir"])
        elif self.ComboAnalysis.current() == 1:
            plot, file = funcs.getQuantStatistics(self.root, self, DB.db,
                                self.config["def_graph_dir"])
        elif self.ComboAnalysis.current() == 2:
            plot, file = funcs.getBar(self.root, self, DB.db,
                                self.config["def_graph_dir"])
        elif self.ComboAnalysis.current() == 3:  # add analysis here
            plot, file = funcs.getHist(self.root, self, DB.db,
                                 self.config["def_graph_dir"])
        elif self.ComboAnalysis.current() == 4:
            plot, file = funcs.getBoxWhisker(self.root, self, DB.db,
                                       self.config["def_graph_dir"])
        elif self.ComboAnalysis.current() == 5:
            plot, file = funcs.getScatterplot(self.root, self, DB.db,
                                        self.config["def_graph_dir"])
        elif self.ComboAnalysis.current() == 6:
            plot, file = funcs.getPivotStatistics(self.root, self, DB.db,
                                        self.config["def_graph_dir"])
        if file and plot:
            plt.show(plot)
        else:
            funcs.message(self.root,
                    "Не удалось построить диаграмму,\nпопробуйте выбрать\n" +
                    "другие данные", msgtype="error").fade()

    def exportReport(self):
        """
        Сохранение выбранного отчета в файл

        :Автор(ы): Константинов, Сидоров, Березуцкий
        """
        if self.paramsValid():
            funcs.message(self.root, "Не выбран элемент", msgtype="warning").fade()
            return
        if self.ComboAnalysis.current() == 0:
            plot, file = funcs.getQualityStatistics(self.root, self, DB.db,
                                self.config["def_graph_dir"])
        elif self.ComboAnalysis.current() == 1:
            plot, file = funcs.getQuantStatistics(self.root, self, DB.db,
                                self.config["def_graph_dir"])
        elif self.ComboAnalysis.current() == 2:
            plot, file = funcs.getBar(self.root, self, DB.db,
                                self.config["def_graph_dir"])
        elif self.ComboAnalysis.current() == 3:  # add analysis here
            plot, file = funcs.getHist(self.root, self, DB.db,
                                 self.config["def_graph_dir"])
        elif self.ComboAnalysis.current() == 4:
            plot, file = funcs.getBoxWhisker(self.root, self, DB.db,
                                       self.config["def_graph_dir"])
        elif self.ComboAnalysis.current() == 5:
            plot, file = funcs.getScatterplot(self.root, self, DB.db,
                                        self.config["def_graph_dir"])
        elif self.ComboAnalysis.current() == 6:
            plot, file = funcs.getPivotStatistics(self.root, self, DB.db,
                                        self.config["def_graph_dir"])
        if file and plot:
            plot.savefig(file)
            funcs.message(self.root, "Файл сохранён", msgtype="success").fade()  # TODO show path
            # need to change label to text in message
        else:
            funcs.message(self.root,
                    "Не удалось построить диаграмму,\nпопробуйте выбрать\n" +
                    "другие данные", msgtype="error").fade()

    def saveAsExcel(self, event=None):
        """
        Сохранение текущей таблицы в файл .xlsx
        
        :param event: объект события
        :Автор(ы): Константинов
        """
        funcs.saveAsExcel(self.root, self.tables[self.Data.index("current")])

    def moveSelection1(self, event):
        """
        Синхронное передвижение выделения в Filter_List1

        :Автор(ы): Сидоров
        """
        global select
        select = list(self.Filter_List2.curselection())
        self.Filter_List1.select_clear(0, 'end')
        self.Filter_List1.selection_set(select[0])
        self.Filter_List1.select_anchor(select[0])

    def moveSelection2(self, event):
        """
        Синхронное передвижение выделения в Filter_List2

        :Автор(ы): Сидоров
        """
        global select
        select = list(self.Filter_List1.curselection())
        self.Filter_List2.select_clear(0, 'end')
        self.Filter_List2.selection_set(select[0])
        self.Filter_List2.select_anchor(select[0])

    def scrollList1(self, event):
        """
        Синхронное передвижение скроллбара в Filter_List1

        :Автор(ы): Сидоров
        """
        self.Filter_List1.yview_scroll(int(-4*(event.delta/120)), "units")

    def scrollList2(self, event):
        """
        Синхронное передвижение скроллбара в Filter_List2

        :Автор(ы): Сидоров
        """
        self.Filter_List2.yview_scroll(int(-4*(event.delta/120)), "units")

    def updateCombos(self):
        """
        Вроде как бесполезная, надо не забыть удалить перед релизом

        :Автор(ы): Константинов
        """
        pass
#        self.ComboQuant.set('')
#        self.ComboQual.set('')
#        if len(self.ComboQuant["values"]) == 0:
#            self.ComboQuant.configure(state="disabled")
#        else:
#            self.ComboQuant.configure(state="normal")
#        if len(self.ComboQual["values"]) == 0:
#            self.ComboQual.configure(state="disabled")
#        else:
#            self.ComboQual.configure(state="normal")

    def tabChoice(self, event):
        """
        Отображение checkbox'ов и фильтров в зависимости от таблицы

        :Автор(ы): Сидоров
        """
        global selected_tab
        selected_tab = event.widget.select()
        tab = event.widget.index(selected_tab)
        for i in self.tables[tab].get_children():
            self.tables[tab].delete(i)
        for j in DB.db[tab].index:
            items = []
            for title in DB.db[tab].columns:
                items.append(DB.db[tab][title][j])
            self.tables[tab].add("", values=items)
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
        """
        Скрывает ненужные checbox'ы

        :Автор(ы): Сидоров
        """
        for i in self.Cboxes:
            for j in i:
                j.grid_forget()

    def insertCheckBoxes(self, tab):
        """
        Вставляет нужные checbox'ы

        :param tab: номер таблицы
        :type df: int

        :Автор(ы): Сидоров
        """
        self.hideAll()
        for i in range(len(self.Cboxes[tab])):
            self.Cboxes[tab][i].grid(row=i, column=0, sticky='W')

    def removeColumns(self):
        """
        Скрывает столбцы, выбранные в фильтрах

        :Автор(ы): Сидоров
        """
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
        """
        Сбрасывает все фильры

        :Автор(ы): Сидоров
        """
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
        for j in DB.db[tab].index:
            items = []
            for title in DB.db[tab].columns:
                items.append(DB.db[tab][title][j])
            self.tables[tab].add("", values=items)
        if select != []:
            self.Filter_List1.selection_set(select[0])
            self.Filter_List1.select_anchor(select[0])
            self.Filter_List2.selection_set(select[0])
            self.Filter_List2.select_anchor(select[0])

    def parInsert(self, tab):
        """
        Удаление старых параметров из listbox и вставка новых

        :param tab: номер таблицы
        :type df: int

        :Автор(ы): Сидоров
        """
        self.Filter_List1.delete(0, 'end')
        self.Filter_List2.delete(0, 'end')
        cols = list(DB.db[tab].columns)
        for i in range(len(cols)-1):
            self.Filter_List1.insert('end', cols[i+1])
            self.Filter_List2.insert('end', "")

    def openDialog(self):
        """
        Открывает окно изменения парметра,
        получает значение из этого окна и запускает фильтрацию
        Или выводит сообщение об ошибке.

        :Автор(ы): Сидоров
        """
        global newPar, select
        if len(select) != 0:
            newPar = funcs.ChangeDialog(self.root, self.config,
                                  "Введите новое значение:").show()
            self.Filter_List2.delete(select[0])
            self.Filter_List2.insert(select[0], newPar)
            self.Filter_List2.selection_set(select[0])
            self.Filter_List2.select_anchor(select[0])
            self.filterTable()
        else:
            funcs.message(self.root, "Не выбран элемент", msgtype="warning").fade()

    def filterTable(self):
        """
        Выполняет фильтрацию по заданным значениям параметров

        :Автор(ы): Сидоров
        """
        global selcted_tab
        filters = []
        tab = self.Data.index(selected_tab)
        cols = list(DB.db[tab].columns)
        cols = cols[1:]
        df = DB.db[tab]
        df.index = np.arange(len(df))
        check = True
        # refresh
        for i in self.tables[tab].get_children():
            self.tables[tab].delete(i)
        for j in DB.db[tab].index:
            items = []
            for title in DB.db[tab].columns:
                items.append(DB.db[tab][title][j])
            self.tables[tab].add("", values=items)

        for i in range(len(cols)):
            filters.append(self.Filter_List2.get(i))
        for fil in filters:
            if fil != "":
                check = False
        if check:
            for i in self.tables[tab].get_children():
                self.tables[tab].delete(i)
            for j in DB.db[tab].index:
                items = []
                for title in DB.db[tab].columns:
                    items.append(DB.db[tab][title][j])
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
        """
        Выбор всех строк в текущей таблице

        :param event: объект события
        :Автор(ы): Константинов
        """
        self.tables[self.Data.index("current")].selectAll()

    def modRecord(self, event=None):
        """
        Изменение выбранной записи в текущей таблице

        :param event: объект события
        :Автор(ы): Константинов
        """
        self.tables[self.Data.index("current")].modRecord()
        self.updateTitle()

    def addRecord(self, event=None):
        """
        Добавление записи в текущую таблицу

        :param event: объект события
        :Автор(ы): Константинов
        """
        self.tables[self.Data.index("current")].addRecord()
        self.updateTitle()

    def deleteRecords(self, event=None):
        """
        Удаление выбранных строк текущей таблицы

        :param event: объект события
        :Автор(ы): Константинов
        """
        self.tables[self.Data.index("current")].deleteRecords()
        self.updateTitle()

    def newDatabase(self, event=None):
        """
        Создание пустой базы данных
        
        :param event: объект события
        :Автор(ы): Константинов
        """
        if DB.modified:
            ans = tk.messagebox.askyesnocancel("Несохраненные изменения",
                                               "Хотите сохранить изменения перед закрытием?")
            if ans:
                self.save()
            elif ans is None:
                return
        DB.db, DB.modified, DB.currentFile = createEmptyDatabase()
        self.updateTitle()
        self.loadTables()

    def loadTables(self):
        """
        Загрузка строк в таблицу из объекта pandas.DataFrame

        :Автор(ы): Константинов
        """
        for tree in self.tables:
            for item in tree.get_children():
                tree.delete(item)
        for i in range(len(self.tables)):
            for j in DB.db[i].index:
                items = []
                for title in DB.db[i].columns:
                    items.append(DB.db[i][title][j])
                self.tables[i].add("", values=items)

    def exit(self, event=None):
        """
        Выход из приложения
        
        :param event: объект события
        :Автор(ы): Константинов
        """
        if DB.modified:
            ans = tk.messagebox.askyesnocancel("Несохраненные изменения",
                                               "Хотите сохранить изменения перед закрытием?")
            if ans:
                self.save()
            elif ans is None:
                return
        self.root.destroy()
        exit()

    def open(self, event=None):
        """
        Открытие базы данных из файла .xlsx или бинарного файла pickle
        
        :param event: объект события
        :Автор(ы): Константинов
        """
        if DB.modified:
            ans = tk.messagebox.askyesnocancel("Несохраненные изменения",
                                               "Хотите сохранить изменения перед закрытием?")
            if ans:
                self.save()
            elif ans is None:
                return
        file = filedialog.askopenfilename(filetypes=[("pickle files",
                                                      "*.pickle"),
                                                     ("Excel files",
                                                      "*.xls *.xlsx")])
        DB.db, DB.modified, DB.currentFile = funcs.openFromFile(file, DB.db,
                                                          DB.modified,
                                                          DB.currentFile,
                                                          self.newDatabase)
        self.updateTitle()
        self.loadTables()

    def save(self, event=None):
        """
        Сохранение базы данных в бинарный файл pickle
        
        :param event: объект события
        :Автор(ы): Константинов
        """
        if (DB.currentFile != ''):
            funcs.saveToPickle(DB.currentFile, DB.db)
        else:
            self.saveas()
        self.updateTitle()

    def saveas(self, event=None):
        """
        Сохранение базы данных в новый бинарный файл pickle
        
        :param event: объект события
        :Автор(ы): Константинов
        """
        filename = filedialog.asksaveasfilename(filetypes=[],
                                                defaultextension=".pickle")
        DB.currentFile = filename
        DB.modified = False
        self.updateTitle()
        funcs.saveToPickle(filename, DB.db)

    def statusUpdate(self, event=None):
        """
        Обновление строки состояния
        
        :param event: объект события
        :Автор(ы): Константинов
        """
        curTable = self.tables[self.Data.index(self.Data.select())]
        status = "Elements: "
        selected = len(curTable.selection())
        if selected == 0:
            status += str(len(curTable.get_children()))
        else:
            status += ("%d out of %d" % (selected,
                                         len(curTable.get_children())))
        self.statusbar.config(text=status)
        self.statusbar.update_idletasks()

    def treeSort(self, treeview, col, reverse):
        """
        Сортировка таблиц по вазврастанию/убыванию при нажатии на заголовок вкладки

        :param treeview: таблица для сортировки
        :param col: колонка по которой происходит сортировка
        :param reverse: показатель прямой или обратной сортировки
        :Автор(ы): Березуцкий
        """

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

        treeview.heading(col, text=col+char,
                         command=lambda: self.treeSort(treeview, col,
                                                       not reverse))

    def treeCheckForDigit(self, string):
        """
        Проверка строки на число

        :param string: строка для проверки
        :return: Возвращает True, если string число, иначе False
        :rtype: boolean
        :Автор(ы): Березуцкий
        """
        # print(string, type(string))
        if string.isdigit():
            return True
        else:
            try:
                float(string)
                return True
            except ValueError:
                return False

    def customizeGUI(self, event=None):
        """
        Вызов диалога настройки приложения

        :param event: объект события
        :Автор(ы): Константинов
        """
        funcs.CustomizeGUIDialog(self.root).show()


class TreeViewWithPopup(ttk.Treeview):
    """
    Класс виджета ttk.Treeview с добавленным к нему контекстным меню

    :Автор(ы): Константинов
    """
    def __init__(self, parent, config, *args, **kwargs):
        """
        Инициализация виджета

        :param parent: родительский виджет
        :param config: словарь настроек
        :type config: dict
        :param *args: список неименованных документов
        :param **kwargs: список именованных документов
        :Автор(ы): Константинов
        """
        ttk.Treeview.__init__(self, parent, *args, **kwargs)
        self.config = config
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
        self.root = parent
        self.globalCounter = 0

    def add(self, parent, values):
        """
        Добавление строки

        :param parent: родительский виджет
        :param values: список значений столбцов
        :type values: list
        :Автор(ы): Константинов
        """
        self.insert("", "end", iid=self.globalCounter, values=values)
        self.globalCounter += 1

    def popup(self, event):
        """
        Отображение контекстного меню

        :param event: объект события
        :Автор(ы): Константинов
        """
        try:
            self.popup_menu.tk_popup(event.x_root, event.y_root)
        finally:
            self.popup_menu.grab_release()

    def selectAll(self, event=None):
        """
        Выделение всех строк
        
        :param event: объект события
        :Автор(ы): Константинов
        """
        self.selection_set(tuple(self.get_children()))

    def genUID(self):
        """
        Генерация UID для новой строки строки

        :return: UID
        :rtype: integer
        :Автор(ы): Константинов
        """
        uid = 1
        ids = self.get_children()
        ids = set([int(self.item(item)["values"][0]) for item in ids])
        while uid in ids:
            uid += 1
        return uid


    def addRecord(self):
        """
        Добавление новой строки

        :Автор(ы): Константинов
        """
        nb = self.master.master
        nb = nb.index(nb.select())
        dic = funcs.askValuesDialog(self.root, self.config, DB.db[nb].columns).show()
        values = list(dic.values())
        keys = list(dic.keys())
        if (len(values)):
            values = [item.get() for item in values]
            values[0] = str(self.genUID())
            DB.modified = True

            DB.db[nb] = DB.db[nb].append(
                    pd.DataFrame([[np.int64(item) if item.isdigit() else item for item in values]],
                                 columns=keys), ignore_index=True)
            self.add("", values=values)

    def deleteRecords(self, event=None):
        """
        Удаление выделенных строк

        :param event: объект события
        :Автор(ы): Константинов
        """
        nb = self.master.master
        nb = nb.index(nb.select())
        selected = [int(i) for i in self.selection()]
        if not len(selected):
            funcs.message(self.root, "Не выбран элемент", msgtype="warning").fade()
        else:
            DB.modified = True
            for item in selected:
                itemId = int(self.item(item)['values'][0])
                DB.db[nb] = DB.db[nb].drop(DB.db[nb].index[DB.db[nb]['Код'] == itemId])
                self.delete(self.selection()[0])

    def modRecord(self):
        """
        Редактирование записи

        :param parent: родительский виджет
        :param values: список значений столбцов
        :type values: list
        :Автор(ы): Константинов
        """
        nb = self.master.master
        nb = nb.index(nb.select())
        selected = self.selection()
        if not selected:
            funcs.message(self.root, "Не выбран элемент", msgtype="warning").fade()
        else:
            selected = int(selected[0])
            itemId = np.int64(self.item(selected)['values'][0])
            itemValues = DB.db[nb][DB.db[nb]['Код'] == itemId].values[0].tolist()
            dic = funcs.askValuesDialog(self.root, self.config, DB.db[nb].columns,
                                  currValues=itemValues).show()
            keys = list(dic.keys())
            values = list(dic.values())
            if (len(values)):
                values = [item.get() for item in values]
                values[0] = itemId
                DB.modified = True
                for i in range(len(keys)):
                    self.item(selected, values=values)
                    DB.db[nb].loc[itemId-1, keys[i]] = values[i]
