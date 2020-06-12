import sys
import tkinter as tk
import tkinter.ttk as ttk
import pandas as pd
import pickle as pk
import numpy as np
import matplotlib.pyplot as plt
from tkinter import filedialog
from tkinter import messagebox as mb
import matplotlib
from os import listdir
from os.path import isfile, join


colorDict = {"warning": ["yellow", "lightyellow"],
             "error": ["red", "mistyrose"],
             "success": ["lawngreen", "palegreen"],
             "info": ["cornflowerblue", "lightsteelblue"]}


quantParams = set(["Сумма", "Код работника", "Дата выплаты", "Код должности", \
                   "Отделение","Норма (ч)", "Ставка (ч)", "Номер договора"])

quantComboValues = ["Сумма", "Дата выплаты", "Отделение", "Норма (ч)", "Ставка (ч)"]
qualComboValues = ["Тип выплаты", "Должность", "Образование", "Отдел"]


def getConfig(configFile="../Library/config.txt"):
    f = open(configFile, 'r')
    config = dict()
    for line in f:
        line = line.strip()
        if len(line) and line[0] != '#':
            eq = line.find('=')
            config[line[:eq]] = line[eq+1:]
    f.close()
    return config


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
        

def openFromFile(filename, db, modified, currentFile, createEmptyDatabase):
    if not filename:
        return db, modified, currentFile
    if (filename[-6::] == "pickle"):
        try:
            dbf = open(filename, "rb")
        except FileNotFoundError:
            mb.showerror(title="Файл не найден!", message="По указанному пути не удалось открыть файл. Будет создана пустая база данных.")
            createEmptyDatabase()
        else:
            currentFile = filename
            db = pk.load(dbf)
            dbf.close()
            modified = False
            return db, modified, currentFile
    else:
        try:
            xls = pd.ExcelFile(filename)  # your repository
        except FileNotFoundError:
            mb.showerror(title="Файл не найден!", message="По указанному пути не удалось открыть файл. Будет создана пустая база данных.")
            createEmptyDatabase()
        else:
            db = pd.read_excel(xls, list(range(5)))
            currentFile = ''
            modified = True
            return db, modified, currentFile


def getUID(s):
    if (s.find(".png") == -1):
        return -1
    uid = 0
    i = s.rfind('_')
    if i == -1:
        return -1
    i += 1
    while i < len(s) and s[i] != '.':
        uid = uid * 10 + int(s[i])
        i += 1
    return uid


def createUniqueFilename(specs, extension, directory):
    newUID = 1
    specs.append(str(newUID))
    filename = '_'.join(specs).replace(' ', '_') + extension
    onlyfiles = [f for f in listdir(directory) if isfile(join(directory, f))]
    uids = set([getUID(file) for file in onlyfiles if filename[:filename.rfind('_')] == file[:filename.rfind('_')]])
    while newUID in uids:
        newUID += 1
    specs[-1] = str(newUID)
    filename = directory + '_'.join(specs).replace(' ', '_') + extension
    return filename


def getBoxWhisker(root, window, fdf):
    qual = window.ComboQual.get()
    quant = window.ComboQuant.get()
    fig, ax1 = plt.subplots(figsize=(8, 4))
    data = []
    names = []
    for i in range(5):
        fdf[i].index = np.arange(len(fdf[i]))
    size = len(fdf[1].index)
    csize = len(fdf[0].index)
    if (qual == "Должность" and quant == "Сумма"):
        lprof = len(fdf[2].index)
        names = fdf[2]["Название"].tolist()
        for i in range(lprof - 1):
            fdata = []
            fnums = []
            for j in range(size - 1):
                if fdf[1].loc[j + 1, "Код должности"] == fdf[2].loc[i + 1, "Код"]:
                    fnums.append(fdf[1].loc[j + 1, "Код"])
            for j in range(csize - 1):
                if fdf[0].loc[j + 1, "Код работника"] in fnums:
                    fdata.append(float(fdf[0].loc[j + 1, "Сумма"]))
            data.append(fdata)
    elif (qual == "Образование" and quant == "Сумма"):
        names = ["Высшее", "Неоконченное высшее", "Среднее профессиональное",
                 "Студент", "Среднее общее"]
        for i in names:
            fdata = []
            fnums = []
            for j in range(size - 1):
                if fdf[3].loc[j + 1, "Образование"] == i:
                    fnums.append(fdf[3].loc[j + 1, "Код"])
            for j in range(csize - 1):
                if fdf[0].loc[j + 1, "Код работника"] in fnums:
                    fdata.append(float(fdf[0].loc[j + 1, "Сумма"]))
            data.append(fdata)
    elif (qual == "Отдел" and quant == "Сумма"):
        ldep = len(fdf[4].index)
        names = fdf[4]["Название"].tolist()
        for i in range(ldep - 1):
            fdata = []
            fnums = []
            for j in range(size - 1):
                if fdf[1].loc[j + 1, "Отделение"] == fdf[4].loc[i + 1, "Код"]:
                    fnums.append(fdf[1].loc[j + 1, "Код"])
            for j in range(csize - 1):
                if fdf[0].loc[j + 1, "Код работника"] in fnums:
                    fdata.append(float(fdf[0].loc[j + 1, "Сумма"]))
            data.append(fdata)
    else:
        return None, None
    ax1.boxplot(data, 0, '')
    ax1.set_xticklabels(names, rotation=45, fontsize=8)
    ax1.set_title('Диаграмма: $' + str(quant) + '$ от $' + str(qual) + '$')
    ax1.set_xlabel('$' + str(qual) + '$')
    ax1.set_ylabel('$' + str(quant) + '$')
    if len(ax1.get_xticks()) > 30:
        for tick in ax1.xaxis.get_majorticklabels():
            tick.set_horizontalalignment("right")
        plt.xticks(rotation=45, fontsize=6)
    else:
        labels_formatted = [str(label) if i % 2 == 0 else '\n' + str(label) for i, label in enumerate(names)]
        ax1.set_xticklabels(labels_formatted)
    plt.tight_layout()

    filename = createUniqueFilename(['БокВис', quant, qual], '.png', '../Graphics/')
    return fig, filename

def getBar(root, window, df):
    qual = window.ComboQual.get()
    quant = window.ComboQuant.get()
    quals = []
    xlabels = set()
    data = None
    if qual == "Должность" and quant == "Образование":
        xlabels = set(df[3]["Образование"].tolist()) # degrees
        quals = df[2]["Название"].tolist()
        data = [None] * len(quals)
        profs = df[2]["Код"].tolist()
        for i in range(len(profs)):
            workerIDs = df[1].loc[df[1]["Код должности"] == int(profs[i])]["Код"].tolist()
            edus = []
            for worker in workerIDs:
                found = df[3].loc[df[3]["Код"] == worker]["Образование"].tolist()
                if len(found):
                    edus.append(str(found[0]))
            data[i] = [edus.count(edu) for edu in xlabels]
    elif qual == "Образование" and quant == "Должность":
        quals = list(set(df[3]["Образование"].tolist()))
        xlabels = df[2]["Название"].tolist() # profs
        data = [None] * len(quals)
        print(data)
        for i in range(len(quals)):
            data[i] = [0] * len(xlabels)
            workerIDs = df[3].loc[df[3]["Образование"] == quals[i]]["Код"].tolist()
            for j in range(len(workerIDs)):
                profs = df[1].loc[df[1]["Код"] == workerIDs[j]]["Код должности"].tolist()
                if len(profs):
                    data[i][profs[0]-1] += 1
            print("data[", i, "]: ", data[i], sep="")
        print('\n\n\n')
        print(data)
    else:
        return None, None
    fig, ax1 = plt.subplots(figsize=(8, 4))
    ax1.set_xlabel('$' + str(quant) +'$')
    ax1.set_ylabel('$Частота$')
    colors = ['red', 'tan', 'lime', 'grey', 'black', 'blue', 'cyan', 'magenta']
    for i in range(len(data)):
        ax1.bar(list(xlabels), data[i], width=.8-.1*i, color=colors[i], label=quals, edgecolor='black', alpha=1)
    ax1.legend(quals, prop={'size': 8})
    ax1.set_title('Диаграмма $' + quant + '$ x $' + qual + '$')
    for tick in ax1.xaxis.get_majorticklabels():
        tick.set_horizontalalignment("right")
    plt.xticks(rotation=45, fontsize=6)
    plt.tight_layout()
    filename = createUniqueFilename(['гист', quant, qual], '.png', '../Graphics/')
    return fig, filename


def getHist(root, window, df):
    qual = window.ComboQual.get()
    quant = window.ComboQuant.get()
    quals = []
    data = None
    if qual == "Должность" and quant == "Сумма":
        quals = set(df[2]["Название"].tolist())
        data = [None] * len(quals)
        profs = df[2]["Код"].tolist()
        for i in range(len(profs)):
            workerIDs = df[1].loc[df[1]["Код должности"] == int(profs[i])]["Код"].tolist()
            sals = []
            for worker in workerIDs:
                found = df[0].loc[df[0]["Код работника"] == worker]["Сумма"]
                if not found.empty:
                    sals.append(float(found))
            data[i] = sals
    elif qual == "Образование" and quant == "Сумма":
        quals = set(df[3]["Образование"].tolist())
        data = [None] * len(quals)
        profs = df[2]["Код"].tolist()
        i = 0
        for edu in quals:
            workerIDs = df[3].loc[df[3]["Образование"] == edu]["Код"].tolist()
            sals = []
            for worker in workerIDs:
                found = df[0].loc[df[0]["Код работника"] == worker]["Сумма"]
                if not found.empty:
                    sals.append(float(found))
            data[i] = sals
            i += 1
    elif qual == "Отдел" and quant == "Сумма":
        quals = set(df[4]["Название"].tolist())
        data = [None] * len(quals)
        deps = df[4]["Код"].tolist()
        for i in range(len(deps)):
            workerIDs = df[1].loc[df[1]["Отделение"] == int(deps[i])]["Код"].tolist()
            sals = []
            for worker in workerIDs:
                found = df[0].loc[df[0]["Код работника"] == worker]["Сумма"]
                if not found.empty:
                    sals.append(float(found))
            data[i] = sals
    elif qual == "Тип выплаты" and quant == "Сумма":
        quals = set(df[0]["Тип выплаты"].tolist())
        data = [None] * len(quals)
        i = 0
        for stype in quals:
            workerIDs = df[0].loc[df[0]["Тип выплаты"] == stype]["Код"].tolist()
            sals = []
            for worker in workerIDs:
                sals.append(float(df[0].loc[df[0]["Код работника"] == worker]["Сумма"]))
            data[i] = sals
            i += 1
    else:
        return None, None
    fig, ax1 = plt.subplots(figsize=(8, 4))
    ax1.set_xlabel('$' + str(quant) +'$')
    ax1.set_ylabel('$Частота$')
    colors = ['red', 'tan', 'lime', 'grey', 'black', 'blue', 'cyan', 'magenta']
    ax1.hist(data, 10, density=False, histtype='bar', color=colors[:len(data)], label=quals, edgecolor='black')
    ax1.legend(prop={'size': 10})
    ax1.set_title('Диаграмма $' + quant + '$ x $' + qual + '$')

    filename = createUniqueFilename(['гист', quant, qual], '.png', '../Graphics/')
    return fig, filename


def cutName(s):
    words = s.split()
    shortName = words[0]
    if len(words) > 1:
        shortName += (' ' + words[1][0] + '.')
    if len(words) > 2:
        shortName += (' ' + words[2][0] + '.')
    return shortName


def configureWidgets(scr, top):
    scr.Pick_Analysis = tk.LabelFrame(top, bg=scr.config["def_frame_color"], fg=scr.config["def_frame_fg_color"])
    scr.Pick_Analysis.place(relx=0.023, rely=0.017, relheight=0.33,
                            relwidth=0.207)
    scr.Pick_Analysis.configure(text="Анализ")
    scr.Pick_Analysis.configure(cursor="arrow")

    scr.ComboAnalysis = ttk.Combobox(scr.Pick_Analysis, values=['Качественный параметр',
                                                                'Количественный параметр',
                                                                'Столбчатая диаграмма',
                                                                'Гистограмма',
                                                                'Диаграмма Бокса-Вискера',
                                                                'Диаграмма рассеивания',])
    scr.ComboAnalysis.place(relx=.05, rely=.35, height=20, relwidth=.9,
                            bordermode='ignore')

    scr.ShowAnalysisBtn = tk.Button(scr.Pick_Analysis, text="Показать", command=scr.showReport, bg=scr.config["def_btn_color"], fg=scr.config["def_btn_fg_color"])
    scr.ShowAnalysisBtn.place(relx=.048, rely=.5, height=32, relwidth=.9,
                              bordermode='ignore')
    scr.ShowAnalysisBtn.configure(cursor="hand2")

    scr.ExportAnalysisBtn = tk.Button(scr.Pick_Analysis, text="Экспорт", command=scr.exportReport, bg=scr.config["def_btn_color"], fg=scr.config["def_btn_fg_color"])
    scr.ExportAnalysisBtn.place(relx=.048, rely=.7, height=32, relwidth=.9,
                                bordermode='ignore')
    scr.ExportAnalysisBtn.configure(cursor="hand2")

    scr.Choice_Label = tk.Label(scr.Pick_Analysis, text="Вид отчета", bg=scr.config["def_frame_color"], fg=scr.config["def_frame_fg_color"])
    scr.Choice_Label.place(relx=.05, rely=.2, height=25, width=127,
                           bordermode='ignore')

    scr.Analysis_Frame = tk.LabelFrame(top, text="Параметры отчета", bg=scr.config["def_frame_color"], fg=scr.config["def_frame_fg_color"])
    scr.Analysis_Frame.place(relx=.24, rely=.017, relheight=.33,
                             relwidth=.201)

    scr.LabelQual = tk.Label(scr.Analysis_Frame, text="Качественный: ", anchor="w", bg=scr.config["def_frame_color"], fg=scr.config["def_frame_fg_color"])
    scr.LabelQual.place(relx=.05, rely=.2, height=25, width=127,
                            bordermode='ignore')

    scr.ComboQual = ttk.Combobox(scr.Analysis_Frame)
    scr.ComboQual.place(relx=.05, rely=.3, height=20, relwidth=.9,
                          bordermode='ignore')
    scr.ComboQual.configure(values = qualComboValues)

    scr.LabelQuant = tk.Label(scr.Analysis_Frame, text="Количественный: ", anchor="w", bg=scr.config["def_frame_color"], fg=scr.config["def_frame_fg_color"])
    scr.LabelQuant.place(relx=.05, rely=.4, height=25, width=127,
                            bordermode='ignore')

    scr.ComboQuant = ttk.Combobox(scr.Analysis_Frame)
    scr.ComboQuant.place(relx=.05, rely=.5, height=20, relwidth=.9,
                         bordermode='ignore')
    scr.ComboQuant.configure(values = quantComboValues)

    scr.Filter_Frame = tk.LabelFrame(top, text="Фильтры", bg=scr.config["def_frame_color"], fg=scr.config["def_frame_fg_color"])
    scr.Filter_Frame.place(relx=0.45, rely=0.017, relheight=0.33,
                           relwidth=0.532)

    scr.Data = ttk.Notebook(top)
    scr.Data.place(relx=0.023, rely=0.374, relheight=.528, relwidth=0.96)
    #  scr.Data.configure(takefocus="")

    scr.Data_t1 = tk.Frame(scr.Data)
    scr.Data.add(scr.Data_t1, padding=3)
    scr.Data.tab(0, text="Учёт")

    scr.Data_t2 = tk.Frame(scr.Data)
    scr.Data.add(scr.Data_t2, padding=3)
    scr.Data.tab(1, text="Работники")

    scr.Data_t3 = tk.Frame(scr.Data)
    scr.Data.add(scr.Data_t3, padding=3)
    scr.Data.tab(2, text="Должности")

    scr.Data_t4 = tk.Frame(scr.Data)
    scr.Data.add(scr.Data_t4, padding=3)
    scr.Data.tab(3, text="Информация")

    scr.Data_t5 = tk.Frame(scr.Data)
    scr.Data.add(scr.Data_t5, padding=3)
    scr.Data.tab(4, text="Отдел")

    #  configure filter lists
    scr.Filter_List1 = tk.Listbox(scr.Filter_Frame, exportselection=0)
    scr.Filter_List1.place(relx=0.019, rely=0.268, relheight=0.46,
                           relwidth=0.301, bordermode='ignore')

    scr.Filter_List2 = tk.Listbox(scr.Filter_Frame, exportselection=0)
    scr.Filter_List2.place(relx=0.338, rely=0.268, relheight=0.46,
                           relwidth=0.301, bordermode='ignore')

    scr.Filter_List1.insert('end', "Тип выплаты")
    scr.Filter_List1.insert('end', "Дата выплаты")
    scr.Filter_List1.insert('end', "Сумма")
    scr.Filter_List1.insert('end', "Код работника")
    for i in range(4):
        scr.Filter_List2.insert('end', "")

    scr.Filter_scroll = tk.Scrollbar(scr.Filter_List1)
    scr.Filter_List1.config(yscrollcommand=scr.Filter_scroll.set)
    scr.Filter_List1.bind("<MouseWheel>", scr.scrollList2)
    scr.Filter_List2.config(yscrollcommand=scr.Filter_scroll.set)
    scr.Filter_List2.bind("<MouseWheel>", scr.scrollList1)

    scr.Change_Button = tk.Button(scr.Filter_Frame, bg=scr.config["def_btn_color"], fg=scr.config["def_btn_fg_color"])
    scr.Change_Button.place(relx=0.357, rely=0.804, height=32, width=148,
                            bordermode='ignore')
    scr.Change_Button.configure(cursor="hand2")
    scr.Change_Button.configure(text="Изменить значения", command=scr.open_dialog)

    scr.Reset_Button = tk.Button(scr.Filter_Frame, text="Сбросить выбор", command=scr.reset, bg=scr.config["def_btn_color"], fg=scr.config["def_btn_fg_color"])
    scr.Reset_Button.place(relx=0.03, rely=0.804, height=32, width=148,
                           bordermode='ignore')
    scr.Reset_Button.configure(cursor="hand2")

    scr.Param_Label = tk.Label(scr.Filter_Frame, text="Параметры", bg=scr.config["def_frame_color"], fg=scr.config["def_frame_fg_color"])
    scr.Param_Label.place(relx=0.075, rely=0.134, height=25, width=97,
                          bordermode='ignore')

    scr.Values_Label = tk.Label(scr.Filter_Frame, text="Значения", bg=scr.config["def_frame_color"], fg=scr.config["def_frame_fg_color"])
    scr.Values_Label.place(relx=0.414, rely=0.152, height=15, width=83,
                           bordermode='ignore')

    # Checkboxes
    scr.Boxes_Frame = tk.LabelFrame(scr.Filter_Frame, text="Столбцы", bg=scr.config["def_frame_color"], fg=scr.config["def_frame_fg_color"])
    scr.Boxes_Frame.place(relx=0.658, rely=0.130, relheight=0.8,
                           relwidth=0.32, bordermode='ignore')
    scr.Cvars = []
    scr.Cvars0 = []
    scr.Cvars.append(scr.Cvars0)
    scr.Cvars1 = []
    scr.Cvars.append(scr.Cvars1)
    scr.Cvars2 = []
    scr.Cvars.append(scr.Cvars2)
    scr.Cvars3 = []
    scr.Cvars.append(scr.Cvars3)
    scr.Cvars4 = []
    scr.Cvars.append(scr.Cvars4)

    for i in range(4):
        scr.Cvars0.append(tk.BooleanVar())
        scr.Cvars0[i].set(1)
    for i in range(2):
        scr.Cvars1.append(tk.BooleanVar())
        scr.Cvars1[i].set(1)
    for i in range(3):
        scr.Cvars2.append(tk.BooleanVar())
        scr.Cvars2[i].set(1)
    for i in range(5):
        scr.Cvars3.append(tk.BooleanVar())
        scr.Cvars3[i].set(1)
    for i in range(2):
        scr.Cvars4.append(tk.BooleanVar())
        scr.Cvars4[i].set(1)

    scr.names = []
    scr.Cboxes = []
    scr.Cboxes0 = []
    scr.Cboxes.append(scr.Cboxes0)
    scr.Cboxes1 = []
    scr.Cboxes.append(scr.Cboxes1)
    scr.Cboxes2 = []
    scr.Cboxes.append(scr.Cboxes2)
    scr.Cboxes3 = []
    scr.Cboxes.append(scr.Cboxes3)
    scr.Cboxes4 = []
    scr.Cboxes.append(scr.Cboxes4)
    scr.Cbox0 = tk.Checkbutton(scr.Boxes_Frame, command=scr.removeColumns, bg=scr.config["def_frame_color"], fg=scr.config["def_frame_fg_color"])
    scr.Cbox0.grid(row=0, column=0, sticky='W')
    scr.Cbox0.configure(justify='left')
    scr.Cbox0.configure(text="Тип выплаты", variable=scr.Cvars0[0])
    scr.names.append(scr.Cbox0.cget("text"))
    scr.Cboxes0.append(scr.Cbox0)

    scr.Cbox1 = tk.Checkbutton(scr.Boxes_Frame, command=scr.removeColumns, bg=scr.config["def_frame_color"], fg=scr.config["def_frame_fg_color"])
    scr.Cbox1.grid(row=1, column=0, sticky='W')
    scr.Cbox1.configure(justify='left')
    scr.Cbox1.configure(text="Дата выплаты", variable=scr.Cvars0[1])
    scr.names.append(scr.Cbox1.cget("text"))
    scr.Cboxes0.append(scr.Cbox1)

    scr.Cbox2 = tk.Checkbutton(scr.Boxes_Frame, command=scr.removeColumns, bg=scr.config["def_frame_color"], fg=scr.config["def_frame_fg_color"])
    scr.Cbox2.grid(row=2, column=0, sticky='W')
    scr.Cbox2.configure(justify='left')
    scr.Cbox2.configure(text="Сумма", variable=scr.Cvars0[2])
    scr.names.append(scr.Cbox2.cget("text"))
    scr.Cboxes0.append(scr.Cbox2)

    scr.Cbox3 = tk.Checkbutton(scr.Boxes_Frame, command=scr.removeColumns, bg=scr.config["def_frame_color"], fg=scr.config["def_frame_fg_color"])
    scr.Cbox3.grid(row=3, column=0, sticky='W')
    scr.Cbox3.configure(justify='left')
    scr.Cbox3.configure(text="Код работника", variable=scr.Cvars0[3])
    scr.names.append(scr.Cbox3.cget("text"))
    scr.Cboxes0.append(scr.Cbox3)

    scr.Cbox4 = tk.Checkbutton(scr.Boxes_Frame, command=scr.removeColumns, bg=scr.config["def_frame_color"], fg=scr.config["def_frame_fg_color"])
    scr.Cbox4.grid(row=0, column=0, sticky='W')
    scr.Cbox4.configure(justify='left')
    scr.Cbox4.configure(text="Код должности", variable=scr.Cvars1[0])
    scr.Cbox4.grid_forget()
    scr.names.append(scr.Cbox4.cget("text"))
    scr.Cboxes1.append(scr.Cbox4)

    scr.Cbox5 = tk.Checkbutton(scr.Boxes_Frame, command=scr.removeColumns, bg=scr.config["def_frame_color"], fg=scr.config["def_frame_fg_color"])
    scr.Cbox5.grid(row=1, column=0, sticky='W')
    scr.Cbox5.configure(justify='left')
    scr.Cbox5.configure(text="Отделение", variable=scr.Cvars1[1])
    scr.Cbox5.grid_forget()
    scr.names.append(scr.Cbox5.cget("text"))
    scr.Cboxes1.append(scr.Cbox5)

    scr.Cbox6 = tk.Checkbutton(scr.Boxes_Frame, command=scr.removeColumns, bg=scr.config["def_frame_color"], fg=scr.config["def_frame_fg_color"])
    scr.Cbox6.grid(row=0, column=0, sticky='W')
    scr.Cbox6.configure(justify='left')
    scr.Cbox6.configure(text="Название", variable=scr.Cvars2[0])
    scr.Cbox6.grid_forget()
    scr.names.append(scr.Cbox6.cget("text"))
    scr.Cboxes2.append(scr.Cbox6)

    scr.Cbox7 = tk.Checkbutton(scr.Boxes_Frame, command=scr.removeColumns, bg=scr.config["def_frame_color"], fg=scr.config["def_frame_fg_color"])
    scr.Cbox7.grid(row=1, column=0, sticky='W')
    scr.Cbox7.configure(justify='left')
    scr.Cbox7.configure(text="Норма (ч)", variable=scr.Cvars2[1])
    scr.Cbox7.grid_forget()
    scr.names.append(scr.Cbox7.cget("text"))
    scr.Cboxes2.append(scr.Cbox7)

    scr.Cbox8 = tk.Checkbutton(scr.Boxes_Frame, command=scr.removeColumns, bg=scr.config["def_frame_color"], fg=scr.config["def_frame_fg_color"])
    scr.Cbox8.grid(row=2, column=0, sticky='W')
    scr.Cbox8.configure(justify='left')
    scr.Cbox8.configure(text="Ставка (ч)", variable=scr.Cvars2[2])
    scr.Cbox8.grid_forget()
    scr.names.append(scr.Cbox8.cget("text"))
    scr.Cboxes2.append(scr.Cbox8)

    scr.Cbox9 = tk.Checkbutton(scr.Boxes_Frame, command=scr.removeColumns, bg=scr.config["def_frame_color"], fg=scr.config["def_frame_fg_color"])
    scr.Cbox9.grid(row=0, column=0, sticky='W')
    scr.Cbox9.configure(justify='left')
    scr.Cbox9.configure(text="ФИО", variable=scr.Cvars3[0])
    scr.Cbox9.grid_forget()
    scr.names.append(scr.Cbox9.cget("text"))
    scr.Cboxes3.append(scr.Cbox9)

    scr.Cbox10 = tk.Checkbutton(scr.Boxes_Frame, command=scr.removeColumns, bg=scr.config["def_frame_color"], fg=scr.config["def_frame_fg_color"])
    scr.Cbox10.grid(row=1, column=0, sticky='W')
    scr.Cbox10.configure(justify='left')
    scr.Cbox10.configure(text="Номер договора", variable=scr.Cvars3[1])
    scr.Cbox10.grid_forget()
    scr.names.append(scr.Cbox10.cget("text"))
    scr.Cboxes3.append(scr.Cbox10)

    scr.Cbox11 = tk.Checkbutton(scr.Boxes_Frame, command=scr.removeColumns, bg=scr.config["def_frame_color"], fg=scr.config["def_frame_fg_color"])
    scr.Cbox11.grid(row=2, column=0, sticky='W')
    scr.Cbox11.configure(justify='left')
    scr.Cbox11.configure(text="Телефон", variable=scr.Cvars3[2])
    scr.Cbox11.grid_forget()
    scr.names.append(scr.Cbox11.cget("text"))
    scr.Cboxes3.append(scr.Cbox11)

    scr.Cbox12 = tk.Checkbutton(scr.Boxes_Frame, command=scr.removeColumns, bg=scr.config["def_frame_color"], fg=scr.config["def_frame_fg_color"])
    scr.Cbox12.grid(row=3, column=0, sticky='W')
    scr.Cbox12.configure(justify='left')
    scr.Cbox12.configure(text="Образование", variable=scr.Cvars3[3])
    scr.Cbox12.grid_forget()
    scr.names.append(scr.Cbox12.cget("text"))
    scr.Cboxes3.append(scr.Cbox12)

    scr.Cbox13 = tk.Checkbutton(scr.Boxes_Frame, command=scr.removeColumns, bg=scr.config["def_frame_color"], fg=scr.config["def_frame_fg_color"])
    scr.Cbox13.grid(row=4, column=0, sticky='W')
    scr.Cbox13.configure(justify='left')
    scr.Cbox13.configure(text="Адрес", variable=scr.Cvars3[4])
    scr.Cbox13.grid_forget()
    scr.names.append(scr.Cbox13.cget("text"))
    scr.Cboxes3.append(scr.Cbox13)

    scr.Cbox14 = tk.Checkbutton(scr.Boxes_Frame, command=scr.removeColumns, bg=scr.config["def_frame_color"], fg=scr.config["def_frame_fg_color"])
    scr.Cbox14.grid(row=0, column=0, sticky='W')
    scr.Cbox14.configure(justify='left')
    scr.Cbox14.configure(text="Название", variable=scr.Cvars4[0])
    scr.Cbox14.grid_forget()
    scr.names.append(scr.Cbox14.cget("text"))
    scr.Cboxes4.append(scr.Cbox14)

    scr.Cbox15 = tk.Checkbutton(scr.Boxes_Frame, command=scr.removeColumns, bg=scr.config["def_frame_color"], fg=scr.config["def_frame_fg_color"])
    scr.Cbox15.grid(row=1, column=0, sticky='W')
    scr.Cbox15.configure(justify='left')
    scr.Cbox15.configure(text="Телефон", variable=scr.Cvars4[1])
    scr.Cbox15.grid_forget()
    scr.names.append(scr.Cbox15.cget("text"))
    scr.Cboxes4.append(scr.Cbox15)

    # menu
    menubar = tk.Menu(top)
    filemenu = tk.Menu(menubar, tearoff=0)
    filemenu.add_command(label="Новый", command=scr.newDatabase)
    filemenu.add_command(label="Открыть", command=scr.open)
    filemenu.add_command(label="Сохранить", command=scr.save)
    filemenu.add_command(label="Сохранить как...", command=scr.saveas)
    filemenu.add_command(label="Экспорт", command=scr.saveAsExcel)
    filemenu.add_separator()
    filemenu.add_command(label="Выход", command=scr.exit)
    menubar.add_cascade(label="Файл", menu=filemenu)

    helpmenu = tk.Menu(menubar, tearoff=0)
    helpmenu.add_command(label="Добавить", command=scr.addRecord)
    helpmenu.add_command(label="Удалить", command=scr.deleteRecords)
    helpmenu.add_command(label="Изменить", command=scr.modRecord)
    menubar.add_cascade(label="Правка", menu=helpmenu)
    
    helpmenu = tk.Menu(menubar, tearoff=0)
    helpmenu.add_command(label="Настройка интерфейса", command=scr.addRecord)
    menubar.add_cascade(label="Вид", menu=helpmenu)

    top.config(menu=menubar)

    # status bar
    scr.statusbar = tk.Label(top, text="Oh hi. I didn't see you there...", bd=1,
                             relief=tk.SUNKEN, anchor=tk.W)
    scr.statusbar.pack(side=tk.BOTTOM, fill=tk.X)


def refreshFromExcel(filename):
    xls = pd.ExcelFile(filename)  # your repository
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
        config = getConfig()
        self.geometry("200x90+550+230")
        self.resizable(0, 0)
        self.title("")

        self.var = tk.StringVar()

        self.configure(bg=config["def_bg_color"])
        self.label = tk.Label(self, text=prompt, bg=config["def_bg_color"], fg=config["def_frame_fg_color"])
        self.entry = tk.Entry(self, textvariable=self.var)
        self.ok_button = tk.Button(self, text="OK", command=self.on_ok)
        self.ok_button.configure(bg=config["def_btn_color"], fg=config["def_btn_fg_color"])
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
    

def testVal(inStr,acttyp):
    if acttyp == '1': #insert
        if not inStr.isdigit():
            return False
    return True


class message(tk.Toplevel):
    def __init__(self, parent, prompt="Сообщение", msgtype="info"):
        self.opacity = 3.0
        tk.Toplevel.__init__(self, parent)
        self.resizable(0, 0)
        self.header = tk.Label(self, text="")
        self.header.pack(side="top", fill="x")
        self.label = tk.Label(self, text=prompt, justify="center")
        self.label.pack(side="top", fill="both", expand=1)
        geom = "200x80+" + str(parent.winfo_screenwidth()-260) + "+" + str(parent.winfo_screenheight()-120)
        self.geometry(geom)
        self.setColor(msgtype)
        self.overrideredirect(True)

    def setColor(self, msgtype="info"):
        self.header.configure(background=colorDict[msgtype][0])
        self.label.configure(background=colorDict[msgtype][1])
        self.configure(background=colorDict[msgtype][1])

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
        self.parent = parent
        x = str(parent.winfo_screenwidth() // 2 - 150)
        y = str(parent.winfo_screenheight() // 2 - 200)
        self.geometry("300x400+" + x + "+" + y)
        self.resizable(0, 0)
        self.grab_set()  # make modal
        self.focus()
        self.Labels = [None] * len(labelTexts)
        self.Edits = [None] * len(labelTexts)
        self.retDict = dict()
        config = getConfig()
        self.configure(background=config["def_bg_color"])
        for i in range(len(labelTexts)):
            self.retDict[labelTexts[i]] = tk.StringVar()
            editHeight = .8*400/len(labelTexts)
            self.Labels[i] = tk.Label(self, text=labelTexts[i]+":", anchor='e')
            self.Labels[i].configure(bg=config["def_bg_color"], fg=config["def_frame_fg_color"])
            self.Labels[i].place(relx=.1, y=40+i*editHeight, width=100)
            self.Edits[i] = tk.Entry(self, textvariable=self.retDict[labelTexts[i]], validate="key")
            if labelTexts[i] in quantParams:
                self.Edits[i].configure(validate="key")
                self.Edits[i].configure(validatecommand = (self.Edits[i].register(testVal),'%P','%d'))
            self.Edits[i].place(relx=.5, y=40+i*editHeight, width=100)
            if labelTexts[i] == 'Код':
                self.Edits[i].configure(state='disabled')
            if currValues:
                self.Edits[i].insert(0, currValues[i])

        self.ok_button = tk.Button(self, text="OK", command=self.on_ok)
        self.ok_button.configure(bg=config["def_btn_color"], fg=config["def_btn_fg_color"])
        self.ok_button.place(relx=.5, rely=.9, relwidth=.4, height=30, anchor="c")

        self.bind("<Return>", self.on_ok)
        self.protocol("WM_DELETE_WINDOW", self.exit)

    def exit(self):
        self.retDict.clear()
        self.destroy()

    def on_ok(self, event=None):
        for edit in self.Edits[1:]:
            if edit.get().strip() == '':
                message(self.parent, "Поля не могут быть пустыми.", msgtype="warning").fade()
                return
        self.destroy()

    def show(self):
        self.wm_deiconify()
        self.wait_window()
        return self.retDict
