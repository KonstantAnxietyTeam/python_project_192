"""
Модуль, содержащий универсальные функции, либо требующие минимальных модификаций
для применение в других проектах
"""

import tkinter as tk
import tkinter.ttk as ttk
import pandas as pd
import pickle as pk
import numpy as np
import matplotlib.pyplot as plt
from tkinter import filedialog
from tkinter import messagebox as mb
from tkinter import colorchooser
from os import listdir
from os.path import isfile, join
from collections import Counter

colorDict = {'warning': ['yellow', 'lightyellow'],
             'error': ['red', 'mistyrose'],
             'success': ['lawngreen', 'palegreen'],
             'info': ['cornflowerblue', 'lightsteelblue']}


quantParams = set(['Сумма', 'Код работника', 'Дата выплаты', 'Код должности',
                   'Отделение', 'Норма (ч/мес)', 'Ставка (ч)',
                   'Номер договора', 'Отработано (ч)'])

quantComboValues = ['Сумма', 'Дата выплаты', 'Норма (ч/мес)', 'Ставка (ч)',
                    'Отработано (ч)']
qualComboValues = ['Должность', 'Образование', 'Отдел']


def getDefaultConfig():
    """
    Генерация настроек по умолчанию

    :return: Словарь настроек по умолчанию
    :rtype: dict
    :Автор(ы): Константинов
    """
    config = {
        'def_db': '../Data/db.pickle',
        'def_graph_dir': '../Graphics/',
        'def_output_dir': '../Output/',
        'def_window_width': '1000',
        'def_window_height': '600',
        'fullscreen': '0',
        'maximize': '0',
        'def_bg_color': 'whitesmoke',
        'def_frame_color': 'whitesmoke',
        'def_btn_color': 'whitesmoke',
        'def_frame_fg_color': 'black',
        'def_btn_fg_color': 'black',
    }
    return config


def getConfig(configFile='../Library/config.txt'):
    """
    Загрузка настроек из файла

    :param configFile: путь к файлу настроек
    :type configFile: string
    :return: словарь настроек
    :rtype: dict
    :Автор(ы): Константинов
    """
    f = open(configFile, 'r')
    config = dict()
    config = getDefaultConfig()
    if not f:
        writeConfig(config)
        return config
    for line in f:
        line = line.strip()
        if len(line) and line[0] != '#':
            eq = line.find('=')
            config[line[:eq]] = line[eq+1:]
    f.close()
    return config


def writeConfig(config=None, path='../Library/config.txt'):
    """
    Запись настроек в файл

    :param config: словарь настроек
    :type config: dict
    :param path: путь к файлу для сохранения
    :type path: string
    :Автор(ы): Константинов
    """
    f = open(path, 'w')
    f.write('###### paths\n')
    for key in ['def_db', 'def_graph_dir', 'def_output_dir']:
        f.write(key + '=' + config[key] + '\n')
    f.write('\n')
    f.write('###### geometry\n')
    for key in ['def_window_width', 'def_window_height', 'fullscreen',
                'maximize']:
        f.write(key + '=' + config[key] + '\n')
    f.write('\n')
    f.write('###### colors\n')
    for key in ['def_bg_color', 'def_frame_color', 'def_btn_color',
                'def_frame_fg_color', 'def_btn_fg_color']:
        f.write(key + '=' + config[key] + '\n')
    f.write('\n')
    f.close()


def saveAsExcel(root, tree):
    """
    Сохранение содержимого таблицы treeview в файл .xlsx

    :param tree: таблица
    :type tree: ttk.TreeView
    :param root: корневой объект для вывода сообщения
    :type root: tkinter widget
    :Автор(ы): Константинов
    """
    file = filedialog.asksaveasfilename(title='Select file',
                                        initialdir='../Data/db1.xlsx',
                                        defaultextension='.xlsx',
                                        filetypes=[('Excel file', '*.xlsx')])
    if file:
        ids = tree.get_children()
        #dic = dict([tree.column(i)['id'] for i in tree["displaycolumns"]]) # TODO need to get displayed columns only
        dic = dict.fromkeys(tree['columns'], [])
        keys = list(dic.keys())
        for i in range(len(keys)):
            dic[keys[i]] = []
        for iid in ids:
            for i in range(len(keys)):
                dic[keys[i]].append(tree.item(iid)['values'][i])

        dic = pd.DataFrame.from_dict(dic)
        try:
            dic.to_excel(file, engine='xlsxwriter',index=False)
            message(root, 'Таблица сохранена', msgtype='success').fade()
        except:
            message(root, 'Не удалось сохранить файл!\nВозможно, он открыт\nв другой программе', msgtype='error').fade()
    else:
        pass  # pressed cancel


def openFromFile(filename, db, modified, currentFile, createEmptyDatabase=None):
    """
    Открытик базы данных из файла .xslx или из бинарного файла pickle

    :param filename: путь к файлу для сохранения
    :type filename: string
    :param db: текущая база для возврата в случае ошибочного параметра имени файла
    :type db: pandas.DataFrame
    :modified: настоящее состояние текущей базы данных для возврата в случае ошибочного параметра имени файла
    :type modified: bool
    :param currentFile: пусть к файлу текущей базы данных для возврата в случае ошибочного параметра имени файла
    :type currentFile: string
    :param createEmptyDatabase: функция для создания пустой базы в случае неудачи
    :type createEmptyDatabase: функция
    :Автор(ы): Константинов
    """
    if not filename:
        return db, modified, currentFile
    if (filename[-6::] == 'pickle'):
        try:
            dbf = open(filename, 'rb')
        except FileNotFoundError:
            mb.showerror(title='Файл не найден!', message='По указанному пути не удалось открыть файл. Будет создана пустая база данных.')
            return createEmptyDatabase()
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
            mb.showerror(title='Файл не найден!', message='По указанному пути не удалось открыть файл. Будет создана пустая база данных.')
            return createEmptyDatabase()
        else:
            db = pd.read_excel(xls, list(range(5)))
            currentFile = ''
            modified = True
            return db, modified, currentFile


def getUID(s):
    """
    Изъятие UID из строки формата `info1_info2_UID.png`

    :param s: имя файла
    :type s: string
    :return: UID
    :rtype: `integer`
    :Автор(ы): Константинов
    """
    if (s.find('.png') == -1):
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
    """
    Создание уникального для директории имени файла в формате `spec_spec_spec_UID.ext`

    :param specs: список для формирования информационной части названия
    :type specs: list
    :param extension: расширение файла
    :type extension: string
    :param directory: директория для сохранения
    :type directory: string
    :return: путь к файлу с созданным именем
    :rtype: string
    :Автор(ы): Константинов
    """
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

def getSummaryDf(df):
    """
    Создание общей таблицы

    :param df: База данных
    :type df: pandas.DataFrame
    :return: Общая таблица
    :rtype: pandas.DataFrame
    :Автор(ы): Березуцкий
    """

    plt.close('all')
    totalDf = df[0].merge(df[1], how='left', left_on='Код работника', right_on='Код')
    totalDf = totalDf.merge(df[2], how='left', left_on='Код должности', right_on='Код').drop(['Код'], axis='columns')
    totalDf = totalDf.merge(df[3], how='left', left_on='Код_y', right_on='Код').drop(['Код_y', 'Код'], axis='columns')
    totalDf = totalDf.merge(df[4], how='left', left_on='Отделение', right_on='Код').drop(['Код'], axis='columns')
    totalDf = totalDf.rename(columns={'Код_x': 'Код', 'Название_x':'Должность', 'Название_y':'Отдел'})
    totalDf = totalDf.rename(columns={'Код_x': 'Код', 'Название_x':'Должность', 'Название_y':'Отдел'})

    return totalDf


def getScatterplot(root, window, fdf, directory):
    """
    Создание категоризированной диаграммы рассеивания (доступные комбинации:
    Сумма-Отработано(ч) для: Должность, Образование, Отдел))

    :param window: объект окна содержащего меню с параметрами
    :type window: MainWindow
    :param fdf: База данных
    :type fdf: pandas.DataFrame
    :param directory: путь к папке для сохранения
    :type directory: string
    :return: объект построенной диаграммы
    :rtype: matplotlib.pyplot.figure
    :return: путь к файлу с уникальным именем для сохранения
    :rtype: string
    :Автор(ы): Сидоров
    """
    qual = window.ComboQual.get()
    quant1 = window.ComboQuant.get()
    quant2 = window.ComboQuant2.get()
    size = len(fdf[1].index)
    csize = len(fdf[0].index)
    if qual == 'Должность' and ((quant1 == 'Сумма' and quant2 == 'Отработано (ч)') or (quant2 == 'Сумма' and quant1 == 'Отработано (ч)')):
        fig, ax1 = plt.subplots(figsize=(8, 4))
        profs = fdf[2]['Название'].tolist()
        for i in range(len(profs)):
            fdata1 = []
            fdata2 = []
            fnums = []
            for j in range(size):
                if fdf[1].loc[j, 'Код должности'] == fdf[2].loc[i, 'Код']:
                    fnums.append(fdf[1].loc[j, 'Код'])
            for j in range(csize):
                if fdf[0].loc[j, 'Код работника'] in fnums:
                    fdata1.append(int(fdf[0].loc[j, 'Сумма']))
                    fdata2.append(int(fdf[0].loc[j, 'Отработано (ч)']))
            ax1.scatter(fdata1, fdata2, label=profs[i])
    elif qual == 'Образование' and ((quant1 == 'Сумма' and quant2 == 'Отработано (ч)') or (quant2 == 'Сумма' and quant1 == 'Отработано (ч)')):
        fig, ax1 = plt.subplots(figsize=(8, 4))
        degrees = set(fdf[3]['Образование'].tolist())
        for degree in degrees:
            fdata1 = []
            fdata2 = []
            fnums = []
            for i in range(size):
                if fdf[3].loc[i, 'Образование'] == degree:
                    fnums.append(fdf[3].loc[i, 'Код'])
            for i in range(csize):
                if fdf[0].loc[i, 'Код работника'] in fnums:
                    fdata1.append(int(fdf[0].loc[i, 'Сумма']))
                    fdata2.append(int(fdf[0].loc[i, 'Отработано (ч)']))
            ax1.scatter(fdata1, fdata2, label=degree)
    elif qual == 'Отдел' and ((quant1 == 'Сумма' and
                               quant2 == 'Отработано (ч)') or
                              (quant2 == 'Сумма' and quant1 ==
                               'Отработано (ч)')):
        fig, ax1 = plt.subplots(figsize=(8, 4))
        deps = fdf[4]['Название'].tolist()
        for i in range(len(deps)):
            fdata1 = []
            fdata2 = []
            fnums = []
            for j in range(size):
                if fdf[1].loc[j, 'Отделение'] == fdf[4].loc[i, 'Код']:
                    fnums.append(fdf[1].loc[j, 'Код'])
            for j in range(csize):
                if fdf[0].loc[j, 'Код работника'] in fnums:
                    fdata1.append(int(fdf[0].loc[j, 'Сумма']))
                    fdata2.append(int(fdf[0].loc[j, 'Отработано (ч)']))
            ax1.scatter(fdata1, fdata2, label=deps[i])
    else:
        return None, None
    plt.close('all')
    if quant1 == 'Сумма' and quant2 == 'Отработано (ч)':
        ax1.set_title('Диаграмма: $' + str(quant1) +
                      '$ от $' + str(quant2) + '$')
        ax1.set_xlabel('$' + str(quant1) + '$')
        ax1.set_ylabel('$' + str(quant2) + '$')
    else:
        ax1.set_title('Диаграмма: $' + str(quant2) + '$ от $' +
                      str(quant1) + '$')
        ax1.set_xlabel('$' + str(quant2) + '$')
        ax1.set_ylabel('$' + str(quant1) + '$')
    ax1.legend(loc='upper left', bbox_to_anchor=(1, 0.9))
    plt.tight_layout()

    filename = createUniqueFilename(['Расс', quant1, quant2],
                                    '.png', directory)
    return fig, filename


def getBoxWhisker(root, window, fdf, directory):
    """
    Создание категоризированной диаграммы Бокса-Вискера (доступные комбинации:
    Должность-Сумма, Образование-Сумма, Отдел-Сумма))

    :param window: объект окна содержащего меню с параметрами
    :type window: MainWindow
    :param fdf: База данных
    :type fdf: pandas.DataFrame
    :param directory: путь к папке для сохранения
    :type directory: string
    :return: объект построенной диаграммы
    :rtype: matplotlib.pyplot.figure
    :return: путь к файлу с уникальным именем для сохранения
    :rtype: string
    :Автор(ы): Сидоров
    """
    qual = window.ComboQual.get()
    quant = window.ComboQuant.get()
    for i in range(5):
        fdf[i].index = np.arange(len(fdf[i]))
    data = []
    names = []
    for i in range(5):
        fdf[i].index = np.arange(len(fdf[i]))
    size = len(fdf[1].index)
    csize = len(fdf[0].index)
    if (qual == 'Должность' and quant == 'Сумма'):
        lprof = len(fdf[2].index)
        names = fdf[2]['Название'].tolist()
        for i in range(lprof):
            fdata = []
            fnums = []
            for j in range(size):
                if fdf[1].loc[j, 'Код должности'] == fdf[2].loc[i, 'Код']:
                    fnums.append(fdf[1].loc[j, 'Код'])
            for j in range(csize):
                if fdf[0].loc[j, 'Код работника'] in fnums:
                    fdata.append(int(fdf[0].loc[j, 'Сумма']))
            data.append(fdata)
    elif (qual == 'Образование' and quant == 'Сумма'):
        names = set(fdf[3]['Образование'].tolist())
        for i in names:
            fdata = []
            fnums = []
            for j in range(size):
                if fdf[3].loc[j, 'Образование'] == i:
                    fnums.append(fdf[3].loc[j, 'Код'])
            for j in range(csize):
                if fdf[0].loc[j, 'Код работника'] in fnums:
                    fdata.append(float(fdf[0].loc[j, 'Сумма']))
            data.append(fdata)
    elif (qual == 'Отдел' and quant == 'Сумма'):
        ldep = len(fdf[4].index)
        names = fdf[4]['Название'].tolist()
        for i in range(ldep):
            fdata = []
            fnums = []
            for j in range(size):
                if fdf[1].loc[j, 'Отделение'] == fdf[4].loc[i, 'Код']:
                    fnums.append(fdf[1].loc[j, 'Код'])
            for j in range(csize):
                if fdf[0].loc[j, 'Код работника'] in fnums:
                    fdata.append(float(fdf[0].loc[j, 'Сумма']))
            data.append(fdata)
    else:
        return None, None
    plt.close('all')
    fig, ax1 = plt.subplots(figsize=(8, 4))
    ax1.boxplot(data, 0, '')
    ax1.set_xticklabels(names, rotation=45, fontsize=8)
    ax1.set_title('Диаграмма: $' + str(quant) + '$ от $' + str(qual) + '$')
    ax1.set_xlabel('$' + str(qual) + '$')
    ax1.set_ylabel('$' + str(quant) + '$')
    if len(ax1.get_xticks()) > 30:
        for tick in ax1.xaxis.get_majorticklabels():
            tick.set_horizontalalignment('right')
        plt.xticks(rotation=45, fontsize=6)
    else:
        labels_formatted = [str(label) if i % 2 == 0 else
                            '\n' + str(label) for i, label in enumerate(names)]
        ax1.set_xticklabels(labels_formatted)
    plt.tight_layout()

    filename = createUniqueFilename(['БокВис', quant, qual], '.png', directory)
    return fig, filename


def getBar(root, window, df, directory):
    """
    Создание кластеризованной столбчатой диаграммы (доступные комбинации:
    Должность-Образование, Образование-Должность)

    :param window: объект окна содержащего меню с параметрами
    :type window: MainWindow
    :param df: База данных
    :type df: pandas.DataFrame
    :param directory: путь к папке для сохранение
    :type directory: string
    :return: объект построенной диаграммы
    :rtype: matplotlib.pyplot.figure
    :return: путь к файлу с уникальным именем для сохранения
    :rtype: string
    :Автор(ы): Константинов
    """
    qual = window.ComboQual.get()
    quant = window.ComboQuant.get()
    quals = []
    xlabels = set()
    data = None
    if qual == 'Должность' and quant == 'Образование':
        xlabels = set(df[3]['Образование'].tolist())  # degrees
        quals = df[2]['Название'].tolist()
        data = [None] * len(quals)
        profs = df[2]['Код'].tolist()
        for i in range(len(profs)):
            workerIDs = df[1].loc[df[1]['Код должности'] ==
                                  int(profs[i])]['Код'].tolist()
            edus = []
            for worker in workerIDs:
                found = df[3].loc[df[3]['Код'] ==
                                  worker]['Образование'].tolist()
                if len(found):
                    edus.append(str(found[0]))
            data[i] = [edus.count(edu) for edu in xlabels]
    elif qual == 'Образование' and quant == 'Должность':
        quals = list(set(df[3]['Образование'].tolist()))
        xlabels = df[2]['Название'].tolist()  # profs
        data = [None] * len(quals)
        for i in range(len(quals)):
            data[i] = [0] * len(xlabels)
            workerIDs = df[3].loc[df[3]['Образование'] == quals[i]]['Код'].tolist()
            for j in range(len(workerIDs)):
                profs = df[1].loc[df[1]['Код'] ==
                                  workerIDs[j]]['Код должности'].tolist()
                if len(profs):
                    data[i][profs[0]-1] += 1
    else:
        return None, None
    plt.close('all')
    fig, ax1 = plt.subplots(figsize=(8, 4))
    ax1.set_xlabel('$' + str(quant) + '$')
    ax1.set_ylabel('$Частота$')
    colors = ['red', 'tan', 'lime', 'grey', 'black', 'blue', 'cyan', 'magenta',
              'whitesmoke', 'yellow']
    for i in range(len(data)):
        ax1.bar(list(xlabels), data[i%10], width=.95-.1*i, color=colors[i],
                label=quals, edgecolor='black', alpha=1)
    # ax1.legend(quals, prop={'size': 8})
    ax1.legend(quals, loc='upper left', bbox_to_anchor=(1, 1))
    ax1.set_title('Диаграмма $' + quant + '$ x $' + qual + '$')
    for tick in ax1.xaxis.get_majorticklabels():
        tick.set_horizontalalignment('right')
    plt.xticks(rotation=45, fontsize=6)
    plt.tight_layout()
    filename = createUniqueFilename(['столб', quant, qual], '.png', directory)
    return fig, filename


def getHist(root, window, df, directory):
    """
    Создание гатегаризованной гистограммы (доступные комбинации:
    Должность-Сумма, Образование-Сумма, Отдел-Сумма)

    :param window: объект окна содержащего меню с параметрами
    :type window: MainWindow
    :param df: База данных
    :type df: pandas.DataFrame
    :param directory: путь к папке для сохранение
    :type directory: string
    :return: объект построенной диаграммы
    :rtype: matplotlib.pyplot.figure
    :return: путь к файлу с уникальным именем для сохранения
    :rtype: string
    :Автор(ы): Константинов
    """
    qual = window.ComboQual.get()
    quant = window.ComboQuant.get()
    quals = []
    data = None
    if qual == 'Должность' and quant == 'Сумма':
        quals = set(df[2]['Название'].tolist())
        data = [None] * len(quals)
        profs = df[2]['Код'].tolist()
        for i in range(len(profs)):
            workerIDs = df[1].loc[df[1]['Код должности'] ==
                                  int(profs[i])]['Код'].tolist()
            sals = []
            for worker in workerIDs:
                found = df[0].loc[df[0]['Код работника'] == worker]['Сумма']
                if not found.empty:
                    sals.append(float(found))
            data[i] = sals
    elif qual == 'Образование' and quant == 'Сумма':
        quals = set(df[3]['Образование'].tolist())
        data = [None] * len(quals)
        profs = df[2]['Код'].tolist()
        i = 0
        for edu in quals:
            workerIDs = df[3].loc[df[3]['Образование'] == edu]['Код'].tolist()
            sals = []
            for worker in workerIDs:
                found = df[0].loc[df[0]['Код работника'] == worker]['Сумма']
                if not found.empty:
                    sals.append(float(found))
            data[i] = sals
            i += 1
    elif qual == 'Отдел' and quant == 'Сумма':
        quals = set(df[4]['Название'].tolist())
        data = [None] * len(quals)
        deps = df[4]['Код'].tolist()
        for i in range(len(deps)):
            workerIDs = df[1].loc[df[1]['Отделение'] ==
                                  int(deps[i])]['Код'].tolist()
            sals = []
            for worker in workerIDs:
                found = df[0].loc[df[0]['Код работника'] == worker]['Сумма']
                if not found.empty:
                    sals.append(float(found))
            data[i] = sals
    else:
        return None, None
    plt.close('all')
    fig, ax1 = plt.subplots(figsize=(8, 4))
    ax1.set_xlabel('$' + str(quant) + '$')
    ax1.set_ylabel('$Частота$')
    colors = ['red', 'tan', 'lime', 'grey', 'black', 'blue', 'cyan', 'magenta',
              'whitesmoke', 'yellow']
    try:
        ax1.hist(data, 10, density=False, histtype='bar', color=colors[:len(data)],
                 label=quals, edgecolor='black')
    except:
        message(root, 'Слишком много значений.\nМаксимум: 10', msgtype='warning').fade()
        return None, None
    ax1.legend(loc='upper left', bbox_to_anchor=(1, 1))
    ax1.set_title('Диаграмма $' + quant + '$ x $' + qual + '$')
    plt.tight_layout()
    
    filename = createUniqueFilename(['гист', quant, qual], '.png', directory)
    return fig, filename


def getQualityStatistics(root, window, df, directory):
    """
    Создание качественного отчета

    :param window: объект окна содержащего меню с параметрами
    :type window: MainWindow
    :param df: База данных
    :type df: pandas.DataFrame
    :param directory: путь к папке для сохранение
    :type directory: string
    :return: объект построенной диаграммы
    :rtype: matplotlib.pyplot.figure
    :return: путь к файлу с уникальным именем для сохранения
    :rtype: string
    :Автор(ы): Березуцкий
    """

    sumDf = getSummaryDf(df)
    qual = window.ComboQual.get()
    quals = sumDf[qual].tolist()
    
    columns = ('Переменная', 'Количество', 'Процент')

    fig, ax = plt.subplots(figsize =(12, 12))

    # hide axes
    fig.patch.set_visible(False)
    ax.axis('off')
    ax.axis('tight')
    
    elementCount = Counter(quals).most_common()
    count = len(quals)

    cellTable = []
    for i in elementCount:
        line = []
        line.append(i[0])
        line.append(i[1])
        line.append(i[1]*100/count)
        cellTable.append(line)

    ax.table(cellText=cellTable, colLabels=columns, cellLoc='center', loc='center')
    fig.tight_layout()


    filename = createUniqueFilename(['Качественная', qual], '.png', directory)
    return fig, filename


def getQuantStatistics(root, window, df, directory):
    """
    Создание количественного отчета

    :param window: объект окна содержащего меню с параметрами
    :type window: MainWindow
    :param df: База данных
    :type df: pandas.DataFrame
    :param directory: путь к папке для сохранение
    :type directory: string
    :return: объект построенной диаграммы
    :rtype: matplotlib.pyplot.figure
    :return: путь к файлу с уникальным именем для сохранения
    :rtype: string
    :Автор(ы): Березуцкий
    """
    plt.close('all')
    sumDf = getSummaryDf(df)
    
    quant = []
    tab = window.Data.index('current')
    for i in range(len(window.Cvars[tab])):
        if(window.Cvars[tab][i].get() == 1):
            quant.append(window.tables[tab]['columns'][i+1]) #код пропускается
    
    columns = ('Переменная', 'Минимум', 'Максимум', 'Среднее', 'Дисперсия', 'Стандартное отклонение')
    
    fig, ax = plt.subplots(figsize = (12,12))
    
    # hide axes
    fig.patch.set_visible(False)
    ax.axis('off')
    ax.axis('tight')
    
    cellTable = []
    for column in quant: #todo edit
        quants = sumDf[column].tolist()
        try:
            line = []
            line.append(column)
            line.append(round(min(quants), 2))  #min
            line.append(round(max(quants), 2))  #max
            line.append(round(np.mean(quants), 2))  #avg
            line.append(round(np.var(quants), 2))   #dispersion
            #line.append(round(sum((xi - np.mean(quants)) ** 2 for xi in quants)/len(quants)  , 2))
            line.append(round(np.std(quants), 2))   #stdDeviation
            cellTable.append(line)
        except TypeError:
            return

    ax.table(cellText=cellTable, colLabels=columns, cellLoc='center', loc='center')
    fig.tight_layout()
    
    filename = createUniqueFilename(['Количественная'], '.png', directory)
    
    return fig, filename


def getPivotStatistics(root, window, df, directory):
    """
    Создание сводной таблицы

    :param window: объект окна содержащего меню с параметрами
    :type window: MainWindow
    :param df: База данных
    :type df: pandas.DataFrame
    :param directory: путь к папке для сохранение
    :type directory: string
    :return: объект построенной диаграммы
    :rtype: matplotlib.pyplot.figure
    :return: путь к файлу с уникальным именем для сохранения
    :rtype: string
    :Автор(ы): Березуцкий
    """
    plt.close('all')
    sumDf = getSummaryDf(df)

    qual = window.ComboQual.get()
    quant = window.ComboQuant.get()
    
    tableDf = sumDf.pivot_table(quant, index=qual)
    
    columns = (qual, quant)
    cellTable = []

    for j in tableDf.index:
        line = [j]
        line.append(round(tableDf[tableDf.columns[0]][j], 2))
        cellTable.append(line)

    fig, ax = plt.subplots()

    fig.patch.set_visible(False)
    ax.axis('off')
    ax.axis('tight')
    
    ax.table(cellText=cellTable, colLabels=columns, cellLoc='center', loc='center')
    fig.tight_layout()
    filename = createUniqueFilename(['Сводная', qual, quant], '.png', directory)
    return fig, filename

def cutName(s):
    """
    Сокращение полного имени из двух или трех слов до фамилии с инициалами

    :param s: полное имя
    :type s: string
    :return: сокращенне имя
    :rtype: string
    :Автор(ы): Константинов
    """
    words = s.split()
    shortName = words[0]
    if len(words) > 1:
        shortName += (' ' + words[1][0] + '.')
    if len(words) > 2:
        shortName += (' ' + words[2][0] + '.')
    return shortName


def configureWidgets(scr, top):
    """
    Создание виджетов приложения

    :param scr: объект окна
    :type scr: MainWindow
    :param top: корневой объект
    :type top: tk.Tk
    :Автор(ы): Константинов, Сидоров, Березуцкий
    """
    scr.Pick_Analysis = tk.LabelFrame(top, bg=scr.config['def_frame_color'],
                                      fg=scr.config['def_frame_fg_color'])
    scr.Pick_Analysis.place(relx=0.023, rely=0.017, relheight=0.33,
                            relwidth=0.207)
    scr.Pick_Analysis.configure(text='Анализ')
    scr.Pick_Analysis.configure(cursor='arrow')

    scr.ComboAnalysis = ttk.Combobox(scr.Pick_Analysis,
                                     values=['Качественный параметр',
                                             'Количественный параметр',
                                             'Столбчатая диаграмма',
                                             'Гистограмма',
                                             'Диаграмма Бокса-Вискера',
                                             'Диаграмма рассеивания',
                                             'Сводная таблица'])
    scr.ComboAnalysis.place(relx=.05, rely=.35, height=20, relwidth=.9,
                            bordermode='ignore')

    scr.ShowAnalysisBtn = tk.Button(scr.Pick_Analysis,
                                    text='Показать', command=scr.showReport,
                                    bg=scr.config['def_btn_color'],
                                    fg=scr.config['def_btn_fg_color'])
    scr.ShowAnalysisBtn.place(relx=.048, rely=.5, height=32, relwidth=.9,
                              bordermode='ignore')
    scr.ShowAnalysisBtn.configure(cursor='hand2')

    scr.ExportAnalysisBtn = tk.Button(scr.Pick_Analysis,
                                      text='Экспорт', command=scr.exportReport,
                                      bg=scr.config['def_btn_color'],
                                      fg=scr.config['def_btn_fg_color'])
    scr.ExportAnalysisBtn.place(relx=.048, rely=.7, height=32, relwidth=.9,
                                bordermode='ignore')
    scr.ExportAnalysisBtn.configure(cursor='hand2')

    scr.Choice_Label = tk.Label(scr.Pick_Analysis,text='Вид отчета',
                                bg=scr.config['def_frame_color'],
                                fg=scr.config['def_frame_fg_color'])
    scr.Choice_Label.place(relx=.05, rely=.2, height=25, width=127,
                           bordermode='ignore')

    scr.Analysis_Frame = tk.LabelFrame(top, text='Параметры отчета',
                                       bg=scr.config['def_frame_color'],
                                       fg=scr.config['def_frame_fg_color'])
    scr.Analysis_Frame.place(relx=.24, rely=.017, relheight=.33,
                             relwidth=.201)

    scr.LabelQual = tk.Label(scr.Analysis_Frame, text='Качественный: ',
                             anchor='w', bg=scr.config['def_frame_color'],
                             fg=scr.config['def_frame_fg_color'])
    scr.LabelQual.place(relx=.05, rely=.2, height=25, width=127,
                        bordermode='ignore')

    scr.ComboQual = ttk.Combobox(scr.Analysis_Frame)
    scr.ComboQual.place(relx=.05, rely=.3, height=20, relwidth=.9,
                        bordermode='ignore')
    scr.ComboQual.configure(values=qualComboValues)

    scr.LabelQuant = tk.Label(scr.Analysis_Frame, text='Количественный: ',
                              anchor='w', bg=scr.config['def_frame_color'],
                              fg=scr.config['def_frame_fg_color'])
    scr.LabelQuant.place(relx=.05, rely=.4, height=25, width=127,
                         bordermode='ignore')

    scr.ComboQuant = ttk.Combobox(scr.Analysis_Frame)
    scr.ComboQuant.place(relx=.05, rely=.5, height=20, relwidth=.9,
                         bordermode='ignore')
    scr.ComboQuant.configure(values=quantComboValues)

    scr.LabelQuant2 = tk.Label(scr.Analysis_Frame, text='Количественный: ',
                               anchor='w', bg=scr.config['def_frame_color'],
                               fg=scr.config['def_frame_fg_color'])
    scr.LabelQuant2.place(relx=.05, rely=.6, height=25, width=127,
                          bordermode='ignore')

    scr.ComboQuant2 = ttk.Combobox(scr.Analysis_Frame)
    scr.ComboQuant2.place(relx=.05, rely=.7, height=20, relwidth=.9,
                          bordermode='ignore')
    scr.ComboQuant2.configure(values=quantComboValues)
    scr.ComboQuant2.configure(state='disabled')

    scr.Filter_Frame = tk.LabelFrame(top, text='Фильтры',
                                     bg=scr.config['def_frame_color'],
                                     fg=scr.config['def_frame_fg_color'])
    scr.Filter_Frame.place(relx=0.45, rely=0.017, relheight=0.33,
                           relwidth=0.532)

    scr.Data = ttk.Notebook(top)
    scr.Data.place(relx=0.023, rely=0.374, relheight=.528, relwidth=0.96)
    #  scr.Data.configure(takefocus="")

    scr.Data_t1 = tk.Frame(scr.Data)
    scr.Data.add(scr.Data_t1, padding=3)
    scr.Data.tab(0, text='Учёт')

    scr.Data_t2 = tk.Frame(scr.Data)
    scr.Data.add(scr.Data_t2, padding=3)
    scr.Data.tab(1, text='Работники')

    scr.Data_t3 = tk.Frame(scr.Data)
    scr.Data.add(scr.Data_t3, padding=3)
    scr.Data.tab(2, text='Должности')

    scr.Data_t4 = tk.Frame(scr.Data)
    scr.Data.add(scr.Data_t4, padding=3)
    scr.Data.tab(3, text='Информация')

    scr.Data_t5 = tk.Frame(scr.Data)
    scr.Data.add(scr.Data_t5, padding=3)
    scr.Data.tab(4, text='Отдел')

    #  configure filter lists
    scr.Filter_List1 = tk.Listbox(scr.Filter_Frame, exportselection=0)
    scr.Filter_List1.place(relx=0.019, rely=0.268, relheight=0.46,
                           relwidth=0.301, bordermode='ignore')

    scr.Filter_List2 = tk.Listbox(scr.Filter_Frame, exportselection=0)
    scr.Filter_List2.place(relx=0.338, rely=0.268, relheight=0.46,
                           relwidth=0.301, bordermode='ignore')

    scr.Filter_List1.insert('end', 'Отработано (ч)')
    scr.Filter_List1.insert('end', 'Дата выплаты')
    scr.Filter_List1.insert('end', 'Сумма')
    scr.Filter_List1.insert('end', 'Код работника')
    for i in range(4):
        scr.Filter_List2.insert('end', '')

    scr.Filter_scroll = tk.Scrollbar(scr.Filter_List1)
    scr.Filter_List1.config(yscrollcommand=scr.Filter_scroll.set)
    scr.Filter_List1.bind('<MouseWheel>', scr.scrollList2)
    scr.Filter_List2.config(yscrollcommand=scr.Filter_scroll.set)
    scr.Filter_List2.bind('<MouseWheel>', scr.scrollList1)

    scr.Change_Button = tk.Button(scr.Filter_Frame,
                                  bg=scr.config['def_btn_color'],
                                  fg=scr.config['def_btn_fg_color'])
    scr.Change_Button.place(relx=0.357, rely=0.804, height=32, width=148,
                            bordermode='ignore')
    scr.Change_Button.configure(cursor='hand2')
    scr.Change_Button.configure(text='Изменить значения',
                                command=scr.openDialog)

    scr.Reset_Button = tk.Button(scr.Filter_Frame, text='Сбросить выбор',
                                 command=scr.reset,
                                 bg=scr.config['def_btn_color'],
                                 fg=scr.config['def_btn_fg_color'])
    scr.Reset_Button.place(relx=0.03, rely=0.804, height=32, width=148,
                           bordermode='ignore')
    scr.Reset_Button.configure(cursor='hand2')

    scr.Param_Label = tk.Label(scr.Filter_Frame, text='Параметры',
                               bg=scr.config['def_frame_color'],
                               fg=scr.config['def_frame_fg_color'])
    scr.Param_Label.place(relx=0.075, rely=0.134, height=25, width=97,
                          bordermode='ignore')

    scr.Values_Label = tk.Label(scr.Filter_Frame, text='Значения',
                                bg=scr.config['def_frame_color'],
                                fg=scr.config['def_frame_fg_color'])
    scr.Values_Label.place(relx=0.414, rely=0.152, height=15, width=83,
                           bordermode='ignore')

    # Checkboxes
    scr.Boxes_Frame = tk.LabelFrame(scr.Filter_Frame, text='Столбцы',
                                    bg=scr.config['def_frame_color'],
                                    fg=scr.config['def_frame_fg_color'])
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
    scr.Cbox0 = tk.Checkbutton(scr.Boxes_Frame, command=scr.removeColumns,
                               bg=scr.config['def_frame_color'],
                               fg=scr.config['def_frame_fg_color'])
    scr.Cbox0.grid(row=0, column=0, sticky='W')
    scr.Cbox0.configure(justify='left')
    scr.Cbox0.configure(text='Отработано (ч)', variable=scr.Cvars0[0])
    scr.names.append(scr.Cbox0.cget('text'))
    scr.Cboxes0.append(scr.Cbox0)

    scr.Cbox1 = tk.Checkbutton(scr.Boxes_Frame, command=scr.removeColumns,
                               bg=scr.config['def_frame_color'],
                               fg=scr.config['def_frame_fg_color'])
    scr.Cbox1.grid(row=1, column=0, sticky='W')
    scr.Cbox1.configure(justify='left')
    scr.Cbox1.configure(text='Дата выплаты', variable=scr.Cvars0[1])
    scr.names.append(scr.Cbox1.cget('text'))
    scr.Cboxes0.append(scr.Cbox1)

    scr.Cbox2 = tk.Checkbutton(scr.Boxes_Frame, command=scr.removeColumns,
                               bg=scr.config['def_frame_color'],
                               fg=scr.config['def_frame_fg_color'])
    scr.Cbox2.grid(row=2, column=0, sticky='W')
    scr.Cbox2.configure(justify='left')
    scr.Cbox2.configure(text='Сумма', variable=scr.Cvars0[2])
    scr.names.append(scr.Cbox2.cget('text'))
    scr.Cboxes0.append(scr.Cbox2)

    scr.Cbox3 = tk.Checkbutton(scr.Boxes_Frame, command=scr.removeColumns,
                               bg=scr.config['def_frame_color'],
                               fg=scr.config['def_frame_fg_color'])
    scr.Cbox3.grid(row=3, column=0, sticky='W')
    scr.Cbox3.configure(justify='left')
    scr.Cbox3.configure(text='Код работника', variable=scr.Cvars0[3])
    scr.names.append(scr.Cbox3.cget('text'))
    scr.Cboxes0.append(scr.Cbox3)

    scr.Cbox4 = tk.Checkbutton(scr.Boxes_Frame, command=scr.removeColumns,
                               bg=scr.config['def_frame_color'],
                               fg=scr.config['def_frame_fg_color'])
    scr.Cbox4.grid(row=0, column=0, sticky='W')
    scr.Cbox4.configure(justify='left')
    scr.Cbox4.configure(text='Код должности', variable=scr.Cvars1[0])
    scr.Cbox4.grid_forget()
    scr.names.append(scr.Cbox4.cget('text'))
    scr.Cboxes1.append(scr.Cbox4)

    scr.Cbox5 = tk.Checkbutton(scr.Boxes_Frame, command=scr.removeColumns,
                               bg=scr.config['def_frame_color'],
                               fg=scr.config['def_frame_fg_color'])
    scr.Cbox5.grid(row=1, column=0, sticky='W')
    scr.Cbox5.configure(justify='left')
    scr.Cbox5.configure(text='Отделение', variable=scr.Cvars1[1])
    scr.Cbox5.grid_forget()
    scr.names.append(scr.Cbox5.cget('text'))
    scr.Cboxes1.append(scr.Cbox5)

    scr.Cbox6 = tk.Checkbutton(scr.Boxes_Frame, command=scr.removeColumns,
                               bg=scr.config['def_frame_color'],
                               fg=scr.config['def_frame_fg_color'])
    scr.Cbox6.grid(row=0, column=0, sticky='W')
    scr.Cbox6.configure(justify='left')
    scr.Cbox6.configure(text='Название', variable=scr.Cvars2[0])
    scr.Cbox6.grid_forget()
    scr.names.append(scr.Cbox6.cget('text'))
    scr.Cboxes2.append(scr.Cbox6)

    scr.Cbox7 = tk.Checkbutton(scr.Boxes_Frame, command=scr.removeColumns,
                               bg=scr.config['def_frame_color'],
                               fg=scr.config['def_frame_fg_color'])
    scr.Cbox7.grid(row=1, column=0, sticky='W')
    scr.Cbox7.configure(justify='left')
    scr.Cbox7.configure(text='Норма (ч/мес)', variable=scr.Cvars2[1])
    scr.Cbox7.grid_forget()
    scr.names.append(scr.Cbox7.cget('text'))
    scr.Cboxes2.append(scr.Cbox7)

    scr.Cbox8 = tk.Checkbutton(scr.Boxes_Frame, command=scr.removeColumns,
                               bg=scr.config['def_frame_color'],
                               fg=scr.config['def_frame_fg_color'])
    scr.Cbox8.grid(row=2, column=0, sticky='W')
    scr.Cbox8.configure(justify='left')
    scr.Cbox8.configure(text='Ставка (ч)', variable=scr.Cvars2[2])
    scr.Cbox8.grid_forget()
    scr.names.append(scr.Cbox8.cget('text'))
    scr.Cboxes2.append(scr.Cbox8)

    scr.Cbox9 = tk.Checkbutton(scr.Boxes_Frame, command=scr.removeColumns,
                               bg=scr.config['def_frame_color'],
                               fg=scr.config['def_frame_fg_color'])
    scr.Cbox9.grid(row=0, column=0, sticky='W')
    scr.Cbox9.configure(justify='left')
    scr.Cbox9.configure(text='ФИО', variable=scr.Cvars3[0])
    scr.Cbox9.grid_forget()
    scr.names.append(scr.Cbox9.cget('text'))
    scr.Cboxes3.append(scr.Cbox9)

    scr.Cbox10 = tk.Checkbutton(scr.Boxes_Frame, command=scr.removeColumns,
                                bg=scr.config['def_frame_color'],
                                fg=scr.config['def_frame_fg_color'])
    scr.Cbox10.grid(row=1, column=0, sticky='W')
    scr.Cbox10.configure(justify='left')
    scr.Cbox10.configure(text='Номер договора', variable=scr.Cvars3[1])
    scr.Cbox10.grid_forget()
    scr.names.append(scr.Cbox10.cget('text'))
    scr.Cboxes3.append(scr.Cbox10)

    scr.Cbox11 = tk.Checkbutton(scr.Boxes_Frame, command=scr.removeColumns,
                                bg=scr.config['def_frame_color'],
                                fg=scr.config['def_frame_fg_color'])
    scr.Cbox11.grid(row=2, column=0, sticky='W')
    scr.Cbox11.configure(justify='left')
    scr.Cbox11.configure(text='Телефон', variable=scr.Cvars3[2])
    scr.Cbox11.grid_forget()
    scr.names.append(scr.Cbox11.cget('text'))
    scr.Cboxes3.append(scr.Cbox11)

    scr.Cbox12 = tk.Checkbutton(scr.Boxes_Frame, command=scr.removeColumns,
                                bg=scr.config['def_frame_color'],
                                fg=scr.config['def_frame_fg_color'])
    scr.Cbox12.grid(row=3, column=0, sticky='W')
    scr.Cbox12.configure(justify='left')
    scr.Cbox12.configure(text='Образование', variable=scr.Cvars3[3])
    scr.Cbox12.grid_forget()
    scr.names.append(scr.Cbox12.cget('text'))
    scr.Cboxes3.append(scr.Cbox12)

    scr.Cbox13 = tk.Checkbutton(scr.Boxes_Frame, command=scr.removeColumns,
                                bg=scr.config['def_frame_color'],
                                fg=scr.config['def_frame_fg_color'])
    scr.Cbox13.grid(row=4, column=0, sticky='W')
    scr.Cbox13.configure(justify='left')
    scr.Cbox13.configure(text='Адрес', variable=scr.Cvars3[4])
    scr.Cbox13.grid_forget()
    scr.names.append(scr.Cbox13.cget('text'))
    scr.Cboxes3.append(scr.Cbox13)

    scr.Cbox14 = tk.Checkbutton(scr.Boxes_Frame, command=scr.removeColumns,
                                bg=scr.config['def_frame_color'],
                                fg=scr.config['def_frame_fg_color'])
    scr.Cbox14.grid(row=0, column=0, sticky='W')
    scr.Cbox14.configure(justify='left')
    scr.Cbox14.configure(text='Название', variable=scr.Cvars4[0])
    scr.Cbox14.grid_forget()
    scr.names.append(scr.Cbox14.cget('text'))
    scr.Cboxes4.append(scr.Cbox14)

    scr.Cbox15 = tk.Checkbutton(scr.Boxes_Frame, command=scr.removeColumns,
                                bg=scr.config['def_frame_color'],
                                fg=scr.config['def_frame_fg_color'])
    scr.Cbox15.grid(row=1, column=0, sticky='W')
    scr.Cbox15.configure(justify='left')
    scr.Cbox15.configure(text='Телефон', variable=scr.Cvars4[1])
    scr.Cbox15.grid_forget()
    scr.names.append(scr.Cbox15.cget('text'))
    scr.Cboxes4.append(scr.Cbox15)

    # menu
    menubar = tk.Menu(top)
    filemenu = tk.Menu(menubar, tearoff=0)
    filemenu.add_command(label='Новый', command=scr.newDatabase, accelerator='Ctrl+N')
    filemenu.add_command(label='Открыть', command=scr.open, accelerator='Ctrl+O')
    filemenu.add_command(label='Сохранить', command=scr.save, accelerator='Ctrl+S')
    filemenu.add_command(label='Сохранить как...', command=scr.saveas, accelerator='Ctrl+Shift+S')
    filemenu.add_command(label='Экспорт в xlsx', command=scr.saveAsExcel, accelerator='Ctrl+Q')
    filemenu.add_separator()
    filemenu.add_command(label='Выход', command=scr.exit, accelerator='Ctrl+Esc')
    menubar.add_cascade(label='Файл', menu=filemenu)

    helpmenu = tk.Menu(menubar, tearoff=0)
    helpmenu.add_command(label='Добавить', command=scr.addRecord, accelerator='Ctrl+Shift+A')
    helpmenu.add_command(label='Удалить', command=scr.deleteRecords, accelerator='Delete')
    helpmenu.add_command(label='Изменить', command=scr.modRecord, accelerator='Ctrl+R')
    menubar.add_cascade(label='Правка', menu=helpmenu)

    viewmenu = tk.Menu(menubar, tearoff=0)
    viewmenu.add_command(label='Пути и интерфейс', command=scr.customizeGUI, accelerator='Ctrl+P')
    menubar.add_cascade(label='Настройки', menu=viewmenu)

    top.config(menu=menubar)

    # status bar
    scr.statusbar = tk.Label(top, text="Oh hi. I didn't see you there...",
                             bd=1, relief=tk.SUNKEN, anchor=tk.W)
    scr.statusbar.pack(side=tk.BOTTOM, fill=tk.X)


def saveToPickle(filename, obj):
    """
    Сохранение объекта в бинарный файл pickle

    :param filename: путь к файлу для сохранения
    :type filename: string
    :param obj: объект для сохранения
    :type obj: любой объект
    :Автор(ы): Константинов
    """
    if (filename):
        db = open(filename, 'wb')
        pk.dump(obj, db)
        db.close()


class ChangeDialog(tk.Toplevel):
    """
    Класс диалога для изменения параметров фильтров
        :Автор(ы): Сидоров
    """
    def __init__(self, parent, config, prompt):
        """
        Инициализация диалогового окна

        :param parent: родительский объект
        :type parent: MainWindow или tk.Tk
        :param config: словарь настроек для получения цветов
        :type config: dict
        :param prompt: текст запроса
        :type prompt: string
        :Автор(ы): Сидоров
        """
        tk.Toplevel.__init__(self, parent)
        self.geometry('200x90+550+230')
        self.resizable(0, 0)
        self.title('')

        self.var = tk.StringVar()

        self.configure(bg=config['def_bg_color'])
        self.label = tk.Label(self, text=prompt, bg=config['def_bg_color'],
                              fg=config['def_frame_fg_color'])
        self.entry = tk.Entry(self, textvariable=self.var)
        self.ok_button = tk.Button(self, text='OK', command=self.Close)
        self.ok_button.configure(bg=config['def_btn_color'],
                                 fg=config['def_btn_fg_color'], cursor='hand2')
        self.ok_button.pack(side='bottom', fill='x', padx=60, pady=10)

        self.label.pack(side='top', fill='x')
        self.entry.pack(side='top', fill='x', padx=10)

        self.entry.bind('<Return>', self.Close)

    def Close(self, event=None):
        """
        Закрытие окна

        :Автор(ы): Сидоров
        """
        self.destroy()

    def show(self):
        """
        Отображение окна и получение данных после закрытия окна

        :return: введенное значение
        :rtype: string
        :Автор(ы): Сидоров
        """
        self.wm_deiconify()
        self.entry.focus_force()
        self.grab_set()
        self.wait_window()
        return self.var.get()


def testVal(inStr, acttyp):
    if acttyp == '1':  # insert
        if not inStr.isdigit():
            return False
    return True


class message(tk.Toplevel):
    """
    Класс всплывающего сообщения
        :Автор(ы): Константинов
    """
    def __init__(self, parent, prompt='Сообщение', msgtype='info'):
        """
        Инициализация окна сообщения

        :param parent: любой родительский объект
        :type parent: MainWindow или tk.Tk
        :param prompt: сообщение
        :type prompt: string
        :param msgtype: тип сообщения [warning, error, success, info]; по умолчанию info
        :type msgtype: string
        :Автор(ы): Константинов
        """
        self.opacity = 3.0
        tk.Toplevel.__init__(self, parent)
        self.resizable(0, 0)
        self.header = tk.Label(self, text='')
        self.header.pack(side='top', fill='x')
        self.label = tk.Label(self, text=prompt, justify='center')
        self.label.pack(side='top', fill='both', expand=1)
        geom = '200x80+' + str(parent.winfo_screenwidth()-260) + '+' + \
            str(parent.winfo_screenheight()-120)
        self.geometry(geom)
        self.setColor(msgtype)
        self.overrideredirect(True)

    def setColor(self, msgtype='info'):
        """
        Задача цвета элементам сообщения в соответствии со словарем colorDict

        :param msgtype: тип сообщения [warning, error, success, info]; по умолчанию info
        :type msgtype: string
        :Автор(ы): Константинов
        """
        self.header.configure(background=colorDict[msgtype][0])
        self.label.configure(background=colorDict[msgtype][1])
        self.configure(background=colorDict[msgtype][1])

    def fade(self):
        """
        Затухание окна
        :Автор(ы): Константинов
        """
        self.opacity -= 0.01
        if self.opacity <= 0.05:
            self.destroy()
            return
        self.wm_attributes('-alpha', self.opacity)
        self.after(10, self.fade)


class askValuesDialog(tk.Toplevel):
    """
    Класс диалога для ввода данных
    
    :Автор(ы): Константинов
    """
    def __init__(self, parent, config, labelTexts, currValues=None):
        """
        Инициализация окна диалога

        :param parent: любой родительский объект
        :type parent: MainWindow или tk.Tk
        :param config: словарь настроек для получения цветов
        :type config: dict
        :param labelTexts: имена вводимых параметров
        :type labelTexts: list
        :param currValues: текущие значения параметров
        :type currValues: list
        :Автор(ы): Константинов
        """
        tk.Toplevel.__init__(self, parent)
        self.parent = parent
        x = str(parent.winfo_screenwidth() // 2 - 150)
        y = str(parent.winfo_screenheight() // 2 - 200)
        self.title('Введите значения')
        self.geometry('300x400+' + x + '+' + y)
        self.resizable(0, 0)
        self.grab_set()  # make modal
        self.focus()
        self.Labels = [None] * len(labelTexts)
        self.Edits = [None] * len(labelTexts)
        self.retDict = dict()
        self.configure(background=config['def_bg_color'])
        for i in range(len(labelTexts)):
            self.retDict[labelTexts[i]] = tk.StringVar()
            editHeight = .8*400/len(labelTexts)
            self.Labels[i] = tk.Label(self, text=labelTexts[i]+':', anchor='e')
            self.Labels[i].configure(bg=config['def_bg_color'],
                                     fg=config['def_frame_fg_color'])
            self.Labels[i].place(relx=.1, y=40+i*editHeight, width=100)
            self.Edits[i] = tk.Entry(self,
                                     textvariable=self.retDict[labelTexts[i]],
                                     validate='key')
            if labelTexts[i] in quantParams:
                self.Edits[i].configure(validate='key')
                self.Edits[i].configure(validatecommand = (self.Edits[i].register(testVal),'%P','%d'))
            self.Edits[i].place(relx=.5, y=40+i*editHeight, width=100)
            if labelTexts[i] == 'Код':
                self.Edits[i].configure(state='disabled')
            if currValues:
                self.Edits[i].insert(0, currValues[i])

        self.ok_button = tk.Button(self, text='OK', command=self.on_ok)
        self.ok_button.configure(bg=config['def_btn_color'],
                                 fg=config['def_btn_fg_color'])
        self.ok_button.place(relx=.5, rely=.9, relwidth=.4,
                             height=30, anchor='c')

        self.bind('<Return>', self.on_ok)
        self.protocol('WM_DELETE_WINDOW', self.exit)

    def exit(self):
        """
        Закрытие окна диалога
        :Автор(ы): Константинов
        """
        self.retDict.clear()
        self.destroy()

    def on_ok(self, event=None):
        """
        Закрытие окна диалога с подтверждением данных
        :Автор(ы): Константинов
        """
        for edit in self.Edits[1:]:
            if edit.get().strip() == '':
                message(self.parent, 'Поля не могут быть пустыми.',
                        msgtype='warning').fade()
                return
        self.destroy()

    def show(self):
        """
        Получение данных после закрытия окна (вызвать для отображения)

        :return: словарь введенных значений параметров
        :rtype: dict
        :Автор(ы): Константинов
        """
        self.wm_deiconify()
        self.wait_window()
        return self.retDict


class CustomizeGUIDialog(tk.Toplevel):
    """
    Класс диалога настройки приложения
        :Автор(ы): Константинов
    """
    def __init__(self, parent):
        """
        Инициализация окна даилога

        :param parent: любой родительский объект
        :type parent: MainWindow или tk.Tk
        :Автор(ы): Константинов
        """
        tk.Toplevel.__init__(self, parent)
        self.parent = parent
        x = str(parent.winfo_screenwidth() // 2 - 150)
        y = str(parent.winfo_screenheight() // 2 - 200)
        self.geometry('500x500+' + x + '+' + y)
        self.title('Настройки')
        self.resizable(0, 0)
        self.grab_set()  # make modal
        self.focus()
        self.retDict = dict()
        self.config = getConfig()
        self.configure(bg=self.config['def_bg_color'])

        self.Frame = tk.LabelFrame(self)
        self.Frame.place(relx=.6, rely=.1, relheight=0.33,
                         relwidth=.5, anchor='n')
        self.Frame.configure(text='Предпросмотр', cursor='arrow')

        self.Label = tk.Label(self.Frame, text='Надпись')
        self.Label.place(relx=.5, rely=.2, anchor='c')
        self.Label.configure(bg=self.config['def_frame_color'],
                             fg=self.config['def_frame_fg_color'])

        self.ButtonTextSample = tk.Button(self.Frame, text='Кнопочка раз')
        self.ButtonTextSample.place(relx=.048, rely=.4, height=32, relwidth=.9,
                                    bordermode='ignore')
        self.ButtonTextSample.configure(cursor='hand2')

        self.ButtonBgSample = tk.Button(self.Frame, text='Кнопочка два')
        self.ButtonBgSample.place(relx=.048, rely=.7, height=32, relwidth=.9,
                                  bordermode='ignore')
        self.ButtonBgSample.configure(cursor='hand2')

        self.BtnBg = tk.Button(self, text='Приложение')
        self.BtnBg.place(relx=.04, rely=.1, height=30, width=130,
                         bordermode='ignore')
        self.BtnBg.configure(cursor='hand2',
                             command=lambda: self.pickColor(event='Bg'))

        self.BtnFrame = tk.Button(self, text='Раздел')
        self.BtnFrame.place(relx=.04, rely=.18, height=30, width=130,
                            bordermode='ignore')
        self.BtnFrame.configure(cursor='hand2',
                                command=lambda: self.pickColor(event='Frame'))

        self.BtnText = tk.Button(self, text='Текст')
        self.BtnText.place(relx=.04, rely=.26, height=30, width=130,
                           bordermode='ignore')
        self.BtnText.configure(cursor='hand2',
                               command=lambda: self.pickColor(event='Text'))

        self.BtnBgBtn = tk.Button(self, text='Фон кнопки')
        self.BtnBgBtn.place(relx=.04, rely=.34, height=30, width=130,
                            bordermode='ignore')
        self.BtnBgBtn.configure(cursor='hand2',
                                command=lambda: self.pickColor(event='BtnBg'))

        self.BtnTextBtn = tk.Button(self, text='Текст кнопки')
        self.BtnTextBtn.place(relx=.04, rely=.42, height=30, width=130,
                              bordermode='ignore')
        self.BtnTextBtn.configure(cursor='hand2',
                                  command=lambda: self.pickColor(event='BtnText'))

        self.BtnDBPath = tk.Button(self, text='База по умолчанию')
        self.BtnDBPath.place(relx=.04, rely=.5, height=30, width=130,
                             bordermode='ignore')
        self.BtnDBPath.configure(cursor='hand2',
                                 command=lambda: self.pickDir(event='db'))

        self.BtnGraphPath = tk.Button(self, text='Папка для диаграмм')
        self.BtnGraphPath.place(relx=.04, rely=.58, height=30, width=130,
                                bordermode='ignore')
        self.BtnGraphPath.configure(cursor='hand2', command=lambda: self.pickDir(event='graph'))

        self.BtnOutPath = tk.Button(self, text='Папка для отчетов')
        self.BtnOutPath.place(relx=.04, rely=.66, height=30, width=130,
                              bordermode='ignore')
        self.BtnOutPath.configure(cursor='hand2',
                                  command=lambda: self.pickDir(event='out'))

        self.fsvar = tk.IntVar()
        self.Boxfs = tk.Checkbutton(self, command=self.switchFs,
                                    text='Запускать в полноэкранном режиме',
                                    variable=self.fsvar)
        self.Boxfs.place(relx=.35, rely=.5, anchor='w')

        self.maxvar = tk.IntVar()
        self.Boxmax = tk.Checkbutton(self, command=self.switchMax, text='Запускать в режиме растянутого окна', variable=self.maxvar)
        self.Boxmax.place(relx=.35, rely=.6, anchor='w')

        self.BtnRestore = tk.Button(self, text='По умолчанию')
        self.BtnRestore.place(relx=.48, rely=.9, relwidth=.32, height=32,
                              bordermode='ignore', anchor='e')
        self.BtnRestore.configure(cursor='hand2', command=lambda: self.pickColor(event='RestoreDefaults'))

        self.ok_button = tk.Button(self, text='Сохранить изменения',
                                   command=self.on_ok)
        self.ok_button.place(relx=.52, rely=.9, relwidth=.32, height=32,
                             anchor='w')
        self.ok_button.configure(cursor='hand2')

        self.updateAll()
        self.bind('<Return>', self.on_ok)
        self.protocol('WM_DELETE_WINDOW', self.exit)

    def switchFs(self):
        """
        Запись параметра fullscreen в словарь настроек
        
        :Автор(ы): Константинов
        """
        self.config['fullscreen'] = str(int(self.fsvar.get()))

    def switchMax(self):
        """
        Запись параметра maximize в словарь настроек
        
        :Автор(ы): Константинов
        """
        self.config['maximize'] = str(int(self.maxvar.get()))

    def updateAll(self):
        """
        Обновление элементов окна диалога в соответствии с настоящим словарем настроек
        
        :Автор(ы): Константинов
        """
        self.configure(bg=self.config['def_bg_color'])
        self.Frame.configure(bg=self.config['def_frame_color'],
                             fg=self.config['def_frame_fg_color'])
        self.Label.configure(bg=self.config['def_frame_color'],
                             fg=self.config['def_frame_fg_color'])
        self.ButtonTextSample.configure(bg=self.config['def_btn_color'],
                                        fg=self.config['def_btn_fg_color'])
        self.ButtonBgSample.configure(bg=self.config['def_btn_color'],
                                      fg=self.config['def_btn_fg_color'])
        self.BtnBg.configure(bg=self.config['def_btn_color'],
                             fg=self.config['def_btn_fg_color'])
        self.BtnFrame.configure(bg=self.config['def_btn_color'],
                                fg=self.config['def_btn_fg_color'])
        self.BtnText.configure(bg=self.config['def_btn_color'],
                               fg=self.config['def_btn_fg_color'])
        self.BtnBgBtn.configure(bg=self.config['def_btn_color'],
                                fg=self.config['def_btn_fg_color'])
        self.BtnTextBtn.configure(bg=self.config['def_btn_color'],
                                  fg=self.config['def_btn_fg_color'])
        self.BtnRestore.configure(bg=self.config['def_btn_color'],
                                  fg=self.config['def_btn_fg_color'])
        self.ok_button.configure(bg=self.config['def_btn_color'],
                                 fg=self.config['def_btn_fg_color'])
        self.Boxfs.config(bg=self.config['def_bg_color'],
                          fg=self.config['def_frame_fg_color'])
        self.Boxmax.config(bg=self.config['def_bg_color'],
                           fg=self.config['def_frame_fg_color'])
        self.BtnDBPath.configure(bg=self.config['def_btn_color'],
                                 fg=self.config['def_btn_fg_color'])
        self.BtnGraphPath.configure(bg=self.config['def_btn_color'],
                                    fg=self.config['def_btn_fg_color'])
        self.BtnOutPath.configure(bg=self.config['def_btn_color'],
                                  fg=self.config['def_btn_fg_color'])

        self.fsvar.set(self.config['fullscreen'])
        self.maxvar.set(self.config['maximize'])

    def pickDir(self, event=None):
        """
        Выбор файлов и папок по нажатию на соответствующие кнопки
        
        :Автор(ы): Константинов
        """
        if event == 'db':
            path = filedialog.askopenfilename(filetypes=[('pickle files', '*.pickle'), ('Excel files', '*.xls *.xlsx')])
            self.config['def_db'] = path + '/'
        elif event == 'graph':
            path = filedialog.askdirectory()
            self.config['def_graph_dir'] = path + '/'
        elif event == 'out':
            path = filedialog.askdirectory()
            self.config['def_output_dir'] = path + '/'

    def pickColor(self, event=None):
        """
        Выбор цвета по нажатию на соответствующую кнопку
        
        :Автор(ы): Константинов
        """
        if event == 'BtnText':
            trash, color = colorchooser.askcolor(color=self.BtnText.cget('fg'))
            self.config['def_btn_fg_color'] = color
        elif event == 'BtnBg':
            trash, color = colorchooser.askcolor(color=self.BtnText.cget('bg'))
            self.config['def_btn_color'] = color
        elif event == 'Text':
            trash, color = colorchooser.askcolor(color=self.Label.cget('fg'))
            self.config['def_frame_fg_color'] = color
        elif event == 'Frame':
            trash, color = colorchooser.askcolor(color=self.Frame.cget('bg'))
            self.config['def_frame_color'] = color
        elif event == 'Bg':
            trash, color = colorchooser.askcolor(color=self.cget('bg'))
            self.config['def_bg_color'] = color
        elif event == 'RestoreDefaults':
            self.config = getDefaultConfig()
        self.updateAll()

    def exit(self):
        """
        Закрытие окна диалога
        
        :Автор(ы): Константинов
        """
        self.retDict.clear()
        self.destroy()

    def on_ok(self, event=None):
        """
        Закрытие окна диалога с подтверждением
        
        :Автор(ы): Константинов
        """
        writeConfig(self.config)
        message(self.parent, 'Сохранено\nИзменения будут применены\nпри следующем запуске\nприложения', msgtype='success').fade()
        self.destroy()

    def show(self):
        """
        Отображение окна диалога
        
        :Автор(ы): Константинов
        """
        self.wm_deiconify()
        self.wait_window()
        return
