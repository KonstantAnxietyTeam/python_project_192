import sys
import tkinter as tk
import tkinter.ttk as ttk
import pandas as pd
import pickle as pk

#import main_support

def start_gui():
    """Starting point when module is the main routine."""
    global val, w, root
    root = tk.Tk()
    top = MainWindow (root)
    #main_support.init(root, top)
    root.mainloop()

w = None
def create_MainWindow(rt, *args, **kwargs):
    """Starting point when module is imported by another module.
       Correct form of call: 'create_MainWindow(root, *args, **kwargs)' ."""
    global w, w_win, root
    #rt = root
    root = rt
    w = tk.Toplevel (root)
    top = MainWindow (w)
    #main_support.init(w, top, *args, **kwargs)
    return (w, top)

def destroy_MainWindow():
    global w
    w.destroy()
    w = None

def refreshFromExcel():
    xls = pd.ExcelFile('../Data/db.xlsx')
    p = pd.read_excel(xls, list(range(5)))
    db = open("../Data/db.pickle", "wb")
    pk.dump(p, db)
    db.close()

class MainWindow:
    def addRecord(self):
        table = self.Data.index(self.Data.select())
        t = pd.DataFrame(columns = self.db[table].columns)
        for c in t.columns:
            t.at[0, c] = input(c+': ')
        self.db[table] = self.db[table].append(t, ignore_index=True)
        print(self.db[table])
        items = []
        for title in self.db[table].columns:
            items.append(self.db[table][title][len(self.db[table].index)-1])
        self.tables[table].insert("", "end", values=items)

    def __init__(self, top=None):
        """This class configures and populates the toplevel window.
           top is the toplevel containing window."""
        '''load db from pickle'''
        #refreshFromExcel()
        dbf = open("../Data/db.pickle", "rb")
        self.db = pk.load(dbf)
        dbf.close()
        
        top.geometry("1000x600+150+30")
        top.resizable(0, 0)
        top.title("База Данных")

        self.Table_Frame = tk.LabelFrame(top)
        self.Table_Frame.place(relx=0.023, rely=0.017, relheight=0.373
                , relwidth=0.207)
        self.Table_Frame.configure(text='''Таблица''')
        self.Table_Frame.configure(cursor="arrow")

        self.Add_Button = tk.Button(self.Table_Frame, command=self.addRecord)
        self.Add_Button.place(relx=0.048, rely=0.625, height=32, width=88,
                              bordermode='ignore')
        self.Add_Button.configure(cursor="hand2")
        self.Add_Button.configure(text='''Добавить''')

        self.ExpTab_Button = tk.Button(self.Table_Frame)
        self.ExpTab_Button.place(relx=0.531, rely=0.799, height=32, width=88,
                                 bordermode='ignore')
        self.ExpTab_Button.configure(cursor="hand2")
        self.ExpTab_Button.configure(text='''Экспорт''')

        self.Delete_Button = tk.Button(self.Table_Frame)
        self.Delete_Button.place(relx=0.048, rely=0.799, height=32, width=88,
                                 bordermode='ignore')
        self.Delete_Button.configure(cursor="hand2")
        self.Delete_Button.configure(text='''Удалить''')

        self.Edit_Button = tk.Button(self.Table_Frame)
        self.Edit_Button.place(relx=0.531, rely=0.625, height=32, width=88,
                               bordermode='ignore')
        self.Edit_Button.configure(cursor="hand2")
        self.Edit_Button.configure(text='''Правка''')

        self.Choice_Label = tk.Label(self.Table_Frame)
        self.Choice_Label.place(relx=0.386, rely=0.089, height=25, width=65,
                                bordermode='ignore')
        self.Choice_Label.configure(text='''Выбор''')

        self.Analysis_Frame = tk.LabelFrame(top)
        self.Analysis_Frame.place(relx=0.24, rely=0.017, relheight=0.373
                , relwidth=0.201)
        self.Analysis_Frame.configure(text='''Анализ''')

        self.Method_Label = tk.Label(self.Analysis_Frame)
        self.Method_Label.place(relx=0.199, rely=0.134, height=17, width=127,
                                bordermode='ignore')
        self.Method_Label.configure(text='''Метод Анализа''')

        self.Analysis_Button = tk.Button(self.Analysis_Frame)
        self.Analysis_Button.place(relx=0.05, rely=0.795, height=32, width=88,
                                   bordermode='ignore')
        self.Analysis_Button.configure(cursor="hand2")
        self.Analysis_Button.configure(text='''Анализ''')

        self.ExpAn_Button = tk.Button(self.Analysis_Frame)
        self.ExpAn_Button.place(relx=0.547, rely=0.795, height=32, width=78,
                                bordermode='ignore')
        self.ExpAn_Button.configure(cursor="hand2")
        self.ExpAn_Button.configure(text='''Экспорт''')

        self.Analysis_List = tk.Listbox(self.Analysis_Frame)
        self.Analysis_List.place(relx=0.05, rely=0.268, relheight=0.46,
                                 relwidth=0.871, bordermode='ignore')

        self.Table_List = tk.Listbox(top)
        self.Table_List.place(relx=0.033, rely=0.097, relheight=0.128,
                              relwidth=0.188)

        self.Filter_Frame = tk.LabelFrame(top)
        self.Filter_Frame.place(relx=0.45, rely=0.017, relheight=0.373,
                                relwidth=0.532)
        self.Filter_Frame.configure(text='''Фильтры''')

        self.Filter_List1 = tk.Listbox(self.Filter_Frame)
        self.Filter_List1.place(relx=0.019, rely=0.268, relheight=0.46,
                                relwidth=0.301, bordermode='ignore')

        self.Filter_List2 = tk.Listbox(self.Filter_Frame)
        self.Filter_List2.place(relx=0.338, rely=0.268, relheight=0.46,
                                relwidth=0.301, bordermode='ignore')

        self.Change_Button = tk.Button(self.Filter_Frame)
        self.Change_Button.place(relx=0.357, rely=0.804, height=32, width=148
                , bordermode='ignore')
        self.Change_Button.configure(cursor="hand2")
        self.Change_Button.configure(text='''Изменить значения''')

        self.Reset_Button = tk.Button(self.Filter_Frame)
        self.Reset_Button.place(relx=0.677, rely=0.800, height=32, width=148,
                                bordermode='ignore')
        self.Reset_Button.configure(cursor="hand2")
        self.Reset_Button.configure(text='''Сбросить выбор''')

        self.Param_Label = tk.Label(self.Filter_Frame)
        self.Param_Label.place(relx=0.075, rely=0.134, height=25, width=97,
                               bordermode='ignore')
        self.Param_Label.configure(text='''Параметры''')

        self.Values_Label = tk.Label(self.Filter_Frame)
        self.Values_Label.place(relx=0.414, rely=0.152, height=15, width=83,
                                bordermode='ignore')
        self.Values_Label.configure(text='''Значения''')

        self.Columns_Label = tk.Label(self.Filter_Frame)
        self.Columns_Label.place(relx=0.752, rely=0.134, height=24, width=86,
                                 bordermode='ignore')
        self.Columns_Label.configure(text='''Столбцы''')

        self.Filter_List3 = tk.Listbox(self.Filter_Frame)
        self.Filter_List3.place(relx=0.658, rely=0.268, relheight=0.46,
                                relwidth=0.32, bordermode='ignore')
        self.Filter_List3.configure(foreground="#000000")

        self.Filter_Button = tk.Button(self.Filter_Frame)
        self.Filter_Button.place(relx=0.038, rely=0.804, height=32, width=148,
                                 bordermode='ignore')
        self.Filter_Button.configure(cursor="hand2")
        self.Filter_Button.configure(text='''Фильтровать''')

        self.Data = ttk.Notebook(top)
        self.Data.place(relx=0.03, rely=0.417, relheight=0.558, relwidth=0.952)
        #self.Data.configure(takefocus="")

        self.Data_t1 = tk.Frame(self.Data)
        self.Data.add(self.Data_t1, padding=3)
        self.Data.tab(0, text="Товары")

        self.Data_t2 = tk.Frame(self.Data)
        self.Data.add(self.Data_t2, padding=3)
        self.Data.tab(1, text="Компоненты")

        self.Data_t3 = tk.Frame(self.Data)
        self.Data.add(self.Data_t3, padding=3)
        self.Data.tab(2, text="Производители")

        self.Data_t4 = tk.Frame(self.Data)
        self.Data.add(self.Data_t4, padding=3)
        self.Data.tab(3, text="Полный список")
        
        tabs = [self.Data_t1, self.Data_t2, self.Data_t3, self.Data_t4]
        self.tables = [1, 2, 3, 4]
        for i in range(len(tabs)):
            self.tables[i] = ttk.Treeview(tabs[i])
            self.tables[i].place(relwidth=1.0, relheight=1.0)
            self.tables[i]["columns"] = list(self.db[i].columns)
            self.tables[i]['show'] = 'headings'
            cols = list(self.db[i].columns)
            self.tables[i].column("#0", width=0, minwidth=0)
            self.tables[i].heading("#0", text="")
            for j in range(0, len(cols)):
                self.tables[i].heading(cols[j], text=cols[j])
                self.tables[i].column(cols[j], width=5)
            for j in self.db[i].index:
                items = []
                for title in self.db[i].columns:
                    items.append(self.db[i][title][j])
                self.tables[i].insert("", "end", values=items)
                

if __name__ == '__main__':
    start_gui()
