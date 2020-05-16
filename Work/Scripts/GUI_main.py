import sys
import tkinter as tk
import tkinter.ttk as ttk
import pandas as pd
import pickle as pk
from tkinter import filedialog

#  import main_support


def start_gui():
    """Starting point when module is the main routine."""
    global val, w, root
    root = tk.Tk()
    top = MainWindow(root)
    #  main_support.init(root, top)
    root.mainloop()


w = None


#def create_MainWindow(rt, *args, **kwargs):
#    """Starting point when module is imported by another module.
#       Correct form of call: 'create_MainWindow(root, *args, **kwargs)' ."""
#    w, w_win, root
#    #  rt = root
#    root = rt
#    w = tk.Toplevel(root)
#    top = MainWindow(w)
#    #  main_support.init(w, top, *args, **kwargs)
#    return (w, top)


def destroy_MainWindow():
    global w
    w.destroy()
    w = None


def refreshFromExcel():
    xls = pd.ExcelFile('../db.xlsx')  #  your repository
    p = pd.read_excel(xls, list(range(5)))
    saveToPickle("../db.pickle", p)


def saveToPickle(filename, obj):
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


class MainWindow:
    db = None
    currentFile = ''
    modified = False
    def __init__(self, top=None):
        """This class configures and populates the toplevel window.
           top is the toplevel containing window."""
        #  refreshFromExcel()  #  use once for db.pickle
        dbf = open("../Data/db.pickle", "rb")
        MainWindow.currentFile = "../Data/db.pickle"
        MainWindow.db = pk.load(dbf)
        dbf.close()

        top.geometry("1000x600+150+30")
        top.resizable(0, 0)
        top.title("База Данных")

        self.Table_Frame = tk.LabelFrame(top)
        self.Table_Frame.place(relx=0.023, rely=0.017, relheight=0.373,
                               relwidth=0.207)
        self.Table_Frame.configure(text='''Таблица''')
        self.Table_Frame.configure(cursor="arrow")

        self.Add_Button = tk.Button(self.Table_Frame)
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
        self.Analysis_Frame.place(relx=0.24, rely=0.017, relheight=0.373,
                                  relwidth=0.201)
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
        self.Change_Button.place(relx=0.357, rely=0.804, height=32, width=148,
                                 bordermode='ignore')
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
            cols = list(MainWindow.db[i].columns)
            self.tables[i].column("#0", minwidth=5, width=5, stretch=tk.NO)
            self.tables[i].heading("#0", text="")
            for j in range(0, len(cols)):
                self.tables[i].heading(cols[j], text=cols[j])
                self.Data.update()
                self.tables[i].column(cols[j], width=int((self.Data.winfo_width()-30)/(len(cols)-1)), stretch=tk.NO)
            self.tables[i].column(cols[0], width=30, stretch=tk.NO)
            for j in MainWindow.db[i].index:
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
        filemenu.add_command(label="New", command=self.menuFunc)
        filemenu.add_command(label="Open", command=self.open)
        filemenu.add_command(label="Save", command=self.save)
        filemenu.add_command(label="Save as...", command=self.saveas)
        filemenu.add_separator()
        filemenu.add_command(label="Exit", command=self.exit)
        menubar.add_cascade(label="File", menu=filemenu)
        filemenu.entryconfig("New", state="disabled")
        
        helpmenu = tk.Menu(menubar, tearoff=0)
        helpmenu.add_command(label="Add record", command=self.menuFunc)
        helpmenu.add_command(label="Delete record", command=self.menuFunc)
        helpmenu.add_command(label="Modify record", command=self.menuFunc)
        menubar.add_cascade(label="Edit", menu=helpmenu)

        top.config(menu=menubar)

        # status bar
        self.statusbar = tk.Label(top, text="Oh hi. I didn't see you there...", bd=1,
                             relief=tk.SUNKEN, anchor=tk.W)
        self.statusbar.pack(side=tk.BOTTOM, fill=tk.X)
        self.statusUpdate()
        
    def exit(self):
        print(CustomDialog(root, "Enter something:").show())
        
    def open(self):
        filename = filedialog.askopenfilename(filetypes = (("pickle files","*.pickle"),("Excel files", "*.xls *.xlsx")))
        print(filename)
        # TODO do you want to close? save? cancel?
        
    def save(self):
        pass
        # save to MainWindow.currentFile
        # TODO
        
    def saveas(self):
        filename = filedialog.askopenfilename(filetypes = (("pickle files","*.pickle"),("Excel files", "*.xls *.xlsx")))
        print(filename)
        # TODO
   
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

    def menuFunc(self):
        pass
    
    
class CustomDialog(tk.Toplevel):
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
        self.popup_menu.add_command(label="Delete",
                                    command=self.deleteRecords)
        self.bind("<Delete>", self.deleteRecords)
        self.bind("<Control-a>", self.selectAll)
        self.popup_menu.add_command(label="Select All",
                                    command=self.selectAll)
        self.bind("<Button-3>", self.popup)
        
    def popup(self, event):
        try:
            self.popup_menu.tk_popup(event.x_root, event.y_root, 0)
        finally:
            self.popup_menu.grab_release()
            
    def selectAll(self, event=None):
        self.selection_set(tuple(self.get_children()))  
            
    def deleteRecords(self, event=None):
        nb = self.master.master
        nb = nb.index(nb.select())
        selected = [int(i[1::], 16)-1 for i in self.selection()]
        MainWindow.db[nb] = MainWindow.db[nb].drop(selected)
        for item in self.selection():
            self.delete(item)

if __name__ == '__main__':
    start_gui()
