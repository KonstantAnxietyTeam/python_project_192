import pandas as pd
import pickle as pk

class AccountingDB:
    def __init__(self):
        xls = pd.ExcelFile('D:\ВШЭ МИЭМ ИВТ\Python\project1\db.xlsx')
        #self.df1 = pd.read_excel(xls, 'accounting')
        #self.df2 = pd.read_excel(xls, 'employee')
        #self.df3 = pd.read_excel(xls, 'info')
        #self.df4 = pd.read_excel(xls, 'position')
        #self.df5 = pd.read_excel(xls, 'department')
        self.df1 = pd.read_excel(xls, list(range(5)))

def main():
    p = AccountingDB()
    #db = open("db.pickle","wb")
    #pk.dump(p, db)
    #db.close()
    
    #db = open("db.pickle","rb")
    #p = pk.load(db)
    
    print(p.df1[0])
    print()
    print(p.df1[1]["Код"][1])
    
if __name__ == "__main__":
    main()