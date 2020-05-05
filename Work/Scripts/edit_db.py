import pandas as pd
import pickle as pk

class AccountingDB:
    def __init__(self):
        xls = pd.ExcelFile('db.xlsx')
        self.df1 = pd.read_excel(xls, list(range(5)))
        
def refreshFromExcel():
    p = AccountingDB()
    db = open("db.pickle", "wb")
    pk.dump(p, db)
    db.close()
    
def main():
    refreshFromExcel()
    db = open("db.pickle","rb")
    p = pk.load(db)
    db.close()
    
    db = open("db.pickle", "wb")
    print(p.df1[0])
    cmd = ''
    while (cmd != "end"):
        print("\n[del/add/mod/save/end]")
        cmd = input()
        if cmd == "del":
            table = int(input("table: "))
            rows = list(map(int, input("rows: ").split()))
            p.df1[table] = p.df1[table].drop(rows)
            print(p.df1[table])
        elif cmd == "add":
            table = int(input("table: "))
            t = pd.DataFrame(columns = p.df1[table].columns)
            for c in t.columns:
                t.at[0, c] = input(c+': ')
            p.df1[table] = p.df1[table].append(t, ignore_index=True)
            print(p.df1[table])
        elif cmd == "mod":
            table = int(input("table: "))
            col = input("col: ")
            row = int(input("row: "))
            print(p.df1[table][col][row])
            nval = input("new value: ")
            p.df1[table].at[row, col] = nval
            print(p.df1[table])
        elif cmd == "save":
            pk.dump(p, db)
    db.close()

if __name__ == "__main__":
    main()