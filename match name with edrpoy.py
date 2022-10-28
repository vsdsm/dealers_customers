import sqlite3 as sq
import os
import openpyxl

#db start
con = sq.connect('customers.db')
cur = con.cursor()
path = r"C:\Users\Vadym\Documents\projects\dealers_customers\Trash.xlsx"
book = openpyxl.open(path)
sheet = book.active



name = ''
egrpou = 0

for i in range(2, sheet.max_row + 1):
    print(f"Завантажую {i}")
    if sheet[i][1].value == None or sheet[i][1].value == "":
        name = "none"
    else:
        name = sheet[i][1].value.strip()
    egrpou = sheet[i][0].value
    if sheet[i][0].value == None or sheet[i][0].value == '':
        continue
    cur.execute("UPDATE customers SET name = :new_name WHERE egrpou = :egrpou", {'new_name': name, 'egrpou': egrpou})
con.commit()


#closing DB
con.close()
