import sqlite3 as sq
import os
import openpyxl

#db start
con = sq.connect('customers.db')
cur = con.cursor()
path = r"C:\Users\Vadym\Documents\projects\dealers_customers\Proza+Liman+kvitka.xlsx"
book = openpyxl.open(path)
sheet = book.active



name = ''
egrpou = 0
dealer_code = ''

for i in range(2, sheet.max_row + 1):
    print(f"Завантажую {i}")
    if sheet[i][0].value == None or sheet[i][0].value == "":
        name = "none"
    else:
        name = sheet[i][0].value.strip()
    egrpou = sheet[i][1].value
    dealer_code = sheet[i][2].value.strip()

    cur.execute("INSERT INTO customers(name, egrpou, dealer_code) VALUES(:name, :egrpou, :dealer_code)", {'name': name, 'egrpou': egrpou,
                                                                                                      'dealer_code': dealer_code})
con.commit()


#closing DB
con.close()
