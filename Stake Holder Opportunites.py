'''
This program will find the Stake Holder Opportunites.

Note: -> Key Contact, Projects and Team Countries Exports mandatory
'''


# Importing Libraries
import sqlite3
import pandas as pd
import win32com.client
import xlwings as xw
import tkinter as tk
from tkinter import filedialog
from tkinter import messagebox
from datetime import datetime
import time
import os

# Filedialog box
fPath = []
for x in ['Key Contact Export','Projects Export','CIC_EmpDetails_TeamCountries.xlsx']:
    root = tk.Tk()
    root.withdraw()
    fPath.append(filedialog.askopenfilename(title='Select {} file'.format(x), filetypes=(('Excel','*xlsx'),('All','*.*'))))

# MessageBox
root = tk.Tk()
root.withdraw()
infobox = tk.messagebox.askyesno(title="Message Box", message='''Do you want to Exclude "Planning Authority" ?''', icon ='warning')

def DelDatabase():
    try:
        os.remove('KeyContact.db')
        print('KeyContact.db Deleted')
    except:
        pass
        print('Pass')

DelDatabase()

if "/Key Contacts Export_" in fPath[0]:
    sTime = datetime.now()
    wb = xw.Book(fPath[0])
    print('Key Contacts Export workbook opened {}'.format(datetime.now()-sTime))
    sht = wb.sheets('Sheet1')
    try:
        sht.api.ShowAllData()
    except:
        pass
    sht.range('B:C,E:G,K:V,X:AF').api.Delete()
    sht.range('J1').value = "=COUNT($A:$A)+1"
    col1 = ['ProjectId', 'ProjectStatus', 'KeyContact', 'ContactType', 'KeypersonName', 'companyID']
    if sht.range('A1:F1').value == col1:
        print('Same Columns, redy to Insert Data')
        db = sqlite3.connect('KeyContact.db')
        print ('DataBase Opened')

        db.execute('''create table if not exists Table1 ({} integer, "{}" text, "{}" text, "{}" text, "{}" text, {} integer)'''.format(*col1))
        time.sleep(0.5)
        print('Table1 Created')

        sTime = datetime.now()
        for rows in sht.range('A2:F{}'.format(int(sht.range('J1').value))).value:
            db.execute('''insert into Table1 values(?,?,?,?,?,?)''', tuple(rows))
        print ('Inserting Data to SQL Completed {}'.format(datetime.now()-sTime))
        wb.close()
        time.sleep(0.5)

        db.execute('''delete from Table1 where ProjectStatus="Construction Complete" or ProjectStatus="Cancelled"''')
        time.sleep(0.5)
        print("Deleted 'Construction Completed' & 'Cancelled' from 'ProjectStage'")

        if infobox:
            db.execute('''delete from Table1 where KeyContact="Planning Authority"''')
            time.sleep(0.5)
            print("Deleted 'Planning Authority' from 'KeyContact''")

        db.execute('''create table Table2 as select *, ProjectId || companyID as "crit" from Table1 where KeypersonName is not null''')
        time.sleep(0.5)
        print("Table2 with 'KeypersonName' is not null")

        db.execute('''create table Table3 as select *, ProjectId || companyID as "crit" from Table1 where KeypersonName is null''')
        time.sleep(0.5)
        print("Table3 with 'KeypersonName' is null")

        db.execute('''create table Table4 as 
        select Table3.ProjectId, Table3.ProjectStatus, Table3.KeyContact, Table3.ContactType, Table3.companyID from Table3
        left join Table2 on Table3.crit = Table2.crit
        where Table2.ProjectId is null''')
        time.sleep(0.5)
        print("Table4 with Zero Stake Holders Opportunites")

        db.execute('''create table Table5 as select * from Table4 group by Table4.ProjectId, Table4.companyID, Table4.KeyContact''')
        time.sleep(0.5)
        print("Table5 with unique values on ProjectId,CompanyID,KeyContact")

        Status1 = True

    else:
        print('Different Columns')
        wb.close()
time.sleep(1)

if "/ProjectsExport_" in fPath[1] and Status1:
    sTime = datetime.now()
    wb = xw.Book(fPath[1])
    print ('Projects Export Workbook Opened {}'.format(datetime.now()-sTime))
    # app = xw.apps.active
    sht = wb.sheets('Sheet1')
    try:
        sht.api.ShowAllData()
    except:
        pass
    sht.range("B:K,M:AO").api.Delete()
    sht.range('J1').value = "=COUNT($A:$A)+1"
    col2 = ['ProjectID','Country']
    if sht.range('A1:B1').value == col2:
        print('Same Columns, redy to Insert Data')

        # db = sqlite3.connect('KeyContact.db')
        # print('DataBase Opened')
        db.execute('''create table Table6 ({} integer, "{}" text)'''.format(*col2))
        time.sleep(0.5)
        print('Table6 Created')

        sTime = datetime.now()
        for rows in sht.range('A2:B{}'.format(int(sht.range('J1').value))).value:
            db.execute('''insert into Table6 values(?,?)''', tuple(rows))
        print ('Inserting Data to SQL Completed {}'.format(datetime.now()-sTime))
        wb.close()
        time.sleep(0.5)

        Status2 = True
    else:
        print('Different Columns')
        wb.close()
time.sleep(1)

if "CIC_EmpDetails_TeamCountries.xlsx" in fPath[2] and Status1 and Status2:
    sTime = datetime.now()
    xlApp = win32com.client.Dispatch("Excel.Application")
    xlwb = xlApp.Workbooks.Open(fPath[2], False, True, None, "cicemp2k18")
    xlApp.Visible = True
    wb = xw.books.active
    print ('Excel Workbook Opened {}'.format(datetime.now()-sTime))
    # app = xw.apps.active
    sht = wb.sheets('Team Countries')
    sht.range('Q1').value = "=COUNTA($J:$J)"
    col3 = ['Country', 'Team']
    if sht.range('J1:K1').value == col3:
        print('Same Columns, redy to Insert Data')
        # db = sqlite3.connect('KeyContact.db')
        # print('DataBase Opened')
        db.execute('''create table Table7 ("{}" text, "{}" text)'''.format(*col3))
        time.sleep(0.5)
        print('Table3 Created')

        sTime = datetime.now()
        for rows in sht.range('J2:K{}'.format(int(sht.range('Q1').value))).value:
            db.execute('''insert into Table7 values(?,?)''', tuple(rows))
        print ('Inserting Data to SQL Completed {}'.format(datetime.now()-sTime))
        wb.close()
        time.sleep(0.5)

        db.execute('''create table Table8 as 
        select Table5.ProjectId, Table5.ProjectStatus, Table5.KeyContact, Table5.ContactType, Table5.companyID, Table6.Country, Table7.Team from Table5
        join Table6 on Table5.ProjectId = Table6.ProjectID
        join Table7 on Table6.Country = Table7.Country''')
        time.sleep(0.5)
        print('Table8 with final result')

        Status3 = True
    else:
        print('Different Columns')
        wb.close()
time.sleep(1)

if Status1 and Status2 and Status3:
    # db = sqlite3.connect('KeyContact.db')
    # print('DataBase Opened')
    xw.view(pd.read_sql_query('''select * from Table8''',db))
    wb = xw.books.active
    sht = wb.sheets.active
    sht.range('A:A').api.Delete()
    sht.range('A1:G1').api.Font.Bold = True
    if infobox:
        sht.range('I1').value='''Stake Holder Opportunities by Company Role Excluding "Planning Authority"'''
        sht.range('I1').api.Font.Bold = True
        sht.range('I1').api.Font.Size = 16
    else:
        sht.range('I1').value = '''Stake Holder Opportunities by Company Role Including "Planning Authority"'''
        sht.range('I1').api.Font.Bold = True
        sht.range('I1').api.Font.Size = 16
    sht.range('J2').value = 'Key Contact Export : {}'.format(os.path.basename(fPath[0]))
    sht.range('J3').value = 'Daily Export : {}'.format(os.path.basename(fPath[1]))
    sht.range('A:G').column_width = 10.71
    db.close()
    print('DataBase Closed')
    DelDatabase()
