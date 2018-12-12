'''
This program will find the Zero Contacts in an company.

Note: -> Key Contact, Projects and Team Countries Exports mandatory
'''


# Importing Libraries
import sqlite3
import xlwings as xw
import tkinter as tk
from tkinter import filedialog
from datetime import datetime
import win32com.client
import pandas as pd

# FileDialog Box
root = tk.Tk()
root.withdraw()
iDir = "T:\\WMI\\WMI Projects & Tender\\Construction Project\\2018\\Export\\Key Contact Export"
fPath1 = filedialog.askopenfilename(initialdir=iDir,title='Select Key Contact Export file', filetypes=(('Excel','*.xlsx'),('All','*.*')))

# FileDialog Box
root = tk.Tk()
root.withdraw()
iDir = "T:\\WMI\\WMI Projects & Tender\\Construction Project\\2018\\Export\\Daily Export"
fPath2 = filedialog.askopenfilename(initialdir=iDir,title='Select Projects Export file', filetypes=(('Excel','*.xlsx'),('All','*.*')))
fPath3 = "T:\\WMI\\WMI Projects & Tender\\Construction Project\\2018\\Management Trackers\\Stnd_Ref\\CIC_EmpDetails_TeamCountries.xlsx"

if "/Key Contacts Export_" in fPath1:
    sTime = datetime.now()
    wb = xw.Book(fPath1)
    print ('Excel Workbook Opened {}'.format(datetime.now()-sTime))
    app = xw.apps.active
    sht = wb.sheets('Sheet1')
    sht.range("B:C,E:G,K:V,X:AF").api.Delete()
    sht.range('J1').value = "=COUNT($A:$A)+1"
    if sht.range('A1:F1').value == ['ProjectId','ProjectStatus','KeyContact','ContactType','KeypersonName','companyID']:
        print('Same Columns, redy to Insert Data')
        db = sqlite3.connect('T:/WMI\WMI Projects & Tender/Construction Project/2018/Management Trackers/Stnd_Ref/KeyContact.db')
        print ('DataBase Opened')
        sTime = datetime.now()
        for rows in sht.range('A2:F{}'.format(int(sht.range('J1').value))).value:
            db.execute('''insert into Table1 values(?,?,?,?,?,?)''', tuple(rows))
        print ('Inserting Data to SQL Completed {}'.format(datetime.now()-sTime))
        app.quit()

        db.execute('''delete from Table1 where ProjectStatus="Construction Complete" or ProjectStatus="Cancelled"''')
        print("Deleted 'Construction Completed' & 'Cancelled' from 'ProjectStage'")

        db.execute('''delete from Table1 where KeyContact="Planning Authority"''')
        print("Deleted 'Planning Authority' from 'KeyContact''")

        db.execute('''update Table1 set KeypersonName = "Yes" where KeypersonName is not null''')
        print("Replaced all Non Empty cells to 'Yes' in 'KeypersonName'")

        db.execute('''update Table1 set KeypersonName = "No" where KeypersonName is null''')
        print("Replaced all Empty cells to 'No' in 'KeypersonName'")

        db.execute('''CREATE view view1 as SELECT * FROM Table1 GROUP BY ProjectId,CompanyID''')
        print("Removed Duplicates on ProjectId,CompanyID and Created view1")

        db.execute('''create view view2 as select ProjectID, count(*) as "TSH" from view1 group by ProjectID''')
        print("view2 with Total Stake Holders")

        db.execute('''create view view3 as select ProjectID, count(*) as "TNullSH" from view1 where KeypersonName = "No" group by ProjectID''')
        print("view3 with Total Null Stake Holders")

        db.execute('''create view view4 as select view2.ProjectId, view2.TSH, view3.TNullSH, (view2.TSH - view3.TNullSH) as "zCheck" from view2 left join view3 on view2.ProjectId = view3.ProjectId where zCheck=0''')
        print("view4 with Check column")

        db.commit()
        print ('Changes Committed')

        db.close()
        print('DataBase Closed')
        Status1 = True
    else:
        print('Different Columns')
        app.quit()

if "/ProjectsExport_" in fPath2:
    sTime = datetime.now()
    wb = xw.Book(fPath2)
    print ('Excel Workbook Opened {}'.format(datetime.now()-sTime))
    app = xw.apps.active
    sht = wb.sheets('Sheet1')
    sht.range("B:K,M:AO").api.Delete()
    sht.range('J1').value = "=COUNT($A:$A)+1"
    if sht.range('A1:B1').value == ['ProjectID','Country']:
        print('Same Columns, redy to Insert Data')
        db = sqlite3.connect('T:/WMI\WMI Projects & Tender/Construction Project/2018/Management Trackers/Stnd_Ref/KeyContact.db')
        print('DataBase Opened')

        sTime = datetime.now()
        for rows in sht.range('A2:B{}'.format(int(sht.range('J1').value))).value:
            db.execute('''insert into Table2 values(?,?)''', tuple(rows))
        print ('Inserting Data to SQL Completed {}'.format(datetime.now()-sTime))
        app.quit()

        db.execute('''create table Table4 as select view4.ProjectId as "ProjectId",Table2.Country as "Country" from view4 left join Table2 on view4.ProjectId=Table2.ProjectID where Table2.Country is not null''')
        print("Table5 with Country")

        db.commit()
        print ('Changes Committed')

        db.close()
        print('DataBase Closed')
        Status2 = True
    else:
        print('Different Columns')
        app.quit()

if Status1 and Status2:
    sTime = datetime.now()
    xlApp = win32com.client.Dispatch("Excel.Application")
    xlwb = xlApp.Workbooks.Open(fPath3, False, True, None, "cicemp2k18")
    xlApp.Visible = True
    wb = xw.books.active
    print ('Excel Workbook Opened {}'.format(datetime.now()-sTime))
    app = xw.apps.active
    sht = wb.sheets('Team Countries')
    sht.range('Q1').value = "=COUNTA($J:$J)"
    if sht.range('J1:K1').value == ['Country','Team']:
        print('Same Columns, redy to Insert Data')
        db = sqlite3.connect('T:/WMI\WMI Projects & Tender/Construction Project/2018/Management Trackers/Stnd_Ref/KeyContact.db')
        print('DataBase Opened')

        sTime = datetime.now()
        for rows in sht.range('J2:K{}'.format(int(sht.range('Q1').value))).value:
            db.execute('''insert into Table3 values(?,?)''', tuple(rows))
        print ('Inserting Data to SQL Completed {}'.format(datetime.now()-sTime))
        app.quit()

        db.execute('''create table Table5 as select Table4.ProjectId as "ProjectId", Table4.Country as "Country", Table3.Team as "Team" from Table4 join Table3 on Table4.Country = Table3.Country''')
        print("Table5 with Team")

        db.commit()
        print ('Changes Committed')

        db.close()
        print('DataBase Closed')
        Status3 = True
    else:
        print('Different Columns')
        app.quit()

if Status1 and Status2 and Status3:
    db = sqlite3.connect('T:/WMI\WMI Projects & Tender/Construction Project/2018/Management Trackers/Stnd_Ref/KeyContact.db')
    print('DataBase Opened')
    wb = xw.Book()
    sht = wb.sheets('Sheet1')
    sht.range('A1').value= pd.read_sql_query('''select * from Table5''',db)
    db.close()
    print('DataBase Closed')
