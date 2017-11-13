import pandas as pd
import tkinter as tk
from tkinter.filedialog import askopenfilename
from tkinter import ttk as tws
import csv
import sqlalchemy
import urllib
import pyodbc
from datetime import datetime
import os

class MainGui:

    recordArray = []
    recordCount = 1
    dataDictionary = {"Excel":[],"SQL_Connection":[],"SQL_Script":[]}

    def __init__(self,master):
        self.master = master
        master.title("Excel To SQL")
        
        connectionsFrame = tk.Frame(self.master)
        connectionsFrame.grid()
        cT = self.connectionsTab(self.recordArray, self.recordCount, self.dataDictionary, self.master)
        cT.Main(connectionsFrame)

        finishedButton = tk.Button(master,text="Finished", command=self.finished)
        finishedButton.grid()

    def finished(self):

        engine = None
        strDatabaseName = self.dataDictionary["SQL_Connection"][0][2]
        strDatabaseUserName = self.dataDictionary["SQL_Connection"][0][0]
        strDatabasePassword = self.dataDictionary["SQL_Connection"][0][1]
        strServerLocation = self.dataDictionary["SQL_Connection"][0][3]
        strServerPort = self.dataDictionary["SQL_Connection"][0][4]

        objConn = "mssql+pymssql://{0}:{1}@{3}:{4}/{2}".format(strDatabaseUserName, strDatabasePassword, strDatabaseName, strServerLocation, strServerPort)

        try:
            
            engine = sqlalchemy.create_engine(objConn)
            cnxn = engine.connect()

            print(objConn)

            for item in self.dataDictionary['Excel']:
                xl = pd.ExcelFile(item[0])
                df = xl.parse(item[1])
                df.to_sql(item[2], engine, if_exists='replace')

            for item in self.dataDictionary['SQL_Script']:
                strFileName = ''
                p = finishedPopup(self.master,strFileName, item)
                self.master.wait_window(p.top)
                strFileName = p.strFileName
                print(strFileName)
                strDirectory = tk.filedialog.askdirectory()
                datestring = datetime.strftime(datetime.now(), '%Y-%m-%d_%H-%M-%S')
                strFullFileName = strFileName + '_' + datestring + '.csv'
                filePath = os.path.join(strDirectory, strFullFileName)
                
                with open(item, 'r') as fd:
                    strSQL = fd.read()
                    results = cnxn.execute(strSQL)

                    outfile = open(filePath ,'w')
                    outcsv = csv.writer(outfile)
                    columnNames = results.keys()

                    outcsv.writerow(columnNames)
                                       
                    for row in results:
                        outcsv.writerow(row)
                    outfile.close()
                
        except Exception as e:
            print(str(e))

        finally:
            cnxn.close()
        
            

    class connectionsTab:

        dataEntryFrame = None
        SQLScripts = False
        
        def __init__(self, recordArray, recordCount, dataDictionary, master):

            self.recordArray = recordArray
            self.recordCount = recordCount
            self.dataDictionary = dataDictionary
            self.master = master
        
        def Main(self,mainFrame):
            
            self.locationLabel = tk.Label(mainFrame, text="Enter Database Name")
            self.locationLabel.grid(row=2, column=0, stick=tk.W)

            self.entryFrame = tk.Frame(mainFrame, width= 200, height=25)
            self.entryFrame.grid(row=2, column=1, columnspan=2, stick=tk.W)
            self.entryFrame.columnconfigure(0, weight=10)
            self.entryFrame.grid_propagate(False)

            self.locationValue = tk.StringVar()
            self.locationEntry = tk.Entry(self.entryFrame)
            self.locationEntry.grid(sticky="we")

            self.pathPickerFrame = tk.Frame(mainFrame)
            self.pathPickerFrame.grid(row=2, column=1, columnspan=2, stick=tk.E)

            self.pathPickerButton = tk.Button(self.pathPickerFrame, text="Pick the Path", command=self.filePicker)
            self.pathPickerButton.grid()

            self.sqlentryFrame = tk.Frame(mainFrame)
            self.sqlentryFrame.grid(row=3, columnspan=2, stick=tk.E)
            
            sqlUserNameLabel = tk.Label(self.sqlentryFrame, text="Sql User Name: ")
            sqlUserNameLabel.grid(stick=tk.E)

            sqlUserNameEntryFrame = tk.Frame(self.sqlentryFrame, width = 200, height = 25)
            sqlUserNameEntryFrame.grid(row=0, column=1,stick=tk.E)
            sqlUserNameEntryFrame.columnconfigure(0, weight=10)
            sqlUserNameEntryFrame.grid_propagate(False)

            self.sqlUserNameEntry = tk.Entry(sqlUserNameEntryFrame)
            self.sqlUserNameEntry.grid()

            sqlPasswordLabel = tk.Label(self.sqlentryFrame, text="Sql Password: ")
            sqlPasswordLabel.grid(stick=tk.E)

            sqlPasswordEntryFrame = tk.Frame(self.sqlentryFrame, width = 200, height = 25)
            sqlPasswordEntryFrame.grid(row=1, column=1,stick=tk.E)
            sqlPasswordEntryFrame.columnconfigure(0, weight=10)
            sqlPasswordEntryFrame.grid_propagate(False)

            self.sqlPasswordEntry = tk.Entry(sqlPasswordEntryFrame)
            self.sqlPasswordEntry.grid()

            sqlServerLocationLabel = tk.Label(self.sqlentryFrame, text="Server Path: ")
            sqlServerLocationLabel.grid(row=2,column=0)
            
            sqlServerLocationFrame = tk.Frame(self.sqlentryFrame, width = 200, height = 25)
            sqlServerLocationFrame.grid(row=2, column=1,stick=tk.E)
            sqlServerLocationFrame.columnconfigure(0, weight=10)
            sqlServerLocationFrame.grid_propagate(False)

            self.sqlServerLocationEntry = tk.Entry(sqlServerLocationFrame)
            self.sqlServerLocationEntry.grid()

            sqlServerPortLabel = tk.Label(self.sqlentryFrame, text="Server Port: ")
            sqlServerPortLabel.grid(row=3,column=0)

            sqlServerPortFrame = tk.Frame(self.sqlentryFrame, width = 200, height = 25)
            sqlServerPortFrame.grid(row=3, column=1,stick=tk.E)
            sqlServerPortFrame.columnconfigure(0, weight=10)
            sqlServerPortFrame.grid_propagate(False)

            self.sqlServerPortEntry = tk.Entry(sqlServerPortFrame)
            self.sqlServerPortEntry.grid()

            self.addRecordButton = tk.Button(mainFrame, text="Add New Record",command=self.addRecord)
            self.addRecordButton.grid(row=4, columnspan=2)

            self.variable = tk.StringVar(mainFrame)
            self.variable.trace("w", self.callback)
            self.variable.set("Excel Spreadsheet")

            self.typeSelection = tk.OptionMenu(mainFrame, self.variable, "Excel Spreadsheet", "SQL Server", "SQL Script")        
            self.typeSelection.grid(row=1,columnspan=2, stick=tk.W)

            textFrame = tk.Frame(mainFrame,width=400, height=100)
            textFrame.grid(row=5,columnspan=2)
            textFrame.columnconfigure(0, weight=10)
            textFrame.rowconfigure(0, weight=1)
            textFrame.grid_propagate(False)

            self.recordList = tk.Listbox(textFrame)
            self.recordList.grid(row=0, column=0, sticky=(tk.N, tk.S, tk.E, tk.W))
                       
        def save(self):
            print("Saving")
            with open('connectionsCSV.csv', 'w', newline='') as myfile:
                wr = csv.writer(myfile, quoting=csv.QUOTE_ALL)
                wr.writerows(self.recordArray)
            
        def load(self):
            print("Loading")
            with open('connectionsCSV.csv','r') as myfile:
                reader = csv.reader(myfile)
                self.recordArray = list(reader)
            print(self.recordArray)

        def callback(self, *args):
            if (self.variable.get() == "SQL Server"):
                self.pathPickerFrame.grid_remove()
                self.sqlentryFrame.grid()
                self.entryFrame.grid()
                self.addRecordButton.grid()
                self.SQLScripts = False
                print ("Showing SQL Information " + str(self.SQLScripts))
            elif self.variable.get() == "Excel Spreadsheet":
                self.sqlentryFrame.grid_remove()
                self.entryFrame.grid_remove()
                self.addRecordButton.grid_remove()
                self.pathPickerFrame.grid()
                self.SQLScripts = False
                print ("Showing SQL Information " + str(self.SQLScripts))
            elif self.variable.get() == "SQL Script":
                self.sqlentryFrame.grid_remove()
                self.entryFrame.grid_remove()
                self.addRecordButton.grid_remove()
                self.pathPickerFrame.grid()
                self.SQLScripts = True
                print ("Showing SQL Information " + str(self.SQLScripts))

        def filePicker(self):
            if not self.SQLScripts:
                self.fileName = askopenfilename(filetypes=(("Excel Files", "*.xlsx"), ("Macro Excel Files", "*.xlsm")))
            else:
                self.fileName = askopenfilename(filetypes=(("SQL Scripts", "*.sql"),("All Files","*.*")))
            if self.fileName:
               
                if not self.SQLScripts:

                    tempArray =[]
                    tempArray.append(self.fileName)
                    p = popup(self.master,tempArray)
                    self.master.wait_window(p.top)
                    self.dataDictionary["Excel"].append(tempArray)
                    self.recordList.insert(tk.END, "Excel :" + self.fileName)
                else:
                    self.dataDictionary["SQL_Script"].append(self.fileName)
                    self.recordList.insert(tk.END, "SQL Script :" + self.fileName)            
            print(self.dataDictionary)

        def addRecord(self):
            if (self.sqlUserNameEntry.get() and self.sqlPasswordEntry.get() and self.locationEntry.get() and self.sqlServerLocationEntry.get() and self.sqlServerPortEntry.get()):
                self.dataDictionary["SQL_Connection"].append([self.sqlUserNameEntry.get(), self.sqlPasswordEntry.get(), self.locationEntry.get(),self.sqlServerLocationEntry.get(),self.sqlServerPortEntry.get()])
                self.recordList.insert(tk.END, "SQL Server :" + self.locationEntry.get() + " Username :" + self.sqlUserNameEntry.get()+ " Password :" + self.sqlPasswordEntry.get())
                print(self.dataDictionary)




class finishedPopup:
    def __init__(self, parent, strFileName, sqlFileName):
        self.strFileName = strFileName
        top = self.top = tk.Toplevel(parent)
        tk.Label(top, text="File Name for: " + sqlFileName).grid(row=0, column=0)
        fileNameEntryFrame = tk.Frame(top, width = 100, height= 25)
        fileNameEntryFrame.grid(row = 0, column = 1)
        fileNameEntryFrame.columnconfigure(0, weight=10)
        fileNameEntryFrame.grid_propagate(False)

        self.fileNameEntry = tk.Entry(fileNameEntryFrame)
        self.fileNameEntry.grid()

        buttonOk = tk.Button(top, text="Ok", command=self.finished)
        buttonOk.grid(row=1)

    def finished(self):
        if (self.fileNameEntry.get()):
            self.strFileName = self.fileNameEntry.get()
            self.top.destroy()


class popup:
    def __init__(self, parent, excelArray):

        self.excelArray = excelArray
        top = self.top = tk.Toplevel(parent)
        tk.Label(top, text="Sheet Name").grid(row=0, column=0)

        sheetEntryFrame = tk.Frame(top, width = 100, height = 25)
        sheetEntryFrame.grid(row=0, column=1,stick=tk.E)
        sheetEntryFrame.columnconfigure(0, weight=10)
        sheetEntryFrame.grid_propagate(False)

        self.sheetNameEntry = tk.Entry(sheetEntryFrame)
        self.sheetNameEntry.grid()

        tk.Label(top, text="Table Name").grid(row=1, column=0)

        tableEntryFrame = tk.Frame(top, width = 100, height = 25)
        tableEntryFrame.grid(row=1, column=1,stick=tk.E)
        tableEntryFrame.columnconfigure(0, weight=10)
        tableEntryFrame.grid_propagate(False)

        self.tableNameEntry = tk.Entry(tableEntryFrame)
        self.tableNameEntry.grid()

        buttonOk = tk.Button(top, text="OK", command=self.finished)
        buttonOk.grid(row=2)

    def finished(self):
        if self.sheetNameEntry.get() and self.tableNameEntry.get():
            self.excelArray.append(self.sheetNameEntry.get())
            self.excelArray.append(self.tableNameEntry.get())
            self.top.destroy()
            
root = tk.Tk()
my_gui = MainGui(root)
root.mainloop()
