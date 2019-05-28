import os
import xlrd
import tkinter as tk
from tkinter import filedialog

print("Please wait while program initializes.")
application_window = tk.Tk()
application_window.withdraw()
fileDir = os.path.normpath(filedialog.askdirectory(parent=application_window, initialdir=os.path.dirname(
    os.getcwd()), title="Please select the folder with the Excel files:"))
importDir = os.path.join(os.path.normpath(fileDir), "import")
if not os.path.exists(importDir) and not os.path.isdir(importDir):
    os.mkdir(importDir)
for file in os.listdir(fileDir):
    if file.endswith(".xlsx" or ".xls"):
        print("\nFound file: ", file)
        print("Extracting...")
        fileName = ("IMPORT-" + file.split(".xls")[0])
        wb = xlrd.open_workbook(os.path.join(fileDir, file))
        sheet = wb.sheet_by_index(0)
        rows = (sheet.nrows)
        f2 = open("%s\\%s.csv" % (fileDir, fileName), "w+")
        titleRow = (sheet.cell_value(0, 11) + "," + sheet.cell_value(
            0, 12) + "," + sheet.cell_value(0, 13))
        f2.write("%s\n" % (titleRow,))
        i = 1
        while i < rows:
            columnD = int(sheet.cell_value(i, 3))
            columnK = (sheet.cell_value(i, 10))
            columnL = (sheet.cell_value(i, 11))
            columnM = (sheet.cell_value(i, 12))
            columnN = (sheet.cell_value(i, 13))
            f = open("%s\\%d.html" % (importDir, columnD), "w+")
            f.write(columnK)
            f2.write("%s,%s,%s\n" % (columnL, columnM, columnN))
            i += 1
        print("---HTML(s): COMPLETE\n---CSV:     COMPLETE")
        f.close()
        f2.close()
print("\nDone with files in: ", fileDir)
input("Press ENTER to close.")
