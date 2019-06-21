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
        try:
            wb = xlrd.open_workbook(os.path.join(fileDir, file))
        except PermissionError:
            print("Unable to open %s, it might be open by a different program.\nPlease close and re-run the program." % file)
        else:
            sheet = wb.sheet_by_index(0)
            rows = (sheet.nrows)
            try:
                f2 = open("%s\\%s.csv" % (fileDir, fileName), "w+")
            except PermissionError:
                print("Unable to open %s.csv, it might be open by a different program.\nPlease close and re-run the program." % fileName)
            else:
                titleRow = (sheet.cell_value(0, 11) + "," + sheet.cell_value(0, 12) + "," + sheet.cell_value(0, 13) + "," +
                            sheet.cell_value(0, 14) + "," + sheet.cell_value(0, 15) + "," + sheet.cell_value(0, 16) + "," + sheet.cell_value(0, 17))
                f2.write("\"%s\"\n" % (titleRow,))
                i = 1
                while i < rows:
                    columnD = int(sheet.cell_value(i, 3))
                    columnK = (sheet.cell_value(i, 10))
                    columnL = (sheet.cell_value(i, 11))
                    columnM = (sheet.cell_value(i, 12))
                    columnN = (sheet.cell_value(i, 13))
                    columnO = (sheet.cell_value(i, 14))
                    columnP = (sheet.cell_value(i, 15))
                    columnQ = (sheet.cell_value(i, 16))
                    columnR = (sheet.cell_value(i, 17))
                    f = open("%s\\%d.html" % (importDir, columnD), "w+")
                    f.write(columnK)
                    f2.write("\"%s\",\"%s\",\"%s\",\"%s\",\"%s\",\"%s\",\"%s\"\n" % (
                        columnL, columnM, columnN, columnO, columnP, columnQ, columnR))
                    i += 1
                print("---HTML(s): COMPLETE\n---CSV:     COMPLETE")
                f.close()
                f2.close()
print("\nDone with files in: ", fileDir)
input("Press ENTER to close.")
