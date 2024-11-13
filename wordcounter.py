import csv
import re
import xlsxwriter # type: ignore
import tkinter as tk
from tkinter import filedialog

root = tk.Tk()
root.withdraw()

filepath = filedialog.askopenfilename()
#filepath = input('Type file path relative to script location: ')

rowArray = []

with open(filepath, encoding='utf8', newline='') as file:
    filereader = csv.reader(file, dialect='excel', delimiter=' ', quotechar='|')
    for row in filereader:
        rowArray.append(row)
        
    rowArray.pop(0)
    for i in range(len(rowArray)):
        for j in range(len(rowArray[i])):
            rowArray[i][j] = re.sub(r'[^a-zA-Z0-9]', '', rowArray[i][j])
    
    counts = dict()
    for row in rowArray:
        for word in row:
            counts[word] = counts.get(word, 0) + 1
    
    workbook = xlsxwriter.Workbook(filepath + '-count.xlsx')
    worksheet = workbook.add_worksheet()
    
    row = 0
    
    for word in counts:
        worksheet.write(row, 0, word)
        worksheet.write(row, 1, counts[word])
        row+=1
        
    workbook.close()