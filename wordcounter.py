import csv
import re
import os
import xlsxwriter # type: ignore
import tkinter as tk
from tkinter import filedialog

root = tk.Tk()
root.withdraw()

filepath = os.path.abspath(filedialog.askopenfilename())
folder = os.path.dirname(os.path.abspath(filepath))
filename = os.path.basename(filepath).split('.')[0]

outputFolder = os.path.join(folder, 'Count')
try:
    os.mkdir(outputFolder)
except FileExistsError as error:
    print(error)

rowArray = []

with open(filepath, encoding='utf8', newline='') as file:
    filereader = csv.reader(file, dialect='excel', delimiter=' ', quotechar='|')
    for row in filereader:
        rowArray.append(row)
        
    for i in range(len(rowArray)):
        for j in range(len(rowArray[i])):
            rowArray[i][j] = re.sub(r'[^a-zA-Z0-9]', '', rowArray[i][j]).lower()
    
    counts = dict()
    for row in rowArray:
        for word in row:
            counts[word] = counts.get(word, 0) + 1
    
    workbook = xlsxwriter.Workbook(os.path.join(outputFolder, filename + '-count.xlsx'))
    worksheet = workbook.add_worksheet()
    
    row = 0
    
    for word in counts:
        worksheet.write(row, 0, word)
        worksheet.write(row, 1, counts[word])
        row+=1
        
    workbook.close()