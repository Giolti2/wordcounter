import csv
import re
import xlsxwriter # type: ignore

rowArray = []
filepath = input('Type file path relative to script location: ')
filename = filepath.split('/')[-1]
folder = filepath.replace(filename, '')
filename = filename.split('.')[0]

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
    
    workbook = xlsxwriter.Workbook(folder + filename + '-count.xlsx')
    worksheet = workbook.add_worksheet()
    
    row = 0
    
    for word in counts:
        worksheet.write(row, 0, word)
        worksheet.write(row, 1, counts[word])
        row+=1
        
    workbook.close()