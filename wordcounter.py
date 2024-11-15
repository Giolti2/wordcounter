import csv
import re
import os
import xlsxwriter  # type: ignore
import tkinter as tk
from tkinter import filedialog

# Initialize the tkinter GUI (hidden)
root = tk.Tk()
root.withdraw()

# Initialize cumulative counts dictionary
cumulative_counts = dict()

# Main loop to process files
while True:
    # File selection dialog
    filepath = os.path.abspath(filedialog.askopenfilename())
    if not filepath:
        print("No file selected. Exiting.")
        break

    folder = os.path.dirname(os.path.abspath(filepath))
    filename = os.path.basename(filepath).split('.')[0]

    # Create output folder if it doesn't exist
    outputFolder = os.path.join(folder, 'Count')
    try:
        os.mkdir(outputFolder)
    except FileExistsError:
        pass

    # Read and process the file
    rowArray = []
    with open(filepath, encoding='utf8', newline='') as file:
        filereader = csv.reader(file, dialect='excel', delimiter=' ', quotechar='|')
        for row in filereader:
            rowArray.append(row)

        for i in range(len(rowArray)):
            for j in range(len(rowArray[i])):
                rowArray[i][j] = re.sub(r'[^a-zA-Z0-9]', '', rowArray[i][j]).lower()

        # Count occurrences and update cumulative counts
        for row in rowArray:
            for word in row:
                cumulative_counts[word] = cumulative_counts.get(word, 0) + 1

    # Ask if the user wants to process another file
    choice = input("Do you want to process another file? (yes/no): ").strip().lower()
    if choice not in ['yes', 'y']:
        break

# Write the cumulative counts to an Excel file
if cumulative_counts:
    workbook = xlsxwriter.Workbook(os.path.join(outputFolder, 'cumulative-count.xlsx'))
    worksheet = workbook.add_worksheet()

    row = 0
    for word, count in cumulative_counts.items():
        worksheet.write(row, 0, word)
        worksheet.write(row, 1, count)
        row += 1

    workbook.close()
    print(f"Cumulative count saved to {os.path.join(outputFolder, 'cumulative-count.xlsx')}")

print("Program ended.")