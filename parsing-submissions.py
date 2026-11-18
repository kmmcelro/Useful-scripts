# Author: Kaitlin McElroy
# Purpose: To parse spreadsheet file from a Google form's submissions, so they are in a readable format to share easily
# If you rerun it for a spreadsheet file, it will overwrite the old text file.
# Date: November 17, 2026

import openpyxl

workbookName = input("Please input the name of the file your form results are in: ") # Getting the file name for the parsing
workbookName += input("Please input the file extension: \nSupported formats are: .xlsx,.xlsm,.xltx,.xltm ") # getting the file format ".xlsx"
dataframe = openpyxl.load_workbook(workbookName)
responses = dataframe.active


fieldName = []
fieldNum = int(input("How many fields are there? Please overshoot if estimating. ")) + 1 # include the timestamp as a field to get the correct endpoint
responseNum = int(input("How many responses are there? Please overshoot if estimating. ")) + 1

# Saving a list of the column values and finding the column for storygraph links
for col in responses.iter_cols(2, fieldNum):
    if col[0].value is None:
        break
    fieldName.append(col[0].value)

# Opening/creating reformatted response file
responseFile = workbookName.split(".")[0] + "_formatted.txt"
with open(responseFile, "w", encoding='utf-8') as f:
    
  # Reading each entry and reformatting it
  for row in responses.iter_rows(2,responseNum): 
      i = 1
      s = ""
      if row[1].value == None:
          break
      for title in fieldName:
          s += title + ":\n" + str(row[i].value) + " \n\n"
          i += 1
      s += "\n"
      f.write(s)
    
f.close()