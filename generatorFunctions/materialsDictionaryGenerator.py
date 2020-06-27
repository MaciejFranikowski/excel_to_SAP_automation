import openpyxl
import os
import re
import pprint

csubValueRegex = re.compile(r'(\d.(\d)?\d.\d(\d)?)')
unsortedCsubValues = []
currentPosition = 2
currentSliderPosition = 1
sliderY = 272
xMousePosition = 1029
yMousePosition = 272
startingY = 272

print('Give materials excel file current directory')
excelDirectory = input()
os.chdir(excelDirectory)

print('Give the name of the file (without the extension)')
fileName = input()
workBook = openpyxl.load_workbook(fileName + '.xlsx')
workSheet = workBook['Sheet1']
#productCode = list(workSheet.columns)[1]
#print(productCode)
#input()

dictonary = {}
for row in workSheet.iter_rows(min_row=1, max_col=3, max_row=96, values_only = True):
    t = tuple(row[:2])
    dictonary[t] = row[2]

fileObj = open('materialsDictionary.py', 'w')
fileObj.write('materials = ' + pprint.pformat(dictonary) + '\n')
fileObj.close()
