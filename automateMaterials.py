import openpyxl
import os
import re
import materialsDictionary

materials = materialsDictionary.materials.copy()

print('Give materials excel file current directory')
excelDirectory = input()
os.chdir(excelDirectory)

print('Give the name of the file (without the extension)')
fileName = input()
workBook = openpyxl.load_workbook(fileName + '.xlsx')
workSheet = workBook['Materiały OPL']
productCode = list(workSheet.columns)[1]

materialsDescriptions = materials.values()


for product in productCode:
    for material in materialsDescriptions:
        tuple(material)
        if product.value in material[0] or product.value in material[1]:
            print(product.value)

print(productCode)
input()
