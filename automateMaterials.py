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

materialsDescriptions = materials.keys()

for product in productCode:
    for material in materialsDescriptions:
        shortDescr = str(material[0])
        longDescr = str(material[1])
        if product.value is not None:
            if product.value in shortDescr or product.value in longDescr:
                rowSap = product.row
                columnSap = product.column + 2
                workSheet.cell(row=rowSap, column=columnSap).value = materials.get(material)

workBook.save(fileName + 'MATERIALY.xlsx')
