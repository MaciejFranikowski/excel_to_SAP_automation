import openpyxl
import os
import re
import materialsDictionary

sapNumberRegex = re.compile(r'(\d\d\d\d\d\d\d\d\d\d)')
sapNumberQuantityDictionary = {}

materials = materialsDictionary.materials.copy()

print('Give materials excel file current directory')
excelDirectory = input()
os.chdir(excelDirectory)

print('Give the name of the file (without the extension)')
fileName = input()
workBook = openpyxl.load_workbook(fileName + '.xlsx')
workSheet = workBook['Materiały OPL']
productCode = list(workSheet.columns)[1]

print('Do you want to just input materials or get SAP codes as well?')
print('1 - Just input')
print('2 - Input and SAP codes into excel')

choice = int(input())

print('In which column is the quantity of the material?')
quantityCoulmn = int(input())

# Getting the SAP numbers from the dictionary to their correct descriptions.
if choice == 2:
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

# Creating a list contating SAP numbers with their according quantities.
for currentRow in workSheet.values:
    for value in currentRow:
        mo = sapNumberRegex.search(str(value))
        if(mo is not None):
            if(mo.group() not in sapNumberQuantityDictionary):
                sapNumberQuantityDictionary[mo.group()] = currentRow[4]
            else:
                value = sapNumberQuantityDictionary.pop(mo.group())
                value += currentRow[4]
                sapNumberQuantityDictionary[mo.group()] = value

print(sapNumberQuantityDictionary)
workBook.save(fileName + 'MATERIALY.xlsx')
