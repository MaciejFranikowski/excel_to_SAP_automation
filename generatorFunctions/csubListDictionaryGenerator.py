import openpyxl
import os
import re
import pprint
from operator import itemgetter

numberRegex = re.compile(r'(\d.(\d)?\d.\d(\d)?)')

def sortList(myList):
    for i in range(0,len(unsortedList),1):
        temp = unsortedList[i].split('.')
        temp[0] = int(temp[0])
        temp[1] = int(temp[1])
        temp[2] = int(temp[2])
        unsortedList[i] = temp
    x = sorted(unsortedList, key = itemgetter(2))
    y = sorted(x, key = lambda z: z[1])
    v = sorted(y, key = lambda z: z[0])
    for i in range(0,len(v),1):
        temp = v[i]
        temp[0] = str(temp[0])
        temp[1] = str(temp[1])
        temp[2] = str(temp[2])
        temp = unsortedList[i]
        temp2 = ".".join(v[i])
        v[i] = temp2
    return v


print('Give excel file current directory')
excelDirectory = input()
os.chdir(excelDirectory)

print('Give the name of the file (without the extension)')
fileName = input()
workBook = openpyxl.load_workbook(fileName + '.xlsx')

workSheet = workBook.active
firstColumnList = list(workSheet.columns)[0]

unsortedList = []

for cellObject in firstColumnList:
    mo = numberRegex.search(str(cellObject.value))
    if(mo is not None):
        unsortedList.append(mo.group())


sortedList = sortList(unsortedList)
dictonary = {}
i = 0
for strings in sortedList:
    i = i + 1
    dictonary[strings]= i
for x, y in dictonary.items():
    print(x, y)

fileObj = open('csubDictionary.py', 'w')
fileObj.write('csub = ' + pprint.pformat(dictonary) + '\n')
fileObj.close()
