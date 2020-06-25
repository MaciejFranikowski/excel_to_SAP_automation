import openpyxl
import os
import re
import csubDictionary
import sortListFunction
import pyautogui
import time

csubValueRegex = re.compile(r'(\d.(\d)?\d.\d(\d)?)')
unsortedCsubValues = []
currentPosition = 2
currentSliderPosition = 1
sliderY = 272
xMousePosition = 1029
yMousePosition = 272
firstRowY = 235
sliderScaling = 2.55
numberOfRows = 27

print('Give excel file current directory')
excelDirectory = input()
os.chdir(excelDirectory)

print('Give the name of the file (without the extension)')
fileName = input()
workBook = openpyxl.load_workbook(fileName + '.xlsx')
workSheet = workBook['Szczegółowy bez mat.']
firstColumnList = list(workSheet.columns)[0]

# Adding the values from the first column of the excel file to the
# unsorted list that'll contatin all csub values.
for cellObject in firstColumnList:
    mo = csubValueRegex.search(str(cellObject.value))
    if(mo is not None and mo.group() not in unsortedCsubValues):
        unsortedCsubValues.append(mo.group())


sortedCsubValues = sortListFunction.sortList(unsortedCsubValues)
for strings in sortedCsubValues:
    print(strings)
    print(csubDictionary.csub.get(strings))

time.sleep(5)
pyautogui.moveTo(xMousePosition, yMousePosition)

for strings in sortedCsubValues:
    # Calculate the difference between the ordering number of the Csub Value
    # and the current position in the SAP list.
    differencePosition = csubDictionary.csub.get(strings) - currentPosition
    # Calculate the physical pixel difference, by multiplying the position
    # in the csub file by the height of a SAP column.
    physicalDifference = differencePosition * 27
    currentPosition += differencePosition
    print('differencePosition = ' + str(differencePosition))
    print('currentPosition - currentSliderPosition ' + str(currentPosition - currentSliderPosition))
    if(currentPosition - currentSliderPosition >= numberOfRows):
        yMousePosition = firstRowY
        pyautogui.moveTo(xMousePosition, sliderY)
        sliderY += (currentPosition - currentSliderPosition) * sliderScaling
        pyautogui.drag(0, (currentPosition - currentSliderPosition) * sliderScaling, duration = 0.2)
        time.sleep(3)
        currentSliderPosition = currentPosition
        pyautogui.click(30, yMousePosition)
    else:
        pyautogui.click(30, yMousePosition + physicalDifference)
        yMousePosition += physicalDifference

    print('Current Position = ' + str(currentPosition))
    print('Current Slider Position = ' + str(currentSliderPosition) + '\n')
