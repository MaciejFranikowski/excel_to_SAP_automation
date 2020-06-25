import openpyxl
import os
import re
import csubDictionary
import sortListFunction
import pyautogui
import time

numberRegex = re.compile(r'(\d.(\d)?\d.\d(\d)?)')
unsortedList = []
currentPosition = 2
currentSliderPosition = 1
sliderY = 261
xMousePosition = 1029
yMousePosition = 261
startingY = 261

print('Give excel file current directory')
excelDirectory = input()
os.chdir(excelDirectory)

print('Give the name of the file (without the extension)')
fileName = input()
workBook = openpyxl.load_workbook(fileName + '.xlsx')
workSheet = workBook['Szczegółowy bez mat.']
firstColumnList = list(workSheet.columns)[0]

for cellObject in firstColumnList:
    mo = numberRegex.search(str(cellObject.value))
    if(mo is not None and mo.group() not in unsortedList):
        unsortedList.append(mo.group())

sortedList = sortListFunction.sortList(unsortedList)
for strings in sortedList:
    print(strings)
    print(csubDictionary.csub.get(strings))

time.sleep(5)
pyautogui.moveTo(xMousePosition, yMousePosition)

for strings in sortedList:
    differenceInt = csubDictionary.csub.get(strings) - currentPosition
    difference = differenceInt * 27
    currentPosition += differenceInt
    print('DifferenceInt = ' + str(differenceInt))
    print('currentPosition - currentSliderPosition ' + str(currentPosition - currentSliderPosition))
    if(currentPosition - currentSliderPosition >= 25):
        yMousePosition = startingY - 20
        pyautogui.moveTo(xMousePosition, sliderY)
        sliderY += (currentPosition - currentSliderPosition) * 2.888888888
        pyautogui.drag(0, (currentPosition - currentSliderPosition) * 2.888888888, duration = 0.2)
        time.sleep(3)
        currentSliderPosition = currentPosition
        pyautogui.click(30, yMousePosition)
    else:
        pyautogui.click(30, yMousePosition + difference)
        yMousePosition += difference

    print('Current Position = ' + str(currentPosition))
    print('Current Slider Position = ' + str(currentSliderPosition) + '\n')
