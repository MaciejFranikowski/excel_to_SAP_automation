import pyautogui
import pyperclip
import time

sapNumberX = 530
sapNumberY = 340
columnYDiff = 27
quantityX = 800
warehouseX = 1020


def inputSapNumberQuantity(materialsDictionary, numberOfRows):
    time.sleep(3)
    yCoord = sapNumberY
    materialsDictionaryValues = list(materialsDictionary.values())
    materialsDictionaryKeys = list(materialsDictionary.keys())
    pyautogui.click(sapNumberX, sapNumberY)
    time.sleep(3)
    pyautogui.click(sapNumberX, sapNumberY)

    for i in range(numberOfRows):
        key = materialsDictionaryKeys[i]
        pyperclip.copy(key)
        pyautogui.keyDown('ctrl')
        pyautogui.keyDown('v')
        pyautogui.keyUp('v')
        pyautogui.keyUp('ctrl')
        yCoord += columnYDiff
        pyautogui.click(sapNumberX, yCoord)
        if i + 1 == len(materialsDictionary):
            break

    yCoord = sapNumberY
    pyautogui.click(quantityX, sapNumberY)
    for i in range(numberOfRows):
        value = materialsDictionaryValues[i]
        pyperclip.copy(value)
        pyautogui.keyDown('ctrl')
        pyautogui.keyDown('v')
        pyautogui.keyUp('v')
        pyautogui.keyUp('ctrl')
        yCoord += columnYDiff
        pyautogui.click(quantityX, yCoord)
        if i + 1 == len(materialsDictionary):
            break


def inputWarehouse(wareHouseNumber, numberOfRows, materialsDictionaryLength):
    yCoord = sapNumberY
    pyautogui.moveTo(warehouseX, sapNumberY)
    pyautogui.click(warehouseX, sapNumberY)
    for i in range(numberOfRows):

        pyperclip.copy(wareHouseNumber)
        pyautogui.keyDown('ctrl')
        pyautogui.keyDown('v')
        pyautogui.keyUp('v')
        pyautogui.keyUp('ctrl')
        yCoord += columnYDiff
        if i + 1 == materialsDictionaryLength:
            break
        pyautogui.click(warehouseX, yCoord)
