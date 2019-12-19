import itertools

import Spreadsheet as data  # importing the library as 'data'
import re

# ---------------------------#
# Created by Chris 29/11/2018
# ---------------------------#

spreadsheet = data.sheet  # reference the JCI sheet as 'spreadsheet'
finalSpreadSheet = data.FINALSHEET

# Create Lists with all data from columns
brand = [x for x in spreadsheet.range("B37:B38") if x]
ganged = [x for x in spreadsheet.range("C37:C42") if x]
size = [x for x in spreadsheet.range("D37:D39") if x]
material = [x for x in spreadsheet.range("E37:E37") if x]
insulation = [x for x in spreadsheet.range("F37:F38") if x]
actuator = [x for x in spreadsheet.range("G37:G38") if x]
tType = [x for x in spreadsheet.range("H37:H37") if x]
airflow = [x for x in spreadsheet.range("I37:I39") if x]
pressure = [x for x in spreadsheet.range("J37:J38") if x]
controller = [x for x in spreadsheet.range("K37:K38") if x]

# Create empty lists to store the cell values later
brandConv = []
gangedConv = []
sizeConv = []
materialConv = []
insulationConv = []
actuatorConv = []
tTypeConv = []
airflowConv = []
pressureConv = []
controllerConv = []


# Loops to convert the list objects into strings and append them to the new, empty lists        #todo, ensure these still correctly split data
def loopList():
    for x in brand:
        xString = str(x)
        xString = xString.split("'")
        brandConv.append(xString[1])

    for x in ganged:
        xString = str(x)
        xString = xString.split("'")
        gangedConv.append(xString[1])

    for x in size:
        xString = str(x)
        xString = xString.split("'")
        sizeConv.append(xString[1])

    for x in material:
        xString = str(x)
        xString = xString.split("'")
        materialConv.append(xString[1])

    for x in insulation:
        xString = str(x)
        xString = xString.split("'")
        insulationConv.append(xString[1])

    for x in actuator:
        xString = str(x)
        xString = xString.split("'")
        actuatorConv.append(xString[1])

    for x in tType:
        xString = str(x)
        xString = xString.split("'")
        tTypeConv.append(xString[1])

    for x in airflow:
        xString = str(x)
        xString = xString.split("'")
        airflowConv.append(xString[1])

    for x in pressure:
        xString = str(x)
        xString = xString.split("'")
        pressureConv.append(xString[1])

    for x in controller:
        xString = str(x)
        xString = xString.split("'")
        controllerConv.append(xString[1])


def CVUException(x):
    # print(x)
    if x.endswith("U"):
        if "CV," in x:
            return True
        else:
            return False
    else:
        return False


def FANException(x):
    # print(x)
    if x.endswith("N"):
        if "FA," in x:
            return True
        else:
            return False
    else:
        return False


def fourTeenFException(x):
    # print(x)
    if "14," in x:
        if "FS," in x:
            return True
        else:
            return False
    else:
        return False


def concatinationTxt():
    # create all possible combinations via means of Cartesian Product

    c = list(itertools.product(brandConv, gangedConv, sizeConv, materialConv, insulationConv, actuatorConv, tTypeConv,
                               airflowConv, pressureConv, controllerConv))

    # Concatenate each tuple in C to a single string
    cFinal = list(''.join(c) for c in c)
    itemsToRemove = list(set())

    for x in cFinal:
        if ("2x08" in x or "3x08" in x
                or "4x08" in x or "6x08" in x or CVUException(x) or FANException(x)):
            itemsToRemove.append(x)

    # print(cFinal)
    print("ShortDescriptionListLengthBefore")
    print(cFinal.__len__())

    # print(itemsToRemove)

    bFinal = [x for x in cFinal if x not in itemsToRemove]

    cFinal = [x.replace('SS,', '').replace('ON,', '').replace('IN,', '').replace(',H,', ',')
                  .replace(',U,', ',').replace(',D,', ',').replace('LPU', 'LP')
                        .replace('LPN', 'LP').replace('MPU', 'MP').replace('MPN', 'MP')
              for x in bFinal]

    print("ShortDescriptionListLengthAfter")
    print(cFinal.__len__())
    print("\n")
    return "\n".join(cFinal)


def concatinationSheet():
    # create all possible combinations via means of Cartesian Product

    c = list(itertools.product(brandConv, gangedConv, sizeConv, materialConv, insulationConv, actuatorConv, tTypeConv,
                               airflowConv, pressureConv, controllerConv))

    # Concatenate each tuple in C to a single string
    cFinal = list(''.join(c) for c in c)
    itemsToRemove = list(set())
    for x in cFinal:
        if ("2x08" in x or "3x08" in x
                or "4x08" in x or "6x08" in x or CVUException(x) or FANException(x)):
            itemsToRemove.append(x)


    bFinal = [x for x in cFinal if x not in itemsToRemove]

    cFinal = [x.replace('AL,', '').replace('HR,', '').replace('ON,', '').replace('IN,', '').replace(',H,', ',')
                  .replace(',U,', ',').replace(',D,', ',').replace('LPU', 'LP')
                        .replace('LPN', 'LP').replace('MPU', 'MP').replace('MPN', 'MP')
              for x in bFinal]

    return cFinal


def writeModelNumbers(newFileData):
    cell_list = finalSpreadSheet.range('F2:F1681')
    cell_values = newFileData
    print("SDG cell list length: ", cell_list.__len__())
    print("SDG cell value length: ", cell_values.__len__())

    for cell, x in enumerate(cell_values):
        cell_list[cell].value = x

    finalSpreadSheet.update_cells(cell_list)


def writeModelNumbersToFile(newFileData):
    file = open("outputs/SS316Output/SS316_ShortDescription.txt", "w")
    file.write(newFileData)


def run():
    loopList()
    # writeModelNumbers(concatinationSheet())
    writeModelNumbersToFile(concatinationTxt())
    print("SS316 Short Description Number Write success")


# run()
