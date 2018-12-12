import itertools

import Spreadsheet as data  # importing the library as 'data'

# ---------------------------#
# Created by Chris 29/11/2018
# ---------------------------#

spreadsheet = data.sheet  # reference the JCI sheet as 'spreadsheet'
finalSpreadSheet = data.FINALSHEET

# Create Lists with all data from columns                           #todo, change these cell values
brand = [x for x in spreadsheet.range("B34:B35") if x]
ganged = [x for x in spreadsheet.range("C34:C37") if x]
size = [x for x in spreadsheet.range("D34:D37") if x]
material = [x for x in spreadsheet.range("E34:E35") if x]
insulation = [x for x in spreadsheet.range("F34:F35") if x]
actuator = [x for x in spreadsheet.range("G34:G35") if x]
tType = [x for x in spreadsheet.range("H34:H35") if x]
airflow = [x for x in spreadsheet.range("I34:I36") if x]
pressure = [x for x in spreadsheet.range("J34:J35") if x]
controller = [x for x in spreadsheet.range("K34:K35") if x]

# Create empty lists to store the cell values later                 #todo, change these listVariable Names
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

def concatinationTxt():
    # create all possible combinations via means of Cartesian Product
    c = list(itertools.product(brandConv, gangedConv, sizeConv, materialConv, insulationConv, actuatorConv, tTypeConv,
                               airflowConv, pressureConv, controllerConv))

    # Concatenate each tuple in C to a single string
    cFinal = list(''.join(c) for c in c)
    itemsToRemove = list(set())

    for x in cFinal:
        if ("D,08" in x or "T,08" in x
                or "Q,08" in x or "CV,FS" in x or CVUException(x) or FANException(x)):
            itemsToRemove.append(x)

    cFinal = [x for x in cFinal if x not in itemsToRemove]

    return "\n".join(cFinal)


def concatinationSheet():
    # create all possible combinations via means of Cartesian Product
    c = list(itertools.product(brandConv, gangedConv, sizeConv, materialConv, insulationConv, actuatorConv, tTypeConv,
                               airflowConv, pressureConv, controllerConv))

    # Concatenate each tuple in C to a single string
    cFinal = list(''.join(c) for c in c)
    itemsToRemove = list(set())

    for x in cFinal:
        if ("D,08" in x or "T,08" in x
                or "Q,08" in x or "CV,FS" in x or CVUException(x) or FANException(x)):
            itemsToRemove.append(x)

    cFinal = [x for x in cFinal if x not in itemsToRemove]

    return cFinal




def writeModelNumbers(newFileData):

    cell_list = finalSpreadSheet.range('F2:F1873')
    cell_values = newFileData
    print("SDG cell list length: ", cell_list.__len__())
    print("SDG cell value length: ", cell_values.__len__())

    for cell, x in enumerate(cell_values):
        cell_list[cell].value = x

    finalSpreadSheet.update_cells(cell_list)

def writeModelNumbersToFile(newFileData):

    file = open("outputs/ShortDescription.txt", "w")
    file.write(newFileData)

def run():
    loopList()
    writeModelNumbers(concatinationSheet())
    writeModelNumbersToFile(concatinationTxt())

# run()