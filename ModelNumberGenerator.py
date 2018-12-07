import itertools

import Spreadsheet as data  # importing the library as 'data'

# ---------------------------#
# Created by Chris 29/11/2018
# ---------------------------#

spreadsheet = data.sheet  # reference the JCI sheet as 'spreadsheet'
finalSpreadSheet = data.FINALSHEET

# Create Lists with all data from columns
brand = [x for x in spreadsheet.range("B19:B20") if x]
ganged = [x for x in spreadsheet.range("C19:C22") if x]
size = [x for x in spreadsheet.range("D19:D22") if x]
material = [x for x in spreadsheet.range("E19:E20") if x]
insulation = [x for x in spreadsheet.range("F19:F20") if x]
actuator = [x for x in spreadsheet.range("G19:G20") if x]
tType = [x for x in spreadsheet.range("H19:H20") if x]
airflow = [x for x in spreadsheet.range("I19:I21") if x]
pressure = [x for x in spreadsheet.range("J19:J20") if x]
controller = [x for x in spreadsheet.range("K19:K20") if x]

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


# Loops to convert the list objects into strings and append them to the new, empty lists
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
        if "CV" in x:
            return True
        else:
            return False
    else:
        return False

def concatinationTxt():             #This function writes to the local txt file
    # create all possible combinations via means of Cartesian Product
    c = list(itertools.product(brandConv, gangedConv, sizeConv, materialConv, insulationConv, actuatorConv, tTypeConv,
                               airflowConv, pressureConv, controllerConv))

    # Concatenate each tuple in C to a single string
    cFinal = list(''.join(c) for c in c)
    itemsToRemove = list(set())

    for x in cFinal:
        if ("208" in x or "308" in x
                or "408" in x or "CVF" in x or CVUException(x)):
            itemsToRemove.append(x)

    cFinal = [x for x in cFinal if x not in itemsToRemove]

    return "\n".join(cFinal)
    # return cFinal

def concatinationSheet():           #This function writes to the sheet
    # create all possible combinations via means of Cartesian Product
    c = list(itertools.product(brandConv, gangedConv, sizeConv, materialConv, insulationConv, actuatorConv, tTypeConv,
                               airflowConv, pressureConv, controllerConv))

    # Concatenate each tuple in C to a single string
    cFinal = list(''.join(c) for c in c)
    itemsToRemove = list(set())

    for x in cFinal:
        if ("208" in x or "308" in x
                or "408" in x or "CVF" in x or CVUException(x)):
            itemsToRemove.append(x)

    cFinal = [x for x in cFinal if x not in itemsToRemove]

    # return "\n".join(cFinal)
    return cFinal

def writeModelNumbers(newFileData):

    cell_list = finalSpreadSheet.range('B2:B3121')
    cell_values = newFileData

    print("MNG cell list length: ", cell_list.__len__())
    print("MNG cell value length: ",cell_values.__len__())

    for cell, x in enumerate(cell_values):
        cell_list[cell].value = x

    finalSpreadSheet.update_cells(cell_list)

def writeModelNumbersToFile(newFileData):

    file = open("outputs/ModelNumbers.txt", "w")
    file.write(newFileData)

def run():
    loopList()
    writeModelNumbers(concatinationSheet())
    writeModelNumbersToFile(concatinationTxt())

# run()