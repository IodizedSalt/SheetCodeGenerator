import itertools

import Spreadsheet as data  # importing the library as 'data'

# ---------------------------#
# Created by Chris 29/11/2018
# ---------------------------#

spreadsheet = data.sheet  # reference the JCI sheet as 'spreadsheet'
finalSpreadSheet = data.FINALSHEET

# Create Lists with all data from columns
brand = [x for x in spreadsheet.range("B47:B48") if x]
ganged = [x for x in spreadsheet.range("C47:C50") if x]
size = [x for x in spreadsheet.range("D47:D50") if x]
material = [x for x in spreadsheet.range("E47:E48") if x]
insulation = [x for x in spreadsheet.range("F47:F48") if x]
actuator = [x for x in spreadsheet.range("G47:G48") if x]
tType = [x for x in spreadsheet.range("H47:H48") if x]
airflow = [x for x in spreadsheet.range("I47:I49") if x]
pressure = [x for x in spreadsheet.range("J47:J48") if x]
controller = [x for x in spreadsheet.range("K47:K48") if x]

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
    if x.endswith("UVM"):
        if "CONSTANT VOLUME," in x:
            return True
        else:
            return False
    else:
        return False


def FANException(x):
    # print(x)
    if x.endswith("NO CTRL"):
        if "FAST ACTING," in x:
            return True
        else:
            return False
    else:
        return False

def fourTeenFException(x):
    # print(x)
    if "14in," in x:
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
        if ("2x08in" in x or "3x08in" in x
                or "4x08in" in x or "CONSTANT VOLUME,FS," in x or CVUException(x)or FANException(x) or fourTeenFException(x)):
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
        if ("2x08in" in x or "3x08in" in x
                or "4x08in" in x or "CONSTANT VOLUME,FS," in x or CVUException(x)or FANException(x) or fourTeenFException(x)):
            itemsToRemove.append(x)

    cFinal = [x for x in cFinal if x not in itemsToRemove]

    return cFinal



def writeModelNumbers(newFileData):

    cell_list = finalSpreadSheet.range('G2:G1681')
    cell_values = newFileData
    print("LDG cell list length: ", cell_list.__len__())
    print("LDG cell value length: ", cell_values.__len__())

    for cell, x in enumerate(cell_values):
        cell_list[cell].value = x

    finalSpreadSheet.update_cells(cell_list)

def writeModelNumbersToFile(newFileData):

    file = open("outputs/LongDescription2.0.txt", "w")
    file.write(newFileData)

def run():
    loopList()
    writeModelNumbers(concatinationSheet())
    writeModelNumbersToFile(concatinationTxt())
    print("Write Success")



run()