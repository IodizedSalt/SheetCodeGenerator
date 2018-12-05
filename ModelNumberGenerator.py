import itertools

import Spreadsheet as data  # importing the library as 'data'

# ---------------------------#
# Created by Chris 29/11/2018
# ---------------------------#

spreadsheet = data.sheet  # reference the JCI sheet as 'spreadsheet'

# Create Lists with all data from columns
brand = [x for x in spreadsheet.range("B18:B19") if x]
ganged = [x for x in spreadsheet.range("C18:C21") if x]
size = [x for x in spreadsheet.range("D18:D21") if x]
material = [x for x in spreadsheet.range("E18:E19") if x]
insulation = [x for x in spreadsheet.range("F18:F19") if x]
actuator = [x for x in spreadsheet.range("G18:G19") if x]
tType = [x for x in spreadsheet.range("H18:H19") if x]
airflow = [x for x in spreadsheet.range("I18:I20") if x]
pressure = [x for x in spreadsheet.range("J18:J19") if x]
controller = [x for x in spreadsheet.range("K18:K19") if x]

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


def concatination():
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


def writeModelNumbers(newFileData):                         #todo, write directly to excel sheet
    file = open("outputs/ModelNumbers.txt", "w")
    file.write(newFileData)




def run():
    loopList()
    writeModelNumbers(concatination())

