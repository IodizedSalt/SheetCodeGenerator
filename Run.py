import itertools

import Spreadsheet as data  # importing the library as 'data'

# ---------------------------#
# Created by Chris 29/11/2018
# ---------------------------#

spreadsheet = data.sheet  # reference the JCI sheet as 'spreadsheet'

# Create Lists with all data from columns
brand = [x for x in spreadsheet.range("B25:B26") if x]
ganged = [x for x in spreadsheet.range("C25:C28") if x]
size = [x for x in spreadsheet.range("D25:D28") if x]
material = [x for x in spreadsheet.range("E25:E26") if x]
insulation = [x for x in spreadsheet.range("F25:F26") if x]
actuator = [x for x in spreadsheet.range("G25:G26") if x]
tType = [x for x in spreadsheet.range("H25:H26") if x]
flanges = [x for x in spreadsheet.range("I25:I26") if x]
airflow = [x for x in spreadsheet.range("J25:J27") if x]
pressure = [x for x in spreadsheet.range("K25:K26") if x]
controller = [x for x in spreadsheet.range("L25:L25") if x]

# Create empty lists to store the cell values later
brandConv = []
gangedConv = []
sizeConv = []
materialConv = []
insulationConv = []
actuatorConv = []
tTypeConv = []
flangesConv = []
airflowConv = []
pressureConv = []
controllerConv = []

# Loops to convert the list objects into strings and append them to the new, empty lists
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

for x in flanges:
    xString = str(x)
    xString = xString.split("'")
    flangesConv.append(xString[1])

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

    # Debugging by printing lines to the console because im a fucking pleb
# print("brand: ", brandConv)
# print("ganged: ", gangedConv)
# print("size: ", sizeConv)
# print("material: ", materialConv)
# print("insulation: ", insulationConv)
# print("actuator: ", actuatorConv)
# print("type: ", tTypeConv)
# print("flanges: ", flangesConv)
# print("airflow: ", airflowConv)
# print("pressure: ", pressureConv)
# print("controller: ", controllerConv)

# create all possible combinations via means of Cartesian Product
c = list(itertools.product(brandConv, gangedConv, sizeConv, materialConv, insulationConv, actuatorConv, tTypeConv,
                           flangesConv, airflowConv, pressureConv, controllerConv))

# Concatenate each tuple in C to a single string
cFinal = list((''.join(c) for c in c))

# Print each concatenated string on a new line
print("\n".join(cFinal))
