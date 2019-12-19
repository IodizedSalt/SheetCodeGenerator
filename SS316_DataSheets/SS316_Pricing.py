import json
import time
import sys, os
from os.path import dirname, join, abspath


sys.path.insert(0, abspath(join(dirname(__file__), '../outputs/SS316Output')))
import Spreadsheet as data

finalSpreadSheet = data.FINALSHEET
priceSheet = data.price

modelNumInvalidFormat = finalSpreadSheet.range('B2:B673')
# colA08PA = priceSheet.acell("G9").value       #todo, decide between implement this or static values

modelNum = {}
priceList = {}


def loopList():
    for x in modelNumInvalidFormat:
        xString = str(x)
        xString = xString.split("'")
        # priceList[(xString[1])] = 0.0
        modelNum[(xString[1])] = 0.0
        # modelNum = dict(zip(priceList.keys(), [float(value) for value in priceList.values()]))


def tempLoopList():
    with open("outputs/SS316Output/Pricing.txt") as f:
        for line in f:
            modelNum[str(line)] = 0.0


def priceColA():
    for x in modelNum:
        if "N" in x[2] or "F" in x[2]:
            ifColA(x)
        else:
            ifColAGanged(x)
    return "\n".join(modelNum)


def ifColA(x):
    if "P" in x[9]:
        if "08" in x[3:5]:
            a = modelNum[x]
            a += 200.74
            modelNum.update({x: a})
        elif "10" in x[3:5]:
            a = modelNum[x]
            a += 250.40
            modelNum.update({x: a})
        elif "12" in x[3:5]:
            a = modelNum[x]
            a += 252.74
            modelNum.update({x: a})

def ifColAGanged(x):
    if "08" in x[3:5]:
        if "2" in x[2]:
            a = modelNum[x]
            a += (200.74 * 2)
            modelNum.update({x: a})
        if "3" in x[2]:
            a = modelNum[x]
            a += (200.74 * 3)
            modelNum.update({x: a})
        if "4" in x[2]:
            a = modelNum[x]
            a += (200.74 * 4)
            modelNum.update({x: a})
        if "6" in x[2]:
            a = modelNum[x]
            a += (200.74 * 6)
            modelNum.update({x: a})
    elif "10" in x[3:5]:
        if "2" in x[2]:
            a = modelNum[x]
            a += (250.40 * 2)
            modelNum.update({x: a})
        if "3" in x[2]:
            a = modelNum[x]
            a += (250.40 * 3)
            modelNum.update({x: a})
        if "4" in x[2]:
            a = modelNum[x]
            a += (250.40 * 4)
            modelNum.update({x: a})
        if "6" in x[2]:
            a = modelNum[x]
            a += (250.40 * 6)
            modelNum.update({x: a})
    elif "12" in x[3:5]:
        if "2" in x[2]:
            a = modelNum[x]
            a += (252.74* 2)
            modelNum.update({x: a})
        if "3" in x[2]:
            a = modelNum[x]
            a += (252.74 * 3)
            modelNum.update({x: a})
        if "4" in x[2]:
            a = modelNum[x]
            a += (252.74 * 4)
            modelNum.update({x: a})
        if "6" in x[2]:
            a = modelNum[x]
            a += (252.74 * 6)
            modelNum.update({x: a})

def priceColB():
    for x in modelNum:
        if "N" in x[2] or "F" in x[2]:
            ifColB(x)
            # pass #x1 for all values
        else:
            ifColBGanged(x)
            # pass  #x2 for all values

    return "\n".join(modelNum)


def ifColB(x):
    # for x in modelNum:
    if "08" in x[3:5]:
        a = modelNum[x]
        a += 102.46
        modelNum.update({x: a})
    elif "10" in x[3:5]:
        a = modelNum[x]
        a += 107.21
        modelNum.update({x: a})
    elif "12" in x[3:5]:
        if "L" in x[11]:
            a = modelNum[x]
            a += 116.05
            modelNum.update({x: a})
        elif "M" in x[11]:
            a = modelNum[x]
            a += 115.73
            modelNum.update({x: a})

def ifColBGanged(x):
        if "08" in x[3:5]:
            if "2" in x[2]:
                a = modelNum[x]
                a += (102.46 * 2)
                modelNum.update({x: a})
            if "3" in x[2]:
                a = modelNum[x]
                a += (102.46 * 3)
                modelNum.update({x: a})
            if "4" in x[2]:
                a = modelNum[x]
                a += (102.46 * 4)
                modelNum.update({x: a})
            if "6" in x[2]:
                a = modelNum[x]
                a += (102.46 * 6)
                modelNum.update({x: a})
        elif "10" in x[3:5]:
            if "2" in x[2]:
                a = modelNum[x]
                a += (107.21 * 2)
                modelNum.update({x: a})
            if "3" in x[2]:
                a = modelNum[x]
                a += (107.21 * 3)
                modelNum.update({x: a})
            if "4" in x[2]:
                a = modelNum[x]
                a += (107.21 * 4)
                modelNum.update({x: a})
            if "6" in x[2]:
                a = modelNum[x]
                a += (107.21 * 6)
                modelNum.update({x: a})
        elif "12" in x[3:5] and "L" in x[11]:
            if "2" in x[2]:
                a = modelNum[x]
                a += (116.05 * 2)
                modelNum.update({x: a})
            if "3" in x[2]:
                a = modelNum[x]
                a += (116.05 * 3)
                modelNum.update({x: a})
            if "4" in x[2]:
                a = modelNum[x]
                a += (116.05 * 4)
                modelNum.update({x: a})
            if "6" in x[2]:
                a = modelNum[x]
                a += (116.05 * 6)
                modelNum.update({x: a})
        elif "12" in x[3:5] and "M" in x[11]:
            if "2" in x[2]:
                a = modelNum[x]
                a += (115.73 * 2)
                modelNum.update({x: a})
            if "3" in x[2]:
                a = modelNum[x]
                a += (115.73 * 3)
                modelNum.update({x: a})
            if "4" in x[2]:
                a = modelNum[x]
                a += (115.73 * 4)
                modelNum.update({x: a})
            if "6" in x[2]:
                a = modelNum[x]
                a += (115.73 * 6)
                modelNum.update({x: a})

def priceColC():
    for x in modelNum:
        if "N" in x[2] or "F" in x[2]:
            ifColC(x)
            # pass #x1 for all values
        else:
            ifColCGanged(x)
            # pass  #x2 for all values

    return "\n".join(modelNum)


def ifColC(x):
    # for x in modelNum:
    if "CV" in x[7:9]:
        if "08" in x[3:5]:
            a = modelNum[x]
            a += 55.75
            modelNum.update({x: a})
        elif "10" in x[3:5]:
            a = modelNum[x]
            a += 56.55
            modelNum.update({x: a})
        elif "12" in x[3:5]:
            a = modelNum[x]
            a += 57.55
            modelNum.update({x: a})

    elif "FA" in x[7:9]:
        if "H" in x[10]:
            if "08" in x[3:5]:
                a = modelNum[x]
                a += 211.80
                modelNum.update({x: a})
            elif "10" in x[3:5]:
                a = modelNum[x]
                a += 212.57
                modelNum.update({x: a})
            elif "12" in x[3:5]:
                a = modelNum[x]
                a += 213.58
                modelNum.update({x: a})

        elif "U" in x[10] or "D" in x[10]:
            if "08" in x[3:5]:
                a = modelNum[x]
                a += 217.75
                modelNum.update({x: a})
            elif "10" in x[3:5]:
                a = modelNum[x]
                a += 216.94
                modelNum.update({x: a})
            elif "12" in x[3:5]:
                a = modelNum[x]
                a += 228.89
                modelNum.update({x: a})


def ifColCGanged(x):
    # for x in modelNum:
    if "CV" in x[7:9]:
        if "08" in x[3:5]:
            if "2" in x[2]:
                a = modelNum[x]
                a += (55.75 * 2)
                modelNum.update({x: a})
            if "3" in x[2]:
                a = modelNum[x]
                a += (55.75 * 3)
                modelNum.update({x: a})
            if "4" in x[2]:
                a = modelNum[x]
                a += (55.75 * 4)
                modelNum.update({x: a})
            if "6" in x[2]:
                a = modelNum[x]
                a += (55.75 * 6)
                modelNum.update({x: a})
        elif "10" in x[3:5]:
            if "2" in x[2]:
                a = modelNum[x]
                a += (56.55 * 2)
                modelNum.update({x: a})
            if "3" in x[2]:
                a = modelNum[x]
                a += (56.55 * 3)
                modelNum.update({x: a})
            if "4" in x[2]:
                a = modelNum[x]
                a += (56.55 * 4)
                modelNum.update({x: a})
            if "6" in x[2]:
                a = modelNum[x]
                a += (56.55 * 6)
        elif "12" in x[3:5]:
            if "2" in x[2]:
                a = modelNum[x]
                a += (57.55 * 2)
                modelNum.update({x: a})
            if "3" in x[2]:
                a = modelNum[x]
                a += (57.55 * 3)
                modelNum.update({x: a})
            if "4" in x[2]:
                a = modelNum[x]
                a += (57.55 * 4)
                modelNum.update({x: a})
            if "6" in x[2]:
                a = modelNum[x]
                a += (57.55 * 6)
    elif "FA" in x[7:9]:
        if "H" in x[10]:
            if "08" in x[3:5]:
                if "2" in x[2]:
                    a = modelNum[x]
                    a += (211.80 * 2)
                    modelNum.update({x: a})
                if "3" in x[2]:
                    a = modelNum[x]
                    a += (211.80 * 3)
                    modelNum.update({x: a})
                if "4" in x[2]:
                    a = modelNum[x]
                    a += (211.80 * 4)
                    modelNum.update({x: a})
                if "6" in x[2]:
                    a = modelNum[x]
                    a += (211.80 * 6)
                    modelNum.update({x: a})
            elif "10" in x[3:5]:
                if "2" in x[2]:
                    a = modelNum[x]
                    a += (212.57 * 2)
                    modelNum.update({x: a})
                if "3" in x[2]:
                    a = modelNum[x]
                    a += (212.57 * 3)
                    modelNum.update({x: a})
                if "4" in x[2]:
                    a = modelNum[x]
                    a += (212.57 * 4)
                    modelNum.update({x: a})

                if "6" in x[2]:
                    a = modelNum[x]
                    a += (212.57 * 6)
                    modelNum.update({x: a})
            elif "12" in x[3:5]:
                if "2" in x[2]:
                    a = modelNum[x]
                    a += (213.58 * 2)
                    modelNum.update({x: a})
                if "3" in x[2]:
                    a = modelNum[x]
                    a += (213.58 * 3)
                    modelNum.update({x: a})
                if "4" in x[2]:
                    a = modelNum[x]
                    a += (213.58 * 4)
                    modelNum.update({x: a})
                if "6" in x[2]:
                    a = modelNum[x]
                    a += (213.58 * 6)
                    modelNum.update({x: a})

        elif "U" in x[10] or "D" in x[10]:
            if "08" in x[3:5]:
                if "2" in x[2]:
                    a = modelNum[x]
                    a += (217.75 * 2)
                    modelNum.update({x: a})
                if "3" in x[2]:
                    a = modelNum[x]
                    a += (217.75 * 3)
                    modelNum.update({x: a})
                if "4" in x[2]:
                    a = modelNum[x]
                    a += (217.75 * 4)
                    modelNum.update({x: a})
                if "6" in x[2]:
                    a = modelNum[x]
                    a += (217.75 * 6)
                    modelNum.update({x: a})
            elif "10" in x[3:5]:
                if "2" in x[2]:
                    a = modelNum[x]
                    a += (216.94 * 2)
                    modelNum.update({x: a})
                if "3" in x[2]:
                    a = modelNum[x]
                    a += (216.94 * 3)
                    modelNum.update({x: a})
                if "4" in x[2]:
                    a = modelNum[x]
                    a += (216.94 * 4)
                    modelNum.update({x: a})
                if "6" in x[2]:
                    a = modelNum[x]
                    a += (216.94 * 6)
                    modelNum.update({x: a})
            elif "12" in x[3:5]:
                if "2" in x[2]:
                    a = modelNum[x]
                    a += (228.89 * 2)
                    modelNum.update({x: a})
                if "3" in x[2]:
                    a = modelNum[x]
                    a += (228.89 * 3)
                    modelNum.update({x: a})
                if "4" in x[2]:
                    a = modelNum[x]
                    a += (228.89 * 4)
                    modelNum.update({x: a})
                if "6" in x[2]:
                    a = modelNum[x]
                    a += (228.89 * 6)
                    modelNum.update({x: a})

def priceColD():
    for x in modelNum:
        if "N" in x[2] or "F" in x[2]:
            ifColD(x)
            # pass #x1 for all values
        else:
            ifColDGanged(x)
            # pass  #x2 for all values

    return "\n".join(modelNum)


def ifColD(x):
    if "H" in x[10]:
        a = modelNum[x]
        a += 15.60
        modelNum.update({x: a})
    elif "U" in x[10] or "D" in x[10]:
        a = modelNum[x]
        a += 15.38
        modelNum.update({x: a})


def ifColDGanged(x):
    # for x in modelNum:
    if "H" in x[10]:
        if "08" in x[3:5]:
            if "2" in x[2]:
                a = modelNum[x]
                a += (15.60 * 2)
                modelNum.update({x: a})
            if "3" in x[2]:
                a = modelNum[x]
                a += (15.60 * 3)
                modelNum.update({x: a})
            if "4" in x[2]:
                a = modelNum[x]
                a += (15.60 * 4)
                modelNum.update({x: a})
            if "6" in x[2]:
                a = modelNum[x]
                a += (15.60 * 6)
                modelNum.update({x: a})
        elif "10" in x[3:5]:
            if "2" in x[2]:
                a = modelNum[x]
                a += (15.60 * 2)
                modelNum.update({x: a})
            if "3" in x[2]:
                a = modelNum[x]
                a += (15.60 * 3)
                modelNum.update({x: a})
            if "4" in x[2]:
                a = modelNum[x]
                a += (15.60 * 4)
                modelNum.update({x: a})
            if "6" in x[2]:
                a = modelNum[x]
                a += (15.60 * 6)
                modelNum.update({x: a})
        elif "12" in x[3:5]:
            if "2" in x[2]:
                a = modelNum[x]
                a += (15.60 * 2)
                modelNum.update({x: a})
            if "3" in x[2]:
                a = modelNum[x]
                a += (15.60 * 3)
                modelNum.update({x: a})
            if "4" in x[2]:
                a = modelNum[x]
                a += (15.60 * 4)
                modelNum.update({x: a})
            if "6" in x[2]:
                a = modelNum[x]
                a += (15.60 * 6)
                modelNum.update({x: a})

    elif "U" in x[10] or "D" in x[10]:
        if "08" in x[3:5]:
            if "2" in x[2]:
                a = modelNum[x]
                a += (15.38 * 2)
                modelNum.update({x: a})
            if "3" in x[2]:
                a = modelNum[x]
                a += (15.38 * 3)
                modelNum.update({x: a})
            if "4" in x[2]:
                a = modelNum[x]
                a += (15.38 * 4)
                modelNum.update({x: a})
            if "6" in x[2]:
                a = modelNum[x]
                a += (15.38 * 6)
                modelNum.update({x: a})
        elif "10" in x[3:5]:
            if "2" in x[2]:
                a = modelNum[x]
                a += (15.38 * 2)
                modelNum.update({x: a})
            if "3" in x[2]:
                a = modelNum[x]
                a += (15.38 * 3)
                modelNum.update({x: a})
            if "4" in x[2]:
                a = modelNum[x]
                a += (15.38 * 4)
                modelNum.update({x: a})
            if "6" in x[2]:
                a = modelNum[x]
                a += (15.38 * 6)
                modelNum.update({x: a})
        elif "12" in x[3:5]:
            if "2" in x[2]:
                a = modelNum[x]
                a += (15.38 * 2)
                modelNum.update({x: a})
            if "3" in x[2]:
                a = modelNum[x]
                a += (15.38 * 3)
                modelNum.update({x: a})
            if "4" in x[2]:
                a = modelNum[x]
                a += (15.38 * 4)
                modelNum.update({x: a})
            if "6" in x[2]:
                a = modelNum[x]
                a += (15.38 * 6)
                modelNum.update({x: a})

def colE():
    for x in modelNum:
        if "F" in x[2]:
            if "08" in x[3:5]:
                a = modelNum[x]
                a += 120
                modelNum.update({x: a})
            elif "10" in x[3:5]:
                a = modelNum[x]
                a += 120
                modelNum.update({x: a})
            elif "12" in x[3:5]:
                a = modelNum[x]
                a += 120
                modelNum.update({x: a})


def priceColG():
    for x in modelNum:
        if "N" in x[6]:
            pass
        elif "I" in x[6]:
            ifColG(x)

    return "\n".join(modelNum)


def ifColG(x):
    if "N" in x[2] or "F" in x[2]:
        if "08" in x[3:5]:
            a = modelNum[x]
            a += (16.87)
            modelNum.update({x: a})
        elif "10" in x[3:5]:
            a = modelNum[x]
            a += (27.02)
            modelNum.update({x: a})
        elif "12" in x[3:5]:
            a = modelNum[x]
            a += (24.52)
            modelNum.update({x: a})

    if "2" in x[2]:
        if "08" in x[3:5]:
            a = modelNum[x]
            a += (16.87 * 2)
            modelNum.update({x: a})
        elif "10" in x[3:5]:
            a = modelNum[x]
            a += (27.02 * 2)
            modelNum.update({x: a})
        elif "12" in x[3:5]:
            a = modelNum[x]
            a += (24.52 * 2)
            modelNum.update({x: a})

    elif "3" in x[2]:
        if "08" in x[3:5]:
            a = modelNum[x]
            a += (16.87 * 3)
            modelNum.update({x: a})
        elif "10" in x[3:5]:
            a = modelNum[x]
            a += (27.02 * 3)
            modelNum.update({x: a})
        elif "12" in x[3:5]:
            a = modelNum[x]
            a += (24.52 * 3)
            modelNum.update({x: a})

    elif "4" in x[2]:
        if "08" in x[3:5]:
            a = modelNum[x]
            a += (16.87 * 4)
            modelNum.update({x: a})
        elif "10" in x[3:5]:
            a = modelNum[x]
            a += (27.02 * 4)
            modelNum.update({x: a})
        elif "12" in x[3:5]:
            a = modelNum[x]
            a += (24.52 * 4)
            modelNum.update({x: a})

    elif "6" in x[2]:
        if "08" in x[3:5]:
            a = modelNum[x]
            a += (16.87 * 6)
            modelNum.update({x: a})
        elif "10" in x[3:5]:
            a = modelNum[x]
            a += (27.02 * 6)
            modelNum.update({x: a})
        elif "12" in x[3:5]:
            a = modelNum[x]
            a += (24.52 * 6)
            modelNum.update({x: a})


def priceColHorI():
    for x in modelNum:
        if "N" in x[2] or "F" in x[2]:
            pass
        else:
            ifColHorI(x)

    return "\n".join(modelNum)


def ifColHorI(x):
    if "2" in x[2]:
        if "08" in x[3:5]:
            pass
        elif "10" in x[3:5]:
            a = modelNum[x]
            a += 170
            modelNum.update({x: a})
        elif "12" in x[3:5]:
            a = modelNum[x]
            a += 170
            modelNum.update({x: a})

    elif "3" in x[2]:
        if "08" in x[3:5]:
            pass
        elif "10" in x[3:5]:
            a = modelNum[x]
            a += 180.00
            modelNum.update({x: a})
        elif "12" in x[3:5]:
            a = modelNum[x]
            a += 188.39
            modelNum.update({x: a})

    elif "4" in x[2]:
        if "08" in x[3:5]:
            pass
        elif "10" in x[3:5]:
            a = modelNum[x]
            a += 241.19
            modelNum.update({x: a})

        elif "12" in x[3:5]:
            a = modelNum[x]
            a += 304.62
            modelNum.update({x: a})
    elif "6" in x[2]:
        if "08" in x[3:5]:
            pass
        elif "10" in x[3:5]:
            a = modelNum[x]
            a += (180*2)
            modelNum.update({x: a})
        elif "12" in x[3:5]:
            a = modelNum[x]
            a += (188.39*2)
            modelNum.update({x: a})

def uvmCost():
    for x in modelNum:
        if "U" in x[12]:
            a = modelNum[x]
            a += 116.96
            modelNum.update({x: a})
            if (x == "VV412SNFAPHLU"):
                print(modelNum[x])

    return "\n".join(modelNum)


def colK():
    for x in modelNum:
        if "2" in x[2]:
            a = modelNum[x]
            a += (10.00 * 2)
            modelNum.update({x: a})
        elif "3" in x[2]:
            a = modelNum[x]
            a += (10.00 * 3)
            modelNum.update({x: a})
        elif "4" in x[2]:
            a = modelNum[x]
            a += (10.00 * 4)
            modelNum.update({x: a})
        elif "6" in x[2]:
            a = modelNum[x]
            a += (10.00 * 6)
            modelNum.update({x: a})
        else:
            a = modelNum[x]
            a += 10.00
            modelNum.update({x: a})
def colL():
    for x in modelNum:
        if "2" in x[2]:
            a = modelNum[x]
            a += (10.00 * 2)
            modelNum.update({x: a})
        elif "3" in x[2]:
            a = modelNum[x]
            a += (10.00 * 3)
            modelNum.update({x: a})
        elif "4" in x[2]:
            a = modelNum[x]
            a += (10.00 * 4)
            modelNum.update({x: a})
        elif "6" in x[2]:
            a = modelNum[x]
            a += (10.00 * 6)
            modelNum.update({x: a})
        else:
            a = modelNum[x]
            a += 10.00
            modelNum.update({x: a})



def writeModelNumbersToFile(newFileData):
    file = open("../outputs/SS316Output/Pricing.txt", "w")
    file.write(newFileData)


def writeCostDictionaryToFile(newFileData):
    # with open('outputs/dictionary.txt', 'w') as f:
    #     print(newFileData, file=f)
    with open("../outputs/SS316Output/Pricing.txt", "w") as file:
        json.dump(newFileData, file, indent=2)


def writeModelNumberstoSheet(newFileData):
    cell_list = finalSpreadSheet.range('Q2:Q673')
    cell_values = newFileData

    print("Pricing cell list length: ", cell_list.__len__())
    print("Pricing cell value length: ", cell_values.__len__())

    for cell, x in enumerate(cell_values):
        cell_list[cell].value = x

    finalSpreadSheet.update_cells(cell_list)


def main():
    loopList()
    # modelNum.clear()
    # tempLoopList()
    priceColA()
    priceColB()
    priceColC()
    priceColD()
    colE()
    priceColG()
    priceColHorI()
    uvmCost()
    colK()
    colL()

    # print(modelNum)

    priceList = list(modelNum.values())
    print(priceList)

    writeCostDictionaryToFile(modelNum)  # writes to dictionary.json
    # writeModelNumbersToFile(modelNum)

    writeModelNumberstoSheet(priceList)  # writes to sheet
    print("SS316 Pricing Write Success")


if __name__ == "__main__":
    main()
