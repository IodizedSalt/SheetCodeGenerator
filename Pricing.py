import json
import time

import Spreadsheet as data

finalSpreadSheet = data.FINALSHEET
priceSheet = data.price

modelNumInvalidFormat = finalSpreadSheet.range('B2:B1681')
# colA08PA = priceSheet.acell("G9").value       #todo, decide between implement this or static values

modelNum = {}
priceList = {}


def loopList():
    for x in modelNumInvalidFormat:
        xString = str(x)
        xString = xString.split("'")
        # priceList[(xString[1])] = 0.0
        modelNum[(xString[1])] = 0.0
        # modelNum = dict(zip(priceList.keys(), [float(value) for value in priceList.values()]


def tempLoopList():
    with open("outputs/Pricing.txt") as f:
        for line in f:
            modelNum[str(line)] = 0.0


def priceColA():
    for x in modelNum:
        if "N" in x[2]:
            ifColA(x)
        else:
            ifColAGanged(x)
    return "\n".join(modelNum)


def ifColA(x):
    if "A" in x[5]:
        if "P" in x[9]:
            if "08" in x[3:5]:
                a = modelNum[x]
                a += 87.69
                modelNum.update({x: a})
            elif "10" in x[3:5]:
                a = modelNum[x]
                a += 88.23
                modelNum.update({x: a})
            elif "12" in x[3:5]:
                a = modelNum[x]
                a += 99.12
                modelNum.update({x: a})
            elif "14" in x[3:5]:
                a = modelNum[x]
                a += 168.73
                modelNum.update({x: a})
        elif "F" in x[9]:
            if "08" in x[3:5]:
                a = modelNum[x]
                a += 93.08
                modelNum.update({x: a})
            elif "10" in x[3:5]:
                a = modelNum[x]
                a += 97.84
                modelNum.update({x: a})
            elif "12" in x[3:5]:
                a = modelNum[x]
                a += 111.20
                modelNum.update({x: a})
            elif "14" in x[3:5]:
                pass
    elif "H" in x[5]:
        if "P" in x[9]:
            if "08" in x[3:5]:
                a = modelNum[x]
                a += 120.19
                modelNum.update({x: a})
            elif "10" in x[3:5]:
                a = modelNum[x]
                a += 129.94
                modelNum.update({x: a})
            elif "12" in x[3:5]:
                a = modelNum[x]
                a += 145.12
                modelNum.update({x: a})
            elif "14" in x[3:5]:
                a = modelNum[x]
                a += 214.60
                modelNum.update({x: a})
        elif "F" in x[9]:
            if "08" in x[3:5]:
                a = modelNum[x]
                a += 131.62
                modelNum.update({x: a})
            elif "10" in x[3:5]:
                a = modelNum[x]
                a += 136.59
                modelNum.update({x: a})
            elif "12" in x[3:5]:
                a = modelNum[x]
                a += 160.16
                modelNum.update({x: a})
            elif "14" in x[3:5]:
                pass


def ifColAGanged(x):
    if "A" in x[5]:
        if "P" in x[9]:
            if "08" in x[3:5]:
                if "2" in x[2]:
                    a = modelNum[x]
                    a += (87.69 * 2)
                    modelNum.update({x: a})
                if "3" in x[2]:
                    a = modelNum[x]
                    a += (87.69 * 3)
                    modelNum.update({x: a})
                if "4" in x[2]:
                    a = modelNum[x]
                    a += (87.69 * 4)
                    modelNum.update({x: a})
            elif "10" in x[3:5]:
                if "2" in x[2]:
                    a = modelNum[x]
                    a += (88.23 * 2)
                    modelNum.update({x: a})
                if "3" in x[2]:
                    a = modelNum[x]
                    a += (88.23 * 3)
                    modelNum.update({x: a})
                if "4" in x[2]:
                    a = modelNum[x]
                    a += (88.23 * 4)
                    modelNum.update({x: a})
            elif "12" in x[3:5]:
                if "2" in x[2]:
                    a = modelNum[x]
                    a += (99.12 * 2)
                    modelNum.update({x: a})
                if "3" in x[2]:
                    a = modelNum[x]
                    a += (99.12 * 3)
                    modelNum.update({x: a})
                if "4" in x[2]:
                    a = modelNum[x]
                    a += (99.12 * 4)
                    modelNum.update({x: a})
            elif "14" in x[3:5]:
                if "2" in x[2]:
                    a = modelNum[x]
                    a += (168.73 * 2)
                    modelNum.update({x: a})
                if "3" in x[2]:
                    a = modelNum[x]
                    a += (168.73 * 3)
                    modelNum.update({x: a})
                if "4" in x[2]:
                    a = modelNum[x]
                    a += (168.73 * 4)
                    modelNum.update({x: a})
        elif "F" in x[9]:
            if "08" in x[3:5]:
                if "2" in x[2]:
                    a = modelNum[x]
                    a += (93.08 * 2)
                    modelNum.update({x: a})
                if "3" in x[2]:
                    a = modelNum[x]
                    a += (93.08 * 3)
                    modelNum.update({x: a})
                if "4" in x[2]:
                    a = modelNum[x]
                    a += (93.08 * 4)
                    modelNum.update({x: a})
            elif "10" in x[3:5]:
                if "2" in x[2]:
                    a = modelNum[x]
                    a += (97.84 * 2)
                    modelNum.update({x: a})
                if "3" in x[2]:
                    a = modelNum[x]
                    a += (97.84 * 3)
                    modelNum.update({x: a})
                if "4" in x[2]:
                    a = modelNum[x]
                    a += (97.84 * 4)
                    modelNum.update({x: a})
            elif "12" in x[3:5]:
                if "2" in x[2]:
                    a = modelNum[x]
                    a += (111.20 * 2)
                    modelNum.update({x: a})
                if "3" in x[2]:
                    a = modelNum[x]
                    a += (111.20 * 3)
                    modelNum.update({x: a})
                if "4" in x[2]:
                    a = modelNum[x]
                    a += (111.20 * 4)
                    modelNum.update({x: a})
            elif "14" in x[3:5]:
                pass
    elif "H" in x[5]:
        if "P" in x[9]:
            if "08" in x[3:5]:
                if "2" in x[2]:
                    a = modelNum[x]
                    a += (120.19 * 2)
                    modelNum.update({x: a})
                if "3" in x[2]:
                    a = modelNum[x]
                    a += (120.19 * 3)
                    modelNum.update({x: a})
                if "4" in x[2]:
                    a = modelNum[x]
                    a += (120.19 * 4)
                    modelNum.update({x: a})
            elif "10" in x[3:5]:
                if "2" in x[2]:
                    a = modelNum[x]
                    a += (129.94 * 2)
                    modelNum.update({x: a})
                if "3" in x[2]:
                    a = modelNum[x]
                    a += (129.94 * 3)
                    modelNum.update({x: a})
                if "4" in x[2]:
                    a = modelNum[x]
                    a += (129.94 * 4)
                    modelNum.update({x: a})
            elif "12" in x[3:5]:
                if "2" in x[2]:
                    a = modelNum[x]
                    a += (145.14 * 2)
                    modelNum.update({x: a})
                if "3" in x[2]:
                    a = modelNum[x]
                    a += (145.14 * 3)
                    modelNum.update({x: a})
                if "4" in x[2]:
                    a = modelNum[x]
                    a += (145.14 * 4)
                    modelNum.update({x: a})
            elif "14" in x[3:5]:
                if "2" in x[2]:
                    a = modelNum[x]
                    a += (214.60 * 2)
                    modelNum.update({x: a})
                if "3" in x[2]:
                    a = modelNum[x]
                    a += (214.60 * 3)
                    modelNum.update({x: a})
                if "4" in x[2]:
                    a = modelNum[x]
                    a += (214.60 * 4)
                    modelNum.update({x: a})
        elif "F" in x[9]:
            if "08" in x[3:5]:
                if "2" in x[2]:
                    a = modelNum[x]
                    a += (131.62 * 2)
                    modelNum.update({x: a})
                if "3" in x[2]:
                    a = modelNum[x]
                    a += (131.62 * 3)
                    modelNum.update({x: a})
                if "4" in x[2]:
                    a = modelNum[x]
                    a += (131.62 * 4)
                    modelNum.update({x: a})
            elif "10" in x[3:5]:
                if "2" in x[2]:
                    a = modelNum[x]
                    a += (136.59 * 2)
                    modelNum.update({x: a})
                if "3" in x[2]:
                    a = modelNum[x]
                    a += (136.59 * 3)
                    modelNum.update({x: a})
                if "4" in x[2]:
                    a = modelNum[x]
                    a += (136.59 * 4)
                    modelNum.update({x: a})
            elif "12" in x[3:5]:
                if "2" in x[2]:
                    a = modelNum[x]
                    a += (160.16 * 2)
                    modelNum.update({x: a})
                if "3" in x[2]:
                    a = modelNum[x]
                    a += (160.16 * 3)
                    modelNum.update({x: a})
                if "4" in x[2]:
                    a = modelNum[x]
                    a += (160.16 * 4)
                    modelNum.update({x: a})
            elif "14" in x[3:5]:
                pass


def priceColB():
    for x in modelNum:
        if "N" in x[2]:
            ifColB(x)
            # pass #x1 for all values
        else:
            ifColBGanged(x)
            # pass  #x2 for all values

    return "\n".join(modelNum)


def ifColB(x):
    # for x in modelNum:
    if "A" in x[5]:
        if "P" in x[9]:
            if "08" in x[3:5]:
                a = modelNum[x]
                a += 44.99
                modelNum.update({x: a})
            elif "10" in x[3:5]:
                a = modelNum[x]
                a += 52.26
                modelNum.update({x: a})
            elif "12" in x[3:5]:
                if "L" in x[11]:
                    a = modelNum[x]
                    a += 43.36
                    modelNum.update({x: a})
                elif "M" in x[11]:
                    a = modelNum[x]
                    a += 47.28
                    modelNum.update({x: a})
            elif "14" in x[3:5]:
                a = modelNum[x]
                a += 53.48
                modelNum.update({x: a})
        elif "F" in x[9]:
            if "08" in x[3:5]:
                a = modelNum[x]
                a += 102.08
                modelNum.update({x: a})
            elif "10" in x[3:5]:
                a = modelNum[x]
                a += 98.66
                modelNum.update({x: a})
            elif "12" in x[3:5]:
                if "L" in x[11]:
                    a = modelNum[x]
                    a += 97.57
                    modelNum.update({x: a})
                elif "M" in x[11]:
                    a = modelNum[x]
                    a += 89.73
                    modelNum.update({x: a})
            elif "14" in x[3:5]:
                pass
    elif "H" in x[5]:
        if "P" in x[9]:
            if "08" in x[3:5]:
                a = modelNum[x]
                a += 57.74
                modelNum.update({x: a})
            elif "10" in x[3:5]:
                a = modelNum[x]
                a += 61.19
                modelNum.update({x: a})
            elif "12" in x[3:5]:
                if "L" in x[11]:
                    a = modelNum[x]
                    a += 63.12
                    modelNum.update({x: a})
                elif "M" in x[11]:
                    a = modelNum[x]
                    a += 67.02
                    modelNum.update({x: a})
            elif "14" in x[3:5]:
                a = modelNum[x]
                a += 67.67
                modelNum.update({x: a})
        elif "F" in x[9]:
            if "08" in x[3:5]:
                a = modelNum[x]
                a += 116.65
                modelNum.update({x: a})
            elif "10" in x[3:5]:
                a = modelNum[x]
                a += 120.66
                modelNum.update({x: a})
            elif "12" in x[3:5]:
                if "L" in x[11]:
                    a = modelNum[x]
                    a += 127.55
                    modelNum.update({x: a})
                elif "M" in x[11]:
                    a = modelNum[x]
                    a += 121.42
                    modelNum.update({x: a})
            elif "14" in x[3:5]:
                pass


def ifColBGanged(x):
    if "A" in x[5]:
        if "P" in x[9]:
            if "08" in x[3:5]:
                if "2" in x[2]:
                    a = modelNum[x]
                    a += (44.99 * 2)
                    modelNum.update({x: a})
                if "3" in x[2]:
                    a = modelNum[x]
                    a += (44.99 * 3)
                    modelNum.update({x: a})
                if "4" in x[2]:
                    a = modelNum[x]
                    a += (44.99 * 4)
                    modelNum.update({x: a})
            elif "10" in x[3:5]:
                if "2" in x[2]:
                    a = modelNum[x]
                    a += (52.26 * 2)
                    modelNum.update({x: a})
                if "3" in x[2]:
                    a = modelNum[x]
                    a += (52.26 * 3)
                    modelNum.update({x: a})
                if "4" in x[2]:
                    a = modelNum[x]
                    a += (52.26 * 4)
                    modelNum.update({x: a})
            elif "12" in x[3:5] and "L" in x[11]:
                if "2" in x[2]:
                    a = modelNum[x]
                    a += (43.36 * 2)
                    modelNum.update({x: a})
                if "3" in x[2]:
                    a = modelNum[x]
                    a += (43.36 * 3)
                    modelNum.update({x: a})
                if "4" in x[2]:
                    a = modelNum[x]
                    a += (43.36 * 4)
                    modelNum.update({x: a})
            elif "12" in x[3:5] and "M" in x[11]:
                if "2" in x[2]:
                    a = modelNum[x]
                    a += (47.28 * 2)
                    modelNum.update({x: a})
                if "3" in x[2]:
                    a = modelNum[x]
                    a += (47.28 * 3)
                    modelNum.update({x: a})
                if "4" in x[2]:
                    a = modelNum[x]
                    a += (47.28 * 4)
                    modelNum.update({x: a})
            elif "14" in x[3:5]:
                if "2" in x[2]:
                    a = modelNum[x]
                    a += (53.48 * 2)
                    modelNum.update({x: a})
                if "3" in x[2]:
                    a = modelNum[x]
                    a += (53.48 * 3)
                    modelNum.update({x: a})
                if "4" in x[2]:
                    a = modelNum[x]
                    a += (53.48 * 4)
                    modelNum.update({x: a})
        elif "F" in x[9]:
            if "08" in x[3:5]:
                if "2" in x[2]:
                    a = modelNum[x]
                    a += (102.08 * 2)
                    modelNum.update({x: a})
                if "3" in x[2]:
                    a = modelNum[x]
                    a += (102.08 * 3)
                    modelNum.update({x: a})
                if "4" in x[2]:
                    a = modelNum[x]
                    a += (102.08 * 4)
                    modelNum.update({x: a})
            elif "10" in x[3:5]:
                if "2" in x[2]:
                    a = modelNum[x]
                    a += (98.66 * 2)
                    modelNum.update({x: a})
                if "3" in x[2]:
                    a = modelNum[x]
                    a += (98.66 * 3)
                    modelNum.update({x: a})
                if "4" in x[2]:
                    a = modelNum[x]
                    a += (98.66 * 4)
                    modelNum.update({x: a})
            elif "12" in x[3:5] and "L" in x[11]:
                if "2" in x[2]:
                    a = modelNum[x]
                    a += (97.57 * 2)
                    modelNum.update({x: a})
                if "3" in x[2]:
                    a = modelNum[x]
                    a += (97.57 * 3)
                    modelNum.update({x: a})
                if "4" in x[2]:
                    a = modelNum[x]
                    a += (97.57 * 4)
                    modelNum.update({x: a})
            elif "12" in x[3:5] and "M" in x[11]:
                if "2" in x[2]:
                    a = modelNum[x]
                    a += (89.73 * 2)
                    modelNum.update({x: a})
                if "3" in x[2]:
                    a = modelNum[x]
                    a += (89.73 * 3)
                    modelNum.update({x: a})
                if "4" in x[2]:
                    a = modelNum[x]
                    a += (89.73 * 4)
                    modelNum.update({x: a})
            elif "14" in x[3:5]:
                pass
    elif "H" in x[5]:
        if "P" in x[9]:
            if "08" in x[3:5]:
                if "2" in x[2]:
                    a = modelNum[x]
                    a += (57.74 * 2)
                    modelNum.update({x: a})
                if "3" in x[2]:
                    a = modelNum[x]
                    a += (57.74 * 3)
                    modelNum.update({x: a})
                if "4" in x[2]:
                    a = modelNum[x]
                    a += (57.74 * 4)
                    modelNum.update({x: a})
            elif "10" in x[3:5]:
                if "2" in x[2]:
                    a = modelNum[x]
                    a += (61.19 * 2)
                    modelNum.update({x: a})
                if "3" in x[2]:
                    a = modelNum[x]
                    a += (61.19 * 3)
                    modelNum.update({x: a})
                if "4" in x[2]:
                    a = modelNum[x]
                    a += (61.19 * 4)
                    modelNum.update({x: a})

            elif "12" in x[3:5] and "L" in x[11]:
                if "2" in x[2]:
                    a = modelNum[x]
                    a += (63.12 * 2)
                    modelNum.update({x: a})
                if "3" in x[2]:
                    a = modelNum[x]
                    a += (63.12 * 3)
                    modelNum.update({x: a})
                if "4" in x[2]:
                    a = modelNum[x]
                    a += (63.12 * 4)
                    modelNum.update({x: a})
            elif "12" in x[3:5] and "M" in x[11]:
                if "2" in x[2]:
                    a = modelNum[x]
                    a += (67.02 * 2)
                    modelNum.update({x: a})
                if "3" in x[2]:
                    a = modelNum[x]
                    a += (67.02 * 3)
                    modelNum.update({x: a})
                if "4" in x[2]:
                    a = modelNum[x]
                    a += (67.02 * 4)
                    modelNum.update({x: a})
            elif "14" in x[3:5]:
                if "2" in x[2]:
                    a = modelNum[x]
                    a += (67.67 * 2)
                    modelNum.update({x: a})
                if "3" in x[2]:
                    a = modelNum[x]
                    a += (67.67 * 3)
                    modelNum.update({x: a})
                if "4" in x[2]:
                    a = modelNum[x]
                    a += (67.67 * 4)
                    modelNum.update({x: a})
        elif "F" in x[9]:
            if "08" in x[3:5]:
                if "2" in x[2]:
                    a = modelNum[x]
                    a += (116.65 * 2)
                    modelNum.update({x: a})
                if "3" in x[2]:
                    a = modelNum[x]
                    a += (116.65 * 3)
                    modelNum.update({x: a})
                if "4" in x[2]:
                    a = modelNum[x]
                    a += (116.65 * 4)
                    modelNum.update({x: a})
            elif "10" in x[3:5]:
                if "2" in x[2]:
                    a = modelNum[x]
                    a += (120.66 * 2)
                    modelNum.update({x: a})
                if "3" in x[2]:
                    a = modelNum[x]
                    a += (120.66 * 3)
                    modelNum.update({x: a})
                if "4" in x[2]:
                    a = modelNum[x]
                    a += (120.66 * 4)
                    modelNum.update({x: a})

            elif "12" in x[3:5] and "L" in x[11]:
                if "2" in x[2]:
                    a = modelNum[x]
                    a += (127.55 * 2)
                    modelNum.update({x: a})
                if "3" in x[2]:
                    a = modelNum[x]
                    a += (127.55 * 3)
                    modelNum.update({x: a})
                if "4" in x[2]:
                    a = modelNum[x]
                    a += (127.55 * 4)
                    modelNum.update({x: a})
            elif "12" in x[3:5] and "M" in x[11]:
                if "2" in x[2]:
                    a = modelNum[x]
                    a += (121.42 * 2)
                    modelNum.update({x: a})
                if "3" in x[2]:
                    a = modelNum[x]
                    a += (121.42 * 3)
                    modelNum.update({x: a})
                if "4" in x[2]:
                    a = modelNum[x]
                    a += (121.42 * 4)
                    modelNum.update({x: a})
            elif "14" in x[3:5]:
                pass


def priceColC():
    for x in modelNum:
        if "N" in x[2]:
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
        elif "14" in x[3:5]:
            a = modelNum[x]
            a += 59.75
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
            elif "14" in x[3:5]:
                a = modelNum[x]
                a += 194.54
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
            elif "14" in x[3:5]:
                a = modelNum[x]
                a += 198.87
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
        elif "14" in x[3:5]:
            if "2" in x[2]:
                a = modelNum[x]
                a += (59.75 * 2)
                modelNum.update({x: a})
            if "3" in x[2]:
                a = modelNum[x]
                a += (59.75 * 3)
                modelNum.update({x: a})
            if "4" in x[2]:
                a = modelNum[x]
                a += (59.75 * 4)
                modelNum.update({x: a})
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
            elif "14" in x[3:5]:
                if "2" in x[2]:
                    a = modelNum[x]
                    a += (194.54 * 2)
                    modelNum.update({x: a})
                if "3" in x[2]:
                    a = modelNum[x]
                    a += (194.54 * 3)
                    modelNum.update({x: a})
                if "4" in x[2]:
                    a = modelNum[x]
                    a += (194.54 * 4)
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
            elif "14" in x[3:5]:
                if "2" in x[2]:
                    a = modelNum[x]
                    a += (198.87 * 2)
                    modelNum.update({x: a})
                if "3" in x[2]:
                    a = modelNum[x]
                    a += (198.87 * 3)
                    modelNum.update({x: a})
                if "4" in x[2]:
                    a = modelNum[x]
                    a += (198.87 * 4)
                    modelNum.update({x: a})


def priceColD():
    for x in modelNum:
        if "N" in x[2]:
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
        elif "14" in x[3:5]:
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
        elif "14" in x[3:5]:
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


def priceColG():
    for x in modelNum:
        if "N" in x[6]:
            pass
        elif "I" in x[6]:
            ifColG(x)

    return "\n".join(modelNum)


def ifColG(x):
    if "N" in x[2]:
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
        elif "14" in x[3:5]:
            a = modelNum[x]
            a += (32.22)
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
        elif "14" in x[3:5]:
            a = modelNum[x]
            a += (32.22 * 2)
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
        elif "14" in x[3:5]:
            a = modelNum[x]
            a += (32.22 * 3)
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
        elif "14" in x[3:5]:
            a = modelNum[x]
            a += (32.22 * 4)
            modelNum.update({x: a})


def priceColHorI():
    for x in modelNum:
        if "N" in x[2]:
            pass
        else:
            ifColHorI(x)

    return "\n".join(modelNum)


def ifColHorI(x):
    if "A" in x[5]:
        if "2" in x[2]:
            if "08" in x[3:5]:
                pass
            elif "10" in x[3:5]:
                a = modelNum[x]
                a += 60.94
                modelNum.update({x: a})
            elif "12" in x[3:5]:
                a = modelNum[x]
                a += 61.11
                modelNum.update({x: a})
            elif "14" in x[3:5]:
                a = modelNum[x]
                a += 60.94
                modelNum.update({x: a})
        elif "3" in x[2]:
            if "08" in x[3:5]:
                pass
            elif "10" in x[3:5]:
                a = modelNum[x]
                a += 84.69
                modelNum.update({x: a})
            elif "12" in x[3:5]:
                a = modelNum[x]
                a += 78.41
                modelNum.update({x: a})
            elif "14" in x[3:5]:
                a = modelNum[x]
                a += 118.91
                modelNum.update({x: a})
        elif "4" in x[2]:
            if "08" in x[3:5]:
                pass
            elif "10" in x[3:5]:
                a = modelNum[x]
                a += 121.22
                modelNum.update({x: a})
            elif "12" in x[3:5]:
                a = modelNum[x]
                a += 122.22
                modelNum.update({x: a})
            elif "14" in x[3:5]:
                a = modelNum[x]
                a += 121.22
                modelNum.update({x: a})
    elif "H" in x[5]:
        if "2" in x[2]:
            if "08" in x[3:5]:
                pass
            elif "10" in x[3:5]:
                a = modelNum[x]
                a += 64.44
                modelNum.update({x: a})
            elif "12" in x[3:5]:
                a = modelNum[x]
                a += 57.94
                modelNum.update({x: a})
            elif "14" in x[3:5]:
                a = modelNum[x]
                a += 92.11
                modelNum.update({x: a})
        elif "3" in x[2]:
            if "08" in x[3:5]:
                pass
            elif "10" in x[3:5]:
                a = modelNum[x]
                a += 84.69
                modelNum.update({x: a})
            elif "12" in x[3:5]:
                a = modelNum[x]
                a += 109.16
                modelNum.update({x: a})
            elif "14" in x[3:5]:
                a = modelNum[x]
                a += 150.92
                modelNum.update({x: a})
        elif "4" in x[2]:
            if "08" in x[3:5]:
                pass
            elif "10" in x[3:5]:
                a = modelNum[x]
                a += 196.22
                modelNum.update({x: a})
            elif "12" in x[3:5]:
                a = modelNum[x]
                a += 179.92
                modelNum.update({x: a})
            elif "14" in x[3:5]:
                a = modelNum[x]
                a += 184.22
                modelNum.update({x: a})


def uvmCost():
    for x in modelNum:
        if "U" in x[12]:
            if "2" in x[2]:
                a = modelNum[x]
                a += (116.96*2)
                modelNum.update({x: a})
            elif "3" in x[2]:
                a = modelNum[x]
                a += (116.96*3)
                modelNum.update({x: a})
            elif "4" in x[2]:
                a = modelNum[x]
                a += (116.96*4)
                modelNum.update({x: a})
            else:
                a = modelNum[x]
                a += 116.96
                modelNum.update({x: a})

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
        else:
            a = modelNum[x]
            a += 10.00
            modelNum.update({x: a})



def writeModelNumbersToFile(newFileData):
    file = open("outputs/Pricing.txt", "w")
    file.write(newFileData)


def writeCostDictionaryToFile(newFileData):
    # with open('outputs/dictionary.txt', 'w') as f:
    #     print(newFileData, file=f)
    with open('outputs/dictionary.json', 'w') as file:
        json.dump(newFileData, file, indent=2)


def writeModelNumberstoSheet(newFileData):
    cell_list = finalSpreadSheet.range('Q2:Q1681')
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
    priceColG()
    priceColHorI()
    uvmCost()
    colK()
    colL()

    # print(modelNum)

    priceList = list(modelNum.values())
    print(priceList)

    writeCostDictionaryToFile(modelNum)  # writes to dictionary.json

    writeModelNumberstoSheet(priceList)  # writes to sheet
    print("Write Success")


if __name__ == "__main__":
    main()
