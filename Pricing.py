import time

import Spreadsheet as data

finalSpreadSheet = data.FINALSHEET
priceSheet = data.price

modelNumInvalidFormat = finalSpreadSheet.range('B2:B1873')
# colA08PA = priceSheet.acell("G9").value       #todo, decide between implement this or static values

modelNum = {}
priceList = {}


def loopList():
    for x in modelNumInvalidFormat:
        xString = str(x)
        xString = xString.split("'")
        modelNum[(xString[1])] = ""


def tempLoopList():
    with open("outputs/Pricing.txt") as f:
        for line in f:
            modelNum[str(line)] = ""


def priceColA():
    for x in modelNum:
        if "A" in x[5]:
            if "P" in x[9]:
                if "08" in x[3:5]:
                    # modelNum.update({x: colA08PA})
                    modelNum.update({x: 87.69})
                    # modelNum.update({x: priceSheet.acell("G9").value})
                elif "10" in x[3:5]:
                    modelNum.update({x: "88.23"})           #todo, convert these to ints
                    # time.sleep(.05)
                    # modelNum.update({x: priceSheet.acell("G10").value})
                elif "12" in x[3:5]:
                    modelNum.update({x: "99.12"})
                    # time.sleep(.05)
                    # modelNum.update({x: priceSheet.acell("G11").value})
                elif "14" in x[3:5]:
                    modelNum.update({x: "168.73"})
            elif "F" in x[9]:
                if "08" in x[3:5]:
                    modelNum.update({x: "93.08"})
                elif "10" in x[3:5]:
                    modelNum.update({x: "97.84"})
                elif "12" in x[3:5]:
                    modelNum.update({x: "111.20"})
                elif "14" in x[3:5]:
                    pass  # todo, define N/A in sheet
        elif "H" in x[5]:
            if "P" in x[9]:
                if "08" in x[3:5]:
                    modelNum.update({x: "120.19"})
                elif "10" in x[3:5]:
                    modelNum.update({x: "129.94"})
                elif "12" in x[3:5]:
                    modelNum.update({x: "145.12"})
                elif "14" in x[3:5]:
                    modelNum.update({x: "214.60"})
            elif "F" in x[9]:
                if "08" in x[3:5]:
                    modelNum.update({x: "131.62"})
                elif "10" in x[3:5]:
                    modelNum.update({x: "136.59"})
                elif "12" in x[3:5]:
                    modelNum.update({x: "160.16"})
                elif "14" in x[3:5]:
                    pass  # todo, define N/A in sheet

    return "\n".join(modelNum)

def priceColB():                #Todo, change values
    for x in modelNum:
        if "A" in x[5]:
            if "P" in x[9]:
                if "08" in x[3:5]:
                    a = modelNum[x]
                    a += 44.99
                    modelNum.update({x: a})
                elif "10" in x[3:5]:            #todo perform addition on cost
                    modelNum.update({x: "88.23"})
                elif "12" in x[3:5]:
                    modelNum.update({x: "99.12"})
                elif "14" in x[3:5]:
                    modelNum.update({x: "168.73"})
            elif "F" in x[9]:
                if "08" in x[3:5]:
                    modelNum.update({x: "93.08"})
                elif "10" in x[3:5]:
                    modelNum.update({x: "97.84"})
                elif "12" in x[3:5]:
                    modelNum.update({x: "111.20"})
                elif "14" in x[3:5]:
                    pass  # todo, define N/A in sheet
        elif "H" in x[5]:
            if "P" in x[9]:
                if "08" in x[3:5]:
                    modelNum.update({x: "120.19"})
                elif "10" in x[3:5]:
                    modelNum.update({x: "129.94"})
                elif "12" in x[3:5]:
                    modelNum.update({x: "145.12"})
                elif "14" in x[3:5]:
                    modelNum.update({x: "214.60"})
            elif "F" in x[9]:
                if "08" in x[3:5]:
                    modelNum.update({x: "131.62"})
                elif "10" in x[3:5]:
                    modelNum.update({x: "136.59"})
                elif "12" in x[3:5]:
                    modelNum.update({x: "160.16"})
                elif "14" in x[3:5]:
                    pass  # todo, define N/A in sheet

    return "\n".join(modelNum)


def writeModelNumbersToFile(newFileData):
    file = open("outputs/Pricing.txt", "w")
    file.write(newFileData)


def main():
    # loopList()
    tempLoopList()
    priceColA()
    priceColB()       #todo, update these values

    print(modelNum)


    # writeModelNumbersToFile(priceColA())


if __name__ == "__main__":
    main()
