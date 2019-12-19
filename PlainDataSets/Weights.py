import Spreadsheet as data

finalSpreadSheet = data.FINALSHEET
priceSheet = data.price

modelNumInvalidFormat = finalSpreadSheet.range('B2:B1681')

modelNum = {}


def loopList():
    for x in modelNumInvalidFormat:
        xString = str(x)
        xString = xString.split("'")
        modelNum[(xString[1])] = 0

def weightValues():
    for x in modelNum:
        if "N" in x[2]:
            weightSingle()
        elif "2" or "3" or "4" in x[2]:
            weightGanged()


def weightSingle():
    for x in modelNum:
        if "08" in x[3:5]:
            modelNum.update({x: 15})
        if "10" in x[3:5]:
            modelNum.update({x: 20})
        if "12" in x[3:5]:
            modelNum.update({x: 20})
        if "14" in x[3:5]:
            modelNum.update({x: 25})

def weightGanged():
    for x in modelNum:
        if "10" in x[3:5]:
            if "2" in x[2]:
                modelNum.update({x: 40})
            if "3" in x[2]:
                modelNum.update({x: 60})
            if "4" in x[2]:
                modelNum.update({x: 100})
        if "12" in x[3:5]:
            if "2" in x[2]:
                modelNum.update({x: 60})
            if "3" in x[2]:
                modelNum.update({x: 80})
            if "4" in x[2]:
                modelNum.update({x: 100})
        if "14" in x[3:5]:
            if "2" in x[2]:
                modelNum.update({x: 50})
            if "3" in x[2]:
                a = modelNum[x]
                a += (120.19 * 3)
                modelNum.update({x: 75})
            if "4" in x[2]:
                a = modelNum[x]
                a += (120.19 * 4)
                modelNum.update({x: 120})

def writeModelNumberstoSheet(newFileData):
    cell_list = finalSpreadSheet.range('BB2:BB1681')
    cell_values = newFileData

    print("Pricing cell list length: ", cell_list.__len__())
    print("Pricing cell value length: ", cell_values.__len__())

    for cell, x in enumerate(cell_values):
        cell_list[cell].value = x

    finalSpreadSheet.update_cells(cell_list)

def main():
    loopList()
    weightValues()
    print(modelNum)

    weightList = list(modelNum.values())
    print(weightList)
    writeModelNumberstoSheet(weightList)

if __name__ == '__main__':
    main()