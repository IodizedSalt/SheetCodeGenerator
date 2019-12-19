# -*- coding: utf-8 -*-
import threading
import gspread
import sys, os
from os.path import dirname, join, abspath


sys.path.insert(0, abspath(join(dirname(__file__), './PlainDataSets')))
sys.path.insert(0, abspath(join(dirname(__file__), './SS316_DataSheets')))


# import ModelNumberGenerator as MNG
# import ShortDescriptionGenerator as SDG
# import LongDescriptionGenerator as LDG

import SS316_ModelNumber as SSMNG
import SS316_ShortDescription as SSSDG
import SS316_LongDescription as SSLDG



import Spreadsheet as data

spreadsheet = data.sheet  # reference the JCI sheet as 'spreadsheet'
finalSpreadSheet = data.FINALSHEET



def main():
    # For the PlainDataSets

    # a = threading.Thread(target=MNG.run())
    # b = threading.Thread(target=SDG.run())
    # c = threading.Thread(target=LDG.run())
    #
    # a.start()
    # a.join()
    # b.start()
    # b.join()
    # c.start()
    # c.join()

    # For the SS316 DataSets
    a = threading.Thread(target=SSMNG.run())
    b = threading.Thread(target=SSSDG.run())
    c = threading.Thread(target=SSLDG.run())


    a.start()
    a.join()

    b.start()
    b.join()

    c.start()
    c.join()


if __name__ == "__main__":
    main()
