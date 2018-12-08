# -*- coding: utf-8 -*-


import threading
import gspread
import ModelNumberGenerator as MNG
import ShortDescriptionGenerator as SDG
import LongDescriptionGenerator as LDG

import Spreadsheet as data

spreadsheet = data.sheet  # reference the JCI sheet as 'spreadsheet'
finalSpreadSheet = data.FINALSHEET

def main():
    a = threading.Thread(target=MNG.run())
    b = threading.Thread(target=SDG.run())
    c = threading.Thread(target=LDG.run())

    a.start()
    a.join()
    b.start()
    b.join()
    c.start()
    c.join()


if __name__ == "__main__":
    main()