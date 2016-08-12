# -*- coding: utf-8 -*-
from openpyxl import Workbook
import sys
wb = Workbook()
ws = wb.active
ws.title = "Pipe_Mto"
inputfile = sys.argv[1]
#inputfile = 'a.txt'
outputfile = sys.argv[2]
with open(inputfile) as infile:
    for line in infile:
        ws.append(line.decode('cp936').split("@"))


wb.save(outputfile)       
