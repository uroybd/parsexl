"""This Module Parse xls/xslx files and return data in json."""

import xlrd, datetime
from collections import OrderedDict
import simplejson as json

""" This function take 5 arguments:
    inp = Input file
    outp = Output file
    sheet = Worksheet to work with in input file.
    start = Starting row
    end = Ending row
    fields = A list of field-names to be used in json."""

def xlparse(inp, outp, sheet, start, end, fields):
    inpt = inp
    outpt = outp
    wb = xlrd.open_workbook(inpt)
    sh = wb.sheet_by_name(sheet)
    json_list = []

    for rownum in range(start - 1, end):
        dicto = OrderedDict()
        row_values = sh.row_values(rownum)
        counter = 0
        for i in fields:
            if i.find('date') != -1:
                try:
                    timestr = xlrd.xldate_as_tuple(row_values[counter], wb.datemode)
                    dicto[i] =  str(datetime.datetime(*timestr)).split(' ')[0]
                except:
                    dicto[i] = row_values[counter]
            else:
                dicto[i] = row_values[counter]
            counter = counter + 1

        json_list.append(dicto)
    out = json.dumps(json_list)
    with open(outpt, 'w') as f:
        f.write(out)
