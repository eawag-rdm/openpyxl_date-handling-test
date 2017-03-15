#!/usr/bin/env python
# _*_ coding: utf-8 _*_

import sys
import datetime
import openpyxl

target = 'testdates_' + openpyxl.__version__ + '.xlsx'

def writetest():
    dd = range(0, 3) + range(57, 61)
    dates = [datetime.date(1900, 1, 1) + datetime.timedelta(days=d) for d in dd]
    isodates = [d.isoformat() for d in dates]
    tt = [(h, m, s) for h, m, s in zip(range(0, 4) + range(21, 24),
                                       range(0, 4) + range(57, 60),
                                       range(0, 4) + range(57, 60))]
    times = [datetime.time(h, m, s) for h, m, s in  tt]
    isotimes = [t.isoformat() for t in times]

    if openpyxl.__version__ == '2.5.0':
        wb = openpyxl.Workbook(iso_dates=True)
        print('Using openpyxl 2.5.0')
    else:
        wb = openpyxl.Workbook()

    ws = wb.active
    ws.append(['Ord', 'string_date', 'written_date',
               'string_time', 'written_time', 'read_date', 'read_time'])
    for d in zip([x + 1 for x in dd], isodates, dates, isotimes, times):
        ws.append(d)
    wb.save(target)
    wb.close()

def readwritetest():
    wb = openpyxl.load_workbook(target)
    ws = wb.active
    dates = [d[0].value for d in ws['C2':'C8']]
    times = [d[0].value for d in ws['E2':'E8']]
    for i, row in enumerate(range(2, 9)):
        ws.cell(row=row, column=6).value = dates[i]
        ws.cell(row=row, column=7).value = times[i]
    wb.save(target)
    wb.close()

if __name__ == '__main__':
    arg = sys.argv[1]
    if arg == 'write':
        writetest()
    if arg == 'read':
        readwritetest()
