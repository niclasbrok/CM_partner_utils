import xlrd
import math
import os
from os import listdir

def ReadWriteXLSMatrix(fpath, sname, oname):
    workbook = xlrd.open_workbook(fpath)
    worksheet = workbook.sheet_by_name(sname)
    ofile = open('Temp/' + oname, 'w')
    print('Calculating: ' + fpath[:(len(fpath1) - 4)])
    print('--Number of rows: ' + str(worksheet.nrows))
    print('--Number of cols: ' + str(worksheet.ncols))
    for row_id in range(0, worksheet.nrows):
        potential_date = str(worksheet.cell(row_id, 0).value)
        flag = False
        try:
            if (len(potential_date) >= 4):
                str_date = potential_date[:4]
                try:
                    date = int(str_date)
                    flag = True
                except:
                    pass
            if flag:
                for col_id in range(1, 13):
                    potential_value = worksheet.cell(row_id, col_id).value
                    try:
                        value = float(potential_value)
                        if (value != 0.0):
                            ofile.write(str(date) + '-' + str(col_id).zfill(2) + '-' + '01' + ',' + str(value) + '\n')
                    except:
                        pass
        except:
            pass
    ofile.close()

def ReadWriteXLSColumn(fpath, sname, oname, col_id):
    workbook = xlrd.open_workbook(fpath)
    worksheet = workbook.sheet_by_name(sname)
    ofile = open('Temp/' + oname, 'w')
    print('--Number of rows: ' + str(worksheet.nrows))
    print('--Number of cols: ' + str(worksheet.ncols))
    for row_id in range(0, worksheet.nrows):
        potential_date = str(worksheet.cell(row_id, 0).value)
        flag = False
        try:
            if (len(potential_date) >= 4):
                str_date = potential_date[:4]
                try:
                    date = int(str_date)
                    flag = True
                except:
                    pass
                potential_value = worksheet.cell(row_id, col_id - 1).value
                try:
                    value = float(potential_value)
                except:
                    pass
            if flag:
                for nmonth in range(1, 13):
                    if (value != 0.0):
                        ofile.write(str(date) + '-' + str(nmonth).zfill(2) + '-' + '01' + ',' + str(value) + '\n')
        except:
            pass
    ofile.close()

def CombineTXTData(fpath1, fpath2, oname, operation, period):
    print('Calculating: ' + fpath1[:(len(fpath1) - 4)] + ' ' + operation + ' ' + fpath2[:(len(fpath2) - 4)])
    f1 = open(fpath1, 'r').readlines()
    f2 = open(fpath2, 'r').readlines()
    print('--' + fpath1[:(len(fpath1) - 4)] + ' contains ' + str(len(f1)) + ' observation')
    print('--' + fpath2[:(len(fpath2) - 4)] + ' contains ' + str(len(f2)) + ' observation')
    dates = []
    values = []
    for l1 in f1:
        data1 = l1.strip().split(",")
        d1 = data1[0]
        v1 = data1[1]
        for l2 in f2:
            data2 = l2.strip().split(",")
            d2 = data2[0]
            v2 = data2[1]
            if (d1 == d2):
                dates.append(d1)
                if (operation == '+'):
                    values.append(float(v1) + float(v2))
                elif (operation == '/'):
                    values.append(float(v1) / float(v2))
                elif (operation == '*'):
                    values.append(float(v1) * float(v2))
                elif (operation == '-'):
                    values.append(float(v1) - float(v2))
    ofile = open(oname, 'w')
    for k in range(0, math.floor(len(values)/period)):
        ofile.write(str(dates[k * period]) + ',' + str(values[k * period]) + '\n')
    ofile.close()

def MoveTXTData(fpathold, fpathnew, period):
    print('Moving: ' + fpathold[:(len(fpathold) - 4)] + ' To:  ' + fpathnew[:(len(fpathnew) - 4)])
    f = open(fpathold, 'r').readlines()
    ofile = open(fpathnew, 'w')
    for k in range(0, math.floor(len(f)/period)):
        ofile.write(f[k * period])
    ofile.close()

def ReadWriteXLSColumnIF(fpath, sname, countries, oname):
    workbook = xlrd.open_workbook(fpath)
    worksheet = workbook.sheet_by_name(sname[0])
    for country in countries:
        print('Processing: ' + sname[1] + ' printing to: ' + oname + sname[1] + '-' + country[0] + '.txt')
        name = oname + sname[1]  + '-' + country[0] + '.txt'
        ofile = open(name, 'w')
        col_id = country[1]
        is_empty = True
        for row_id in range(10, worksheet.nrows):
            potential_date = str(worksheet.cell(row_id, 0).value)
            flag = False
            try:
                date_tuple = xlrd.xldate_as_tuple(int(potential_date[:5]), workbook.datemode)
                date = str(date_tuple[0]) + '-' + str(date_tuple[1]).zfill(2) + '-' + str(date_tuple[2]).zfill(2)
                potential_value = worksheet.cell(row_id, col_id - 1).value
                value = float(potential_value)
                if (value != 0.0):
                    is_empty = False
                    ofile.write(date + ',' + str(value) + '\n')
            except:
                pass
        ofile.close()
        if is_empty:
            os.remove(os.getcwd() + '/' + name)
