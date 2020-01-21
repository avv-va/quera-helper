#!/usr/bin/env python
# -*- coding: utf-8 -*-

import xlrd, xlwt, glob

def generateNameAndID(writesheet):
    wr = xlrd.open_workbook('kargahvatamrin/kargah/jalase0/assignment_10568_results.xlsx') 
    readsheet = wr.sheet_by_index(0)       
    for i in range(2, readsheet.nrows):
        writesheet.write(i, 0, readsheet.cell_value(i,2)) # esm
        writesheet.write(i, 1, readsheet.cell_value(i,1)) # shomare daneshjuei


def generateFirstRow(writesheet):
    locations = glob.glob(location)
    for i in range(0, len(locations)):
        txt = locations[i].split('/')[2]   
        writesheet.write(0, i+2, txt)

def generateScore(location, writesheet):
    problem_index = 2 
    for loc in glob.glob(location):        
        wr = xlrd.open_workbook(loc) 
        readsheet = wr.sheet_by_index(0) 
        should_print_column = []
        for i in range(readsheet.ncols):
            if 'نمره داوری با تاخیر' == readsheet.cell_value(1, i) :
                should_print_column.append(i)

        writesheet.write(1, problem_index , len(should_print_column)*100) 
        writesheet.write(1,0,'SCORESUM')
        for i in range(2, readsheet.nrows):
            total_score = 0
            for j in should_print_column:
                if readsheet.cell_value(i, j) != '':
                    total_score += int(readsheet.cell_value(i, j))  
                     
            writesheet.write(i, problem_index, total_score)

        problem_index += 1
            
def generatesheet(wb, location, sheet):
    writesheet = wb.add_sheet(sheet, cell_overwrite_ok=True)
    # generate first row
    generateFirstRow(writesheet)
    # generate name and shomare daneshjuee  
    generateNameAndID(writesheet)
    # read and write scores 
    generateScore(location, writesheet)


if __name__ == '__main__':
    wb = xlwt.Workbook() 
    
    # generate kargah_sheets
    location = "kargahvatamrin/kargah/*/*.xlsx"
    generatesheet(wb, location,'کارگاه')

    # generate tamrin_sheets
    location = "kargahvatamrin/tamrin/*/*.xlsx"
    generatesheet(wb, location, 'تمرین')

    wb.save('ta.xls')


