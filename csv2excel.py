import openpyxl
import csv
import os
import sys


def csvToExcelSingle(csvFilename, excelFilename):

    with open(csvFilename) as f:
        csv2DList = [r for r in csv.reader(f)]

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = 'Sheet1'

    for r in csv2DList:
        ws.append(r)

    wb.save(excelFilename)


def csvToExcelMultiple(csvFilenameList, excelFilenameList):

    for csvFilename, excelFilename in zip(csvFilenameList, excelFilenameList):
        csvToExcelSingle(csvFilename, excelFilename)


if __name__ == '__main__':

    commandLineArgument = sys.argv

    if len(commandLineArgument) == 1:
        sys.exit()

    csvFilenameList = commandLineArgument[1:]
    excelFilenameList = []

    for csvFilename in csvFilenameList:
        csvFilename0, csvFilename1 = os.path.splitext(csvFilename)
        if csvFilename1 == '.csv':
            excelFilename = csvFilename0 + '.xlsx'
            excelFilenameList.append(excelFilename)

    csvToExcelMultiple(csvFilenameList, excelFilenameList)
