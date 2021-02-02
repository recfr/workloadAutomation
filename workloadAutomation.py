import pandas as pd
import numpy as np
import headers
import locale
from datetime import datetime, timedelta
import openpyxl
import xlsxwriter

# define headers
myHeader = headers.Headers()


def cleanByDate(_tempRMLS, _tempTermin):
    for row in df[_tempRMLS]:
        if row != None:
            if row > 3:
                rowIndex = next(iter(df[df[_tempRMLS] == row].index), 'no match')
                df.drop(rowIndex, inplace=True)
    for row in df[_tempTermin]:
        if row != None:
            if row > 3:
                rowIndex = next(iter(df[df[_tempTermin] == row].index), 'no match')
                df.drop(rowIndex, inplace=True)


def timeDiff(sheetName, dateRMLS, dateTermin):
    # TODO into the real life app will use to today instead of testDay.
    # time_format = '%Y-%m-%d'
    # today = datetime.now()
    testDay = datetime.strptime('2021-01-27', '%Y-%m-%d')
    daysDiff_Termin = df[dateTermin] - testDay
    sheetName.insert(0, headers.Headers.tempTermin, daysDiff_Termin.dt.days)
    daysDiff_RMLS = df[dateRMLS] - testDay
    sheetName.insert(0, headers.Headers.tempRMLS, daysDiff_RMLS.dt.days)


def dayFinder():
    return datetime.datetime.now().strftime('%A')


def dayTranslator(argument):
    switcher = {
        'Monday': "Pazartesi",
        'Tuesday': "Salı",
        'Wednesday': "Çarşamba",
        'Thursday': "Perşembe",
        'Friday': "Cuma",
        'Saturday': "Cumartesi",
        'Sunday': "Pazar",
    }
    print(switcher.get(argument, "Invalid month"))


# clean KEM numbers
def rowCleaner_KEM(columnName):
    for row in df[columnName]:
        if row[:3] == "WAR" or row[:3] == "EBS":
            rowIndex = next(iter(df[df[columnName] == row].index), 'no match')
            df.drop(rowIndex, inplace=True)


# Combine documentStatus and orderPhase
def combineColumns(sheetName, firstColumn, secondColumn, position):
    combinedColumn = sheetName[firstColumn] + "/" + sheetName[secondColumn]
    sheetName.insert(position, headers.Headers.combinedColumns, combinedColumn)
    del sheetName[firstColumn]
    del sheetName[secondColumn]


def rowCleaner_docStatus(columnNameStatus):
    for row in df[columnNameStatus]:
        if str(row) == "46/T":
            rowIndex = next(iter(df[df[columnNameStatus] == row].index), 'no match')
            df.drop(rowIndex, inplace=True)


def addPivotTableHeaders():
    df[headers.Headers().pivotTableItem1] = " "
    df[headers.Headers().pivotTableItem2] = " "
    df[headers.Headers().pivotTableItem3] = " "


# read excel file
dailyWorkload = 'dailyworkload.xlsx'
df = pd.read_excel(dailyWorkload, sheet_name='Sheet1')

addPivotTableHeaders()
timeDiff(df, headers.Headers.rmls, headers.Headers.kTermin)
cleanByDate(headers.Headers.tempRMLS, headers.Headers.tempTermin)
rowCleaner_KEM(headers.Headers.kem)
combineColumns(df, headers.Headers.docStatus, headers.Headers.orderPhase, 5)
rowCleaner_docStatus(headers.Headers.combinedColumns)

# write output file
writer = pd.ExcelWriter("edited_Workload.xlsx",
                        engine='xlsxwriter',
                        datetime_format='dd.mm.yyyy', )
df.to_excel(writer, 'Sheet1')
writer.save()

# https://pandas.pydata.org/pandas-docs/stable/reference/api/pandas.DataFrame.between_time.html
# TODO Clean marked dates