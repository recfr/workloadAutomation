import pandas as pd
import headers
from datetime import datetime
import datetime

# define headers
headerObj = headers.Headers()


def cleanByDate(sheetName, _tempRMLS, _tempTermin):
    dayName = datetime.date.today().strftime('%A')

    if dayName == 'Monday' or dayName == 'Tuesday':
        for row in sheetName[_tempRMLS]:
            if row != None and row > 3:
                rowIndex = next(iter(sheetName[sheetName[_tempRMLS] == row].index), 'no match')
                sheetName.drop(rowIndex, inplace=True)
        for row in sheetName[_tempTermin]:
            if row != None and row > 3:
                rowIndex = next(iter(sheetName[sheetName[_tempTermin] == row].index), 'no match')
                sheetName.drop(rowIndex, inplace=True)

    elif dayName == 'Wednesday' or dayName == 'Thursday' or dayName == 'Friday':
        for row in sheetName[_tempRMLS]:
            if row != None and row > 4:
                rowIndex = next(iter(sheetName[sheetName[_tempRMLS] == row].index), 'no match')
                sheetName.drop(rowIndex, inplace=True)
        for row in sheetName[_tempTermin]:
            if row != None and row > 4:
                rowIndex = next(iter(sheetName[sheetName[_tempTermin] == row].index), 'no match')
                sheetName.drop(rowIndex, inplace=True)

    del sheetName[_tempRMLS]
    del sheetName[_tempTermin]


def timeDiff(sheetName, dateRMLS, dateTermin):
    # TODO into the real life app will use to today instead of testDay.
    # time_format = '%Y-%m-%d'
    # today = datetime.datetime.now()
    # testDay = datetime.datetime.strptime('2021-01-27', '%Y-%m-%d')
    daysDiff_Termin = df[dateTermin] - datetime.datetime.now()
    sheetName.insert(0, headerObj.tempTermin, daysDiff_Termin.dt.days)
    daysDiff_RMLS = df[dateRMLS] - datetime.datetime.now()
    sheetName.insert(0, headerObj.tempRMLS, daysDiff_RMLS.dt.days)


def rowCleaner_KEM(sheetName, columnName):
    for row in sheetName[columnName]:
        if row[:3] == "WAR" or row[:3] == "EBS":
            rowIndex = next(iter(sheetName[sheetName[columnName] == row].index), 'no match')
            sheetName.drop(rowIndex, inplace=True)


def combineColumns(sheetName, firstColumn, secondColumn, position):
    # Combine documentStatus and orderPhase
    combinedColumn = sheetName[firstColumn].apply(str) + "/" + sheetName[secondColumn]
    sheetName.insert(position, headerObj.combinedColumns, combinedColumn)
    del sheetName[firstColumn]
    del sheetName[secondColumn]


def rowCleaner_docStatus(sheetName, columnNameStatus):
    for row in sheetName[columnNameStatus]:
        if str(row) == "46/T":
            rowIndex = next(iter(sheetName[sheetName[columnNameStatus] == row].index), 'no match')
            sheetName.drop(rowIndex, inplace=True)


def addPivotTableHeaders():
    df[headerObj.pivotTableItem1] = " "
    df[headerObj.pivotTableItem2] = " "
    df[headerObj.pivotTableItem3] = " "


def cleanBy_BBnummer(sheetName, columnName_BBnummer):
    sheetName[columnName_BBnummer] = sheetName[columnName_BBnummer].fillna('-')
    for row in sheetName[columnName_BBnummer]:
        if row == '-':
            rowIndex = next(iter(sheetName[sheetName[headerObj.bbNummer] == row].index), 'no match')
            sheetName.drop(rowIndex, inplace=True)


def rowMark_GEL(sheetName, columnName):
    for row in sheetName[columnName]:
        if row[:3] == "GEL":
            rowIndex = next(iter(sheetName[sheetName[columnName] == row].index), 'no match')
            sheetName.loc[rowIndex, headerObj.pivotTableItem3] = 'AKT'


def rowMark_HKB(sheetName, columnName):
    for row in sheetName[columnName]:
        if row[:3] == "HKB" or row[:3] == "GEN" or row[:3] == "KAT" or row[:3] == 'CO_' or row[:3] == 'SU_':
            rowIndex = next(iter(sheetName[sheetName[columnName] == row].index), 'no match')
            sheetName.loc[rowIndex, headerObj.pivotTableItem3] = 'RMLS'


def splitStatus(sheetName, combinedDocStatus, _docStatus, _orderPhase, pos1, pos2):
    status_list = []
    phases_list = []
    excel_status = df[combinedDocStatus]

    for status in excel_status:
        status, phases = status.split('/', 1)
        status_list.append(status)
        phases_list.append(phases)

    sheetName.insert(pos1, _docStatus, status_list)
    sheetName.insert(pos2, _orderPhase, phases_list)
    del sheetName[combinedDocStatus]


def workingDays(sheetName, dateRMLS, dateTermin):
    weekdays_rmls_list = sheetName[dateRMLS].dt.day_name()
    weekdays_termin_list = sheetName[dateTermin].dt.day_name()
    sheetName['combinedDays'] = weekdays_termin_list.fillna('') + weekdays_rmls_list.fillna('')

    sheetName['combinedDays'] = sheetName['combinedDays'].replace('Monday', 'Pazartesi Çalışılacak')
    sheetName['combinedDays'] = sheetName['combinedDays'].replace('Tuesday', 'Salı Çalışılacak')
    sheetName['combinedDays'] = sheetName['combinedDays'].replace('Wednesday', 'Çarşamba Çalışılacak')
    sheetName['combinedDays'] = sheetName['combinedDays'].replace('Thursday', 'Perşembe Çalışılacak')
    sheetName['combinedDays'] = sheetName['combinedDays'].replace('Friday', 'Cuma Çalışılacak')

    sheetName[headerObj.pivotTableItem1] = sheetName['combinedDays']
    sheetName[headerObj.pivotTableItem2] = sheetName['combinedDays']
    del sheetName['combinedDays']


def rowMark_Fehler(sheetName, columnName):
    for row in sheetName[columnName]:
        if len(row) > 12 and row[:2] == 'ME':
            rowIndex = next(iter(sheetName[sheetName[columnName] == row].index), 'no match')
            sheetName.loc[rowIndex, headerObj.pivotTableItem1] = 'Bugün Çalışılacak'
            sheetName.loc[rowIndex, headerObj.pivotTableItem2] = 'Bugün Çalışılacak'
            sheetName.loc[rowIndex, headerObj.pivotTableItem3] = 'Fehler'


# read excel file
dailyWorkload = 'dailyworkload.xlsx'
df = pd.read_excel(dailyWorkload, sheet_name='Sheet1')

addPivotTableHeaders()
timeDiff(df, headerObj.rmls, headerObj.kTermin)
cleanByDate(df, headerObj.tempRMLS, headerObj.tempTermin)
rowCleaner_KEM(df, headerObj.docType)
combineColumns(df, headerObj.docStatus, headerObj.orderPhase, 5)
rowCleaner_docStatus(df, headerObj.combinedColumns)
cleanBy_BBnummer(df, headerObj.bbNummer)
rowMark_GEL(df, headerObj.docType)
rowMark_HKB(df, headerObj.docType)
workingDays(df, headerObj.rmls, headerObj.kTermin)
rowMark_Fehler(df, headerObj.docType)
splitStatus(df, headerObj.combinedColumns, headerObj.docStatus, headerObj.orderPhase, 5, 6)

# for row in df[headerObj.combinedColumns]:
#     if str(row) == "47/T" or str(row) == "47/E":
#         rowIndex = next(iter(df[df[headerObj.combinedColumns] == row].index), 'no match')
#         df.loc[rowIndex, headerObj.pivotTableItem1] = "47'de"
#         df.loc[rowIndex, headerObj.pivotTableItem2] = "47'de"
#         df.loc[rowIndex, headerObj.pivotTableItem3] = "47'de"


# write output file
writer = pd.ExcelWriter("edited_Workload.xlsx",
                        engine='xlsxwriter',
                        datetime_format='dd.mm.yyyy', )
df.to_excel(writer, 'Sheet1', index=False)
writer.save()

