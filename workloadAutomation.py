import pandas as pd
import headers
import datetime
import openpyxl

# define headers
myHeader = headers.Headers()

# def today():
#     date = datetime.date.today()
#     formatedDate = date.strftime("%d.%m.%Y")
#     return formatedDate
#     # after3days = formatedDate[:2]
#     # # print(int(after3days)+3)

#clean KEM numbers
def rowCleaner_KEM(columnName):
    for row in df[columnName]:
        if row[:3] == "WAR" or row[:3] == "EBS":
            rowIndex = next(iter(df[df[columnName] == row].index), 'no match')
            df.drop(rowIndex, inplace=True)

#Combine documentStatus and orderPhase
def combineColumns(sheetName, firstColumn, secondColumn, position):
    combinedColumn = sheetName[firstColumn] + "/" + sheetName [secondColumn]
    sheetName.insert(position, headers.Headers.combinedColumns, combinedColumn)
    del sheetName[firstColumn]
    del sheetName[secondColumn]

def rowCleaner_docStatus(columnNameStatus):
    for row in df[columnNameStatus]:
        if str(row) == "46/T":
            rowIndex = next(iter(df[df[columnNameStatus] == row].index), 'no match')
            df.drop(rowIndex, inplace=True)

# read excel file
dailyWorkload = 'dailyworkload.xlsx'
df = pd.read_excel(dailyWorkload, sheet_name='Sheet1')

# adding pivot_table arguments
df[headers.Headers().pivotTableItem1] = 0
df[headers.Headers().pivotTableItem2] = 0
df[headers.Headers().pivotTableItem3] = 0


combineColumns(df, headers.Headers.docStatus, headers.Headers.orderPhase, 5)
rowCleaner_KEM(headers.Headers.kem)
rowCleaner_docStatus(headers.Headers.combinedColumns) # cleaning test docStatus correct

# print(df[headers.Headers.rmls][40:55])

writer = pd.ExcelWriter('edited_Workload.xlsx')
df.to_excel(writer, 'Sheet1')
writer.save()
