import pandas as pd
import datetime
import openpyxl

# def today():
#     date = datetime.date.today()
#     formatedDate = date.strftime("%d.%m.%Y")
#     return formatedDate
#     # after3days = formatedDate[:2]
#     # # print(int(after3days)+3)
#
def rowCleaner_KEM(columnName):
    for row in df[columnName]:
        if row[:3] == "WAR" :
            rowIndex = next(iter(df[df[columnName] == row].index), 'no match')
            df.drop(rowIndex, inplace=True)


# def rowCleaner_Status(columnNameStatus, columnNamePhase):
#     for row in df[columnNameStatus]:
#         if row == 46:
#             print('kirk alti')
#             # print(next(iter(df[df[columnNamePhase] == row].index), 'no match'))
#             for row in df[columnNamePhase]:
#                 if row[:1] == 'T':
#                     rowIndex = next(iter(df[df[columnNamePhase] == row].index), 'no match')
#                     df.drop(rowIndex, inplace=True)
#                 else:
#                     continue
#         else:
#             continue


# read excel file
dailyWorkload = 'dailyworkload.xlsx'
df = pd.read_excel(dailyWorkload, sheet_name='Sheet1')

# adding pivot_table arguments
df['Gecikme Nedeni'] = 0
df['Alınacak Aksiyon'] = 0
df['Çalışma Tipi'] = 0

# define headers
rmls = 'SollRückmeldetermin Leitstand'      # dd.mm.YYYY
kTermin = 'Konstruktionstermin Soll'        # dd.mm.YYYY
kem = 'Dokument'                            # clean WAR, write HKB, GEN, GEL
# docStatus = 'Dokumentstatus'                # 46, 47, 42
# orderPhase = 'Auftragsphase'                # E, T


rowCleaner_KEM(kem)
# rowCleaner_Status(docStatus, orderPhase)


# writer = pd.ExcelWriter('edited_Workload.xlsx')
# df.to_excel(writer, 'Sheet1')
# writer.save()
