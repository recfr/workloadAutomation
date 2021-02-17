from datetime import datetime
import datetime

class Headers:

    def __init__(self):
        # define temp headers
        self.tempRMLS = 'daysDiff_RMLS'
        self.tempTermin = 'daysDiff_Termin'
        self.combinedColumns = 'Document Status'  # Combined docStatus & orderPhase

        # define headers
        self.rmls = 'SollRückmeldetermin Leitstand'  # dd.mm.YYYY
        self.kTermin = 'Konstruktionstermin Soll'  # dd.mm.YYYY
        self.docType = 'Dokument'  # clean WAR,EBS & write HKB, GEN, GEL ; SU_?, CO_?
        self.bbNummer = 'BB-Nummer'
        self.docStatus = 'Dokumentstatus'  # 46, 47, 42, FG
        self.orderPhase = 'Auftragsphase'  # E, T
        self.pivotTableItem1 = 'Gecikme Nedeni'
        self.pivotTableItem2 = 'Alınacak Aksiyon'
        self.pivotTableItem3 = 'Çalışma Tipi'

    def addPivotTableHeaders(self, sheetName):
        sheetName[self.pivotTableItem1] = " "
        sheetName[self.pivotTableItem2] = " "
        sheetName[self.pivotTableItem3] = " "

    def timeDiff(self, sheetName):
        daysDiff_Termin = sheetName[self.kTermin] - datetime.datetime.now()
        sheetName.insert(0, self.tempTermin, daysDiff_Termin.dt.days)
        daysDiff_RMLS = sheetName[self.rmls] - datetime.datetime.now()
        sheetName.insert(0, self.tempRMLS, daysDiff_RMLS.dt.days)

    def cleanByDate(self, sheetName):
        self.dayName = datetime.date.today().strftime('%A')

        if self.dayName == 'Monday' or self.dayName == 'Tuesday':
            for row in sheetName[self.tempRMLS]:
                if row != None and row > 3:
                    rowIndex = next(iter(sheetName[sheetName[self.tempRMLS] == row].index), 'no match')
                    sheetName.drop(rowIndex, inplace=True)
            for row in sheetName[self.tempTermin]:
                if row != None and row > 3:
                    rowIndex = next(iter(sheetName[sheetName[self.tempTermin] == row].index), 'no match')
                    sheetName.drop(rowIndex, inplace=True)

        elif self.dayName == 'Wednesday' or self.dayName == 'Thursday' or self.dayName == 'Friday':
            for row in sheetName[self.tempRMLS]:
                if row != None and row > 4:
                    rowIndex = next(iter(sheetName[sheetName[self.tempRMLS] == row].index), 'no match')
                    sheetName.drop(rowIndex, inplace=True)
            for row in sheetName[self.tempTermin]:
                if row != None and row > 4:
                    rowIndex = next(iter(sheetName[sheetName[self.tempTermin] == row].index), 'no match')
                    sheetName.drop(rowIndex, inplace=True)

        del sheetName[self.tempRMLS]
        del sheetName[self.tempTermin]

    def rowCleaner_KEM(self, sheetName):
        for row in sheetName[self.docType]:
            if row[:3] == "WAR" or row[:3] == "EBS":
                rowIndex = next(iter(sheetName[sheetName[self.docType] == row].index), 'no match')
                sheetName.drop(rowIndex, inplace=True)

    def combineColumns(self, sheetName):
        # Combine documentStatus and orderPhase
        combinedColumn = sheetName[self.docStatus].apply(str) + "/" + sheetName[self.orderPhase]
        sheetName.insert(5, self.combinedColumns, combinedColumn)
        del sheetName[self.docStatus]
        del sheetName[self.orderPhase]

    def rowCleaner_docStatus(self, sheetName):
        for row in sheetName[self.combinedColumns]:
            if str(row) == "46/T":
                rowIndex = next(iter(sheetName[sheetName[self.combinedColumns] == row].index), 'no match')
                sheetName.drop(rowIndex, inplace=True)

    def cleanBy_BBnummer(self, sheetName):
        sheetName[self.bbNummer] = sheetName[self.bbNummer].fillna('-')
        for row in sheetName[self.bbNummer]:
            if row == '-':
                rowIndex = next(iter(sheetName[sheetName[self.bbNummer] == row].index), 'no match')
                sheetName.drop(rowIndex, inplace=True)

    def rowMark_GEL(self, sheetName):
        for row in sheetName[self.docType]:
            if row[:3] == "GEL":
                rowIndex = next(iter(sheetName[sheetName[self.docType] == row].index), 'no match')
                sheetName.loc[rowIndex, self.pivotTableItem3] = 'AKT'

    def rowMark_HKB(self, sheetName):
        for row in sheetName[self.docType]:
            if row[:3] == "HKB" or row[:3] == "GEN" or row[:3] == "KAT" or row[:3] == 'CO_' or row[:3] == 'SU_':
                rowIndex = next(iter(sheetName[sheetName[self.docType] == row].index), 'no match')
                sheetName.loc[rowIndex, self.pivotTableItem3] = 'RMLS'

    def rowMark_Mitteilung(self, sheetName):
        for row in sheetName[self.docType]:
            if len(row) == 12 and row[:2] == "ME":
                rowIndex = next(iter(sheetName[sheetName[self.docType] == row].index), 'no match')
                sheetName.loc[rowIndex, self.pivotTableItem3] = 'Bildiri'

    def workingDays(self, sheetName):
        weekdays_rmls_list = sheetName[self.rmls].dt.day_name()
        weekdays_termin_list = sheetName[self.kTermin].dt.day_name()
        sheetName['combinedDays'] = weekdays_termin_list.fillna('') + weekdays_rmls_list.fillna('')

        sheetName['combinedDays'] = sheetName['combinedDays'].replace('Monday', 'Pazartesi Çalışılacak')
        sheetName['combinedDays'] = sheetName['combinedDays'].replace('Tuesday', 'Salı Çalışılacak')
        sheetName['combinedDays'] = sheetName['combinedDays'].replace('Wednesday', 'Çarşamba Çalışılacak')
        sheetName['combinedDays'] = sheetName['combinedDays'].replace('Thursday', 'Perşembe Çalışılacak')
        sheetName['combinedDays'] = sheetName['combinedDays'].replace('Friday', 'Cuma Çalışılacak')
        sheetName['combinedDays'] = sheetName['combinedDays'].replace('Saturday', 'Cuma Çalışılacak')
        sheetName['combinedDays'] = sheetName['combinedDays'].replace('Sunday', 'Cuma Çalışılacak')

        sheetName[self.pivotTableItem1] = sheetName['combinedDays']
        sheetName[self.pivotTableItem2] = sheetName['combinedDays']
        del sheetName['combinedDays']

    def daySwitch(self, argument):
        self.switcher = {
            'Monday': "Pazartesi",
            'Tuesday': "Salı",
            'Wednesday': "Çarşamba",
            'Thursday': "Perşembe",
            'Friday': "Cuma",
            'Saturday': "Cumartesi",
            'Sunday': "Pazar"
        }
        return self.switcher.get(argument, "Invalid month")

    def rowMark_Fehler(self, sheetName):
        for row in sheetName[self.docType]:
            if len(row) > 12 and row[:2] == 'ME':
                rowIndex = next(iter(sheetName[sheetName[self.docType] == row].index), 'no match')
                sheetName.loc[rowIndex, self.pivotTableItem1] = f'{self.daySwitch(self.dayName)} Çalışılacak'
                sheetName.loc[rowIndex, self.pivotTableItem2] = f'{self.daySwitch(self.dayName)} Çalışılacak'
                sheetName.loc[rowIndex, self.pivotTableItem3] = 'Fehler'

    def splitStatus(self, sheetName):
        status_list = []
        phases_list = []
        excel_status = sheetName[self.combinedColumns]

        for status in excel_status:
            status, phases = status.split('/', 1)
            status_list.append(status)
            phases_list.append(phases)

        sheetName.insert(5, self.docStatus, status_list)
        sheetName.insert(6, self.orderPhase, phases_list)

    def rowMark_Status47(self, sheetName):
        for index, value in sheetName[self.docStatus].items():
            # print(f"Index : {index}, Value : {value}")
            if value == "47":
                sheetName.loc[index, self.pivotTableItem1] = f"{value}'de"
                sheetName.loc[index, self.pivotTableItem2] = f"{value}'de"
                sheetName.loc[index, self.pivotTableItem3] = f"{value}'de"
        del sheetName[self.combinedColumns]