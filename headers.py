from datetime import datetime
import datetime
import pandas


class Headers:

    def __init__(self):
        # define temp headers
        self.tempRMLS = 'daysDiff_RMLS'
        self.tempTermin = 'daysDiff_Termin'
        self.combinedColumns = 'Document Status'  # Combined docStatus & orderPhase
        self.combinedDates = 'combi-date'

        # define headers
        self.rmls = 'SollRückmeldetermin Leitstand'  # dd.mm.YYYY
        self.kTermin = 'Konstruktionstermin Soll'  # dd.mm.YYYY
        self.docType = 'Dokument'  # clean WAR,EBS & write HKB, GEN, GEL ; SU_?, CO_?
        self.bbNummer = 'BB-Nummer'
        self.docStatus = 'Dokumentstatus'  # 46, 47, 42, FG
        self.orderPhase = 'Auftragsphase'  # E, T
        self.sachbearbeiter = 'Sachbearbeiter'
        self.pivotTableItem1 = 'Gecikme Nedeni'
        self.pivotTableItem2 = 'Alınacak Aksiyon'
        self.pivotTableItem3 = 'Çalışma Tipi'

    def addPivotTableHeaders(self, sheetname):
        sheetname[self.pivotTableItem1] = " "
        sheetname[self.pivotTableItem2] = " "
        sheetname[self.pivotTableItem3] = " "

    def timeDiff(self, sheetname):
        # demoDate = datetime.datetime(2021, 04, 15).strftime('%d-%m-%Y')
        currentDate = datetime.datetime.now()
        daysDiff_Termin = sheetname[self.kTermin] - currentDate
        sheetname.insert(0, self.tempTermin, daysDiff_Termin.dt.days)
        daysDiff_RMLS = sheetname[self.rmls] - currentDate
        sheetname.insert(0, self.tempRMLS, daysDiff_RMLS.dt.days)

    def cleanByDate(self, sheetname):
        self.dayName = datetime.date.today().strftime('%A')

        if self.dayName == 'Monday' or self.dayName == 'Tuesday':
            for row in sheetname[self.tempRMLS]:
                if row != None and row >= 3:
                    rowIndex = next(iter(sheetname[sheetname[self.tempRMLS] == row].index), 'no match')
                    sheetname.drop(rowIndex, inplace=True)
            for row in sheetname[self.tempTermin]:
                if row != None and row >= 3:
                    rowIndex = next(iter(sheetname[sheetname[self.tempTermin] == row].index), 'no match')
                    sheetname.drop(rowIndex, inplace=True)

        elif self.dayName == 'Wednesday' or self.dayName == 'Thursday' or self.dayName == 'Friday':
            for row in sheetname[self.tempRMLS]:
                if row != None and row > 4:
                    rowIndex = next(iter(sheetname[sheetname[self.tempRMLS] == row].index), 'no match')
                    sheetname.drop(rowIndex, inplace=True)
            for row in sheetname[self.tempTermin]:
                if row != None and row > 4:
                    rowIndex = next(iter(sheetname[sheetname[self.tempTermin] == row].index), 'no match')
                    sheetname.drop(rowIndex, inplace=True)

    def rowCleaner_KEM(self, sheetname):
        for row in sheetname[self.docType]:
            if row[:3] == "WAR" or row[:3] == "EBS":
                rowIndex = next(iter(sheetname[sheetname[self.docType] == row].index), 'no match')
                sheetname.drop(rowIndex, inplace=True)

    def combineColumns(self, sheetname):
        # Combine documentStatus and orderPhase
        combinedColumn = sheetname[self.docStatus].apply(str) + "/" + sheetname[self.orderPhase]
        sheetname.insert(5, self.combinedColumns, combinedColumn)
        del sheetname[self.docStatus]
        del sheetname[self.orderPhase]

    def rowCleaner_docStatus(self, sheetname):
        for row in sheetname[self.combinedColumns]:
            if str(row) == "46/T":
                rowIndex = next(iter(sheetname[sheetname[self.combinedColumns] == row].index), 'no match')
                sheetname.drop(rowIndex, inplace=True)

    def cleanBy_BBnummer(self, sheetname):
        sheetname[self.bbNummer] = sheetname[self.bbNummer].fillna('-')
        for row in sheetname[self.bbNummer]:
            if row == '-':
                rowIndex = next(iter(sheetname[sheetname[self.bbNummer] == row].index), 'no match')
                sheetname.drop(rowIndex, inplace=True)

    def rowMark_GEL(self, sheetname):
        for row in sheetname[self.docType]:
            if row[:3] == "GEL":
                rowIndex = next(iter(sheetname[sheetname[self.docType] == row].index), 'no match')
                sheetname.loc[rowIndex, self.pivotTableItem3] = 'AKT'

    def rowMark_HKB(self, sheetname):
        for row in sheetname[self.docType]:
            if row[:3] == "HKB" or row[:3] == "GEN" or row[:3] == "KAT" or row[:3] == 'CO_' or row[:3] == 'SU_':
                rowIndex = next(iter(sheetname[sheetname[self.docType] == row].index), 'no match')
                sheetname.loc[rowIndex, self.pivotTableItem3] = 'RMLS'

    def rowMark_Mitteilung(self, sheetname):
        for index, value in sheetname[self.docType].items():
            if value[:2] == "ME" and len(value) == 12:
                sheetname.loc[index, self.pivotTableItem3] = 'Bildiri'

    def workingDays(self, sheetname):
        weekdays_rmls_list = sheetname[self.rmls].dt.day_name()
        weekdays_termin_list = sheetname[self.kTermin].dt.day_name()
        sheetname['combinedDays'] = weekdays_termin_list.fillna('') + weekdays_rmls_list.fillna('')

        sheetname['combinedDays'] = sheetname['combinedDays'].replace('Monday', 'Pazartesi Çalışılacak')
        sheetname['combinedDays'] = sheetname['combinedDays'].replace('Tuesday', 'Salı Çalışılacak')
        sheetname['combinedDays'] = sheetname['combinedDays'].replace('Wednesday', 'Çarşamba Çalışılacak')
        sheetname['combinedDays'] = sheetname['combinedDays'].replace('Thursday', 'Perşembe Çalışılacak')
        sheetname['combinedDays'] = sheetname['combinedDays'].replace('Friday', 'Cuma Çalışılacak')
        sheetname['combinedDays'] = sheetname['combinedDays'].replace('Saturday', 'Cuma Çalışılacak')
        sheetname['combinedDays'] = sheetname['combinedDays'].replace('Sunday', 'Cuma Çalışılacak')

        sheetname[self.pivotTableItem1] = sheetname['combinedDays']
        sheetname[self.pivotTableItem2] = sheetname['combinedDays']
        del sheetname['combinedDays']

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
        return self.switcher.get(argument, "Invalid day")

    def rowMark_Fehler(self, sheetname):
        for row in sheetname[self.docType]:
            if len(row) > 12 and row[:2] == 'ME':
                rowIndex = next(iter(sheetname[sheetname[self.docType] == row].index), 'no match')
                sheetname.loc[rowIndex, self.pivotTableItem1] = f'{self.daySwitch(self.dayName)} Çalışılacak'
                sheetname.loc[rowIndex, self.pivotTableItem2] = f'{self.daySwitch(self.dayName)} Çalışılacak'
                sheetname.loc[rowIndex, self.pivotTableItem3] = 'Fehler'

    def splitStatus(self, sheetname):
        status_list = []
        phases_list = []
        excel_status = sheetname[self.combinedColumns]

        for status in excel_status:
            status, phases = status.split('/', 1)
            status_list.append(status)
            phases_list.append(phases)

        sheetname.insert(5, self.docStatus, status_list)
        sheetname.insert(6, self.orderPhase, phases_list)

    def rowMark_Status47(self, sheetname):
        for index, value in sheetname[self.docStatus].items():
            if value == "47":
                sheetname.loc[index, self.pivotTableItem1] = f"{value}'de"
                sheetname.loc[index, self.pivotTableItem2] = f"{value}'de"
                sheetname.loc[index, self.pivotTableItem3] = f"{value}'de"
        del sheetname[self.combinedColumns]

    def isOpenScope(self, sheetname):
        for index, value in sheetname[self.tempRMLS].items():
            if value != None and value < -1:
                sheetname.loc[index, self.pivotTableItem1] = 'Açık Kapsam'
                sheetname.loc[index, self.pivotTableItem2] = 'Açık Kapsam'
                sheetname.loc[index, self.pivotTableItem3] = 'Kontrol Edilecek'

        del sheetname[self.tempRMLS]
        del sheetname[self.tempTermin]

    def create_PivotTable(self, sheetname):
        pivot_table = pandas.pivot_table(sheetname,
                                          index=self.sachbearbeiter,
                                          columns=[self.pivotTableItem2, self.pivotTableItem3],
                                          values=self.orderPhase,
                                          aggfunc='count',
                                          margins=True,
                                          margins_name='Toplam')
        return pivot_table
