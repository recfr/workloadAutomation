import pandas as pd
from PyQt5 import QtCore, QtWidgets
from PyQt5.QtWidgets import QFileDialog, QMessageBox
from datetime import datetime
import datetime
import headers
import outlook
import xlrd
import openpyxl
import xlsxwriter


class Ui_MainWindow(object):

    def __init__(self):
        self.headersObject = headers.Headers()
        self.outlook = outlook.EmailSender()

    def setupUi(self, MainWindow):
        MainWindow.setObjectName("MainWindow")
        MainWindow.setFixedSize(405, 100)
        self.centralwidget = QtWidgets.QWidget(MainWindow)
        self.centralwidget.setObjectName("centralwidget")

        self.textBrowser = QtWidgets.QTextBrowser(self.centralwidget)
        self.textBrowser.setGeometry(QtCore.QRect(60, 8, 245, 20))
        self.textBrowser.setFrameShadow(QtWidgets.QFrame.Sunken)
        self.textBrowser.setVerticalScrollBarPolicy(QtCore.Qt.ScrollBarAlwaysOff)
        self.textBrowser.setObjectName("textBrowser")

        self.label = QtWidgets.QLabel(self.centralwidget)
        self.label.setGeometry(QtCore.QRect(10, 10, 45, 15))
        self.label.setObjectName("label")

        self.cancelButton = QtWidgets.QPushButton(self.centralwidget)
        self.cancelButton.setGeometry(QtCore.QRect(320, 50, 75, 23))
        self.cancelButton.setObjectName("cancelButton")

        self.executeButton = QtWidgets.QPushButton(self.centralwidget)
        self.executeButton.setGeometry(QtCore.QRect(240, 50, 75, 23))
        self.executeButton.setObjectName("executeButton")
        self.executeButton.setEnabled(False)

        self.attachEmailButton = QtWidgets.QPushButton(self.centralwidget)
        self.attachEmailButton.setGeometry(QtCore.QRect(140, 50, 95, 23))
        self.attachEmailButton.setObjectName("attachEmailButton")
        self.attachEmailButton.setEnabled(False)

        self.browseButton = QtWidgets.QPushButton(self.centralwidget)
        self.browseButton.setGeometry(QtCore.QRect(320, 8, 75, 21))
        self.browseButton.setObjectName("browseButton")

        MainWindow.setCentralWidget(self.centralwidget)

        self.menubar = QtWidgets.QMenuBar(MainWindow)
        self.menubar.setGeometry(QtCore.QRect(0, 0, 406, 21))
        self.menubar.setObjectName("menubar")
        MainWindow.setMenuBar(self.menubar)

        self.statusbar = QtWidgets.QStatusBar(MainWindow)
        self.statusbar.setObjectName("statusbar")
        MainWindow.setStatusBar(self.statusbar)

        self.actionAbout = QtWidgets.QAction(MainWindow)
        self.actionAbout.setObjectName("actionAbout")

        self.retranslateUi(MainWindow)
        QtCore.QMetaObject.connectSlotsByName(MainWindow)
        self.initUI()

    def retranslateUi(self, MainWindow):
        _translate = QtCore.QCoreApplication.translate
        MainWindow.setWindowTitle(_translate("MainWindow", "MBT Workload Automation v2.0"))
        self.label.setText(_translate("MainWindow", "File Path"))
        self.cancelButton.setText(_translate("MainWindow", "Cancel"))
        self.executeButton.setText(_translate("MainWindow", "Execute"))
        self.browseButton.setText(_translate("MainWindow", "Browse"))
        self.actionAbout.setText(_translate("MainWindow", "About"))
        self.attachEmailButton.setText(_translate("MainWindow", "Create an E-mail"))

    def initUI(self):
        self.cancelButton.clicked.connect(self.cancelButton_handler)
        self.browseButton.clicked.connect(self.browseButton_handler)
        self.executeButton.clicked.connect(self.executeButton_handler)
        self.attachEmailButton.clicked.connect(self.attachToEmail_handler)

    def cancelButton_handler(self):
        sys.exit(app.exec_())

    def browseButton_handler(self):
        self.openDialogBox()

    def executeButton_handler(self):
        self.executeExcel()

    def attachToEmail_handler(self):
        self.attachEmail()

    def openDialogBox(self):
        path = "C:\\"
        filter = "Excel file (*.xlsx)"
        self.path = QFileDialog.getOpenFileName(QFileDialog(), "Select file", path, filter)
        self.textBrowser.setPlainText(self.path[0])
        if self.path[0] != None and self.path[0] != "":
            self.executeButton.setEnabled(True)

    def show_popUp(self, created_fileName):
        msg = QMessageBox()
        msg.setWindowTitle("Information")
        msg.setText(f'Workload list created. \n{created_fileName}')
        msg.setIcon(QMessageBox.Information)
        self.executeButton.setEnabled(False)
        self.attachEmailButton.setEnabled(True)
        x = msg.exec_()

    def warn_popUp(self):
        msg = QMessageBox()
        msg.setWindowTitle("Warning")
        msg.setText("The input excel file is not compatible.")
        msg.setIcon(QMessageBox.Warning)
        msg.setDetailedText("Columns' order must be as follows. Please check your exported excel file from SAP."
                            "\n\nSollRückmeldetermin Leitstand"
                            "\nKonstruktionstermin Soll"
                            "\nDokument"
                            "\nBeschreibung"
                            "\nBB-Nummer"
                            "\nDokumentstatus"
                            "\nAuftragsphase"
                            "\nBB Beschreibung"
                            "\nKSW-Status"
                            "\nDokumentennummer Maßnahme"
                            "\nSachbearbeiter")
        x = msg.exec_()

    def executeExcel(self):
        correctHeaderSet = ['SollRückmeldetermin Leitstand', 'Konstruktionstermin Soll', 'Dokument',
                            'Beschreibung', 'BB-Nummer', 'Dokumentstatus', 'Auftragsphase',
                            'BB Beschreibung', 'KSW-Status', 'Dokumentennummer Maßnahme',
                            'Sachbearbeiter']

        dailyWorkload: str = self.path[0]
        workBook = pd.read_excel(dailyWorkload, sheet_name='Sheet1')

        pSeries1 = pd.Series(correctHeaderSet)
        pSeries2 = pd.Series(workBook.columns.values)
        isTrue = pSeries1.equals(other=pSeries2)

        # TODO :update: init text Worksheet name must be 'Sheet1'

        if isTrue:
            self.headersObject.addPivotTableHeaders(workBook)
            self.headersObject.timeDiff(workBook)
            self.headersObject.cleanByDate(workBook)
            self.headersObject.rowCleaner_KEM(workBook)
            self.headersObject.combineColumns(workBook)
            self.headersObject.rowCleaner_docStatus(workBook)
            self.headersObject.cleanBy_BBnummer(workBook)
            self.headersObject.rowMark_GEL(workBook)
            self.headersObject.rowMark_HKB(workBook)
            self.headersObject.rowMark_Mitteilung(workBook)
            self.headersObject.workingDays(workBook)
            self.headersObject.rowMark_Fehler(workBook)
            self.headersObject.splitStatus(workBook)
            self.headersObject.isOpenScope(workBook)
            self.headersObject.rowMark_Status47(workBook)
            pivot = self.headersObject.create_PivotTable(workBook)

            # write output file
            self.tempNameData = datetime.date.today().strftime('%d-%m-%Y')
            fileNameData = "~/desktop/" + self.tempNameData + "_Workload.xlsx"
            writer = pd.ExcelWriter(fileNameData, engine='xlsxwriter', datetime_format='dd.mm.yyyy')
            workBook.to_excel(writer, 'Sheet1', index=False)
            pivot.to_excel(writer, 'Pivot Table', index=True)
            writer.save()
            self.show_popUp(fileNameData[1:])
        else:
            self.warn_popUp()

    def attachEmail(self):
        workBook_path = self.tempNameData + "_Workload.xlsx"
        self.outlook.createNewMail(workBook_path)
        self.attachEmailButton.setEnabled(False)


if __name__ == "__main__":
    import sys

    app = QtWidgets.QApplication(sys.argv)
    MainWindow = QtWidgets.QMainWindow()
    ui = Ui_MainWindow()
    ui.setupUi(MainWindow)
    MainWindow.show()
    sys.exit(app.exec_())
