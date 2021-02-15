import pandas as pd
from PyQt5 import QtCore, QtWidgets
from PyQt5.QtWidgets import QFileDialog, QMessageBox
from datetime import datetime
import datetime
import headers
import xlrd
import openpyxl
import xlsxwriter

class Ui_MainWindow(object):

    def __init__(self):
        self.headersObject = headers.Headers()

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

        self.browseButton = QtWidgets.QPushButton(self.centralwidget)
        self.browseButton.setGeometry(QtCore.QRect(320, 8, 75, 23))
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
        MainWindow.setWindowTitle(_translate("MainWindow", "MBT Workload Automation"))
        self.label.setText(_translate("MainWindow", "File Path"))
        self.cancelButton.setText(_translate("MainWindow", "Cancel"))
        self.executeButton.setText(_translate("MainWindow", "Execute"))
        self.browseButton.setText(_translate("MainWindow", "Browse"))
        self.actionAbout.setText(_translate("MainWindow", "About"))

    def initUI(self):
        self.cancelButton.clicked.connect(self.cancelButton_handler)
        self.browseButton.clicked.connect(self.browseButton_handler)
        self.executeButton.clicked.connect(self.executeButton_handler)

    def cancelButton_handler(self):
        print("cancelButton_handler")
        sys.exit(app.exec_())

    def browseButton_handler(self):
        print("browseButton_handler")
        self.openDialogBox()

    def executeButton_handler(self):
        self.executeExcel()

    def openDialogBox(self):
        self.path = QFileDialog.getOpenFileName()
        self.textBrowser.setPlainText(self.path[0])
        if self.path[0] is not None and self.path[0] != "":
            self.executeButton.setEnabled(True)

    def show_popUp(self, created_fileName):
        msg = QMessageBox()
        msg.setWindowTitle("Information")
        msg.setText(f'Workload list \n{created_fileName} \n created.')
        self.executeButton.setEnabled(False)
        x = msg.exec_()

    def executeExcel(self):
        print("executeButton_handler")
        dailyWorkload: str = self.path[0]
        workBook = pd.read_excel(dailyWorkload, sheet_name='Sheet1')

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

        # # TODO docStatus 47 should mark
        # # write output file
        tempNameData = datetime.date.today().strftime('%d-%m-%Y')
        fileNameData = "~/desktop/" + tempNameData + "_Workload.xlsx"
        writer = pd.ExcelWriter(fileNameData, engine='xlsxwriter', datetime_format='dd.mm.yyyy')
        workBook.to_excel(writer, 'Sheet1', index=False)
        writer.save()
        self.show_popUp(fileNameData[1:])


if __name__ == "__main__":
    import sys

    app = QtWidgets.QApplication(sys.argv)
    MainWindow = QtWidgets.QMainWindow()
    ui = Ui_MainWindow()
    ui.setupUi(MainWindow)
    MainWindow.show()
    sys.exit(app.exec_())