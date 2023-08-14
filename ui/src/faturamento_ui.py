import sys
import locale
import asyncio

from faturamento import Faturamento
from PySide6.QtWidgets import QApplication, QMainWindow
from window import Ui_MainWindow
from datetime import datetime


class MainWindow(QMainWindow, Ui_MainWindow):
    billing: Faturamento

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setupUi(self)
        # Create a Billing instance
        self.billing = Faturamento()
        self.billing.getDictionaryExams()
        # Set default year
        self.yearText.setTextMargins(10, 0, 0, 0)
        year = str(datetime.now().year)
        self.yearText.setText(year)

        # Set default month
        locale.setlocale(locale.LC_ALL, 'pt_BR')
        monthText = datetime.now().strftime('%B')

        self.monthText.setTextMargins(10, 0, 0, 0)
        self.monthText.setText(monthText)

        # Set default folder
        folderName = 'Faturamento'
        self.folderText.setText(folderName)
        self.folderText.setTextMargins(10, 0, 0, 0)

        # Set default sheetDetail
        # self.checkBoxDetail.setChecked(True)

        # Set button method
        self.buttonSend.clicked.connect(self.teste)

    def callGeneratedFiles(self):

        isDetail = mainWindow.checkBoxDetail.isChecked()

        self.billing.setParamsBilling(self.yearText.text(),
                                      self.monthText.text(),
                                      self.folderText.text(), isDetail)

        missingCompanys = asyncio.run(self.billing.generatedFiles())

        self.resultView.addItems(missingCompanys)

    def teste(self):
        if (self.folderText.text() == 'save'):
            self.saveExams()
        else:
            self.getExams()

    def saveExams(self):
        self.billing.saveDictionaryExams()

    def getExams(self):
        dictionary_exams = self.billing.getDictionaryExams()
        return dictionary_exams


if __name__ == '__main__':
    app = QApplication(sys.argv)
    mainWindow = MainWindow()
    mainWindow.show()

    app.exec()
