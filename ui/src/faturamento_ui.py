import sys
import locale
import asyncio
import os

from faturamento import Faturamento
from PySide6.QtWidgets import QApplication, QMainWindow
from PySide6.QtCore import Slot
from window import Ui_MainWindow
from datetime import datetime, date, timedelta


class MainWindow(QMainWindow, Ui_MainWindow):
    billing: Faturamento

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setupUi(self)
        # Create a Billing instance
        self.billing = Faturamento()
        # Get a DictionaryExams data from a file
        self.billing.getDictionaryExams()
        # Set default year
        self.yearText.setTextMargins(10, 0, 0, 0)
        year = str(datetime.now().year)
        self.yearText.setText(year)

        # Set default month
        locale.setlocale(locale.LC_ALL, 'pt_BR')
        # Get Last month
        monthText = (date.today().replace(day=1) -
                     timedelta(days=1)).strftime('%B')
        self.monthText.setTextMargins(10, 0, 0, 0)
        self.monthText.setText(monthText)

        # Set default folder
        folderName = 'Faturamento'
        self.folderText.setText(folderName)
        self.folderText.setTextMargins(10, 0, 0, 0)

        # Set default sheetDetail
        # self.checkBoxDetail.setChecked(True)

        # Set button method
        self.buttonSend.clicked.connect(self.callGeneratedFiles)
        self.resultView.currentItemChanged.connect(self.printTeste)

    def callGeneratedFiles(self):

        isDetail = mainWindow.checkBoxDetail.isChecked()

        self.billing.setParamsBilling(self.yearText.text(),
                                      self.monthText.text(),
                                      self.folderText.text(), isDetail)

        self.billing.setEmployeesFile()
        self.progressBar.setRange(0, self.billing.maxEmployees)
        asyncio.run(self.billing.generatedBaseData())

        companys = asyncio.run(self.generateFiles())
        # missingCompanys = [f'Teste {item}' for item in range(100)]
        if len(companys) > 0:
            try:
                os.mkdir(f'{self.folderText.text()}')
            except FileExistsError:
                ...
            except Exception as e:
                print('Error ao criar a pasta', e)

            asyncio.run(self.getAllExams(companys))
            self.resultView.clear()
            self.resultView.addItems('Finalizado')

    def printTeste(self, teste):
        print(f'Troquei de item selecionado {teste.text()}')

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

    @Slot(int, int)
    def on_progress(self, bytesReceived: int, bytesTotal: int):
        """ Update progress bar"""
        self.progressBar.setRange(0, bytesTotal)
        self.progressBar.setValue(bytesReceived)

    async def generateFiles(self):
        companyListAux = []
        for i, companyBilling in enumerate(self.billing.companyList_Billing):
            companyListAux.append(self.billing.getCompanyList(companyBilling))
            self.on_progress(i+1, len(self.billing.companyList_Billing))
        return companyListAux

    async def getAllExams(self, companyList):
        locale.setlocale(locale.LC_ALL, 'pt_BR')
        monthText: str = self.monthText.text()
        monthNumber = datetime.strptime(monthText, '%B').month
        for i, company in enumerate(companyList):
            if len(company) > 0:
                await self.billing.createSheet(company[0], monthText.upper(),
                                               monthNumber)

                self.on_progress(i+1, len(companyList))


if __name__ == '__main__':
    app = QApplication(sys.argv)
    mainWindow = MainWindow()
    mainWindow.show()

    app.exec()
    # When exit app, save dictionary exams
    sys.exit(mainWindow.saveExams())
