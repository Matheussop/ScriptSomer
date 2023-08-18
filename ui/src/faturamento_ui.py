import sys
import locale

from pprint import pprint
from styles import setupTheme
from datetime import datetime, date, timedelta
from PySide6.QtWidgets import QApplication, QMainWindow
from PySide6.QtCore import Slot, QProcess, QThread
from window import Ui_MainWindow
from faturamento import Faturamento


def convertToArray(companyList):
    aux = []
    for company in companyList:
        aux.append(f"\n{company['name']}")
        for exam in company['missingExams'].split(','):
            if exam != '':
                aux.append(f'    Exame faltando -> {exam}')
            else:
                aux.pop()
    return aux


class MainWindow(QMainWindow, Ui_MainWindow):
    billing: Faturamento
    max_row = 0
    text = ''
    p: QProcess

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setupUi(self)
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

        # Billing instance
        self.billing = None  # type: ignore

        # Reset a companyExamNotFound
        self.companyExamNotFound = []

        # Set button method
        self.buttonSend.clicked.connect(self.startWorkerBilling)
        self.resultView.currentItemChanged.connect(self.printTeste)

    def messageTerminal(self, s):
        self.resultView.addItems(s)

    @Slot()
    def startWorkerBilling(self):

        # Create a Billing instance
        self.billing = Faturamento()

        # Get a DictionaryExams data from a file
        self.billing.getDictionaryExams()

        self._thread = QThread()
        isDetail = mainWindow.checkBoxDetail.isChecked()

        self.billing.setParamsBilling(
            self.yearText.text(), self.monthText.text(),
            self.folderText.text(), isDetail)

        worker = self.billing
        thread = self._thread

        # Mover o worker para a thread
        worker.moveToThread(thread)

        # Adicionar a thread ao processo (Run)
        thread.started.connect(worker.callGeneratedFiles)

        worker.finished.connect(thread.quit)

        thread.finished.connect(thread.deleteLater)
        worker.finished.connect(worker.deleteLater)

        worker.started.connect(self.workerBillingStarted)
        worker.progressed.connect(self.workerBillingProgressed)
        worker.finished.connect(self.workerBillingFinished)
        worker.companiesNotFound.connect(self.workerBillingCompaniesNotFound)
        worker.rangeProgress.connect(self.setRangeProgressBar)

        thread.start()

    def handle_state(self, state):
        states = {
            QProcess.NotRunning: 'Processo Finalizado',  # type: ignore
            QProcess.Starting: 'Iniciando Processo',  # type: ignore
            QProcess.Running: 'Gerando Arquivos',  # type: ignore
        }
        state_name = states[state]
        self.messageTerminal([f"Status: {state_name}"])

    @Slot()
    def printTeste(self, teste):
        objectPrint = f'Troquei de item selecionado {teste.text()}'
        pprint(objectPrint)

    def saveExams(self):
        if self.billing:
            self.billing.saveDictionaryExams()

    @Slot()
    def workerBillingStarted(self, value):
        self.buttonSend.setDisabled(True)
        self.messageTerminal([value])
        print('worker iniciado')

    @Slot()
    def workerBillingProgressed(self, value):
        self.progressBar.setValue(value)
        # print('em progresso')

    @Slot()
    def setRangeProgressBar(self, value):
        self.progressBar.setRange(0, value)

    @Slot()
    def workerBillingFinished(self, value):
        self.buttonSend.setDisabled(False)
        if value != []:
            results = convertToArray(value)
            self.messageTerminal(['\nEmpresas com exames não encontrados'])
            self.messageTerminal(results)
            self.messageTerminal(['\nPrograma finalizado'])
        print('worker finalizado')

    @Slot()
    def workerBillingCompaniesNotFound(self, value):
        if value != []:
            self.messageTerminal(value)


if __name__ == '__main__':
    app = QApplication(sys.argv)
    mainWindow = MainWindow()
    setupTheme()
    mainWindow.show()

    app.exec()
    # When exit app, save dictionary exams
    sys.exit(mainWindow.saveExams())
