import sys
import locale
import re

from pprint import pprint
from faturamento import Faturamento
from PySide6.QtWidgets import QApplication, QMainWindow
from PySide6.QtCore import Slot, QProcess, QThread
from window import Ui_MainWindow
from datetime import datetime, date, timedelta
from typing import List
import json

progress_re = re.compile("Total complete: (\d+)")
max_re = re.compile("maxEmployees (\d+)")


class Company:
    name: str
    missingExams: List[str]

    def __init__(self, name, missingExams):
        self.name = name
        self.missingExams = missingExams


def simple_percent_parser(output):
    """
    Matches lines using the progress_re regex,
    returning a single integer for the % progress.
    """
    m = progress_re.search(output)
    if m:
        pc_complete = m.group(1)
        return int(pc_complete)


def simple_max_row(output):
    """
    Matches lines using the max_re regex,
    returning a single integer.
    """
    m = max_re.search(output)
    if m:
        max = m.group(1)
        return int(max)
    else:
        return 0


def convertToClass(output) -> dict:
    subObject = output.split(';')[1]
    if subObject:
        company_exam = subObject.replace("\'", "\"")
        company_exam = json.loads(company_exam)
        return (company_exam)
    else:
        return dict()


class MainWindow(QMainWindow, Ui_MainWindow):
    billing: Faturamento
    max_row = 0
    text = ''
    p: QProcess
    companyExamNotFound: List[Company]

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

    def startWorkerBilling(self):

        # Create a Billing instance
        self.billing = Faturamento()
        # Get a DictionaryExams data from a file
        self.billing.getDictionaryExams()

        self._thread = QThread()
        isDetail = str(mainWindow.checkBoxDetail.isChecked())

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

        worker.started.connect(self.worker1Started)
        worker.progressed.connect(self.worker1Progressed)
        worker.finished.connect(self.worker1Finished)
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

    def printTeste(self, teste):
        objectPrint = f'Troquei de item selecionado {teste.text()}'
        pprint(objectPrint)

    def saveExams(self):
        self.billing.saveDictionaryExams()

    def getExams(self):
        dictionary_exams = self.billing.getDictionaryExams()
        return dictionary_exams

    def addToExamsNotFound(self, company_examString):
        company_examObject = convertToClass(company_examString)
        newObjectCompany = Company(company_examObject['name'],
                                   company_examObject['examsNotFound'])
        self.companyExamNotFound.append(newObjectCompany)

    def worker1Started(self, value):
        self.buttonSend.setDisabled(True)
        self.messageTerminal([value])
        print('worker iniciado')

    def worker1Progressed(self, value):
        self.progressBar.setValue(value)
        print('em progresso')

    def setRangeProgressBar(self, value):
        self.progressBar.setRange(0, value)
        print('Setado limite maximo', value)

    def worker1Finished(self, value):
        self.buttonSend.setDisabled(False)
        self.messageTerminal([value])
        print('worker finalizado')


if __name__ == '__main__':
    app = QApplication(sys.argv)
    mainWindow = MainWindow()
    mainWindow.show()

    app.exec()
    # When exit app, save dictionary exams
    sys.exit(mainWindow.saveExams())
