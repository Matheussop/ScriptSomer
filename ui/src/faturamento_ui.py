import sys
import locale
import re

from faturamento import Faturamento
from PySide6.QtWidgets import QApplication, QMainWindow
from PySide6.QtCore import Slot, QProcess, QByteArray
from window import Ui_MainWindow
from datetime import datetime, date, timedelta


progress_re = re.compile("(\d+)")


def simple_percent_parser(output):
    """
    Matches lines using the progress_re regex,
    returning a single integer for the % progress.
    """
    m = progress_re.search(output)
    if m:
        pc_complete = m.group(1)
        return int(pc_complete)


class MainWindow(QMainWindow, Ui_MainWindow):
    billing: Faturamento
    max_row = 0
    text = ''
    p: QProcess

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

        # Progress bar
        self.p = None  # type: ignore

        # Set button method
        self.buttonSend.clicked.connect(self.start_process)
        self.resultView.currentItemChanged.connect(self.printTeste)

    def messageTerminal(self, s):
        self.resultView.addItems([s])

    def start_process(self):
        if self.p is None:  # No process running.
            # Keep a reference to the QProcess (e.g. on self)
            # while it's running.
            self.p = QProcess()
            self.p.readyReadStandardOutput.connect(self.handle_stdout)
            self.p.readyReadStandardError.connect(self.handle_stderr)
            self.p.stateChanged.connect(self.handle_state)
            # Clean up once complete.
            isDetail = str(mainWindow.checkBoxDetail.isChecked())

            self.p.finished.connect(self.process_finished)
            self.p.start("python", ['faturamento.py', self.yearText.text(),
                                    self.monthText.text(),
                                    self.folderText.text(), isDetail])

    def handle_state(self, state):
        states = {
            QProcess.NotRunning: 'Processo Finalizado',  # type: ignore
            QProcess.Starting: 'Iniciando Processo',  # type: ignore
            QProcess.Running: 'Gerando Arquivos',  # type: ignore
        }
        state_name = states[state]
        self.messageTerminal(f"Status: {state_name}")

    def handle_stdout(self):
        data = self.p.readAllStandardOutput()
        stdout = bytes(data).decode("utf8")  # type: ignore
        self.messageTerminal(stdout)

    def handle_stderr(self):
        data: QByteArray = self.p.readAllStandardError()
        stderr = bytes(data).decode("utf8")  # type: ignore
        # Extract progress if it is in the data.
        progress = simple_percent_parser(stderr)
        if isinstance(progress, int):
            index = int(simple_percent_parser(stderr))  # type: ignore
            if ('maxEmployees' in stderr):
                self.progressBar.setRange(0, index)
                self.max_row = index
            elif index:
                self.progressBar.setValue(index)

    def process_finished(self):
        self.p = None  # type: ignore

    def printTeste(self, teste):
        print(f'Troquei de item selecionado {teste.text()}')

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
    # When exit app, save dictionary exams
    sys.exit(mainWindow.saveExams())
