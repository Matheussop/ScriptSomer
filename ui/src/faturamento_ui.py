import sys
import locale
import logging
from pprint import pprint
from datetime import datetime, date, timedelta
from PySide6.QtWidgets import QApplication, QMainWindow
from PySide6.QtCore import Slot, QProcess, QThread
from styles import setupTheme
from window import Ui_MainWindow
from faturamento import BillingDataProcessor

# Configure logging
log_file_path = "faturamento.log"
logging.basicConfig(level=logging.INFO,
                    format='%(asctime)s - %(levelname)s - %(message)s',
                    handlers=[
                        logging.FileHandler(log_file_path, mode='w'),
                        logging.StreamHandler(sys.stdout)
                    ])


def convertToArray(companyList):
    """
    Converte a lista de empresas para um array de strings formatadas.

    Args:
        companyList (list): Lista de empresas com exames não encontrados.

    Returns:
        list: Lista formatada de empresas e exames faltantes.
    """
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
    """
    Classe principal da interface gráfica.
    """
    billing: BillingDataProcessor
    max_row = 0
    text = ''
    p: QProcess
    inProcess: bool = False

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setupUi(self)
        self.setup_defaults()
        self.setup_signals()
        self.billing = None  # type: ignore

    def setup_defaults(self):
        """
        Configura os valores padrão dos campos de entrada.
        """
        # Set default year
        self.yearText.setTextMargins(10, 0, 0, 0)
        self.yearText.setText(str(datetime.now().year))

        # Set default month
        locale.setlocale(locale.LC_ALL, 'pt_BR')
        monthText = (date.today().replace(day=1) -
                     timedelta(days=1)).strftime('%B')
        self.monthText.setTextMargins(10, 0, 0, 0)
        self.monthText.setText(monthText)

        # Set default folder
        folderName = 'Faturamento'
        self.folderText.setText(folderName)
        self.folderText.setTextMargins(10, 0, 0, 0)

    def setup_signals(self):
        """
        Configura os sinais e slots da interface.
        """
        self.buttonSend.clicked.connect(self.startWorkerBilling)
        self.resultView.currentItemChanged.connect(self.printTeste)

    def keyPressEvent(self, event):
        """
        Permite iniciar o processo pressionando Enter no último campo.

        Args:
            event (QKeyEvent): Evento de tecla pressionada.
        """
        if event.key() == 16777220 and not self.inProcess:  # Enter key
            self.startWorkerBilling()

    def messageTerminal(self, messages):
        """
        Adiciona mensagens ao terminal da interface.

        Args:
            messages (list): Lista de mensagens para exibir.
        """
        if not isinstance(messages, list):
            logging.error(
                "A função messageTerminal esperava uma lista de mensagens.")
            return
        self.resultView.addItems(messages)

    @Slot()
    def startWorkerBilling(self):
        """
        Inicia o processo de faturamento em uma nova thread.
        """
        if self.inProcess:
            logging.warning("Processo de faturamento já está em execução.")
            self.messageTerminal(
                ["Processo de faturamento já está em execução."])
            return

        self.inProcess = True
        self.reset_state()

        try:
            self.billing = BillingDataProcessor()
            self.billing.getDictionaryExams()
            self._thread = QThread()

            isDetail = self.checkBoxDetail.isChecked()
            self.billing.setParamsBilling(
                self.yearText.text(), self.monthText.text(),
                self.folderText.text(), isDetail)

            worker = self.billing
            thread = self._thread

            worker.moveToThread(thread)
            thread.started.connect(worker.process_billing)
            worker.finished.connect(thread.quit)
            thread.finished.connect(thread.deleteLater)
            worker.finished.connect(worker.deleteLater)

            worker.started.connect(self.workerBillingStarted)
            worker.progressed.connect(self.workerBillingProgressed)
            worker.finished.connect(self.workerBillingFinished)
            worker.companies_not_found_signal.connect(
                self.workerBillingCompaniesNotFound)
            worker.range_progress.connect(self.setRangeProgressBar)

            thread.start()
        except Exception as e:
            logging.error(f"Erro ao iniciar o faturamento: {e}")
            self.messageTerminal([f"Erro ao iniciar o faturamento: {e}"])
            self.inProcess = False

    def reset_state(self):
        """
        Reseta o estado do processo de faturamento.
        """
        self.progressBar.setValue(0)
        self.resultView.clear()

    def handle_state(self, state):
        """
        Lida com as mudanças de estado do processo.

        Args:
            state (QProcess.ProcessState): Estado do processo.
        """
        states = {
            QProcess.NotRunning: 'Processo Finalizado',  # type: ignore
            QProcess.Starting: 'Iniciando Processo',  # type: ignore
            QProcess.Running: 'Gerando Arquivos',  # type: ignore
        }
        state_name = states[state]
        self.messageTerminal([f"Status: {state_name}"])

    @Slot()
    def printTeste(self, teste):
        """
        Imprime o item selecionado no terminal.

        Args:
            teste (QListWidgetItem): Item selecionado.
        """
        objectPrint = f'{teste.text()}'
        pprint(objectPrint)

    def saveExams(self):
        """
        Salva os exames no dicionário quando a aplicação é fechada.
        """
        # if self.billing:
        #     self.billing.saveDictionaryExams()

    @Slot()
    def workerBillingStarted(self, value):
        """
        Slot chamado quando o processo de faturamento é iniciado.

        Args:
            value (str): Mensagem de início.
        """
        if value:
            self.buttonSend.setDisabled(True)
            self.messageTerminal([value])
        logging.info('Processo de faturamento iniciado.')

    @Slot()
    def workerBillingProgressed(self, value):
        """
        Slot chamado para atualizar a barra de progresso.

        Args:
            value (int): Valor do progresso.
        """
        self.progressBar.setValue(value)

    @Slot()
    def setRangeProgressBar(self, value):
        """
        Slot chamado para definir o intervalo da barra de progresso.

        Args:
            value (int): Valor máximo do progresso.
        """
        self.progressBar.setRange(0, value)

    @Slot()
    def workerBillingFinished(self, value):
        """
        Slot chamado quando o processo de faturamento é finalizado.

        Args:
            value (list): Resultado do processo.
        """
        self.buttonSend.setDisabled(False)
        if value:
            if 'Error: ' != value[0]:
                results = convertToArray(value)
                if results:
                    self.messageTerminal(
                        ['\nEmpresas com exames não encontrados'])
                    self.messageTerminal(results)
                self.messageTerminal(['\nPrograma finalizado'])
            else:
                logging.error(f'Erro no processo de faturamento: {value}')
                self.messageTerminal([f'\n{value[1]}'])
        logging.info('Processo de faturamento finalizado.')
        self.inProcess = False

    @Slot()
    def workerBillingCompaniesNotFound(self, value):
        """
        Slot chamado quando empresas não são encontradas.

        Args:
            value (list): Lista de empresas não encontradas.
        """
        if value:
            self.messageTerminal(value)


if __name__ == '__main__':
    app = QApplication(sys.argv)
    mainWindow = MainWindow()
    mainWindow.setWindowTitle('Calculadora de Faturamento')
    setupTheme()
    mainWindow.show()

    app.exec()
    # When exiting the app, save dictionary exams
    sys.exit(mainWindow.saveExams())
