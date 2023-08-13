import sys
import locale
import asyncio

from faturamento import Faturamento
from PySide6.QtWidgets import QApplication, QMainWindow
from window import Ui_MainWindow
from datetime import datetime


class MainWindow(QMainWindow, Ui_MainWindow):

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setupUi(self)
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
        self.buttonSend.clicked.connect(self.callGeneratedFiles)

    def callGeneratedFiles(self):
        print('Chamou o callGeneratedFiles')

        isDetail = mainWindow.checkBoxDetail.isChecked()

        faturamento = Faturamento(self.yearText.text(),
                                  'Julho',
                                  self.folderText.text(), isDetail)

        missingCompanys = asyncio.run(faturamento.generatedFiles())

        self.resultView.addItems(missingCompanys)


if __name__ == '__main__':
    app = QApplication(sys.argv)
    mainWindow = MainWindow()
    mainWindow.show()

    app.exec()
