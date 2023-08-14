# -*- coding: utf-8 -*-

################################################################################
# Form generated from reading UI file 'main_window_billing.ui'
##
# Created by: Qt User Interface Compiler version 6.5.2
##
# WARNING! All changes made in this file will be lost when recompiling UI file!
################################################################################

from PySide6.QtCore import (QCoreApplication, QDate, QDateTime, QLocale,
                            QMetaObject, QObject, QPoint, QRect,
                            QSize, QTime, QUrl, Qt)
from PySide6.QtGui import (QBrush, QColor, QConicalGradient, QCursor,
                           QFont, QFontDatabase, QGradient, QIcon,
                           QImage, QKeySequence, QLinearGradient, QPainter,
                           QPalette, QPixmap, QRadialGradient, QTransform)
from PySide6.QtWidgets import (QApplication, QCheckBox, QFrame, QGridLayout,
                               QHBoxLayout, QLabel, QLineEdit, QListWidget,
                               QListWidgetItem, QMainWindow, QMenuBar, QProgressBar,
                               QPushButton, QScrollArea, QSizePolicy, QSpacerItem,
                               QStatusBar, QVBoxLayout, QWidget)


class Ui_MainWindow(object):
    def setupUi(self, MainWindow):
        if not MainWindow.objectName():
            MainWindow.setObjectName(u"MainWindow")
        MainWindow.resize(800, 610)
        self.centralwidget = QWidget(MainWindow)
        self.centralwidget.setObjectName(u"centralwidget")
        self.verticalLayout_2 = QVBoxLayout(self.centralwidget)
        self.verticalLayout_2.setObjectName(u"verticalLayout_2")
        self.gridLayout = QGridLayout()
        self.gridLayout.setObjectName(u"gridLayout")
        self.detailLabel = QLabel(self.centralwidget)
        self.detailLabel.setObjectName(u"detailLabel")
        sizePolicy = QSizePolicy(QSizePolicy.Preferred, QSizePolicy.Minimum)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(
            self.detailLabel.sizePolicy().hasHeightForWidth())
        self.detailLabel.setSizePolicy(sizePolicy)
        self.detailLabel.setMinimumSize(QSize(0, 30))
        self.detailLabel.setMaximumSize(QSize(16777215, 40))
        self.detailLabel.setMargin(5)

        self.gridLayout.addWidget(self.detailLabel, 5, 0, 1, 1)

        self.scrollArea = QScrollArea(self.centralwidget)
        self.scrollArea.setObjectName(u"scrollArea")
        sizePolicy1 = QSizePolicy(QSizePolicy.Preferred, QSizePolicy.Preferred)
        sizePolicy1.setHorizontalStretch(0)
        sizePolicy1.setVerticalStretch(0)
        sizePolicy1.setHeightForWidth(
            self.scrollArea.sizePolicy().hasHeightForWidth())
        self.scrollArea.setSizePolicy(sizePolicy1)
        self.scrollArea.setMinimumSize(QSize(200, 150))
        self.scrollArea.setMaximumSize(QSize(16777215, 200))
        self.scrollArea.setWidgetResizable(True)
        self.scrollArea.setAlignment(Qt.AlignCenter)
        self.scrollAreaWidgetContents = QWidget()
        self.scrollAreaWidgetContents.setObjectName(
            u"scrollAreaWidgetContents")
        self.scrollAreaWidgetContents.setGeometry(QRect(0, 0, 772, 172))
        self.horizontalLayout = QHBoxLayout(self.scrollAreaWidgetContents)
        self.horizontalLayout.setObjectName(u"horizontalLayout")
        self.resultView = QListWidget(self.scrollAreaWidgetContents)
        self.resultView.setObjectName(u"resultView")
        sizePolicy1.setHeightForWidth(
            self.resultView.sizePolicy().hasHeightForWidth())
        self.resultView.setSizePolicy(sizePolicy1)
        self.resultView.setMinimumSize(QSize(200, 0))

        self.horizontalLayout.addWidget(self.resultView)

        self.scrollArea.setWidget(self.scrollAreaWidgetContents)

        self.gridLayout.addWidget(self.scrollArea, 10, 0, 1, 2)

        self.monthLabel = QLabel(self.centralwidget)
        self.monthLabel.setObjectName(u"monthLabel")
        sizePolicy.setHeightForWidth(
            self.monthLabel.sizePolicy().hasHeightForWidth())
        self.monthLabel.setSizePolicy(sizePolicy)
        self.monthLabel.setMinimumSize(QSize(0, 30))
        self.monthLabel.setMaximumSize(QSize(16777215, 40))
        self.monthLabel.setMargin(5)

        self.gridLayout.addWidget(self.monthLabel, 3, 0, 1, 1)

        self.verticalSpacer = QSpacerItem(
            20, 10, QSizePolicy.Minimum, QSizePolicy.Expanding)

        self.gridLayout.addItem(self.verticalSpacer, 8, 0, 1, 2)

        self.monthText = QLineEdit(self.centralwidget)
        self.monthText.setObjectName(u"monthText")
        self.monthText.setMinimumSize(QSize(0, 30))

        self.gridLayout.addWidget(self.monthText, 4, 0, 1, 1)

        self.intro = QLabel(self.centralwidget)
        self.intro.setObjectName(u"intro")
        sizePolicy1.setHeightForWidth(
            self.intro.sizePolicy().hasHeightForWidth())
        self.intro.setSizePolicy(sizePolicy1)
        self.intro.setMinimumSize(QSize(100, 100))
        self.intro.setMaximumSize(QSize(16777215, 100))
        self.intro.setAlignment(Qt.AlignCenter)

        self.gridLayout.addWidget(self.intro, 0, 0, 1, 2)

        self.horizontalLayout_3 = QHBoxLayout()
        self.horizontalLayout_3.setObjectName(u"horizontalLayout_3")
        self.horizontalLayout_3.setContentsMargins(10, -1, -1, -1)
        self.checkBoxDetail = QCheckBox(self.centralwidget)
        self.checkBoxDetail.setObjectName(u"checkBoxDetail")
        self.checkBoxDetail.setLayoutDirection(Qt.LeftToRight)

        self.horizontalLayout_3.addWidget(self.checkBoxDetail)

        self.gridLayout.addLayout(self.horizontalLayout_3, 6, 0, 1, 1)

        self.folderText = QLineEdit(self.centralwidget)
        self.folderText.setObjectName(u"folderText")
        self.folderText.setMinimumSize(QSize(0, 30))

        self.gridLayout.addWidget(self.folderText, 6, 1, 1, 1)

        self.horizontalLayout_2 = QHBoxLayout()
        self.horizontalLayout_2.setObjectName(u"horizontalLayout_2")
        self.buttonSend = QPushButton(self.centralwidget)
        self.buttonSend.setObjectName(u"buttonSend")
        sizePolicy2 = QSizePolicy(QSizePolicy.Minimum, QSizePolicy.Fixed)
        sizePolicy2.setHorizontalStretch(0)
        sizePolicy2.setVerticalStretch(0)
        sizePolicy2.setHeightForWidth(
            self.buttonSend.sizePolicy().hasHeightForWidth())
        self.buttonSend.setSizePolicy(sizePolicy2)
        self.buttonSend.setMinimumSize(QSize(50, 0))
        self.buttonSend.setMaximumSize(QSize(200, 16777215))
        self.buttonSend.setLayoutDirection(Qt.LeftToRight)
        self.buttonSend.setAutoDefault(False)

        self.horizontalLayout_2.addWidget(self.buttonSend)

        self.gridLayout.addLayout(self.horizontalLayout_2, 11, 0, 1, 2)

        self.line = QFrame(self.centralwidget)
        self.line.setObjectName(u"line")
        sizePolicy3 = QSizePolicy(QSizePolicy.Preferred, QSizePolicy.Fixed)
        sizePolicy3.setHorizontalStretch(0)
        sizePolicy3.setVerticalStretch(0)
        sizePolicy3.setHeightForWidth(
            self.line.sizePolicy().hasHeightForWidth())
        self.line.setSizePolicy(sizePolicy3)
        self.line.setFrameShape(QFrame.HLine)
        self.line.setFrameShadow(QFrame.Sunken)

        self.gridLayout.addWidget(self.line, 7, 0, 1, 2)

        self.folderLabel = QLabel(self.centralwidget)
        self.folderLabel.setObjectName(u"folderLabel")
        self.folderLabel.setMinimumSize(QSize(0, 30))
        self.folderLabel.setMaximumSize(QSize(16777215, 40))
        self.folderLabel.setMargin(5)

        self.gridLayout.addWidget(self.folderLabel, 5, 1, 1, 1)

        self.yearText = QLineEdit(self.centralwidget)
        self.yearText.setObjectName(u"yearText")
        self.yearText.setMinimumSize(QSize(0, 30))

        self.gridLayout.addWidget(self.yearText, 4, 1, 1, 1)

        self.yearLabel = QLabel(self.centralwidget)
        self.yearLabel.setObjectName(u"yearLabel")
        self.yearLabel.setMinimumSize(QSize(0, 30))
        self.yearLabel.setMaximumSize(QSize(16777215, 40))
        self.yearLabel.setMargin(5)

        self.gridLayout.addWidget(self.yearLabel, 3, 1, 1, 1)

        self.progressBar = QProgressBar(self.centralwidget)
        self.progressBar.setObjectName(u"progressBar")
        self.progressBar.setValue(0)

        self.gridLayout.addWidget(self.progressBar, 9, 0, 1, 2)

        self.verticalLayout_2.addLayout(self.gridLayout)

        MainWindow.setCentralWidget(self.centralwidget)
        self.menubar = QMenuBar(MainWindow)
        self.menubar.setObjectName(u"menubar")
        self.menubar.setGeometry(QRect(0, 0, 800, 24))
        MainWindow.setMenuBar(self.menubar)
        self.statusbar = QStatusBar(MainWindow)
        self.statusbar.setObjectName(u"statusbar")
        MainWindow.setStatusBar(self.statusbar)
        QWidget.setTabOrder(self.monthText, self.yearText)
        QWidget.setTabOrder(self.yearText, self.checkBoxDetail)
        QWidget.setTabOrder(self.checkBoxDetail, self.folderText)
        QWidget.setTabOrder(self.folderText, self.resultView)
        QWidget.setTabOrder(self.resultView, self.scrollArea)
        QWidget.setTabOrder(self.scrollArea, self.buttonSend)

        self.retranslateUi(MainWindow)

        QMetaObject.connectSlotsByName(MainWindow)
    # setupUi

    def retranslateUi(self, MainWindow):
        MainWindow.setWindowTitle(QCoreApplication.translate(
            "MainWindow", u"MainWindow", None))
        self.detailLabel.setText(QCoreApplication.translate(
            "MainWindow", u"As planilhas devem ser detalhadas ? ", None))
        self.monthLabel.setText(QCoreApplication.translate(
            "MainWindow", u"Sobre qual m\u00eas deseja gerar ? ", None))
        self.intro.setText(QCoreApplication.translate(
            "MainWindow", u"Gerar Arquivos de Faturamento", None))
        self.checkBoxDetail.setText(
            QCoreApplication.translate("MainWindow", u"Sim", None))
        self.buttonSend.setText(
            QCoreApplication.translate("MainWindow", u"Gerar", None))
        self.folderLabel.setText(QCoreApplication.translate(
            "MainWindow", u"Qual o nome da pasta que conter\u00e1 todas os excels gerados ?", None))
        self.yearLabel.setText(QCoreApplication.translate(
            "MainWindow", u"De qual ano? ", None))
    # retranslateUi
