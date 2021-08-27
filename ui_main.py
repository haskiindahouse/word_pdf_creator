import sys

from PyQt5.QtWidgets import *
from PyQt5 import QtCore
from PyQt5.QtGui import QPixmap, QIcon, QColor, QFont, QStandardItem, QStandardItemModel
from PyQt5.QtCore import Qt


class Ui(QWidget):

    def __init__(self, model):
        super().__init__()
        self.model = model
        self.setTableView()
        self.initUi()

    def setTableView(self):

        self.tableView = QTableView()
        self.tableView.setModel(self.model)
        self.tableView.setWordWrap(True)
        self.tableView.setTextElideMode(Qt.ElideMiddle)
        self.tableView.resizeRowsToContents()
        self.tableView.setEditTriggers(QAbstractItemView.AllEditTriggers)
        self.tableView.setStyleSheet("QHeaderView::section { background-color: grey }")
        header = self.tableView.horizontalHeader()
        font = QFont("Arial", 20, QFont.Bold)
        header.setFont(font)
        self.tableView.setColumnWidth(0, 350)
        header.setStretchLastSection(True)

    def initUi(self):
        verticalLayout = QVBoxLayout(self)
        gridLayout = QGridLayout(self)
        hBox = QHBoxLayout(self)
        self.file = ""
        self.fileOpenBtn = QPushButton('Выбрать файл', self)
        font1 = self.fileOpenBtn.font()
        font1.setPointSize(10)
        self.fileOpenBtn.setFont(font1)
        self.fileOpenBtn.setFixedSize(200, 50)
        self.fileOpenBtn.clicked.connect(self.startToListen)

        self.lineEdit = QLineEdit(self)
        self.lineEdit.setText("Заказчик")
        self.addCustomerBtn = QPushButton('Добавить заказчика', self)
        font1 = self.addCustomerBtn.font()
        font1.setPointSize(10)
        self.addCustomerBtn.setFont(font1)
        self.addCustomerBtn.setFixedSize(200, 50)
        self.addCustomerBtn.setEnabled(False)
        self.addCustomerBtn.clicked.connect(self.writeCustomer)

        self.dateEdit = QDateEdit(self)
        self.dateEdit.setLocale(QtCore.QLocale(QtCore.QLocale.Russian, QtCore.QLocale.RussianFederation))
        font = QFont()
        font.setPointSize(12)
        self.dateEdit.setFont(font)
        self.dateEdit.setCalendarPopup(True)
        self.dateEdit.setTimeSpec(QtCore.Qt.LocalTime)
        self.dateEdit.setGeometry(QtCore.QRect(220, 31, 133, 20))

        self.comboBox = QComboBox(self)
        self.comboBox.setMinimumWidth(10)
        self.comboBox.activated[str].connect(self.onActivated)
        hBox.addWidget(self.fileOpenBtn)
        hBox.addWidget(self.addCustomerBtn)
        hBox.addWidget(self.lineEdit)
        hBox.addWidget(self.dateEdit)
        gridLayout.addLayout(hBox, 0, 0)
        verticalLayout.addLayout(gridLayout)
        verticalLayout.addWidget(self.comboBox)

        # ------
        horizontalButtonLayout = QHBoxLayout(self)
        self.addProduct = QPushButton('Добавить продукт', self)
        self.addProduct.setFont(font1)
        self.addProduct.setFixedSize(150, 50)
        self.addProduct.clicked.connect(self.appendProductToModel)

        self.addCategory = QPushButton('Добавить категорию', self)
        self.addCategory.setFont(font1)
        self.addCategory.setFixedSize(150, 50)
        self.addCategory.clicked.connect(self.appendCategoryToModel)

        self.resultBtn = QPushButton('Подсчитать ИТОГО', self)
        self.resultBtn.setFont(font1)
        self.resultBtn.setFixedSize(150, 50)
        self.resultBtn.clicked.connect(self.countResult)

        self.paymentBox = QComboBox(self)
        self.paymentBox.setMinimumWidth(120)
        self.paymentBox.addItem("Оплата б/н")
        self.paymentBox.addItem("Оплата наличными")

        horizontalButtonLayout.addWidget(self.addProduct)
        horizontalButtonLayout.addWidget(self.addCategory)
        horizontalButtonLayout.addWidget(self.resultBtn)
        horizontalButtonLayout.addWidget(self.paymentBox)

        verticalLayout.addLayout(horizontalButtonLayout)
        verticalLayout.addWidget(self.tableView)

        btnHLayout = QHBoxLayout(self)
        self.addFileBtn = QPushButton('Добавить запись в файл', self)
        self.addFileBtn.setFont(font1)
        self.addFileBtn.setFixedSize(150, 50)

        self.downloadBtn = QPushButton('Сохранить файл', self)
        self.downloadBtn.setFont(font1)
        self.downloadBtn.setFixedSize(150, 50)
        self.downloadBtn.clicked.connect(self.formFile)

        btnHLayout.addWidget(self.addFileBtn)
        btnHLayout.addWidget(self.downloadBtn)

        verticalLayout.addLayout(btnHLayout)

        self.setLayout(verticalLayout)
        self.move(300, 300)
        self.setWindowTitle('wordCreator v1.0')
        self.resize(800, 800)
        self.show()

    def appendProductToModel(self):
        items = [QStandardItem("") for i in range(4)]
        self.model.appendRow(items)

    def appendCategoryToModel(self):
        items = [QStandardItem("") for i in range(4)]
        items[0].setBackground(QColor("lightGray"))
        self.model.appendRow(items)

    def countResult(self):
        items = [QStandardItem("") for i in range(4)]
        items[0].setText("ИТОГО")
        items[-1].setText(self.paymentBox.currentText())
        for item in items:
            item.setBackground(QColor(247, 134, 5))
        self.model.appendRow(items)

    def formFile(self):
        name, _ = QFileDialog.getSaveFileName(self,
                                              'Save File',
                                              '',
                                              'Word(*.word);;Pdf(*.pdf)')
        if not name:
            return
        file = open(name, 'w')
        text = self.textEdit.toPlainText()
        file.write(text)
        file.close()

    def startToListen(self):
        self.file, _ = QFileDialog.getOpenFileName(self,
                                                   'Открыть файл',
                                                   './',
                                                   'Поставщики (*.txt)')
        if not self.file:
            return
        with open(self.file) as f:
            lines = f.readlines()
        for line in lines:
            self.comboBox.addItem(line)
        self.addCustomerBtn.setEnabled(True)

    def writeCustomer(self):
        if not self.file or not self.lineEdit.text():
            return
        text = self.lineEdit.text()
        with open(self.file, 'a+') as f:
            f.seek(0)
            lines = f.readlines()
            if text + '\n' in lines:
                return
            f.write(text + '\n')
            self.comboBox.addItem(text)

    def onActivated(self, text):
        return


def initUi():
    app = QApplication(sys.argv)

    model = QStandardItemModel()
    model.setHorizontalHeaderLabels(["Наименование", "К-во", "Цена", "Прим."])

    ex = Ui(model)

    sys.exit(app.exec_())
