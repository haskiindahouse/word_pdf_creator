import sys

from PyQt5.QtWidgets import *
from PyQt5 import QtCore
from PyQt5.QtGui import QPixmap, QIcon, QColor, QFont, QStandardItem, QStandardItemModel
from PyQt5.QtCore import Qt
from document_composer import DocumentComposer


class Ui(QWidget):

    def __init__(self, model):
        super().__init__()
        self.model = model

        self.document = DocumentComposer()

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
        self.fileOpenBtn = QPushButton('Открыть файл с заказчиками', self)
        font1 = self.fileOpenBtn.font()
        font1.setPointSize(10)
        self.fileOpenBtn.setFont(font1)
        self.fileOpenBtn.setFixedSize(200, 50)
        self.fileOpenBtn.clicked.connect(lambda: self.startToListen(flag=True))

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

        comboVLayout = QVBoxLayout(self)
        self.comboBox = QComboBox(self)
        self.comboBox.setMinimumWidth(100)
        self.comboBox.setBaseSize(100, 50)
        comboVLayout.addWidget(self.comboBox)
        self.comboBox.maximumSize()
        hBox.addWidget(self.fileOpenBtn)
        hBox.addWidget(self.addCustomerBtn)
        hBox.addWidget(self.lineEdit)
        hBox.addWidget(self.dateEdit)
        gridLayout.addLayout(hBox, 0, 0)
        verticalLayout.addLayout(gridLayout)
        verticalLayout.addLayout(comboVLayout)

        horizontalButtonLayout = QHBoxLayout(self)

        self.addProduct = QPushButton('Добавить продукт', self)
        self.addProduct.setFont(font1)
        self.addProduct.setFixedSize(150, 50)
        self.addProduct.clicked.connect(self.appendProductToModel)

        self.addCategoryToProduct = QPushButton('Назначить категорию', self)
        self.addCategoryToProduct.setFont(font1)
        self.addCategoryToProduct.setFixedSize(150, 50)
        self.addCategoryToProduct.clicked.connect(self.appendCategoryToProduct)

        self.categoriesFile = ""
        self.categoriesBtn = QPushButton('Открыть файл с категориями', self)
        self.categoriesBtn.setFont(font1)
        self.categoriesBtn.setFixedSize(200, 50)
        self.categoriesBtn.clicked.connect(lambda: self.startToListen(flag=False))

        self.categoriesComboBox = QComboBox(self)

        self.addCategory = QPushButton('Добавить категорию', self)
        self.addCategory.setFont(font1)
        self.addCategory.setEnabled(False)
        self.addCategory.setFixedSize(150, 50)
        self.addCategory.clicked.connect(self.appendCategoryToComboBox)

        self.categoriesLineEdit = QLineEdit(self)
        self.categoriesLineEdit.setText("Категория")
        self.categoriesLineEdit.setAlignment(Qt.AlignCenter)
        self.categoriesLineEdit.setFixedSize(100, 20)

        categoriesLayout = QHBoxLayout(self)
        categoriesLayout.addWidget(self.categoriesBtn)
        categoriesLayout.addWidget(self.addCategory)
        categoriesLayout.addWidget(self.categoriesLineEdit)
        categoriesLayout.addWidget(self.categoriesComboBox)
        verticalLayout.addLayout(categoriesLayout)

        self.resultBtn = QPushButton('Подсчитать ИТОГО', self)
        self.resultBtn.setFont(font1)
        self.resultBtn.setFixedSize(150, 50)
        self.resultBtn.clicked.connect(self.countResult)

        self.paymentBox = QComboBox(self)
        self.paymentBox.setMinimumWidth(120)
        self.paymentBox.addItem("Оплата б/н")
        self.paymentBox.addItem("Оплата наличными")

        horizontalButtonLayout.addWidget(self.addProduct)
        horizontalButtonLayout.addWidget(self.addCategoryToProduct)
        horizontalButtonLayout.addWidget(self.resultBtn)
        horizontalButtonLayout.addWidget(self.paymentBox)

        verticalLayout.addLayout(horizontalButtonLayout)
        verticalLayout.addWidget(self.tableView)

        btnHLayout = QHBoxLayout(self)
        self.addFileBtn = QPushButton('Добавить запись в файл', self)
        self.addFileBtn.setFont(font1)
        self.addFileBtn.setFixedSize(150, 50)
        self.addFileBtn.clicked.connect(self.addFile)

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

    def appendCategoryToProduct(self):
        """
        Назначает текущему товару в таблице категорию из categoryBox.
        Сама категория отображается только на моменте формирования моделей для документа.
        """
        currentCategory = self.categoriesLineEdit.text()
        for row in range(self.model.rowCount()):
            for column in range(self.model.columnCount()):
                self.model.item(row, column).setData(currentCategory, 1)

    def appendProductToModel(self):
        items = [QStandardItem("") for _ in range(4)]
        self.model.appendRow(items)

    def appendCategoryToComboBox(self):
        """
        # Раньше добавлялась в модель категорию для отображения в таблицу,
        # Теперь задается записи и при выгрузке в pdf/word уже добавляются нужные записи о категориях
        items = [QStandardItem("") for _ in range(4)]
        items[0].setBackground(QColor("lightGray"))
        self.model.appendRow(items)
        """
        if not self.categoriesFile or not self.categoriesLineEdit.text():
            return
        text = self.categoriesLineEdit.text()
        with open(self.categoriesFile, 'a+') as f:
            f.seek(0)
            lines = f.readlines()
            if text + '\n' in lines:
                return
            f.write(text + '\n')
            self.categoriesComboBox.addItem(text)

    def countResult(self):
        items = [QStandardItem("") for _ in range(4)]
        items[0].setText("ИТОГО")
        items[-1].setText(self.paymentBox.currentText())
        for item in items:
            item.setBackground(QColor(247, 134, 5))
        self.model.appendRow(items)

    def formFile(self):
        name, a = QFileDialog.getSaveFileName(self,
                                              'Save File',
                                              '',
                                              '.docx;;.pdf')
        if not name:
            return
        print(name)
        print(a)
        self.document.saveToFile(name, a, self.dateEdit.date().toPyDate())

    def startToListen(self, flag=True):
        if flag:
            self.file, _ = QFileDialog.getOpenFileName(self,
                                                       'Открыть файл с заказчиками',
                                                       './',
                                                       'Поставщики (*.txt)')
            fileName = self.file
            comboBox = self.comboBox
            self.addCustomerBtn.setEnabled(True)
        else:
            self.categoriesFile, _ = QFileDialog.getOpenFileName(self,
                                                       'Открыть файл с категориями',
                                                       './',
                                                       'Поставщики (*.txt)')
            fileName = self.categoriesFile
            comboBox = self.categoriesComboBox
            self.addCategory.setEnabled(True)

        if not fileName:
            return
        with open(fileName) as f:
            lines = f.readlines()
        comboBox.clear()
        for line in lines:
            comboBox.addItem(line)

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

    def addFile(self):
        """
        Отправляет модель для дальнейшей записи ее в файл.
        Чистит модель.
        :return:
        """
        self.document.appendDataToTable(self.model)
        self.model.clear()
        self.model.setHorizontalHeaderLabels(["Наименование", "К-во", "Цена", "Прим."])


def initUi():
    app = QApplication(sys.argv)

    model = QStandardItemModel()
    model.setHorizontalHeaderLabels(["Наименование", "К-во", "Цена", "Прим."])

    ex = Ui(model)

    sys.exit(app.exec_())
