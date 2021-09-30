import sys
import os

from PyQt5.QtWidgets import *
from PyQt5 import QtCore
from PyQt5.QtGui import QPixmap, QIcon, QColor, QFont, QStandardItem, QStandardItemModel
from PyQt5.QtCore import Qt, QDate
from PyQt5.QtCore import QTextCodec
from document_composer import DocumentComposer


class Ui(QWidget):

    def __init__(self, model):
        super().__init__()
        QTextCodec.setCodecForLocale(QTextCodec.codecForName("Windows-1251"))
        self.model = model
        self.document = DocumentComposer()

        self.tableView = QTableView()
        self.setTableView()
        self.initUi()
        self.setStartFiles()

    def setTableView(self):
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
        self.fileOpenBtn = QPushButton('Открыть файл с поставщиками', self)
        font1 = self.fileOpenBtn.font()
        font1.setPointSize(10)
        self.fileOpenBtn.setFont(font1)
        self.fileOpenBtn.clicked.connect(lambda: self.startToListen(flag=True))

        self.lineEdit = QLineEdit(self)
        self.lineEdit.setText("Поставщик")

        self.addCustomerBtn = QPushButton('Добавить поставщика', self)
        font1 = self.addCustomerBtn.font()
        font1.setPointSize(12)
        self.addCustomerBtn.setFont(font1)
        self.addCustomerBtn.setEnabled(False)
        self.addCustomerBtn.clicked.connect(self.writeCustomer)

        self.dateEdit = QDateEdit(self)
        self.dateEdit.setLocale(QtCore.QLocale(QtCore.QLocale.Russian, QtCore.QLocale.RussianFederation))
        font = QFont()
        font.setPointSize(12)
        self.dateEdit.setFont(font)
        self.dateEdit.setCalendarPopup(True)
        self.dateEdit.setTimeSpec(QtCore.Qt.LocalTime)
        self.dateEdit.setTimeSpec(QtCore.Qt.LocalTime)
        self.dateEdit.setGeometry(QtCore.QRect(220, 31, 133, 20))
        self.dateEdit.setDate(QDate.currentDate())

        self.spanBtn = QPushButton("Объединить ячейки", self)
        font1 = self.spanBtn.font()
        font1.setPointSize(12)
        self.spanBtn.setFont(font1)
        self.spanBtn.setEnabled(True)
        self.spanBtn.clicked.connect(self.spanRow)

        comboVLayout = QVBoxLayout(self)
        self.comboBox = QComboBox(self)

        comboVLayout.addWidget(self.comboBox)
        hBox.addWidget(self.spanBtn)
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

        self.addProduct.clicked.connect(self.appendProductToModel)

        self.addCategoryToProduct = QPushButton('Назначить категорию', self)
        self.addCategoryToProduct.setFont(font1)

        self.addCategoryToProduct.clicked.connect(self.appendCategoryToProduct)

        self.categoriesFile = ""
        self.categoriesBtn = QPushButton('Открыть файл с категориями', self)

        self.categoriesBtn.setFont(font1)
        self.categoriesBtn.clicked.connect(lambda: self.startToListen(flag=False))

        self.categoriesComboBox = QComboBox(self)

        self.addCategory = QPushButton('Добавить категорию', self)
        self.addCategory.setFont(font1)
        self.addCategory.setEnabled(False)
        self.addCategory.clicked.connect(self.appendCategoryToComboBox)

        self.categoriesLineEdit = QLineEdit(self)
        self.categoriesLineEdit.setText("Категория")
        self.categoriesLineEdit.setAlignment(Qt.AlignCenter)

        categoriesLayout = QHBoxLayout(self)
        categoriesLayout.setAlignment(Qt.AlignLeft)
        categoriesLayout.addWidget(self.categoriesBtn)
        categoriesLayout.addWidget(self.addCategory)
        categoriesLayout.addWidget(self.categoriesLineEdit)
        categoriesLayout.addWidget(self.categoriesComboBox)
        verticalLayout.addLayout(categoriesLayout)

        self.resultBtn = QPushButton('Подсчитать ИТОГО', self)
        self.resultBtn.setFont(font1)
        self.resultBtn.clicked.connect(self.countResult)

        self.paymentBox = QComboBox(self)
        self.paymentBox.addItem("Оплата б/н")
        self.paymentBox.addItem("Оплата наличными")

        horizontalButtonLayout.addWidget(self.addProduct)
        horizontalButtonLayout.addWidget(self.addCategoryToProduct)
        horizontalButtonLayout.addWidget(self.resultBtn)
        horizontalButtonLayout.addWidget(self.paymentBox)

        verticalLayout.addLayout(horizontalButtonLayout)
        verticalLayout.addWidget(self.tableView)

        btnHLayout = QHBoxLayout(self)

        self.deleteBtn = QPushButton('Удалить', self)
        self.deleteBtn.setFont(font1)
        self.deleteBtn.clicked.connect(self.deleteFromModel)

        self.addFileBtn = QPushButton('Добавить запись в файл', self)
        self.addFileBtn.setFont(font1)
        self.addFileBtn.clicked.connect(self.addFile)

        self.downloadBtn = QPushButton('Сохранить файл', self)
        self.downloadBtn.setFont(font1)
        self.downloadBtn.clicked.connect(self.formFile)

        btnHLayout.addWidget(self.deleteBtn)
        btnHLayout.addWidget(self.addFileBtn)
        btnHLayout.addWidget(self.downloadBtn)

        verticalLayout.addLayout(btnHLayout)

        self.setLayout(verticalLayout)
        self.move(300, 300)
        self.setWindowTitle('wordCreator v1.4.2')
        self.resize(1200, 1200)
        self.show()

    def deleteFromModel(self):
        """
        Удаляет выбранные строки из модели.
        """
        selectedRows = self.tableView.selectionModel().selectedRows()
        if not selectedRows:
            return

        self.model.removeRows(selectedRows[0].row(), len(selectedRows))

    def spanRow(self):
        """
        Объединяет строки в моделе.
        """
        selectedRows = self.tableView.selectionModel().selectedRows()

        if not selectedRows or len(selectedRows) == 1:
            return

        delta = selectedRows[-1].row() - selectedRows[0].row()
        self.tableView.setSpan(selectedRows[0].row(), 3, delta + 1, 1)
        index = self.model.index(selectedRows[0].row(), 3)
        self.model.itemFromIndex(index).setData(delta + 1, 5)

    def appendCategoryToProduct(self):
        """
        Назначает текущему товару в таблице категорию из categoryBox.
        Сама категория отображается только на моменте формирования моделей для документа.
        """
        selectedRows = self.tableView.selectionModel().selectedRows()
        if not selectedRows:
            return
        currentCategory = self.categoriesComboBox.currentText()
        currentCategory = str(currentCategory).strip()
        items = [QStandardItem("") for _ in range(4)]
        items[0] = QStandardItem(currentCategory)
        for item in items:
            item.setBackground(QColor(192, 192, 192))
        items[0].setData(4, 4)
        self.model.insertRow(selectedRows[0].row(), items)

    def appendProductToModel(self):
        items = [QStandardItem("") for _ in range(4)]
        self.model.appendRow(items)

    def appendCategoryToComboBox(self):
        """
        Раньше добавлялась в модель категорию для отображения в таблицу,
        Теперь задается записи и при выгрузке в pdf/word уже добавляются нужные записи о категориях.
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

        # Подчет количества товара в текущей группе
        currentCount = 0
        for row in range(self.model.rowCount()):
            index = self.model.index(row, 1)
            value = str(self.model.data(index))
            if not value.isnumeric():
                continue
            currentCount += int(value)

        # Подчет суммы денег в текущей группе
        currentSum = 0
        for row in range(self.model.rowCount()):
            index = self.model.index(row, 2)
            value = str(self.model.data(index))
            if not value.isnumeric():
                continue
            currentSum += int(value)

        items[0].setText("ИТОГО")
        items[1].setText(str(currentCount))
        items[2].setText((str(currentSum)))
        items[3].setText(self.paymentBox.currentText())
        for item in items:
            item.setBackground(QColor(247, 134, 5))
        self.model.appendRow(items)

    def setStartFiles(self):
        """
        Запоминает путь к двум файлам для дальнейшей прогрузки.
        """
        fileName = 'autoStart.txt'
        m = 'w+'
        if os.path.exists(fileName):
            m = 'r+'
        with open(fileName, m) as f:
            lines = f.read().splitlines()

        lines = [row for row in lines if row != '']
        if len(lines) < 2:
            return
        self.file = lines[0].strip()
        self.categoriesFile = lines[1].strip()

        self.addCustomerBtn.setEnabled(True)
        self.addCategory.setEnabled(True)

        with open(self.file) as f:
            lines = f.readlines()
        self.comboBox.clear()

        for line in lines:
            self.comboBox.addItem(line)

        with open(self.categoriesFile) as f:
            lines = f.readlines()
        self.categoriesComboBox.clear()

        for line in lines:
            self.categoriesComboBox.addItem(line)


    def formFile(self):
        name, a = QFileDialog.getSaveFileName(self,
                                              'Save File',
                                              '',
                                              '.docx;;.pdf')
        if not name:
            return
        self.document.saveToFile(name, a, self.dateEdit.date().toPyDate())

    def startToListen(self, flag=True):
        if flag:
            self.file, _ = QFileDialog.getOpenFileName(self,
                                                       'Открыть файл с поставщиками',
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
        if self.file is not None and self.categoriesFile is not None:
            with open('autoStart.txt', 'w') as f:
                f.seek(0)
                f.write(self.file + '\n')
                f.write(self.categoriesFile + '\n')

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
        """
        if not self.model.rowCount():
            return
        self.document.appendDataToTable(self.model, self.comboBox.currentText())
        self.model.clear()
        self.model.setHorizontalHeaderLabels(["Наименование", "К-во", "Цена", "Прим."])
        self.setTableView()


def initUi():
    app = QApplication(sys.argv)

    model = QStandardItemModel()
    model.setHorizontalHeaderLabels(["Наименование", "К-во", "Цена", "Прим."])

    ex = Ui(model)

    sys.exit(app.exec_())
