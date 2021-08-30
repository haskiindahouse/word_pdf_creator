from docx import Document
from docx.shared import Inches
from PyQt5.QtGui import QStandardItemModel
import locale


class DocumentComposer:

    def __init__(self):
        super(DocumentComposer, self).__init__()
        self.document = Document()
        self.data = []
        self.table = self.document.add_table(rows=1, cols=4)
        self.table.style = 'TableGrid'

    def appendHeader(self, date):
        """
        Добавление фразы с датой из QDateEdit -> приходит к нам в формате "2" Августа 2021 г.
        :return:
        """
        locale.setlocale(locale.LC_ALL, "ru_RU.UTF-8")
        date = '"' + str(date.day) + '"' + " " + str(date.strftime('%B')).title() + " " + str(date.year) + " г."
        titleName = "Заявка на " + date
        cell = self.table.rows[0].cells[0]
        cell.text = titleName
        cell.bold = True
        cell.underline = True
        cell.merge(self.table.rows[0].cells[3])

    def appendDataToTable(self, model):
        """
        Добавляет в общий контейнер с данными запись из модели.
        На вход приходит уже преобразованная data from Model.
        Для удобства все хранится в list.
        """
        categories = []
        for row in range(model.rowCount()):
            categories.append(model.item(row, 0).data(1))
        print(categories)
        newTable = []
        headers = []
        for column in range(model.columnCount()):
            headers.append(model.horizontalHeaderItem(column).text())
        newTable.append(headers)
        for row in range(model.rowCount()):
            newTable.append([])
            for column in range(model.columnCount()):
                index = model.index(row, column)
                newTable[row].append(str(model.data(index)))
        newTable = newTable[:len(newTable) - 1]
        self.data.append(newTable)

    def appendTableToFile(self, table):
        """
        Добавляет таблицу в файл
        """
        for row in range(len(table)):
            self.table.add_row()
            for column in range(len(table[-1])):
                cell = self.table.rows[len(self.table.rows) - 1].cells[column]
                cell.text = table[row][column]

    def writeTablesToFile(self):
        """
        Для удобства записи большого количества таблиц в файл.
        """
        for table in self.data:
            self.appendTableToFile(table)

    def saveToFile(self, name, fileFormat, date):
        """
        Общий метод сохранения.
        Содержит в себе вызовы всех вспомогательных методов класса.
        """
        self.appendHeader(date)
        self.writeTablesToFile()
        self.document.save(str(name) + str(fileFormat))
        self.data = []
