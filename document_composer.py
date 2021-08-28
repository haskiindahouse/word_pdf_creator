from docx import Document
from docx.shared import Inches
from PyQt5.QtGui import QStandardItemModel
import locale


class DocumentComposer:

    def __init__(self):
        super(DocumentComposer, self).__init__()
        self.document = Document()
        self.data = []

    def appendHeader(self, date):
        """
        Добавление фразы с датой из QDateEdit -> приходит к нам в формате "2" Августа 2021 г.
        :return:
        """
        locale.setlocale(locale.LC_ALL, "ru_RU.UTF-8")
        date = '"' + str(date.day) + '"' + " " + str(date.strftime('%B')).title() + " " + str(date.year) + " г."
        titleName = "Заявка на " + date
        header = self.document.add_heading(titleName, 0)
        header.alignment = 1
        header.bold = True
        header.underline = True

    def appendDataToTable(self, model):
        """
        Добавляет в общий контейнер с данными запись из модели.
        На вход приходит уже преобразованная data from Model.
        Для удобства все хранится в list.
        :return:
        """
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
        self.data.append(newTable)

    def appendTableToFile(self, table):
        """
        Добавляет таблицу в файл
        :return:
        """
        fileTable = self.document.add_table(rows=len(table), cols=len(table[-1]))
        fileTable.style = 'TableGrid'
        for row in range(len(table)):
            for column in range(len(table[-1])):
                cell = fileTable.rows[row].cells[column]
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
        TODO:
        Изменить метод.
        Все данные теперь хранятся в единой таблице.
        Заголовок перевести на объединенную ячейку.
        """
        self.appendHeader(date)
        self.writeTablesToFile()
        self.document.save(str(name) + str(fileFormat))
