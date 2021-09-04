from docx import Document
from docx.oxml.ns import nsdecls
from docx.oxml import parse_xml
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
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
        self.currentCustomer = None

    def set_result_bg_color(self):
        """
        Задает задний фон для строчек с подсчетом "ИТОГО".
        """

        lastRow = self.table.rows[len(self.table.rows) - 1]
        self.make_rows_bold(lastRow)
        cells = lastRow.cells
        for cell in cells:
            shading_elm_1 = parse_xml(r'<w:shd {} w:fill="ffae00"/>'.format(nsdecls('w')))
            cell._tc.get_or_add_tcPr().append(shading_elm_1)

    def set_header_bg_color(self):
        """
        Задает задний фон для строчек с подсчетом "ИТОГО".
        """

        lastRow = self.table.rows[len(self.table.rows) - 1]
        self.make_rows_bold(lastRow)
        cells = lastRow.cells
        for cell in cells:
            shading_elm_1 = parse_xml(r'<w:shd {} w:fill="878787"/>'.format(nsdecls('w')))
            cell._tc.get_or_add_tcPr().append(shading_elm_1)

    @staticmethod
    def make_rows_bold(*rows):
        """
        Делает строчки жирными.
        """
        for row in rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    for run in paragraph.runs:
                        run.font.bold = True

    @staticmethod
    def make_rows_underline(cells):
        """
        Делает строчки подчеркнутыми.
        """
        for cell in cells:
            for paragraph in cell.paragraphs:
                for run in paragraph.runs:
                    run.font.underline = True

    def appendHeader(self, date):
        """
        Добавление фразы с датой из QDateEdit -> приходит к нам в формате "2" Августа 2021 г.
        """
        locale.setlocale(locale.LC_ALL, "ru_RU.UTF-8")
        date = '"' + str(date.day) + '"' + " " + str(date.strftime('%B')).title() + " " + str(date.year) + " г."
        titleName = "Заявка на " + date
        cell = self.table.rows[0].cells[0]
        cell.text = titleName
        cell.bold = True
        cell.underline = True
        cell.merge(self.table.rows[0].cells[3])
        cell.paragraphs[0].alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
        self.make_rows_bold(self.table.rows[0])
        self.make_rows_underline(self.table.rows[0].cells)

    def appendCustomer(self, customer):
        """
        Добавление нового заказчика в таблицу если такой имеется.
        """
        if customer is not None and self.currentCustomer is None:
            self.currentCustomer = customer
        elif customer == self.currentCustomer:
            return
        else:
            self.currentCustomer = customer
        self.table.add_row()
        cell = self.table.rows[len(self.table.rows) - 1].cells[0]
        cell.text = self.currentCustomer
        cell.bold = True
        cell.underline = True
        self.make_rows_bold(self.table.rows[len(self.table.rows) - 1])
        cell.merge(self.table.rows[len(self.table.rows) - 1].cells[3])

    def appendDataToTable(self, model, customer):
        """
        Добавляет в общий контейнер с данными запись из модели.
        Для удобства все хранится в list.
        """
        newTable = []
        headers = []

        if customer is None:
            return

        self.appendCustomer(customer)

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
        Добавляет таблицу в файл.
        """
        for row in range(len(table)):
            self.table.add_row()
            if row == 0:
                self.set_header_bg_color()
            for column in range(len(table[-1])):
                cell = self.table.rows[len(self.table.rows) - 1].cells[column]
                cell.text = table[row][column]

        self.set_result_bg_color()

    def appendEndTable(self):
        """
        Добавляет в конец документа таблицу согласно требованиям заказчика.
        """
        table = self.document.add_table(rows=6, cols=2)
        table.style = 'TableGrid'
        data = (
            ('расходы', 'зачисления'),
            ('Получено с кассы', ''),
            ('Оплачено ООО', ''),
            ('ГСМ/ком-ные', ''),
            ('Итого расход', ''),
            ('Сдано в кассу', '')
        )
        for row in range(6):
            cells = table.rows[row].cells
            for col in range(2):
                cells[col].text = data[row][col]
        self.make_rows_bold(table.rows[0])

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
        self.appendEndTable()
        self.document.save(str(name) + str(fileFormat))
        self.data = []
