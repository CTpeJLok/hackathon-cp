import tempfile
import io

import openpyxl
import xlrd
from docx import Document
import pdfplumber


class File:
    def __init__(self, file_name: str, file_bytes: bytes):
        self.file_name = file_name
        self.file_bytes = io.BytesIO(file_bytes)

    def get_table(self):
        return self.file_bytes


class XLSXFile(File):
    def __init__(self, file_name: str, file_bytes: bytes):
        super().__init__(file_name, file_bytes)

    def get_table(self):
        # Чтение файла Excel
        wb_x: openpyxl.Workbook = openpyxl.load_workbook(
            self.file_bytes, data_only=True
        )
        sheet_x = wb_x.active
        if not sheet_x:
            return []

        table = []
        for row in sheet_x.iter_rows(values_only=True):
            table.append([])
            for cell in row:
                table[-1].append(cell)

        return table


class XLSFile(File):
    def __init__(self, file_name: str, file_bytes: bytes):
        super().__init__(file_name, file_bytes)

    def get_table(self):
        with tempfile.NamedTemporaryFile(delete=False, suffix=".xls") as temp_file:
            temp_file.write(self.file_bytes.getbuffer())
            temp_file_path = temp_file.name

            # Чтение файла Excel формата .xls
            wb: xlrd.book.Book = xlrd.open_workbook(temp_file_path)
            sheet: xlrd.sheet.Sheet = wb.sheet_by_index(0)

            table = []
            for r in range(sheet.nrows):
                table.append([])

                for c in range(sheet.ncols):
                    table[-1].append(sheet.cell_value(r, c))

            return table


class DOCXFile(File):
    def __init__(self, file_name: str, file_bytes: bytes):
        super().__init__(file_name, file_bytes)

    def get_table(self):
        # Чтение таблицы из файла формата .docx
        doc = Document(self.file_bytes)

        table = []

        doc_table = doc.tables[:1]
        if not doc_table:
            return []

        doc_table = doc.tables[0]
        for r in doc_table.rows:
            table.append([])
            for c in r.cells:
                table[-1].append(c.text)
            table[-1].append(r.cells[0].text)

        return table


class PDFFile(File):
    def __init__(self, file_name: str, file_bytes: bytes):
        super().__init__(file_name, file_bytes)

    def get_table(self):
        with tempfile.NamedTemporaryFile(delete=False, suffix=".xls") as temp_file:
            temp_file.write(self.file_bytes.getbuffer())
            temp_file_path = temp_file.name

            # Чтение таблицы из файла формата .pdf
            with pdfplumber.open(temp_file_path) as pdf:
                t = []
                for page in pdf.pages:
                    all_tables = page.extract_table()
                    if all_tables:
                        t.append(all_tables)

                tables = []

                for t in t:
                    for r in t:
                        if r:
                            tables.append([])
                            for cell in r:
                                tables[-1].append(cell)

                return tables
