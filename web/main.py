from aiohttp import web
import openpyxl
import xlrd
import os
import io
from jinja2 import Template
import tempfile
from docx import Document
import pdfplumber


async def handle(request):
    # Отображаем главную страницу
    return web.Response(
        text=render_template("upload.html", table=""), content_type="text/html"
    )


async def post_handle(request):
    reader = await request.post()
    file = reader.get("file")

    if file:
        file_bytes = io.BytesIO(file.file.read())
        file_name = file.filename

        rows = ""

        if file_name.endswith(".xlsx"):
            # Чтение файла Excel
            wb_x: openpyxl.Workbook = openpyxl.load_workbook(file_bytes, data_only=True)
            sheet_x = wb_x.active

            if not sheet_x:
                return web.Response(
                    text=render_template("upload.html", table=""),
                    content_type="text/html",
                )

            # Генерация строк для HTML-таблицы
            rows = "".join(
                "<tr>" + "".join(f"<td>{cell}</td>" for cell in row) + "</tr>"
                for row in sheet_x.iter_rows(values_only=True)
            )
        elif file_name.endswith(".xls"):
            with tempfile.NamedTemporaryFile(delete=False, suffix=".xls") as temp_file:
                temp_file.write(file_bytes.getbuffer())
                temp_file_path = temp_file.name

                # Чтение файла Excel формата .xls
                wb: xlrd.book.Book = xlrd.open_workbook(temp_file_path)
                sheet: xlrd.sheet.Sheet = wb.sheet_by_index(0)

                rows = "".join(
                    "<tr>"
                    + "".join(
                        f"<td>{sheet.cell_value(r, c)}</td>" for c in range(sheet.ncols)
                    )
                    + "</tr>"
                    for r in range(sheet.nrows)
                )
        elif file_name.endswith(".docx"):
            # Чтение таблицы из файла формата .docx
            doc = Document(file_bytes)
            for table in doc.tables[:1]:
                for row in table.rows:
                    rows += (
                        "<tr>"
                        + "".join(f"<td>{cell.text}</td>" for cell in row.cells)
                        + "</tr>"
                    )
        elif file_name.endswith(".pdf"):
            with tempfile.NamedTemporaryFile(delete=False, suffix=".xls") as temp_file:
                temp_file.write(file_bytes.getbuffer())
                temp_file_path = temp_file.name

                # Чтение таблицы из файла формата .pdf
                with pdfplumber.open(temp_file_path) as pdf:
                    tables = []
                    for page in pdf.pages:
                        t = page.extract_table()
                        if t:
                            tables.append(t)

                    for table in tables[:1]:
                        print(table)
                        for row in table:
                            if row:
                                rows += (
                                    "<tr>"
                                    + "".join(f"<td>{cell}</td>" for cell in row)
                                    + "</tr>"
                                )

        return web.Response(
            text=render_template("table.html", rows=rows), content_type="text/html"
        )

    return web.Response(
        text=render_template("upload.html", table=""), content_type="text/html"
    )


def render_template(template_name, **context):
    template_path = os.path.join("templates", template_name)

    with open(template_path) as f:
        template = Template(f.read())

    return template.render(**context)


app = web.Application(client_max_size=10 * 1024 * 1024)
app.router.add_get("/", handle)
app.router.add_post("/", post_handle)

if __name__ == "__main__":
    web.run_app(app, port=8080)
