from aiohttp import web
import os
from jinja2 import Template

from tensorflow.keras.models import load_model
import numpy as np
import pandas as pd

from File import File, XLSFile, XLSXFile, DOCXFile, PDFFile

model = load_model("/home/andrey/Документы/code/hackathon-cp/web/models/month_1.h5")

EXTS = {
    "xlsx": XLSXFile,
    "xls": XLSFile,
    "docx": DOCXFile,
    "pdf": PDFFile,
}


def preprocess(volumes: pd.DataFrame, prediction_period=1):
    # 1. Загрузка и предобработка данных
    volumes = volumes.fillna("")
    new_headers = volumes.iloc[0].astype(str) + " " + volumes.iloc[1].astype(str)
    volumes.columns = new_headers
    volumes = volumes.drop([0, 1]).reset_index(drop=True)
    volumes.columns = volumes.columns.str.strip()

    summed_volumes = volumes.groupby("ID").sum().iloc[:, 4:].reset_index()
    freight_columns = [
        col for col in summed_volumes.columns if "Провозная плата" in col
    ]
    tonnage_columns = [
        col for col in summed_volumes.columns if "Объем перевозок" in col
    ]

    sorted_freight_columns = sorted(
        freight_columns, key=lambda x: pd.to_datetime(x.split()[0], format="%Y/%m")
    )
    sorted_tonnage_columns = sorted(
        tonnage_columns, key=lambda x: pd.to_datetime(x.split()[0], format="%Y/%m")
    )
    sorted_columns = ["ID"] + sorted_freight_columns + sorted_tonnage_columns
    volumes_sorted = summed_volumes[sorted_columns]

    # 2. Создание признака 'Отток' (при отсутствии активности 12 месяцев)
    volumes_sorted["Отток"] = (
        volumes_sorted.filter(like="Провозная плата")
        .T.rolling(window=12)
        .sum()
        .T.min(axis=1)
        == 0
    ) | (
        volumes_sorted.filter(like="Объем перевозок(тн)")
        .T.rolling(window=12)
        .sum()
        .T.min(axis=1)
        == 0
    )

    # 3. Подготовка данных для временного ряда
    freight_cols = volumes_sorted.filter(like="Провозная плата").columns
    volume_cols = volumes_sorted.filter(like="Объем перевозок(тн)").columns
    sequence_length = 12

    sequences = []
    client_ids = []

    for _, row in volumes_sorted.iterrows():
        freight = row[freight_cols].values.astype("float32")
        v = row[volume_cols].values.astype("float32")
        data = np.stack([freight, v], axis=1)

        if len(data) >= sequence_length + prediction_period:
            sequences.append(data[:sequence_length])
            client_ids.append(row["ID"])

    X = np.array(sequences, dtype="float32")

    return X, volumes_sorted, client_ids


async def handle(request):
    # Отображаем главную страницу
    return web.Response(
        text=render_template("upload.html", table=""), content_type="text/html"
    )


async def post_handle(request):
    reader = await request.post()
    file = reader.get("file")

    if file:
        file_bytes = file.file.read()
        file_name = file.filename

        file_class: type[File] | None = EXTS.get(file_name.split(".")[-1])

        if not file_class:
            return web.Response(
                text=render_template("upload.html", table=""),
                content_type="text/html",
            )

        file_handler = file_class(file_name, file_bytes)

        table = file_handler.get_table()

        for row in table[:5]:
            print(f"\033[32m{row}\033[0m")
            print()

        df = pd.DataFrame(table[1:])

        X, volumes_sorted, client_ids = preprocess(df)

        y = model.predict(X)
        print(y[:10])
        # k = 0
        # for r in table[3:9]:
        #     print(r[0], y[k])
        #     k += 1

        # Сопоставляем тестовые ID с вероятностями и добавляем данные из volumes_sorted
        test_results = pd.DataFrame(
            {"Client_ID": client_ids, "Churn_Probability": y.flatten()}
        )

        # Объединение результатов с исходными данными volumes_sorted на основе Client_ID
        merged_results = test_results.merge(
            volumes_sorted, left_on="Client_ID", right_on="ID"
        )

        # Удаление столбцов 'Отток' и дублирующегося 'ID', перемещение 'Churn_Probability' вправо
        merged_results.drop(columns=["Отток", "ID"], inplace=True)
        merged_results["Churn_Probability"] = merged_results.pop("Churn_Probability")

        # Сортировка по вероятности оттока
        merged_results = merged_results.sort_values(
            by="Churn_Probability", ascending=False
        )

        merged_results.to_excel("result.xlsx", index=False)

        print(merged_results.head())

        # rows = ""
        # for row in table:
        #     rows += "<tr>" + "".join(f"<td>{cell}</td>" for cell in row) + "</tr>"

        rows = ""
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
