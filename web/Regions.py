import pandas as pd

REGION_PATH = {
    "vo": "regions/МС_Владимирская область.xls",
    "ko": "regions/МС_Кировская область.xls",
    "no": "regions/МС_Нижегородская область.xls",
    "rme": "regions/МС_Республика Марий Эл.xls",
    "rm": "regions/МС_Республика Мордовия.xls",
    "rt": "regions/МС_Республика Татарстан.xls",
    "ru": "regions/МС_Республика Удмуртия.xls",
    "rch": "regions/МС_Республика Чувашия.xls",
}

REGION_CODE = {
    "vo": "Владимирская область",
    "ko": "Кировская область",
    "no": "Нижегородская область",
    "rme": "Республика Марий Эл",
    "rm": "Республика Мордовия",
    "rt": "Республика Татарстан",
    "ru": "Республика Удмуртия",
    "rch": "Республика Чувашия",
}


def get_region_df(region: str) -> pd.DataFrame:
    return pd.read_excel(REGION_PATH[region])
