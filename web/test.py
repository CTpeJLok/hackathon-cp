import pdfplumber

with pdfplumber.open(
    "/home/andrey/Загрузки/Telegram Desktop/1_пр_семинар_№_1_командные_запросы_curl.pdf"
) as pdf:
    for page in pdf.pages:
        print(page.extract_table())
