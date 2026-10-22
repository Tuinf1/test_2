from collections import defaultdict
from pathlib import Path

from openpyxl import load_workbook


PROJECT_DIR = Path(__file__).resolve().parent.parent

INPUT_FILE = PROJECT_DIR / "Промежуточная таблица 1.xlsx"
OUTPUT_FILE = PROJECT_DIR / "Промежуточная таблица 2.xlsx"

DATA_SHEET = "Данные"
NEW_SHEET = "2_Контрагент"


def create_intermediate_table_2():
    # Открываем первую промежуточную книгу
    workbook = load_workbook(INPUT_FILE)

    # Читаем первичные данные
    data_sheet = workbook[DATA_SHEET]
    rows = data_sheet.iter_rows(values_only=True)

    headers = next(rows)

    records = [
        dict(zip(headers, row))
        for row in rows
        if any(value is not None for value in row)
    ]

    # Удаляем старую версию листа при повторном запуске
    if NEW_SHEET in workbook.sheetnames:
        del workbook[NEW_SHEET]

    worksheet = workbook.create_sheet(NEW_SHEET)

    # Группируем строки по контрагентам
    grouped_records = defaultdict(list)

    for record in records:
        grouped_records[
            record["Контрагент"]
        ].append(record)

    table_headers = [
        "ID",
        "ДО",
        "Договор",
        "Вид ДЗ",
        "Сумма ДЗ",
        "Сумма КЗ",
        "Остаток ДЗ",
        "Остаток КЗ",
    ]

    for counterparty, counterparty_records in grouped_records.items():
        worksheet.append([
            f"Контрагент: {counterparty}"
        ])

        worksheet.append(table_headers)

        total_receivable = 0
        total_payable = 0

        for record in counterparty_records:
            receivable = record["Сумма ДЗ"] or 0
            payable = record["Сумма КЗ"] or 0

            worksheet.append([
                record["ID"],
                record["ДО"],
                record["Договор"],
                record["Вид ДЗ"],
                receivable,
                payable,
                receivable,
                payable,
            ])

            total_receivable += receivable
            total_payable += payable

        worksheet.append([
            "Итого",
            "",
            "",
            "",
            total_receivable,
            total_payable,
            total_receivable,
            total_payable,
        ])

        # Пустая строка между контрагентами
        worksheet.append([])

    # Сохраняем новую книгу вместе со старыми листами
    workbook.save(OUTPUT_FILE)
    workbook.close()

    print(f"Создан файл: {OUTPUT_FILE}")


if __name__ == "__main__":
    create_intermediate_table_2()