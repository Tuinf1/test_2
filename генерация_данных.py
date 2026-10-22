from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Font

import random


def get_debt_values(do_number: int, counterparty_number: int, record_id: int):
    """Возвращает случайный вид и суммы задолженности."""
    pattern = (do_number + counterparty_number) % 3

    if pattern == 0:
        return "ПДЗ", random.randint(100, 1_000), 0

    if pattern == 1:
        return "ДЗ", random.randint(50, 800), 0

    return "ДЗ", 0, random.randint(50, 800)

# def get_debt_values(do_number: int, counterparty_number: int, record_id: int):
#     """
#     Возвращает вид задолженности, сумму ДЗ и сумму КЗ
#     для одной строки примера.
#     """
#     pattern = (do_number + counterparty_number) % 3

#     if pattern == 0:
#         # Просроченная дебиторская задолженность
#         return "ПДЗ", 100 + record_id, 0

#     if pattern == 1:
#         # Обычная дебиторская задолженность
#         return "ДЗ", 60 + record_id, 0

#     # Кредиторская задолженность
#     return "ДЗ", 0, 50 + record_id


def generate_input_excel(file_path: str = "задолженность.xlsx"):
    """
    Создаёт Excel-файл с листом «Данные».

    В файле будет:
    - 10 ДО;
    - 10 контрагентов;
    - 100 строк: каждая ДО имеет запись по каждому контрагенту.
    """
    path = Path(file_path)
    path.parent.mkdir(parents=True, exist_ok=True)

    workbook = Workbook()
    worksheet = workbook.active
    worksheet.title = "Данные"

    # Заголовки таблицы
    headers = [
        "ID",
        "ДО",
        "Контрагент",
        "Договор",
        "Вид ДЗ",
        "Сумма ДЗ",
        "Сумма КЗ",
    ]
    worksheet.append(headers)

    # Формирование 100 строк данных
    record_id = 1

    for do_number in range(1, 11):
        for counterparty_number in range(1, 11):
            debt_type, receivable, payable = get_debt_values(
                do_number,
                counterparty_number,
                record_id,
            )

            worksheet.append([
                record_id,
                f"ДО{do_number}",
                f"К/а {counterparty_number}",
                f"ДД{record_id:03d}",
                debt_type,
                receivable,
                payable,
            ])

            record_id += 1

    # Простое оформление заголовков
    for cell in worksheet[1]:
        cell.font = Font(bold=True)

    # Фильтр и закрепление первой строки
    worksheet.auto_filter.ref = "A1:G101"
    worksheet.freeze_panes = "A2"

    # Ширина столбцов
    widths = {
        "A": 8,
        "B": 10,
        "C": 18,
        "D": 12,
        "E": 10,
        "F": 14,
        "G": 14,
    }

    for column, width in widths.items():
        worksheet.column_dimensions[column].width = width

    workbook.save(path)
    workbook.close()

    print(f"Файл создан: {path}")


# Запуск генерации
generate_input_excel()