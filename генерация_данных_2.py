from pathlib import Path
import random

from openpyxl import Workbook
from openpyxl.styles import Font


def get_debt_values(
    do_number: int,
    counterparty_number: int,
):
    """
    Возвращает обычную задолженность для одной строки.

    В строке будет либо ДЗ, либо КЗ.
    """
    pattern = (
        do_number + counterparty_number
    ) % 3

    if pattern == 0:
        return (
            "ПДЗ",
            random.randint(100, 1_000),
            0,
        )

    if pattern == 1:
        return (
            "ДЗ",
            random.randint(50, 800),
            0,
        )

    return (
        "ДЗ",
        0,
        random.randint(50, 800),
    )


def generate_input_excel(
    file_path: str = "задолженность.xlsx",
):
    """
    Создаёт тестовый Excel-файл.

    В файле будут:
    - двусторонние задолженности;
    - трёхсторонние задолженности;
    - ДЗ;
    - ПДЗ;
    - КЗ.
    """

    # Чтобы при каждом запуске получались одинаковые суммы
    random.seed(42)

    path = Path(file_path)
    path.parent.mkdir(
        parents=True,
        exist_ok=True,
    )

    workbook = Workbook()
    worksheet = workbook.active
    worksheet.title = "Данные"

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

    record_id = 1

    # Для этих пар будут созданы двусторонние долги.
    # Одинаковая ДО и одинаковый контрагент
    # будут иметь отдельную строку ДЗ и отдельную строку КЗ.
    bilateral_pairs = {
        (1, 1),
        (1, 2),
        (1, 3),
    }

    for do_number in range(1, 11):
        for counterparty_number in range(1, 11):
            pair = (
                do_number,
                counterparty_number,
            )

            if pair in bilateral_pairs:
                # Контрагент должен ДО — дебиторская задолженность
                worksheet.append([
                    record_id,
                    f"ДО{do_number}",
                    f"К/а {counterparty_number}",
                    f"ДД{record_id:03d}",
                    "ПДЗ",
                    random.randint(300, 1_000),
                    0,
                ])

                record_id += 1

                # Эта же ДО должна этому же контрагенту —
                # кредиторская задолженность
                worksheet.append([
                    record_id,
                    f"ДО{do_number}",
                    f"К/а {counterparty_number}",
                    f"ДД{record_id:03d}",
                    "ДЗ",
                    0,
                    random.randint(100, 700),
                ])

                record_id += 1

                continue

            debt_type, receivable, payable = (
                get_debt_values(
                    do_number,
                    counterparty_number,
                )
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

    # Жирные заголовки
    for cell in worksheet[1]:
        cell.font = Font(bold=True)

    # Закрепление первой строки
    worksheet.freeze_panes = "A2"

    # Фильтр на всю таблицу
    worksheet.auto_filter.ref = (
        f"A1:G{worksheet.max_row}"
    )

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
        worksheet.column_dimensions[
            column
        ].width = width

    workbook.save(path)
    workbook.close()

    print(f"Файл создан: {path}")
    print(
        "Двусторонние пары: "
        "ДО1/К/а 1, ДО1/К/а 2, ДО1/К/а 3"
    )
    print(
        f"Количество строк: {record_id - 1}"
    )


if __name__ == "__main__":
    generate_input_excel()