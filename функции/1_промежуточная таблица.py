from collections import defaultdict
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill


PROJECT_DIR = Path(__file__).resolve().parent.parent

INPUT_FILE = (
    PROJECT_DIR
  
    / "задолженность.xlsx"
)

OUTPUT_FILE = (
    PROJECT_DIR

    / "Промежуточная таблица 1.xlsx"
)

DATA_SHEET_NAME = "Данные"
SECOND_SHEET_NAME = "1_ДО"

REQUIRED_COLUMNS = [
    "ID",
    "ДО",
    "Контрагент",
    "Договор",
    "Вид ДЗ",
    "Сумма ДЗ",
    "Сумма КЗ",
]


def read_excel_data(workbook):
    """
    Читает первичные данные из уже открытой книги.
    """

    if DATA_SHEET_NAME not in workbook.sheetnames:
        raise ValueError(
            f"Не найден лист «{DATA_SHEET_NAME}»."
        )

    worksheet = workbook[DATA_SHEET_NAME]

    rows = worksheet.iter_rows(values_only=True)

    headers = next(rows, None)

    if headers is None:
        raise ValueError("Лист с первичными данными пустой.")

    headers = list(headers)

    if headers != REQUIRED_COLUMNS:
        raise ValueError(
            "Неверные столбцы.\n"
            f"Ожидается: {REQUIRED_COLUMNS}\n"
            f"Получено: {headers}"
        )

    records = []

    for values in rows:
        if all(value is None for value in values):
            continue

        record = dict(
            zip(REQUIRED_COLUMNS, values)
        )

        records.append(record)

    return records


def get_debt_direction(record):
    """Возвращает текст: кто кому должен."""

    subsidiary = record["ДО"]
    counterparty = record["Контрагент"]

    receivable = record["Сумма ДЗ"] or 0
    payable = record["Сумма КЗ"] or 0

    if receivable > 0 and payable > 0:
        return (
            f"{counterparty} должен {subsidiary}: "
            f"{receivable}; "
            f"{subsidiary} должно {counterparty}: "
            f"{payable}"
        )

    if receivable > 0:
        return (
            f"{counterparty} должен "
            f"{subsidiary}: {receivable}"
        )

    if payable > 0:
        return (
            f"{subsidiary} должно "
            f"{counterparty}: {payable}"
        )

    return "Задолженность отсутствует"


def add_second_intermediate_sheet(
    workbook,
    records,
):
    """
    Добавляет в книгу лист
    «Вторая промежуточная».
    """

    if SECOND_SHEET_NAME in workbook.sheetnames:
        del workbook[SECOND_SHEET_NAME]

    worksheet = workbook.create_sheet(
        SECOND_SHEET_NAME
    )

    grouped_records = defaultdict(list)

    for record in records:
        subsidiary = record["ДО"]
        grouped_records[subsidiary].append(record)

    headers = [
        "ID",
        "Контрагент",
        "Договор",
        "Вид ДЗ",
        "Сумма ДЗ (контрагент должен ДО)",
        "Сумма КЗ (ДО должно контрагенту)",
        "Кто кому должен",
    ]

    current_row = 1

    for subsidiary, subsidiary_records in grouped_records.items():
        # Название ДО
        worksheet.cell(
            row=current_row,
            column=1,
            value=f"ДО: {subsidiary}",
        )

        worksheet.merge_cells(
            start_row=current_row,
            start_column=1,
            end_row=current_row,
            end_column=len(headers),
        )

        title_cell = worksheet.cell(
            row=current_row,
            column=1,
        )

        title_cell.font = Font(
            bold=True,
            size=12,
        )

        title_cell.fill = PatternFill(
            fill_type="solid",
            fgColor="D9EAF7",
        )

        current_row += 1

        # Заголовки столбцов
        for column_number, header in enumerate(
            headers,
            start=1,
        ):
            cell = worksheet.cell(
                row=current_row,
                column=column_number,
                value=header,
            )

            cell.font = Font(bold=True)

            cell.fill = PatternFill(
                fill_type="solid",
                fgColor="E2F0D9",
            )

        current_row += 1

        total_receivable = 0
        total_payable = 0

        # Все записи текущего ДО
        for record in subsidiary_records:
            receivable = record["Сумма ДЗ"] or 0
            payable = record["Сумма КЗ"] or 0

            values = [
                record["ID"],
                record["Контрагент"],
                record["Договор"],
                record["Вид ДЗ"],
                receivable,
                payable,
                get_debt_direction(record),
            ]

            for column_number, value in enumerate(
                values,
                start=1,
            ):
                worksheet.cell(
                    row=current_row,
                    column=column_number,
                    value=value,
                )

            total_receivable += receivable
            total_payable += payable

            current_row += 1

        # Итоги по ДО
        worksheet.cell(
            row=current_row,
            column=1,
            value="Итого по ДО",
        )

        worksheet.cell(
            row=current_row,
            column=5,
            value=total_receivable,
        )

        worksheet.cell(
            row=current_row,
            column=6,
            value=total_payable,
        )

        for column_number in range(
            1,
            len(headers) + 1,
        ):
            cell = worksheet.cell(
                row=current_row,
                column=column_number,
            )

            cell.font = Font(bold=True)

            cell.fill = PatternFill(
                fill_type="solid",
                fgColor="FFF2CC",
            )

        current_row += 2

    column_widths = {
        "A": 10,
        "B": 20,
        "C": 15,
        "D": 12,
        "E": 34,
        "F": 34,
        "G": 55,
    }

    for column, width in column_widths.items():
        worksheet.column_dimensions[column].width = width

    return worksheet


def create_second_intermediate_file(
    input_path,
    output_path,
):
    """
    Открывает исходную книгу, сохраняет лист «Данные»,
    добавляет второй лист и создаёт новый Excel-файл.
    """

    # Открываем исходную книгу целиком.
    # Поэтому лист «Данные» сохраняется.
    workbook = load_workbook(input_path)

    try:
        records = read_excel_data(workbook)

        add_second_intermediate_sheet(
            workbook=workbook,
            records=records,
        )

        output_path.parent.mkdir(
            parents=True,
            exist_ok=True,
        )

        # Сохраняем книгу под новым именем
        workbook.save(output_path)

    finally:
        workbook.close()


def main():
    if not INPUT_FILE.exists():
        print(f"Не найден файл: {INPUT_FILE}")
        return

    create_second_intermediate_file(
        input_path=INPUT_FILE,
        output_path=OUTPUT_FILE,
    )

    print(f"Создан файл: {OUTPUT_FILE}")
    print("Листы в файле:")
    print("1. Данные")
    print("2. Вторая промежуточная")


if __name__ == "__main__":
    main()