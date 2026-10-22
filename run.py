# run.py
from pathlib import Path

from функции.чтение_1 import read_and_validate_excel
from функции.matcher import match_offsets
from функции.exporter import export_result_excel


INPUT_FILE = Path("input/задолженность.xlsx")
OUTPUT_FILE = Path("output/результат_взаимозачёта.xlsx")


def main():
    """Запускает анализ взаимозачёта по готовому входному файлу."""

    if not INPUT_FILE.exists():
        print(f"Не найден входной файл: {INPUT_FILE}")
        return

    # Чтение и проверка данных
    records = read_and_validate_excel(INPUT_FILE)

    # Подбор операций взаимозачёта и расчёт остатков
    operations, balances = match_offsets(records)

    # Сохранение результата
    export_result_excel(
        output_path=OUTPUT_FILE,
        operations=operations,
        balances=balances,
    )

    print(f"Отчёт создан: {OUTPUT_FILE}")


if __name__ == "__main__":
    main()