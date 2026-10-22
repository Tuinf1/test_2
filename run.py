from pathlib import Path
import subprocess
import sys


PROJECT_DIR = Path(__file__).resolve().parent
FUNCTIONS_DIR = PROJECT_DIR / "функции"

SCRIPTS = [
    FUNCTIONS_DIR / "1_промежуточная таблица.py",
    FUNCTIONS_DIR / "2_функция_2_вторая.py",
    FUNCTIONS_DIR / "3_результат.py",
]


def main():
    for script in SCRIPTS:
        print(f"\nЗапуск: {script.name}")

        subprocess.run(
            [
                sys.executable,
                str(script),
            ],
            check=True,
        )

    print("\nВсе этапы выполнены.")
    print(
        "Итоговый файл: "
        f"{PROJECT_DIR / 'Промежуточная таблица 3.xlsx'}"
    )


if __name__ == "__main__":
    main()