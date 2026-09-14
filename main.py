"""
Лабораторная работа №2: приложение с UI для генератора билетов.

Запуск:
    python main.py

Ожидаемые файлы в той же папке, что и main.py:
    students.xlsx — список студентов (лист на группу)
    tickets.docx  — список билетов
Создаётся при работе:
    results.xlsx  — журнал результатов
"""
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from app.gui import run

if __name__ == "__main__":
    base_dir = os.path.dirname(os.path.abspath(__file__))
    run(base_dir=base_dir)
