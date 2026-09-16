"""
Лабораторная работа №2: генератор билетов.

Запуск веб-интерфейса:
    python main.py

Десктоп (Tkinter):
    python main.py --desktop

Ожидаемые файлы в корне проекта:
    students.xlsx — список студентов (лист на группу)
    tickets.docx  — список билетов
Создаётся при работе:
    results.xlsx  — журнал результатов
"""
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

if __name__ == "__main__":
    base_dir = os.path.dirname(os.path.abspath(__file__))
    if "--desktop" in sys.argv:
        from app.gui import run
        run(base_dir=base_dir)
    else:
        from app.web import run
        run(base_dir=base_dir)
