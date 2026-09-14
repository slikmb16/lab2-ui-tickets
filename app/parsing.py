"""
Парсинг входных файлов: students.xlsx и tickets.docx.
"""
import logging
import re
from dataclasses import dataclass
from typing import Dict, List, Optional

from openpyxl import load_workbook
from docx import Document

logger = logging.getLogger("lab2.parsing")


# ---------------------------------------------------------------------------
# students.xlsx
# ---------------------------------------------------------------------------

@dataclass(frozen=True)
class Student:
    surname: str
    name: str

    @property
    def full_name(self) -> str:
        return f"{self.surname} {self.name}"


def load_students(path: str) -> Dict[str, List[Student]]:
    """
    Читает students.xlsx.
    Один лист на группу, имя листа = имя группы.
    Строка 1 — заголовок, данные с строки 2: A — Фамилия, B — Имя.

    Возвращает {группа: [Student, ...]} (порядок листов и строк сохранён).
    Пустые строки и листы без данных допускаются (группа со пустым списком).
    """
    wb = load_workbook(path, read_only=True, data_only=True)
    groups: Dict[str, List[Student]] = {}

    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        students: List[Student] = []
        for row in ws.iter_rows(min_row=2, max_col=2, values_only=True):
            surname, name = (row + (None, None))[:2]
            surname = (str(surname).strip() if surname is not None else "")
            name = (str(name).strip() if name is not None else "")
            if not surname and not name:
                continue  # пустая строка — пропускаем
            if not surname or not name:
                logger.warning(
                    "Лист '%s': неполная запись студента (Фамилия=%r, Имя=%r) — пропущена",
                    sheet_name, surname, name,
                )
                continue
            students.append(Student(surname=surname, name=name))
        groups[sheet_name] = students

    wb.close()
    return groups


# ---------------------------------------------------------------------------
# tickets.docx
# ---------------------------------------------------------------------------

_TICKET_HEADER_RE = re.compile(r"^\s*Билет\s+(\d+)\s*$", re.IGNORECASE)
_QUESTION_RE = re.compile(r"^\s*(\d+)\s*[.)]\s*(.+?)\s*$")


def load_tickets(path: str) -> Dict[int, List[str]]:
    """
    Парсит tickets.docx.

    Формат:
        Билет 1
        1. Вопрос первый
        2. Вопрос второй
        3. Вопрос третий

        Билет 2
        ...

    Билет считается валидным, если у него ровно >= 3 распознанных
    пронумерованных вопроса. Невалидные билеты (заголовок не распознан,
    меньше 3 вопросов) пропускаются, ошибка логируется — парсинг
    остальных билетов продолжается.

    Возвращает {номер_билета: [вопрос1, вопрос2, вопрос3, ...]}.
    Если у билета больше 3 вопросов, берутся первые 3 (по номеру билета
    достаточно 3 вопросов согласно ТЗ).
    """
    doc = Document(path)
    paragraphs = [p.text for p in doc.paragraphs]

    tickets: Dict[int, List[str]] = {}
    current_number: Optional[int] = None
    current_questions: List[str] = []

    def _flush():
        nonlocal current_number, current_questions
        if current_number is None:
            return
        if len(current_questions) < 3:
            logger.error(
                "Билет %s пропущен: найдено только %d вопрос(а/ов), нужно минимум 3",
                current_number, len(current_questions),
            )
        elif current_number in tickets:
            logger.error(
                "Билет %s пропущен: дублирующийся номер билета", current_number
            )
        else:
            tickets[current_number] = current_questions[:3]
        current_number = None
        current_questions = []

    for raw_line in paragraphs:
        line = raw_line.strip()
        if not line:
            continue

        header_match = _TICKET_HEADER_RE.match(line)
        if header_match:
            _flush()
            current_number = int(header_match.group(1))
            current_questions = []
            continue

        question_match = _QUESTION_RE.match(line)
        if question_match and current_number is not None:
            current_questions.append(question_match.group(2).strip())
            continue

        # Строка не распознана ни как заголовок, ни как вопрос — игнорируем,
        # но логируем на случай, если это "потерянный" билет/вопрос.
        if current_number is None:
            logger.debug("Строка вне билета проигнорирована: %r", line)
        else:
            logger.debug(
                "Билет %s: нераспознанная строка проигнорирована: %r",
                current_number, line,
            )

    _flush()  # последний билет в файле

    if not tickets:
        logger.error("В tickets.docx не найдено ни одного валидного билета")

    return tickets
