"""
Бизнес-логика назначения билета студенту.
"""
import logging
import random
from typing import Dict, List, Tuple

from . import storage

logger = logging.getLogger("lab2.logic")


def assign_ticket(
    results_path: str,
    group: str,
    surname: str,
    name: str,
    tickets: Dict[int, List[str]],
) -> Tuple[int, bool]:
    """
    Определяет номер билета для студента и признак повтора.

    Если студент уже есть в results.xlsx — возвращается номер билета из
    самой первой записи по этому студенту (идемпотентность), repeat=True.

    Если студента ещё нет — билет выбирается случайно среди номеров,
    присутствующих в tickets (то есть только среди валидных билетов),
    repeat=False.

    Возвращает (ticket_number, repeat).
    """
    if not tickets:
        raise ValueError("Нет ни одного валидного билета для генерации")

    existing = storage.find_first_ticket(results_path, group, surname, name)
    if existing is not None:
        logger.info(
            "Студент %s %s (%s) уже в журнале — повторно назначается билет №%d",
            surname, name, group, existing,
        )
        return existing, True

    ticket_number = random.choice(list(tickets.keys()))
    logger.info(
        "Студенту %s %s (%s) назначен новый билет №%d",
        surname, name, group, ticket_number,
    )
    return ticket_number, False
