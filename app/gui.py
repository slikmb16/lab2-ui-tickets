"""
GUI-приложение "Генератор билетов" на Tkinter.
"""
import logging
import os
import tkinter as tk
from tkinter import messagebox, ttk
from typing import Dict, List, Optional

from . import logic, parsing, storage

logger = logging.getLogger("lab2.gui")

STUDENTS_FILE = "students.xlsx"
TICKETS_FILE = "tickets.docx"
RESULTS_FILE = "results.xlsx"


class TicketApp(tk.Tk):
    def __init__(self, base_dir: str = "."):
        super().__init__()
        self.title("Генератор билетов")
        self.geometry("420x220")
        self.resizable(False, False)

        self.base_dir = base_dir
        self.students_path = os.path.join(base_dir, STUDENTS_FILE)
        self.tickets_path = os.path.join(base_dir, TICKETS_FILE)
        self.results_path = os.path.join(base_dir, RESULTS_FILE)

        self.groups: Dict[str, List[parsing.Student]] = {}
        self.tickets: Dict[int, List[str]] = {}

        self.selected_group = tk.StringVar()
        self.selected_student = tk.StringVar()

        self._build_ui()
        self._load_sources()

        self.bind("<Escape>", self._on_escape_global)

    # ------------------------------------------------------------------
    # UI construction
    # ------------------------------------------------------------------
    def _build_ui(self) -> None:
        pad = {"padx": 12, "pady": 8}

        frame = ttk.Frame(self)
        frame.pack(fill="both", expand=True, **pad)

        ttk.Label(frame, text="Группа:").grid(row=0, column=0, sticky="w", pady=4)
        self.group_combo = ttk.Combobox(
            frame, textvariable=self.selected_group, state="readonly", width=30
        )
        self.group_combo.grid(row=0, column=1, pady=4)
        self.group_combo.bind("<<ComboboxSelected>>", self._on_group_selected)

        ttk.Label(frame, text="Студент:").grid(row=1, column=0, sticky="w", pady=4)
        self.student_combo = ttk.Combobox(
            frame, textvariable=self.selected_student, state="disabled", width=30
        )
        self.student_combo.grid(row=1, column=1, pady=4)
        self.student_combo.bind("<<ComboboxSelected>>", self._on_student_selected)

        self.generate_btn = ttk.Button(
            frame,
            text="Сгенерировать билет",
            command=self._on_generate_clicked,
            state="disabled",
        )
        self.generate_btn.grid(row=2, column=0, columnspan=2, pady=20)

        self.status_label = ttk.Label(frame, text="", foreground="#a00", wraplength=380)
        self.status_label.grid(row=3, column=0, columnspan=2, sticky="w")

    # ------------------------------------------------------------------
    # Загрузка исходных данных
    # ------------------------------------------------------------------
    def _load_sources(self) -> None:
        missing = []
        if not os.path.exists(self.students_path):
            missing.append(STUDENTS_FILE)
        if not os.path.exists(self.tickets_path):
            missing.append(TICKETS_FILE)

        if missing:
            msg = (
                "Не найдены необходимые файлы: " + ", ".join(missing) +
                f"\nОжидаются в папке: {os.path.abspath(self.base_dir)}"
            )
            self._show_fatal(msg)
            return

        try:
            self.groups = parsing.load_students(self.students_path)
        except Exception as exc:
            self._show_fatal(f"Не удалось прочитать {STUDENTS_FILE}:\n{exc}")
            return

        try:
            self.tickets = parsing.load_tickets(self.tickets_path)
        except Exception as exc:
            self._show_fatal(f"Не удалось прочитать {TICKETS_FILE}:\n{exc}")
            return

        if not self.tickets:
            self._show_fatal(
                f"В файле {TICKETS_FILE} не найдено ни одного корректного билета."
            )
            return

        try:
            storage.ensure_results_file(self.results_path)
        except Exception as exc:
            self._show_fatal(f"Не удалось создать/открыть {RESULTS_FILE}:\n{exc}")
            return

        self.group_combo["values"] = list(self.groups.keys())
        self.status_label.config(
            foreground="#333",
            text=f"Загружено групп: {len(self.groups)}, билетов: {len(self.tickets)}",
        )

    def _show_fatal(self, message: str) -> None:
        self.group_combo.config(state="disabled")
        self.student_combo.config(state="disabled")
        self.generate_btn.config(state="disabled")
        self.status_label.config(foreground="#a00", text=message)
        messagebox.showerror("Ошибка загрузки данных", message)

    # ------------------------------------------------------------------
    # Обработчики выбора
    # ------------------------------------------------------------------
    def _on_group_selected(self, _event=None) -> None:
        group = self.selected_group.get()
        students = self.groups.get(group, [])

        self.selected_student.set("")
        self.generate_btn.config(state="disabled")

        if not students:
            self.student_combo.config(state="disabled", values=[])
            self.status_label.config(
                foreground="#a00", text=f"В группе «{group}» нет студентов."
            )
            return

        names = [s.full_name for s in students]
        self.student_combo.config(state="readonly", values=names)
        self.status_label.config(foreground="#333", text="")

    def _on_student_selected(self, _event=None) -> None:
        if self.selected_group.get() and self.selected_student.get():
            self.generate_btn.config(state="normal")
        else:
            self.generate_btn.config(state="disabled")

    # ------------------------------------------------------------------
    # Генерация билета
    # ------------------------------------------------------------------
    def _on_generate_clicked(self) -> None:
        group = self.selected_group.get()
        full_name = self.selected_student.get()
        if not group or not full_name:
            return

        surname, _, name = full_name.partition(" ")

        try:
            ticket_number, repeat = logic.assign_ticket(
                self.results_path, group, surname, name, self.tickets
            )
        except Exception as exc:
            messagebox.showerror("Ошибка", f"Не удалось назначить билет:\n{exc}")
            return

        try:
            storage.append_result(
                self.results_path, group, surname, name, ticket_number, repeat
            )
        except storage.ResultsFileLocked as exc:
            messagebox.showerror("Файл занят", str(exc))
            return
        except Exception as exc:
            messagebox.showerror("Ошибка записи", f"Не удалось сохранить результат:\n{exc}")
            return

        questions = self.tickets[ticket_number]
        self._show_result_window(ticket_number, questions)

    # ------------------------------------------------------------------
    # Окно результата
    # ------------------------------------------------------------------
    def _show_result_window(self, ticket_number: int, questions: List[str]) -> None:
        win = tk.Toplevel(self)
        win.title("Результат")
        win.geometry("480x320")
        win.transient(self)
        win.grab_set()

        ttk.Label(
            win, text=f"Ваш билет № {ticket_number}", font=("Segoe UI", 14, "bold")
        ).pack(pady=(16, 8), padx=16, anchor="w")

        for i, question in enumerate(questions[:3], start=1):
            ttk.Label(
                win, text=f"{i}. {question}", wraplength=440, justify="left"
            ).pack(pady=4, padx=16, anchor="w")

        ttk.Label(
            win, text="(ESC — закрыть и вернуться к выбору)", foreground="#666"
        ).pack(pady=(20, 8))

        def close_and_reset(_event=None):
            win.destroy()
            self._reset_selection()

        win.bind("<Escape>", close_and_reset)
        win.protocol("WM_DELETE_WINDOW", close_and_reset)
        win.focus_set()

    def _reset_selection(self) -> None:
        self.selected_student.set("")
        self.student_combo.set("")
        self.generate_btn.config(state="disabled")

    def _on_escape_global(self, _event=None) -> None:
        # ESC на главном окне не должен закрывать приложение.
        pass


def run(base_dir: str = ".") -> None:
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s [%(levelname)s] %(name)s: %(message)s",
    )
    app = TicketApp(base_dir=base_dir)
    app.mainloop()
