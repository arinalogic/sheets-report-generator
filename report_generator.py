"""Окно для создания отчётов о продажах в Google Таблицах."""

import os
import random
import threading
import tkinter as tk
from datetime import date, datetime, timedelta
from tkinter import messagebox, ttk
from typing import Any

from dotenv import load_dotenv

from google_sheets import GoogleSheets, GoogleSheetsError


CLIENTS = [
    "ООО Альта",
    "Вектор Плюс",
    "Северный Порт",
    "Маяк Технологии",
    "Городские Решения",
    "Новый Контур",
    "Орбита Сервис",
    "Точка Роста",
    "Профи Групп",
    "Линия Снабжения",
]
DEAL_NAMES = [
    "Поставка оборудования",
    "Годовое обслуживание",
    "Внедрение платформы",
    "Поставка комплектующих",
    "Консультационные услуги",
    "Расширение лицензий",
    "Обновление инфраструктуры",
    "Сервисный контракт",
]
STATUSES = ["Завершена", "В работе", "Отменена"]


def color(red: int, green: int, blue: int) -> dict[str, float]:
    """Преобразует RGB-цвет в формат Google Sheets API."""
    return {"red": red / 255, "green": green / 255, "blue": blue / 255}


def sheet_range(sheet_name: str, cell_range: str) -> str:
    """Добавляет к диапазону экранированное название листа."""
    return f"{GoogleSheets._quote_sheet_name(sheet_name)}!{cell_range}"


def generate_deals(start_date: date, end_date: date) -> list[list[Any]]:
    """Создаёт случайные сделки за выбранный период."""
    period_days = (end_date - start_date).days
    deals: list[list[Any]] = []
    for number in range(1, random.randint(10, 15) + 1):
        deal_date = start_date + timedelta(days=random.randint(0, period_days))
        deals.append(
            [
                number,
                deal_date.isoformat(),
                random.choice(CLIENTS),
                random.choice(DEAL_NAMES),
                round(random.uniform(2500, 35000), 2),
                random.choice(STATUSES),
            ]
        )
    return deals


def write_report(settings: dict[str, str]) -> str:
    """Создаёт и оформляет лист отчёта в Google Таблице."""
    load_dotenv()
    credentials_path = os.getenv("GOOGLE_CREDENTIALS_PATH")
    spreadsheet_id = os.getenv("SPREADSHEET_ID")
    if not credentials_path or not spreadsheet_id:
        raise GoogleSheetsError(
            "Укажите GOOGLE_CREDENTIALS_PATH и SPREADSHEET_ID в файле .env."
        )

    sheets = GoogleSheets(credentials_path, spreadsheet_id)
    start_date = datetime.strptime(settings["start_date"], "%Y-%m-%d").date()
    end_date = datetime.strptime(settings["end_date"], "%Y-%m-%d").date()
    deals = generate_deals(start_date, end_date)
    revenue = round(sum(row[4] for row in deals if row[5] == "Завершена"), 2)
    deal_count = len(deals)
    client_count = len({row[2] for row in deals})

    created_at = datetime.now()
    sheet_name = ""
    existing_names = sheets.get_sheet_names()
    while True:
        sheet_name = (
            f"Отчёт {settings['start_date']}_{settings['end_date']}_"
            f"{created_at:%H-%M-%S}"
        )
        if sheet_name not in existing_names:
            break
        created_at += timedelta(seconds=1)

    sheets.add_sheet(sheet_name)
    quoted_range = lambda cells: sheet_range(sheet_name, cells)

    sheets.write_range(
        quoted_range("A1:F1"),
        [[f"ОТЧЁТ О ПРОДАЖАХ | {settings['start_date']} — {settings['end_date']}"]],
        value_input_option="RAW",
    )
    sheets.merge_cells(quoted_range("A1:F1"))
    sheets.format_range(
        quoted_range("A1:F1"),
        background_color=color(27, 77, 78),
        text_color=color(255, 255, 255),
        font_size=18,
        bold=True,
        horizontal_alignment="CENTER",
    )

    created_text = created_at.strftime("%Y-%m-%d %H:%M:%S")
    parameter_rows = [
        ["Период", f"{settings['start_date']} — {settings['end_date']}", "", "Тип отчёта", settings["report_type"], ""],
        ["Отдел", settings["department"], "", "Ответственный", settings["responsible"], ""],
        ["Комментарий", settings["comment"], "", "", "", ""],
        ["Дата создания", created_text, "", "", "", ""],
    ]
    sheets.write_range(quoted_range("A3:F6"), parameter_rows, value_input_option="RAW")
    for merge_range in ("B3:C3", "E3:F3", "B4:C4", "E4:F4", "B5:F5", "B6:F6"):
        sheets.merge_cells(quoted_range(merge_range))
    sheets.format_range(
        quoted_range("A3:F6"),
        background_color=color(239, 246, 243),
        text_color=color(40, 58, 57),
    )
    for label_range in ("A3:A6", "D3:D4"):
        sheets.format_range(
            quoted_range(label_range),
            text_color=color(27, 77, 78),
            bold=True,
        )

    table_start = 8
    table_header = ["№", "Дата", "Клиент", "Название сделки", "Сумма", "Статус"]
    sheets.write_range(
        quoted_range(f"A{table_start}:F{table_start}"),
        [table_header],
        value_input_option="RAW",
    )
    sheets.write_range(
        quoted_range(f"A{table_start + 1}:F{table_start + len(deals)}"),
        deals,
        value_input_option="RAW",
    )
    last_deal_row = table_start + len(deals)
    sheets.format_range(
        quoted_range(f"A{table_start}:F{table_start}"),
        background_color=color(47, 112, 105),
        text_color=color(255, 255, 255),
        bold=True,
        horizontal_alignment="CENTER",
    )
    sheets.set_borders(
        quoted_range(f"A{table_start}:F{last_deal_row}"),
        color=color(198, 214, 209),
    )
    sheets.set_number_format(
        quoted_range(f"E{table_start + 1}:E{last_deal_row}"),
        format_type="CURRENCY",
        pattern='"₽" #,##0.00',
    )

    for row_number, deal in enumerate(deals, start=table_start + 1):
        if row_number % 2 == 0:
            sheets.format_range(
                quoted_range(f"A{row_number}:F{row_number}"),
                background_color=color(246, 249, 247),
            )
        status_color = {
            "Завершена": color(34, 113, 76),
            "В работе": color(164, 102, 20),
            "Отменена": color(165, 57, 57),
        }[deal[5]]
        sheets.format_range(
            quoted_range(f"F{row_number}"),
            text_color=status_color,
            bold=True,
        )

    total_row = last_deal_row + 2
    sheets.write_range(
        quoted_range(f"A{total_row}:F{total_row}"),
        [[
            f"ИТОГО: сделок — {deal_count} | клиентов — {client_count} "
            "| выручка завершённых сделок —",
            "",
            "",
            "",
            "",
            revenue,
        ]],
        value_input_option="RAW",
    )
    sheets.merge_cells(quoted_range(f"A{total_row}:E{total_row}"))
    sheets.format_range(
        quoted_range(f"A{total_row}:F{total_row}"),
        background_color=color(221, 237, 230),
        text_color=color(27, 77, 78),
        bold=True,
    )
    sheets.set_number_format(
        quoted_range(f"F{total_row}"),
        format_type="CURRENCY",
        pattern='"₽" #,##0.00',
    )

    for column_range, width in (
        ("A1:A1", 120),
        ("B1:B1", 115),
        ("C1:C1", 175),
        ("D1:D1", 230),
        ("E1:E1", 130),
        ("F1:F1", 155),
    ):
        sheets.set_column_width(quoted_range(column_range), width)

    return sheet_name


class ReportGenerator:
    """Окно ввода параметров и запуска генерации отчёта."""

    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title("Генератор отчётов")
        self.root.resizable(False, False)

        today = date.today()
        self.start_date = tk.StringVar(value=(today - timedelta(days=30)).isoformat())
        self.end_date = tk.StringVar(value=today.isoformat())
        self.report_type = tk.StringVar(value="Ежемесячный")
        self.department = tk.StringVar(value="Продажи")
        self.responsible = tk.StringVar()
        self.comment = tk.StringVar()
        self.status = tk.StringVar(value="Готов к генерации отчета")

        self._build_widgets()

    def _build_widgets(self) -> None:
        frame = ttk.Frame(self.root, padding=18)
        frame.grid(sticky="nsew")
        frame.columnconfigure(1, weight=1)

        fields = [
            ("Дата начала (YYYY-MM-DD)", ttk.Entry(frame, textvariable=self.start_date)),
            ("Дата окончания (YYYY-MM-DD)", ttk.Entry(frame, textvariable=self.end_date)),
            (
                "Тип отчёта",
                ttk.Combobox(
                    frame,
                    textvariable=self.report_type,
                    values=("Ежемесячный", "Еженедельный", "Квартальный"),
                    state="readonly",
                ),
            ),
            (
                "Отдел",
                ttk.Combobox(
                    frame,
                    textvariable=self.department,
                    values=("Продажи", "Маркетинг", "Логистика"),
                    state="readonly",
                ),
            ),
            ("Ответственное лицо", ttk.Entry(frame, textvariable=self.responsible)),
            ("Комментарий", ttk.Entry(frame, textvariable=self.comment)),
        ]
        for row_number, (label, widget) in enumerate(fields):
            ttk.Label(frame, text=label).grid(
                row=row_number, column=0, sticky="w", padx=(0, 12), pady=5
            )
            widget.grid(row=row_number, column=1, sticky="ew", pady=5)

        self.generate_button = ttk.Button(
            frame, text="Сгенерировать отчет", command=self._start_generation
        )
        self.generate_button.grid(
            row=len(fields), column=0, columnspan=2, sticky="ew", pady=(14, 8)
        )
        ttk.Label(frame, textvariable=self.status).grid(
            row=len(fields) + 1, column=0, columnspan=2, sticky="w"
        )

    def _get_settings(self) -> dict[str, str] | None:
        try:
            start_date = datetime.strptime(self.start_date.get(), "%Y-%m-%d").date()
            end_date = datetime.strptime(self.end_date.get(), "%Y-%m-%d").date()
            if start_date.isoformat() != self.start_date.get() or end_date.isoformat() != self.end_date.get():
                raise ValueError
        except ValueError:
            messagebox.showerror(
                "Некорректная дата",
                "Введите даты в формате YYYY-MM-DD, например 2026-10-16.",
                parent=self.root,
            )
            return None

        if start_date > end_date:
            messagebox.showerror(
                "Некорректный период",
                "Дата начала не может быть позже даты окончания.",
                parent=self.root,
            )
            return None
        responsible = self.responsible.get().strip()
        if not responsible:
            messagebox.showerror(
                "Не указано ответственное лицо",
                "Заполните поле «Ответственное лицо».",
                parent=self.root,
            )
            return None

        return {
            "start_date": start_date.isoformat(),
            "end_date": end_date.isoformat(),
            "report_type": self.report_type.get(),
            "department": self.department.get(),
            "responsible": responsible,
            "comment": self.comment.get(),
        }

    def _start_generation(self) -> None:
        settings = self._get_settings()
        if settings is None:
            return

        self.generate_button.configure(state="disabled")
        self.status.set("Генерирую…")
        threading.Thread(
            target=self._generate_in_background,
            args=(settings,),
            daemon=True,
        ).start()

    def _generate_in_background(self, settings: dict[str, str]) -> None:
        try:
            sheet_name = write_report(settings)
        except GoogleSheetsError as error:
            self.root.after(0, self._show_error, str(error))
        except Exception:
            self.root.after(
                0,
                self._show_error,
                "Не удалось создать отчёт. Проверьте подключение и настройки Google Таблиц.",
            )
        else:
            self.root.after(0, self._show_success, sheet_name)

    def _show_error(self, message: str) -> None:
        self.status.set("Ошибка генерации отчета")
        self.generate_button.configure(state="normal")
        messagebox.showerror("Ошибка", message, parent=self.root)

    def _show_success(self, sheet_name: str) -> None:
        self.status.set(f"Отчёт записан на лист «{sheet_name}»")
        self.generate_button.configure(state="normal")


if __name__ == "__main__":
    root = tk.Tk()
    ReportGenerator(root)
    root.mainloop()