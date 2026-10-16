"""Простой интерфейс для работы с Google Таблицами."""

from pathlib import Path
from typing import Any, Sequence

from google.auth.exceptions import GoogleAuthError
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError


class GoogleSheetsError(Exception):
    """Понятная ошибка при работе с Google Таблицами."""


class GoogleSheets:
    """Выполняет операции с Google Таблицей через сервисный аккаунт."""

    SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]

    def __init__(self, credentials_path: str | Path, spreadsheet_id: str) -> None:
        """Создаёт подключение по JSON-ключу и ID таблицы."""
        self.spreadsheet_id = spreadsheet_id

        try:
            credentials = Credentials.from_service_account_file(
                str(credentials_path), scopes=self.SCOPES
            )
        except FileNotFoundError as error:
            raise GoogleSheetsError(
                "Файл JSON-ключа сервисного аккаунта не найден."
            ) from None
        except (OSError, ValueError, GoogleAuthError) as error:
            raise GoogleSheetsError(
                "Не удалось загрузить JSON-ключ сервисного аккаунта. "
                "Проверьте путь и формат файла."
            ) from None

        self.service = build(
            "sheets", "v4", credentials=credentials, cache_discovery=False
        )

    def _execute(self, request: Any) -> dict[str, Any]:
        try:
            return request.execute()
        except HttpError as error:
            status = getattr(error.resp, "status", None)
            if status == 403:
                message = "Нет доступа к таблице. Проверьте доступ сервисного аккаунта."
            elif status == 404:
                message = "Таблица не найдена. Проверьте ID таблицы."
            elif status == 400:
                message = (
                    "Некорректный запрос к Google Sheets API. "
                    "Проверьте диапазон и параметры."
                )
            else:
                message = f"Ошибка Google Sheets API (HTTP {status})."
            raise GoogleSheetsError(message) from None

    def _get_sheet_properties(self) -> list[dict[str, Any]]:
        response = self._execute(
            self.service.spreadsheets().get(
                spreadsheetId=self.spreadsheet_id,
                fields="sheets.properties.title,sheets.properties.sheetId",
            )
        )
        return [sheet["properties"] for sheet in response.get("sheets", [])]

    @staticmethod
    def _clean_values(values: list[list[Any]], clean_nbsp: bool) -> list[list[Any]]:
        if not clean_nbsp:
            return values
        return [
            [value.replace("\xa0", " ") if isinstance(value, str) else value for value in row]
            for row in values
        ]

    @staticmethod
    def _quote_sheet_name(sheet_name: str) -> str:
        return "'" + sheet_name.replace("'", "''") + "'"

    def get_sheet_names(self, clean_nbsp: bool = True) -> list[str]:
        """Возвращает названия листов таблицы."""
        names = [sheet["title"] for sheet in self._get_sheet_properties()]
        if clean_nbsp:
            return [name.replace("\xa0", " ") for name in names]
        return names

    def read_all_cells(
        self, sheet_name: str | None = None, clean_nbsp: bool = True
    ) -> list[list[Any]]:
        """Читает все заполненные ячейки указанного или первого листа."""
        if sheet_name is None:
            sheets = self._get_sheet_properties()
            if not sheets:
                return []
            sheet_name = sheets[0]["title"]
        return self.read_range(
            self._quote_sheet_name(sheet_name), clean_nbsp=clean_nbsp
        )

    def read_range(
        self, range_name: str, clean_nbsp: bool = True
    ) -> list[list[Any]]:
        """Читает значения из указанного диапазона."""
        response = self._execute(
            self.service.spreadsheets()
            .values()
            .get(spreadsheetId=self.spreadsheet_id, range=range_name)
        )
        return self._clean_values(response.get("values", []), clean_nbsp)

    def write_range(
        self,
        range_name: str,
        values: Sequence[Sequence[Any]],
        value_input_option: str = "RAW",
    ) -> dict[str, Any]:
        """Записывает значения в указанный диапазон."""
        return self._execute(
            self.service.spreadsheets()
            .values()
            .update(
                spreadsheetId=self.spreadsheet_id,
                range=range_name,
                valueInputOption=value_input_option,
                body={"values": values},
            )
        )

    def update_cell(self, cell: str, value: Any) -> dict[str, Any]:
        """Изменяет значение одной ячейки."""
        return self.write_range(cell, [[value]])

    def append_row(self, sheet_name: str, values: Sequence[Any]) -> dict[str, Any]:
        """Добавляет строку в конец данных указанного листа."""
        return self._execute(
            self.service.spreadsheets()
            .values()
            .append(
                spreadsheetId=self.spreadsheet_id,
                range=f"{self._quote_sheet_name(sheet_name)}!A1",
                valueInputOption="RAW",
                insertDataOption="INSERT_ROWS",
                body={"values": [list(values)]},
            )
        )

    def clear_range(self, range_name: str) -> dict[str, Any]:
        """Очищает значения в указанном диапазоне."""
        return self._execute(
            self.service.spreadsheets()
            .values()
            .clear(
                spreadsheetId=self.spreadsheet_id,
                range=range_name,
                body={},
            )
        )

    def add_sheet(self, title: str) -> int | None:
        """Добавляет лист и возвращает его внутренний ID."""
        response = self._execute(
            self.service.spreadsheets().batchUpdate(
                spreadsheetId=self.spreadsheet_id,
                body={"requests": [{"addSheet": {"properties": {"title": title}}}]},
            )
        )
        return (
            response.get("replies", [{}])[0]
            .get("addSheet", {})
            .get("properties", {})
            .get("sheetId")
        )

    def delete_sheet(self, title: str) -> None:
        """Удаляет лист с указанным названием."""
        sheet_id = next(
            (
                sheet["sheetId"]
                for sheet in self._get_sheet_properties()
                if sheet["title"] == title
            ),
            None,
        )
        if sheet_id is None:
            raise GoogleSheetsError("Лист с таким названием не найден.")

        self._execute(
            self.service.spreadsheets().batchUpdate(
                spreadsheetId=self.spreadsheet_id,
                body={"requests": [{"deleteSheet": {"sheetId": sheet_id}}]},
            )
        )


def main() -> None:
    """Читает первый лист по настройкам из файла .env."""
    from dotenv import load_dotenv
    import os

    load_dotenv()
    credentials_path = os.getenv("GOOGLE_CREDENTIALS_PATH")
    spreadsheet_id = os.getenv("SPREADSHEET_ID")
    if not credentials_path or not spreadsheet_id:
        print(
            "Ошибка: укажите GOOGLE_CREDENTIALS_PATH и SPREADSHEET_ID в файле .env."
        )
        return

    try:
        sheets = GoogleSheets(credentials_path, spreadsheet_id)
        sheet_names = sheets.get_sheet_names()
        print("Листы:", sheet_names)
        rows = sheets.read_all_cells()
        for number, row in enumerate(rows, start=1):
            print(f"Строка {number}: {row}")
        print(f"Всего строк: {len(rows)}")
    except GoogleSheetsError as error:
        print(f"Ошибка: {error}")


if __name__ == "__main__":
    main()